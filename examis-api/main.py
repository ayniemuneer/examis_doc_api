from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, HttpUrl
from typing import List, Optional
import requests
import io
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from PIL import Image, ImageOps

app = FastAPI(title="Examis AI Document Generator")

# --- 1. PYDANTIC MODELS (Payload Validation) ---
class MarksData(BaseModel):
    mcq_points: float
    short_points: float
    long_points: float
    fib_points: Optional[float] = 1.0

class SubPart(BaseModel):
    question: str
    marks: float

class CustomScenarioItem(BaseModel):
    type: str
    text: str
    marks: Optional[float] = 0.0
    sub_parts: List[SubPart] = []

class MCQItem(BaseModel):
    question: str
    options: List[str]
    target_clo: Optional[str] = None
    image_url: Optional[HttpUrl] = None
    marks: float = 0.0

class ShortQuestionItem(BaseModel):
    question: str
    target_clo: Optional[str] = None
    image_url: Optional[HttpUrl] = None
    sub_parts: List[SubPart] = []
    marks: float = 0.0

class LongQuestionItem(BaseModel):
    question: str
    target_clo: Optional[str] = None
    image_url: Optional[HttpUrl] = None
    sub_parts: List[SubPart] = []
    marks: float = 0.0

class FillInTheBlankItem(BaseModel):
    question: str
    answer: str 
    target_clo: Optional[str] = None
    image_url: Optional[HttpUrl] = None
    marks: float = 0.0

class DiagramQuestionItem(BaseModel):
    question: str
    image_url: HttpUrl
    target_clo: Optional[str] = None
    marks: Optional[float] = 0.0

class ExamData(BaseModel):
    title: str
    department: str
    exam_type: str 
    total_marks: float 
    course_title: str
    credit_hours: str
    paper_type: str
    marks: MarksData
    custom_scenarios: List[CustomScenarioItem] = []
    mcqs: List[MCQItem] = []
    fillInTheBlanks: List[FillInTheBlankItem] = []
    shortQuestions: List[ShortQuestionItem] = []
    longQuestions: List[LongQuestionItem] = []
    diagram_questions: List[DiagramQuestionItem] = [] 


class DocumentRequest(BaseModel):
    template_url: HttpUrl
    show_clo_tags: bool = False
    exam_data: ExamData


# --- 2. CORE LOGIC ---
def process_exam(payload: DocumentRequest) -> io.BytesIO:
    response = requests.get(str(payload.template_url))
    if response.status_code != 200:
        raise HTTPException(status_code=400, detail="Failed to download template")
    
    doc_stream = io.BytesIO(response.content)
    doc = Document(doc_stream)
    exam = payload.exam_data
    show_clo = payload.show_clo_tags

    formatted_total_marks = f"{exam.total_marks:g}" if isinstance(exam.total_marks, float) else str(exam.total_marks)

    replacements = {
        "{{ exam_data.exam_type }}": str(exam.exam_type),
        "{{ exam_data.total_marks }}": formatted_total_marks,
        "{{ exam_data.course_title }}": str(exam.course_title),
        "{{ exam_data.credit_hours }}": str(exam.credit_hours),
        "{{ exam_data.paper_type }}": str(exam.paper_type),
        "{{ department }}": str(exam.department),
        "{{ exam_data.department }}": str(exam.department)
    }

    # --- TEMPLATE VALIDATION CHECK ---
    all_text = " ".join(
        [p.text for section in doc.sections for p in section.header.paragraphs] +
        [p.text for p in doc.paragraphs] +
        [p.text for table in doc.tables for row in table.rows for cell in row.cells for p in cell.paragraphs]
    )
    has_anchor = "{{START_EXAM_HERE}}" in all_text
    has_tags = any(tag in all_text for tag in replacements.keys())

    if not has_anchor and not has_tags:
        raise HTTPException(status_code=400, detail="The template does not have required tags")

    def replace_tags(p):
        for tag, value in replacements.items():
            if tag in p.text:
                p.text = p.text.replace(tag, value)

    for section in doc.sections:
        for header_p in section.header.paragraphs:
            replace_tags(header_p)

    for p in doc.paragraphs:
        replace_tags(p)
        if "{{START_EXAM_HERE}}" in p.text:
            p.text = p.text.replace("{{START_EXAM_HERE}}", "")
            instructions = "[Encircle the correct options. Overwriting will not be entertained. Multiple answers in fill in the blanks will be considered void.]"
            inst_run = p.add_run(instructions)
            inst_run.italic = True

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_tags(p)

    # --- HELPER FORMATTING FUNCTIONS ---
    def add_section_header(title: str):
        p = doc.add_paragraph()
        p.add_run(title).bold = True

    def write_sub_parts(sub_parts: List[SubPart]):
        labels = ['(a)', '(b)', '(c)', '(d)', '(e)', '(f)', '(g)', '(h)', '(i)', '(j)']
        for idx, sp in enumerate(sub_parts):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.5)
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)
            
            label = labels[idx] if idx < len(labels) else '(*)'
            p.add_run(f"{label} {sp.question}")
            p.add_run(f"\t[{sp.marks:g} Marks]").bold = True

    def insert_image_if_exists(img_url):
        if img_url:
            try:
                img_response = requests.get(str(img_url))
                if img_response.status_code == 200:
                    img_stream = io.BytesIO(img_response.content)
                    img = Image.open(img_stream)
                    img = ImageOps.exif_transpose(img)
                    
                    if img.mode in ("RGBA", "P"):
                        img = img.convert("RGB")
                        
                    fixed_stream = io.BytesIO()
                    img.save(fixed_stream, format='PNG')
                    fixed_stream.seek(0)

                    img_paragraph = doc.add_paragraph()
                    img_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    img_run = img_paragraph.add_run()
                    img_run.add_picture(fixed_stream, width=Inches(4.5))
            except Exception as e:
                print(f"Warning: Failed to load image - {e}")

    # ==========================================
    # SECTION LOGIC 
    # ==========================================

    # 1. Write MCQs 
    if exam.mcqs:
        add_section_header("Multiple Choice Questions")
        for i, mcq in enumerate(exam.mcqs, 1):
            p = doc.add_paragraph()
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)

            q_text = f"{i}. {mcq.question}"
            if show_clo and mcq.target_clo:
                q_text += f" [{mcq.target_clo}]"
            p.add_run(q_text).bold = True
            
            if mcq.marks > 0:
                p.add_run(f"\t[{mcq.marks:g} Marks]").bold = True
                
            insert_image_if_exists(mcq.image_url)
            
            labels = ['a)', 'b)', 'c)', 'd)']
            formatted_options = "    ".join([f"{labels[idx]} {opt}" for idx, opt in enumerate(mcq.options) if idx < len(labels)])
            opt_p = doc.add_paragraph(formatted_options)
            opt_p.paragraph_format.left_indent = Inches(0.5)
        doc.add_paragraph() 

    # 2. Write Fill in the Blanks 
    if exam.fillInTheBlanks:
        add_section_header("Fill in the Blanks")
        for i, fib in enumerate(exam.fillInTheBlanks, 1):
            p = doc.add_paragraph()
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)

            q_text = f"{i}. {fib.question}"
            if show_clo and fib.target_clo:
                q_text += f" [{fib.target_clo}]"
            p.add_run(q_text).bold = True

            if fib.marks > 0:
                p.add_run(f"\t[{fib.marks:g} Marks]").bold = True

            insert_image_if_exists(fib.image_url)
        doc.add_paragraph()

    # 3. Write Short Questions 
    if exam.shortQuestions:
        add_section_header("Short Answer Questions")
        for i, sq in enumerate(exam.shortQuestions, 1):
            p = doc.add_paragraph()
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)

            q_text = f"{i}. {sq.question}"
            if show_clo and sq.target_clo:
                q_text += f" [{sq.target_clo}]"
            p.add_run(q_text).bold = True

            if sq.marks > 0:
                p.add_run(f"\t[{sq.marks:g} Marks]").bold = True

            insert_image_if_exists(sq.image_url)
            
            if sq.sub_parts:
                write_sub_parts(sq.sub_parts)
            doc.add_paragraph()  

    # 4. Write Long Questions 
    if exam.longQuestions:
        add_section_header("Long Answer Questions")
        for i, lq in enumerate(exam.longQuestions, 1):
            p = doc.add_paragraph()
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)

            q_text = f"{i}. {lq.question}"
            if show_clo and lq.target_clo:
                q_text += f" [{lq.target_clo}]"
            p.add_run(q_text).bold = True

            if lq.marks > 0:
                p.add_run(f"\t[{lq.marks:g} Marks]").bold = True

            insert_image_if_exists(lq.image_url)
            
            if lq.sub_parts:
                write_sub_parts(lq.sub_parts)
            doc.add_paragraph()  

    # 5. Write Scenarios / Custom Types
    if exam.custom_scenarios:
        grouped_items = {}
        for item in exam.custom_scenarios:
            if item.type not in grouped_items:
                grouped_items[item.type] = []
            grouped_items[item.type].append(item)
            
        for item_type, items in grouped_items.items():
            doc.add_paragraph()
            type_header = doc.add_paragraph()
            type_header.add_run(f"{item_type}").bold = True
            
            for i, scenario in enumerate(items, 1):
                p = doc.add_paragraph()
                tab_stops = p.paragraph_format.tab_stops
                tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)

                p.add_run(f"Question {i}").bold = True
                
                if scenario.marks > 0:
                    p.add_run(f"\t[{scenario.marks:g} Marks]").bold = True
                    
                doc.add_paragraph(scenario.text)
                
                if scenario.sub_parts:
                    write_sub_parts(scenario.sub_parts)
            doc.add_paragraph()

    # 6. Write Diagram Questions 
    if exam.diagram_questions:
        doc.add_paragraph() 
        diag_header = doc.add_paragraph()
        diag_header.add_run("Diagrams & Visuals").bold = True
        
        for i, dq in enumerate(exam.diagram_questions, 1):
            p = doc.add_paragraph()
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)

            q_text = f"{i}. {dq.question}"
            if show_clo and dq.target_clo:
                q_text += f" [{dq.target_clo}]"
            p.add_run(q_text).bold = True
            
            if dq.marks > 0:
                p.add_run(f"\t[{dq.marks:g} Marks]").bold = True

            insert_image_if_exists(dq.image_url)
            doc.add_paragraph()

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0) 
    return output_stream


# --- 3. FASTAPI ENDPOINT ---
@app.post("/api/v1/generate-document")
async def generate_document(request_data: DocumentRequest):
    try:
        final_doc_stream = process_exam(request_data)
        headers = {
            "Content-Disposition": "attachment; filename=generated_exam.docx"
        }
        return StreamingResponse(
            final_doc_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))