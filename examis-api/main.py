from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, HttpUrl
from typing import List, Optional
import requests
import io
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_TAB_ALIGNMENT

app = FastAPI(title="Examis AI Document Generator")

# --- 1. PYDANTIC MODELS (Payload Validation) ---
class MarksData(BaseModel):
    mcq_points: int
    short_points: int
    long_points: int

class MCQItem(BaseModel):
    question: str
    options: List[str]
    target_clo: Optional[str] = None

class ShortQuestionItem(BaseModel):
    question: str
    target_clo: Optional[str] = None

class LongQuestionItem(BaseModel):
    question: str
    target_clo: Optional[str] = None

class ExamData(BaseModel):
    title: str
    marks: MarksData
    mcqs: List[MCQItem] = []
    shortQuestions: List[ShortQuestionItem] = []
    longQuestions: List[LongQuestionItem] = []

class DocumentRequest(BaseModel):
    template_url: HttpUrl
    show_clo_tags: bool = False
    exam_data: ExamData


# --- 2. CORE LOGIC ---
def process_exam(payload: DocumentRequest) -> io.BytesIO:
    # 1. Download the template into RAM
    response = requests.get(str(payload.template_url))
    if response.status_code != 200:
        raise HTTPException(status_code=400, detail="Failed to download template")
    
    doc_stream = io.BytesIO(response.content)
    doc = Document(doc_stream)
    exam = payload.exam_data
    show_clo = payload.show_clo_tags

    # 2. Locate and clear the anchor point
    for p in doc.paragraphs:
        if "{{START_EXAM_HERE}}" in p.text:
            p.text = p.text.replace("{{START_EXAM_HERE}}", "")

    # Helper function to add section headers
    def add_section_header(title: str, points: int, count: int):
        p = doc.add_paragraph()
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)
        marks_str = f"[{points} x {count}]"
        p.add_run(title).bold = True
        p.add_run(f"\t{marks_str}").bold = True

    # 3. Write MCQs
    if exam.mcqs:
        add_section_header("Section A: Multiple Choice Questions", exam.marks.mcq_points, len(exam.mcqs))
        for i, mcq in enumerate(exam.mcqs, 1):
            p = doc.add_paragraph()
            
            # Format question text with optional CLO tag
            q_text = f"{i}. {mcq.question}"
            if show_clo and mcq.target_clo:
                q_text += f" [{mcq.target_clo}]"
                
            p.add_run(q_text).bold = True
            
            labels = ['a)', 'b)', 'c)', 'd)']
            for idx, option in enumerate(mcq.options):
                if idx < len(labels):
                    opt_p = doc.add_paragraph(f"{labels[idx]} {option}")
                    opt_p.paragraph_format.left_indent = Inches(0.5)
        doc.add_paragraph() 

    # 4. Write Short Questions
    if exam.shortQuestions:
        add_section_header("Section B: Short Answer Questions", exam.marks.short_points, len(exam.shortQuestions))
        for i, sq in enumerate(exam.shortQuestions, 1):
            p = doc.add_paragraph()
            
            q_text = f"{i}. {sq.question}"
            if show_clo and sq.target_clo:
                q_text += f" [{sq.target_clo}]"
                
            p.add_run(q_text).bold = True
            for _ in range(3):
                doc.add_paragraph()

    # 5. Write Long Questions
    if exam.longQuestions:
        add_section_header("Section C: Long Answer Questions", exam.marks.long_points, len(exam.longQuestions))
        for i, lq in enumerate(exam.longQuestions, 1):
            p = doc.add_paragraph()
            
            q_text = f"{i}. {lq.question}"
            if show_clo and lq.target_clo:
                q_text += f" [{lq.target_clo}]"
                
            p.add_run(q_text).bold = True
            if i < len(exam.longQuestions):
                doc.add_page_break() 

    # 6. Save modified document to a new memory stream
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
        
        # Switched to StreamingResponse for optimal Serverless memory management
        return StreamingResponse(
            final_doc_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))