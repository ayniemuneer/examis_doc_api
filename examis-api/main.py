from fastapi import FastAPI, HTTPException, Response
from pydantic import BaseModel, HttpUrl
from typing import List
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

class ShortQuestionItem(BaseModel):
    question: str

class LongQuestionItem(BaseModel):
    question: str

class ExamData(BaseModel):
    title: str
    marks: MarksData
    mcqs: List[MCQItem] = []
    shortQuestions: List[ShortQuestionItem] = []
    longQuestions: List[LongQuestionItem] = []

class DocumentRequest(BaseModel):
    template_url: HttpUrl
    exam_data: ExamData


# --- 2. CORE LOGIC ---
def process_exam(payload: DocumentRequest) -> io.BytesIO:
    # 1. Download the template into memory
    response = requests.get(str(payload.template_url))
    if response.status_code != 200:
        raise HTTPException(status_code=400, detail="Failed to download template")
    
    doc_stream = io.BytesIO(response.content)
    doc = Document(doc_stream)
    exam = payload.exam_data

    # 2. Locate and clear the anchor point
    for p in doc.paragraphs:
        if "{{START_EXAM_HERE}}" in p.text:
            p.text = p.text.replace("{{START_EXAM_HERE}}", "")

    # Helper function to add section headers with right-aligned marks
    def add_section_header(title: str, points: int, count: int):
        p = doc.add_paragraph()
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)
        
        marks_str = f"[{points} x {count}]"
        
        run_left = p.add_run(title)
        run_left.bold = True
        
        run_right = p.add_run(f"\t{marks_str}")
        run_right.bold = True

    # 3. Write MCQs
    if exam.mcqs:
        add_section_header("Section A: Multiple Choice Questions", exam.marks.mcq_points, len(exam.mcqs))
        for i, mcq in enumerate(exam.mcqs, 1):
            p = doc.add_paragraph()
            p.add_run(f"{i}. {mcq.question}").bold = True
            
            labels = ['a)', 'b)', 'c)', 'd)']
            for idx, option in enumerate(mcq.options):
                opt_p = doc.add_paragraph(f"{labels[idx]} {option}")
                opt_p.paragraph_format.left_indent = Inches(0.5)
        doc.add_paragraph() 

    # 4. Write Short Questions
    if exam.shortQuestions:
        add_section_header("Section B: Short Answer Questions", exam.marks.short_points, len(exam.shortQuestions))
        for i, sq in enumerate(exam.shortQuestions, 1):
            p = doc.add_paragraph()
            p.add_run(f"{i}. {sq.question}").bold = True
            for _ in range(3):
                doc.add_paragraph()

    # 5. Write Long Questions
    if exam.longQuestions:
        add_section_header("Section C: Long Answer Questions", exam.marks.long_points, len(exam.longQuestions))
        for i, lq in enumerate(exam.longQuestions, 1):
            p = doc.add_paragraph()
            p.add_run(f"{i}. {lq.question}").bold = True
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
        
        return Response(
            content=final_doc_stream.read(),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": "attachment; filename=generated_exam.docx"
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))