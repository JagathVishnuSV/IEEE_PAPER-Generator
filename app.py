from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
import base64
import io
import logging
import re

from utils.ieee_generator import generate_ieee_paper

app = FastAPI()

# Enable CORS for frontend integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Set specific origins in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

LATEX_PATTERN = r"\\[a-zA-Z]+(\{[^}]*\})*"

# ----------- Data Models -----------

class ImageData(BaseModel):
    caption: str
    data: str

class Subsection(BaseModel):
    heading: str
    content: str
    images: Optional[List[ImageData]] = []
    formulas: Optional[List[str]] = []
    tables: Optional[List[List[List[str]]]] = []

class Section(BaseModel):
    heading: str
    content: Optional[str] = ""
    images: Optional[List[ImageData]] = []
    formulas: Optional[List[str]] = []
    tables: Optional[List[List[List[str]]]] = []
    subsections: Optional[List[Subsection]] = []

class PaperData(BaseModel):
    title: str
    authors: List[str]
    affiliations: List[str]
    emails: List[str]
    abstract: str
    keywords: List[str]
    sections: List[Section]
    references: List[str]
    appendix: Optional[List[str]] = []

# ----------- Error Handler -----------

@app.exception_handler(Exception)
async def general_exception_handler(request: Request, exc: Exception):
    logger.error(f"Unhandled error: {exc}")
    return JSONResponse(status_code=500, content={"success": False, "error": str(exc)})

# ----------- Main Endpoint -----------

@app.post("/generate")
async def generate_paper(data: PaperData):
    try:
        validate_data(data)
        word_bytes = generate_ieee_paper(data.dict())
        return StreamingResponse(
            io.BytesIO(word_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=ieee_paper.docx"}
        )
    except Exception as e:
        logger.exception("Error generating document")
        raise HTTPException(status_code=400, detail=str(e))

# ----------- Validation -----------

def validate_data(data: PaperData):
    if not data.title.strip():
        raise ValueError("Title is required")
    if not all(author.strip() for author in data.authors):
        raise ValueError("All authors must be non-empty")
    if not all(aff.strip() for aff in data.affiliations):
        raise ValueError("All affiliations must be non-empty")
    if not all(email.strip() for email in data.emails):
        raise ValueError("All emails must be non-empty")
    if not data.abstract.strip():
        raise ValueError("Abstract is required")
    if not data.keywords:
        raise ValueError("At least one keyword is required")
    if not data.sections:
        raise ValueError("At least one section is required")

    for idx, section in enumerate(data.sections):
        if not section.heading.strip():
            raise ValueError(f"Section {idx + 1} is missing heading")
        has_content = bool(section.content and section.content.strip())
        has_subsections = bool(section.subsections)

        if not has_content and not has_subsections:
            raise ValueError(f"Section {idx + 1} must have content or subsections")

        for img in section.images or []:
            try:
                base64.b64decode(img.data)
            except Exception:
                raise ValueError(f"Invalid base64 image in section '{section.heading}'")

        if section.formulas:
            section.formulas = [f for f in section.formulas if re.match(LATEX_PATTERN, f.strip())]

        for sub_idx, sub in enumerate(section.subsections or []):
            if not sub.heading.strip():
                raise ValueError(f"Subsection {idx + 1}.{sub_idx + 1} is missing heading")
            if not sub.content.strip():
                raise ValueError(f"Subsection {idx + 1}.{sub_idx + 1} is missing content")
            for img in sub.images or []:
                try:
                    base64.b64decode(img.data)
                except Exception:
                    raise ValueError(f"Invalid base64 image in subsection '{sub.heading}'")
            if sub.formulas:
                sub.formulas = [f for f in sub.formulas if re.match(LATEX_PATTERN, f.strip())]
