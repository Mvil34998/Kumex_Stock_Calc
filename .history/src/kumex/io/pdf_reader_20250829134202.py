"""
Простое чтение текста из PDF (все страницы склеены).
"""
from pathlib import Path
import pdfplumber

def read_pdf_text(file_path: str) -> str:
    p = Path(file_path)
    if not p.exists():
        return ""
    chunks = []
    with pdfplumber.open(p) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")
    return "\n".join(chunks)
