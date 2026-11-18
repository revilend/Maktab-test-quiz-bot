"""AI quiz generator from PDF, DOCX, TXT, or topic text using Gemini API."""
from __future__ import annotations

import logging
from pathlib import Path

import aiohttp
from docx import Document
from pypdf import PdfReader

from config import GEMINI_API_KEYS

log = logging.getLogger(__name__)

GEMINI_MODEL = "gemini-2.5-flash"
GEMINI_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)


def extract_text_from_file(path: Path) -> str:
    """Extract plain text from .pdf, .docx, or .txt/.md files."""
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        reader = PdfReader(str(path))
        parts = [page.extract_text() or "" for page in reader.pages]
        return "\n".join(parts).strip()
    if suffix == ".docx":
        doc = Document(str(path))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
                if row_text:
                    parts.append(row_text)
        return "\n".join(parts).strip()
    return path.read_text(encoding="utf-8", errors="ignore").strip()


def _clean_generated_text(raw: str) -> str:
    lines = raw.strip().splitlines()
    if lines and lines[0].strip().startswith("```"):
        lines = lines[1:]
    if lines and lines[-1].strip().startswith("```"):
        lines = lines[:-1]
    return "\n".join(lines).strip()


async def generate_quiz_txt(
    source_text: str,
    num_questions: int = 10,
    difficulty: str = "o'rta",
    language: str = "o'zbek",
) -> str:
    """Call Gemini API (with key rotation) to generate questions in parser.py format."""
    if not GEMINI_API_KEYS:
        raise RuntimeError(
            "GEMINI_API_KEY yoki GEMINI_API_KEYS topilmadi. .env yoki Render sozlamalariga qo'shing."
        )

    trimmed = source_text.strip()[:28000]
    if not trimmed:
        raise ValueError("Manba matni bo'sh.")

    prompt = (
        f"Sen tajribali maktab o'qituvchisi va testologsan.\n"
        f"Quyidagi mavzu/material asosida {num_questions} ta test savoli tuz.\n"
        f"Til: {language}. Qiyinlik darajasi: {difficulty}.\n\n"
        f"QAT'IY FORMAT TALABLARI (boshqa hech qanday kirish yoki izoh yozma!):\n"
        f"1. Savol matni?\n"
        f"A) Birinchi variant\n"
        f"+B) To'g'ri javob varianti\n"
        f"C) Uchinchi variant\n"
        f"D) To'rtinchi variant\n\n"
        f"Qoidalar:\n"
        f"- Har bir savol raqam va nuqta bilan boshlansin (1., 2., ...).\n"
        f"- Har bir savolda aynan 4 ta variant bo'lsin: A), B), C), D).\n"
        f"- Faqat BITTA to'g'ri javob oldidan '+' belgisi qo'yilsin (masalan, +A) yoki +C)).\n"
        f"- Savollar qisqa, aniq va maktab dasturiga mos bo'lsin (Telegram poll limiti: savol <= 250 belgi, variant <= 95 belgi).\n\n"
        f"MATERIAL:\n{trimmed}"
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.4},
    }

    last_err = "Noma'lum xato"
    async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=60)) as session:
        for key in GEMINI_API_KEYS:
            try:
                async with session.post(
                    GEMINI_URL,
                    params={"key": key},
                    json=payload,
                ) as resp:
                    if resp.status == 429:
                        last_err = "API limit (429), keyingi kalitga o'tilmoqda"
                        continue
                    data = await resp.json()
                    if resp.status != 200:
                        last_err = data.get("error", {}).get("message", f"HTTP {resp.status}")
                        continue
                    candidates = data.get("candidates") or []
                    if not candidates:
                        last_err = "AI javob qaytarmadi"
                        continue
                    text = candidates[0]["content"]["parts"][0]["text"]
                    return _clean_generated_text(text)
            except Exception as e:
                last_err = str(e)
                continue

    raise RuntimeError(f"Gemini API xatosi: {last_err}")
