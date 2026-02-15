import base64
import os
import re
import time
import shutil
import threading
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from bisect import bisect_right
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from typing import Any, Dict, List, Tuple, Callable, Optional
from datetime import datetime
import tempfile
import queue
import logging
import sys
import asyncio
import random
import unicodedata
from urllib.parse import urlparse, urlunparse, quote, parse_qs

import pandas as pd
import pdfplumber
import streamlit as st
import streamlit.components.v1 as components  # 👈 NUEVO
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from dataclasses import dataclass

# =========================
# Dependencias PDF / Word / PPT para Broken Link Checker
# =========================
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None

try:
    from pptx import Presentation as PptxPresentation
except ImportError:
    PptxPresentation = None

try:
    import requests
except ImportError:
    requests = None

try:
    import httpx
except ImportError:
    httpx = None

# ======================================================
# LOGGING
# ======================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

# ======================================================
# CONFIG / CONSTANTES (Broken Link Checker)
# ======================================================
APP_TITLE = "UTP - GrammarScan"
APP_ICON = "📚"

MODULES = [
    "Home",
    "Report GrammarScan",
]

DEFAULT_TIMEOUT_S = 15.0
DEFAULT_CONCURRENCY_GLOBAL = 30
DEFAULT_CONCURRENCY_PER_HOST = 6
DEFAULT_RETRIES = 3
DEFAULT_MAX_BYTES = 200_000
DEFAULT_RANGE_BYTES = 131_072

STATUS_BLOCK_SIZE = 200  # número máximo de links por bloque al validar
MAX_ZIP_BLOCK_MB = 200   # tamaño máximo aproximado por bloque de ZIP (MB)

# --------- Límite de seguridad Streamlit Cloud (solo afecta a la descarga masiva desde Excel) ----------
MAX_BULK_URLS_CLOUD = 700


def is_streamlit_cloud() -> bool:
    """Detecta si la app está ejecutándose en Streamlit Cloud.

    Se puede forzar de dos maneras:
    - Configurando st.secrets["is_streamlit_cloud"] = true/false.
    - Definiendo la variable de entorno IS_STREAMLIT_CLOUD=true/false.
    En local, por defecto devolverá False.
    """
    # 1) Intentar leer desde st.secrets (si existe)
    try:
        if "is_streamlit_cloud" in st.secrets:
            val = st.secrets["is_streamlit_cloud"]
        else:
            val = None
    except Exception:
        val = None

    # 2) Si no está en secrets, mirar variable de entorno
    if val is None:
        val = os.environ.get("IS_STREAMLIT_CLOUD", "")

    if isinstance(val, bool):
        return val

    return str(val).strip().lower() in ("1", "true", "yes", "y", "on")


IS_STREAMLIT_CLOUD = is_streamlit_cloud()


# Descarga Masiva
MAX_INTENTOS_DESCARGA = 7
CHUNK_SIZE = 1024 * 256
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; UTP-FileDownloader/1.0)",
    "Accept": "*/*",
    "Accept-Encoding": "identity",
}
DESC_EXT_PERMITIDAS = (".ppt", ".pptx", ".pdf", ".doc", ".docx")

# Patrones de soft-404 (simplificados)
SOFT_404_PATTERNS = [
    r"\b404\b",
    r"page\s+not\s+found",
    r"file\s+not\s+found",
    r"document\s+not\s+found",
    r"p[aá]gina\s+no\s+encontrada",
    r"la\s+p[aá]gina\s+no\s+existe",
    r"no\s+se\s+encontr[oó]",
]
SOFT_404_RE = re.compile("|".join(SOFT_404_PATTERNS), re.IGNORECASE)

# Extensiones binarias
BINARY_EXTS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rar", ".7z",
    ".ppt", ".pptx", ".mp4", ".mp3", ".avi", ".mov", ".wmv", ".mkv",
    ".png", ".jpg", ".jpeg", ".gif", ".tif", ".tiff", ".bmp", ".svg",
}

# Content-Type esperados
EXPECTED_CONTENT_TYPES = {
    ".pdf": ["application/pdf"],
    ".doc": ["application/msword"],
    ".docx": ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
    ".xls": ["application/vnd.ms-excel"],
    ".xlsx": ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
    ".zip": ["application/zip", "application/x-zip-compressed"],
    ".rar": ["application/x-rar-compressed"],
    ".ppt": ["application/vnd.ms-powerpoint"],
    ".pptx": ["application/vnd.openxmlformats-officedocument.presentationml.presentation"],
}

# User-Agents realistas
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_3) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15",
]

# ======================================================
# LÓGICA GRAMMARSCAN (original)
# ======================================================
LINES_PER_TXT_PAGE = 50
PAGE_SEP = "\n\f\n"
ALLOWED_DOC_EXTS = {".pdf", ".docx", ".pptx", ".txt"}

WS_MULTI_RE = re.compile(r"[ \t]+")
NL_3PLUS_RE = re.compile(r"\n{3,}")

def normalize_ws(text: str) -> str:
    text = text.replace("\r", "\n")
    text = WS_MULTI_RE.sub(" ", text)
    text = NL_3PLUS_RE.sub("\n\n", text)
    return text.strip()

def find_java() -> bool:
    return shutil.which("java") is not None

def safe_str(x) -> str:
    try:
        return str(x)
    except Exception:
        return ""

ACCENTED_VOWELS = "áéíóúÁÉÍÓÚ"

def _has_accented_vowel(s: str) -> bool:
    return any(ch in ACCENTED_VOWELS for ch in s or "")

@dataclass
class ModismoPattern:
    modismo: str
    tipo: str
    patron: str
    sugerencia: str
    comentario: str
    regex: re.Pattern

@dataclass
class LogicalFileSource:
    display_name: str
    ext: str
    read_bytes: Callable[[], bytes]

def _normalize_regex_pattern(patron: str) -> str:
    if not isinstance(patron, str):
        patron = str(patron or "")
    patron = patron.replace("\\\\", "\\")
    return patron.strip()

def load_modismos_from_excel(path: str) -> List[ModismoPattern]:
    if not os.path.isfile(path):
        raise FileNotFoundError(f"No se encontró el archivo de modismos: {path}")

    df = pd.read_excel(path)
    required_cols = {"modismo", "tipo", "sugerencia"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Faltan columnas obligatorias en 'modismos_ar.xlsx': {missing}")

    patterns: List[ModismoPattern] = []
    for _, row in df.iterrows():
        modismo = safe_str(row.get("modismo", "")).strip()
        tipo = safe_str(row.get("tipo", "")).strip().lower()
        sugerencia = safe_str(row.get("sugerencia", "")).strip()
        comentario = safe_str(row.get("comentario", "")).strip()
        patron_cfg = safe_str(row.get("patron", "")).strip()

        if not modismo or not sugerencia:
            continue

        if tipo not in ("literal", "regex"):
            tipo = "literal"

        if tipo == "literal":
            patron = r"(?<!\w)" + re.escape(modismo) + r"(?!\w)"
        else:
            base = patron_cfg if patron_cfg else modismo
            patron = _normalize_regex_pattern(base)

        try:
            rx = re.compile(patron, flags=re.IGNORECASE | re.UNICODE)
        except re.error:
            continue

        patterns.append(
            ModismoPattern(
                modismo=modismo,
                tipo=tipo,
                patron=patron,
                sugerencia=sugerencia,
                comentario=comentario,
                regex=rx,
            )
        )

    return patterns

@st.cache_resource(show_spinner=False)
def get_modismos_patterns(modismos_path: str) -> List[ModismoPattern]:
    return load_modismos_from_excel(modismos_path)

HEAD_RE = re.compile(
    r"^\s*(fuentes?\s+bibliogr[aá]ficas?|bibliograf[ií]a|bibliography|"
    r"referencias(\s+bibliogr[aá]ficas?)?|referencias\s+y\s+bibliograf[ií]a|"
    r"works\s+cited|obras\s+citadas|webgraf[ií]a|fuentes\s+de\s+consulta|"
    r"bibliograf[ií]a\s+consultada|citas\s+bibliogr[aá]ficas?)\s*:?\s*$",
    re.IGNORECASE
)

URL_RE = re.compile(r"(https?://\S+|www\.\S+)", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(19|20)\d{2}[a-z]?\b")
DOI_URL_HINT_RE = re.compile(r"(doi:\s*10\.\d{4,9}/\S+|urn:\S+|hdl:\S+|https?://\S+)", re.IGNORECASE)
JOURNAL_HINT_RE = re.compile(
    r"\b(vol\.?|no\.?|nº|n\.\s?o\.?|pp\.?|ed\.|edición|issn|isbn|issue|pages)\b",
    re.IGNORECASE
)
PUBLISHER_HINT_RE = re.compile(
    r"\b(pearson|mcgraw[- ]?hill|elsevier|springer|wiley|cengage|prentice\s*hall|sage|oxford|cambridge|harvard\s*press|"
    r"editorial(?:es)?|ediciones?|universidad(?:\s+de)?\s+\w+)\b",
    re.IGNORECASE
)
ETAL_RE = re.compile(r"\bet\s+al\.?\b", re.IGNORECASE)
BULLET_PREFIX_RE = re.compile(
    r"^\s*(?:[\u2022\u2023\u25E6\u2043\u2219\u25CF\u25AA\u25AB\u25A0\u25A1]|[-–—·•▪●◦‣])\s*"
)

def _strip_bullet(line: str) -> str:
    return BULLET_PREFIX_RE.sub("", line).strip()

BIB_RE_LIST = [
    re.compile(
        r"^[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+,\s(?:[A-Z]\.\s?){1,4}(?:,\s(?:[A-Z]\.\s?)){0,3}"
        r"(?:\s(?:&|y|and)\s[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+,\s(?:[A-Z]\.\s?){1,4})*"
        r"\s\(\d{4}[a-z]?\)\.\s",
        re.UNICODE
    ),
    re.compile(
        r"^[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+,\s(?:[A-Z]\.\s?){1,4}.*\b(19|20)\d{2}[a-z]?\.\s*$",
        re.UNICODE
    ),
]

def is_bibliography_heading(line: str) -> bool:
    return bool(HEAD_RE.match(line.strip()))

def is_reference_line(line: str) -> bool:
    s = _strip_bullet(line)
    if not s:
        return False
    for rx in BIB_RE_LIST:
        if rx.search(s):
            return True
    if YEAR_RE.search(s) and (
        DOI_URL_HINT_RE.search(s) or JOURNAL_HINT_RE.search(s) or PUBLISHER_HINT_RE.search(s)
    ):
        if s.count(",") >= 2 or ETAL_RE.search(s) or URL_RE.search(s):
            return True
    if URL_RE.search(s) and YEAR_RE.search(s):
        return True
    return False

def detect_bibliography_pages(pages: List[Tuple[int, str]]) -> set:
    skip = set()
    in_bib_section = False

    for num, txt in pages:
        lines_raw = [l for l in txt.splitlines() if l.strip()]
        lines = [_strip_bullet(l.strip()) for l in lines_raw if l.strip()]
        if not lines:
            continue

        head_zone = lines[:8]
        has_heading = any(is_bibliography_heading(l) for l in head_zone)

        ref_count = sum(1 for l in lines if is_reference_line(l))
        url_count = sum(1 for l in lines if URL_RE.search(l))
        year_tokens = len(YEAR_RE.findall(" ".join(lines)))
        bullet_lines = sum(1 for l in lines_raw if BULLET_PREFIX_RE.match(l))
        n = max(1, len(lines))

        ref_ratio = ref_count / n
        url_ratio = url_count / n

        if has_heading:
            skip.add(num)
            in_bib_section = True
            continue

        if (
            ref_count >= 2
            or (ref_ratio >= 0.25)
            or (url_count >= 3 and year_tokens >= 2)
            or (bullet_lines >= 3 and (ref_count + url_count) >= 2)
            or (url_ratio >= 0.35 and year_tokens >= 1)
        ):
            skip.add(num)
            in_bib_section = True
            continue

        if in_bib_section and (ref_count >= 1 and (url_count >= 1 or year_tokens >= 1)):
            skip.add(num)
        else:
            in_bib_section = False

    return skip

def is_reference_fragment(text: str) -> bool:
    if not text:
        return False

    for raw in text.splitlines():
        l = _strip_bullet(raw)
        if not l:
            continue
        if is_reference_line(l):
            return True
        if (URL_RE.search(l) and YEAR_RE.search(l)) or (ETAL_RE.search(l) and YEAR_RE.search(l)):
            return True
        if is_bibliography_heading(l):
            return True
        if PUBLISHER_HINT_RE.search(l) and YEAR_RE.search(l):
            return True

    flat = " ".join(_strip_bullet(s) for s in text.split())
    if YEAR_RE.search(flat) and (
        DOI_URL_HINT_RE.search(flat) or JOURNAL_HINT_RE.search(flat) or PUBLISHER_HINT_RE.search(flat)
    ):
        if flat.count(",") >= 2 or ETAL_RE.search(flat) or URL_RE.search(flat):
            return True

    return False

EN_COMMON_WORDS = {
    "the", "and", "or", "but", "if", "then", "else", "when", "where", "who", "what",
    "which", "while", "for", "from", "to", "in", "on", "at", "by", "of", "with",
}

EN_TOKEN_RE = re.compile(r"[a-zA-Z]+")

def is_english_fragment(text: str) -> bool:
    if not text:
        return False

    cleaned = URL_RE.sub(" ", text)
    lower = cleaned.lower()

    if any(ch in lower for ch in "áéíóúüñ"):
        return False

    tokens = EN_TOKEN_RE.findall(lower)
    if len(tokens) < 4:
        return False

    en_hits = sum(1 for t in tokens if t in EN_COMMON_WORDS)
    if en_hits == 0:
        return False

    en_ratio = en_hits / max(1, len(tokens))
    ascii_tokens = [t for t in tokens if re.fullmatch(r"[a-z]+", t)]
    ascii_ratio = len(ascii_tokens) / max(1, len(tokens))

    if ascii_ratio >= 0.8 and en_ratio >= 0.25:
        return True
    if not any(ch in lower for ch in "áéíóúüñ") and en_ratio >= 0.35:
        return True

    return False

LATIN_KEYWORDS = {
    "ius", "honorarium", "civile", "sabinum", "edictum", "aedilium", "curulium",
    "officio", "pronoconsulis", "corpus", "delicti", "mens", "rea",
}

def is_latin_fragment(text: str) -> bool:
    if not text:
        return False

    cleaned = URL_RE.sub(" ", text)
    lower = cleaned.lower()
    tokens = EN_TOKEN_RE.findall(lower)
    if not tokens:
        return False

    hits = sum(1 for t in tokens if t in LATIN_KEYWORDS)

    if hits >= 1 and len(tokens) <= 4:
        return True
    if hits >= 2:
        return True
    if re.search(r"\bius\b", lower) and len(tokens) <= 6:
        return True

    return False

CODE_SYMBOLS = set("{}[]();<>:=+-*/%&|^#!@$\"'\\")
CODE_KEYWORDS = {
    "public ", "private ", "protected ", "class ", "interface ", "implements ",
    "extends ", "static ", "void ", "int ", "double ", "float ", "string ",
}

def is_code_fragment(text: str) -> bool:
    if not text:
        return False

    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return False

    code_like_lines = 0
    total_nonempty = 0

    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        total_nonempty += 1

        lower = s.lower()

        if any(kw in lower for kw in CODE_KEYWORDS):
            code_like_lines += 1
            continue

        if s.endswith((";", "{", "}")):
            code_like_lines += 1
            continue

        non_space = sum(1 for c in s if not c.isspace())
        if non_space >= 6:
            sym_count = sum(1 for c in s if c in CODE_SYMBOLS)
            if sym_count / non_space > 0.3:
                code_like_lines += 1
                continue

    if total_nonempty == 0:
        return False

    if code_like_lines >= 2 and code_like_lines / total_nonempty >= 0.5:
        return True

    all_non_space = sum(1 for c in text if not c.isspace())
    if all_non_space >= 10:
        all_sym = sum(1 for c in text if c in CODE_SYMBOLS)
        if all_sym / all_non_space > 0.35:
            return True

    return False

def read_pdf_pages(bio: BytesIO) -> List[Tuple[int, str]]:
    pages: List[Tuple[int, str]] = []
    with pdfplumber.open(bio) as pdf:
        for i, p in enumerate(pdf.pages, start=1):
            t = p.extract_text() or ""
            t = normalize_ws(t)
            if t:
                pages.append((i, t))
    return pages

def _iter_block_items(doc: Document):
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def _paragraph_has_page_break(para: Paragraph) -> bool:
    try:
        if para.paragraph_format.page_break_before:
            return True
    except Exception:
        pass
    try:
        for run in para.runs:
            if run._element.xpath('.//w:br[@w:type="page"]'):
                return True
    except Exception:
        pass
    return False

def read_docx_pages(bio: BytesIO) -> List[Tuple[int, str]]:
    doc = Document(bio)
    pages: List[Tuple[int, str]] = []
    buff: List[str] = []
    page_no = 1

    def flush():
        nonlocal buff, page_no, pages
        txt = normalize_ws("\n".join(buff))
        if txt:
            pages.append((page_no, txt))
        buff = []

    for b in _iter_block_items(doc):
        if isinstance(b, Paragraph):
            t = b.text.strip()
            if t:
                buff.append(t)
            if _paragraph_has_page_break(b):
                flush()
                page_no += 1
        else:
            for row in b.rows:
                row_text = " | ".join(
                    (c.text or "").strip() for c in row.cells if (c.text or "").strip()
                )
                if row_text:
                    buff.append(row_text)

    flush()

    if not pages:
        flat = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        if flat.strip():
            pages = [(1, normalize_ws(flat))]

    return pages

def _clean_reference_lines_block(text: str) -> str:
    kept = []
    for ln in (text or "").splitlines():
        s = normalize_ws(ln)
        if not s:
            continue
        if (
            is_bibliography_heading(s)
            or is_reference_line(s)
            or (PUBLISHER_HINT_RE.search(s) and YEAR_RE.search(s))
            or is_code_fragment(s)
        ):
            continue
        kept.append(s)
    return "\n".join(kept).strip()

def _read_pptx_via_zip(bio: BytesIO) -> List[Tuple[int, str]]:
    slides: List[Tuple[int, str]] = []
    bio.seek(0)
    with zipfile.ZipFile(bio) as z:
        names = [
            n for n in z.namelist()
            if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        ]

        def slide_no(name: str) -> int:
            m = re.search(r"slide(\d+)\.xml$", name)
            return int(m.group(1)) if m else 0

        names.sort(key=slide_no)

        A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

        for idx, name in enumerate(names, start=1):
            try:
                xml_bytes = z.read(name)
                root = ET.fromstring(xml_bytes)
                texts = [t.text for t in root.iter(f"{A_NS}t") if t.text]
                raw = "\n".join(texts)
                cleaned = _clean_reference_lines_block(normalize_ws(raw))
                if cleaned:
                    slides.append((idx, cleaned))
            except Exception:
                continue
    return slides

def _iter_shape_texts(shape) -> List[str]:
    textos: List[str] = []

    try:
        if getattr(shape, "has_text_frame", False) and shape.text_frame:
            for para in shape.text_frame.paragraphs:
                frags = [run.text for run in para.runs if run.text]
                if frags:
                    textos.append("".join(frags))
    except Exception:
        pass

    try:
        if getattr(shape, "has_table", False) or getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.TABLE:
            tbl = shape.table
            for row in tbl.rows:
                celdas = []
                for cell in row.cells:
                    t = (cell.text or "").strip()
                    if t:
                        celdas.append(t)
                if celdas:
                    textos.append(" | ".join(celdas))
    except Exception:
        pass

    try:
        subshapes = getattr(shape, "shapes", None)
        if subshapes is not None:
            for sub in subshapes:
                textos.extend(_iter_shape_texts(sub))
    except Exception:
        pass

    return textos

def read_pptx_slides(bio: BytesIO) -> List[Tuple[int, str]]:
    try:
        prs = Presentation(bio)
        slides: List[Tuple[int, str]] = []
        slide_h = float(prs.slide_height) if hasattr(prs, "slide_height") else None

        for s_idx, slide in enumerate(prs.slides, start=1):
            chunk: List[str] = []

            for sh in slide.shapes:
                for raw in _iter_shape_texts(sh):
                    if not raw:
                        continue

                    raw_norm = normalize_ws(raw)

                    try:
                        if slide_h and float(getattr(sh, "top", 0)) >= 0.75 * slide_h:
                            if (
                                is_reference_fragment(raw_norm)
                                or URL_RE.search(raw_norm)
                                or (
                                    PUBLISHER_HINT_RE.search(raw_norm)
                                    and YEAR_RE.search(raw_norm)
                                )
                            ):
                                continue
                    except Exception:
                        pass

                    cleaned = _clean_reference_lines_block(raw_norm)
                    if cleaned:
                        chunk.append(cleaned)

            txt = normalize_ws("\n".join(chunk))
            if txt:
                slides.append((s_idx, txt))
        return slides

    except Exception:
        try:
            return _read_pptx_via_zip(bio)
        except Exception as e2:
            raise e2

def read_txt_pages(bio: BytesIO) -> List[Tuple[int, str]]:
    try:
        raw = bio.read()
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                s = raw.decode(enc)
                break
            except Exception:
                continue
        else:
            s = raw.decode("utf-8", errors="ignore")
    except Exception:
        return []

    s = s.replace("\r\n", "\n").replace("\r", "\n")
    lines = s.split("\n")

    pages: List[Tuple[int, str]] = []
    for i in range(0, len(lines), LINES_PER_TXT_PAGE):
        page_num = (i // LINES_PER_TXT_PAGE) + 1
        chunk = normalize_ws("\n".join(lines[i:i + LINES_PER_TXT_PAGE]))
        if chunk:
            pages.append((page_num, chunk))

    if not pages and s.strip():
        pages = [(1, normalize_ws(s))]

    return pages

def extract_pages(file_bytes: bytes, file_name: str) -> Tuple[List[Tuple[int, str]], str]:
    ext = os.path.splitext(file_name)[1].lower()
    bio = BytesIO(file_bytes)

    if ext == ".pdf":
        return read_pdf_pages(bio), "Página"
    if ext == ".docx":
        return read_docx_pages(bio), "Página"
    if ext == ".pptx":
        return read_pptx_slides(bio), "Diapositiva"
    if ext == ".txt":
        return read_txt_pages(bio), "Página"

    return [], "Página"

def build_global_text(
    pages: List[Tuple[int, str]]
) -> Tuple[str, List[int], List[Tuple[int, int, int]]]:
    parts: List[str] = []
    starts: List[int] = []
    bounds: List[Tuple[int, int, int]] = []
    cur = 0

    for idx, (num, txt) in enumerate(pages):
        starts.append(cur)
        start = cur

        parts.append(txt)
        cur += len(txt)

        end = cur
        bounds.append((start, end, num))

        if idx < len(pages) - 1:
            parts.append(PAGE_SEP)
            cur += len(PAGE_SEP)

    return "".join(parts), starts, bounds

def chunk_by_pages(pages: List[Tuple[int, str]], max_chars: int) -> List[Tuple[int, int]]:
    ranges: List[Tuple[int, int]] = []
    i = 0
    while i < len(pages):
        j = i
        total = 0
        while j < len(pages):
            candidate = len(pages[j][1]) + (0 if j == i else len(PAGE_SEP))
            if total + candidate > max_chars and j > i:
                break
            total += candidate
            j += 1
        ranges.append((i, j))
        i = j
    return ranges

def page_for_offset(bounds: List[Tuple[int, int, int]], offset: int) -> int:
    starts = [b[0] for b in bounds]
    idx = bisect_right(starts, offset) - 1
    if idx < 0:
        idx = 0
    return bounds[idx][2]

@st.cache_resource(show_spinner=False)
def get_language_tool(lang_code: str):
    if not find_java():
        raise RuntimeError("Java no detectado. Activa tu JRE/JDK para usar LanguageTool local.")
    import language_tool_python as lt
    return lt.LanguageTool(lang_code)

LT_LOCK = threading.Lock()

def analyze_text(tool, text: str, retries: int = 2, sleep: float = 0.8) -> list:
    if not text.strip():
        return []
    last_err = None
    for _ in range(retries + 1):
        try:
            with LT_LOCK:
                return tool.check(text)
        except Exception as e:
            last_err = e
            msg = str(e)
            if "stdin" in msg or "Connection refused" in msg or "WinError" in msg:
                try:
                    tool = get_language_tool(st.session_state.get("_lang_code", "es"))
                except Exception:
                    pass
            time.sleep(sleep)
    raise RuntimeError(f"Fallo LanguageTool local tras reintentos: {last_err}")

def detect_modismos_in_pages(
    file_name: str,
    pages: List[Tuple[int, str]],
    unit_label: str,
    patterns: List[ModismoPattern],
    skip_pages: set | None = None,
) -> pd.DataFrame:
    if not patterns:
        return pd.DataFrame([])

    skip_pages = skip_pages or set()
    rows: List[Dict[str, Any]] = []

    for page_no, text in pages:
        if page_no in skip_pages:
            continue
        if not text.strip():
            continue

        lower = text.lower()

        for pat in patterns:
            for m in pat.regex.finditer(lower):
                start, end = m.span()
                ctx_start = max(0, start - 60)
                ctx_end = min(len(text), end + 60)
                contexto = text[ctx_start:ctx_end]

                if is_reference_fragment(contexto):
                    continue

                match_text = text[start:end]

                if _has_accented_vowel(pat.modismo) and not _has_accented_vowel(match_text):
                    continue

                mensaje = (
                    f"Uso de modismo argentino «{match_text}». "
                    f"Sugerencia: «{pat.sugerencia}»."
                )

                rows.append({
                    "Archivo": file_name,
                    "Página/Diapositiva": page_no,
                    "BloqueTipo": unit_label,
                    "Mensaje": mensaje,
                    "Sugerencias": pat.sugerencia,
                    "Oración": contexto,
                    "Contexto": contexto,
                    "Regla": "MODISMO_AR",
                    "Categoría": f"UTP_CUSTOM: Modismos argentinos ({pat.modismo})",
                })

    if not rows:
        return pd.DataFrame([])

    df_mod = pd.DataFrame.from_records(rows).drop_duplicates()
    return df_mod

def analyze_file(
    file_name: str,
    file_bytes: bytes,
    lang_code: str,
    max_chars_call: int,
    workers: int,
    excluir_bibliografia: bool = True,
    modismos_patterns: List[ModismoPattern] | None = None,
    analizar_modismos: bool = False,
) -> pd.DataFrame:
    pages, unit_label = extract_pages(file_bytes, file_name)
    if not pages:
        return pd.DataFrame([])

    ext = os.path.splitext(file_name)[1].lower()

    skip_pages = detect_bibliography_pages(pages) if excluir_bibliografia else set()

    if excluir_bibliografia and ext in (".txt", ".docx", ".pptx"):
        skip_pages = set()

    st.session_state["_lang_code"] = lang_code
    _, starts, bounds = build_global_text(pages)
    ranges = chunk_by_pages(pages, max_chars_call)
    tool = get_language_tool(lang_code)

    rows: List[Dict[str, Any]] = []
    lock = threading.Lock()

    def worker(rng: Tuple[int, int]):
        i, j = rng
        group_start = starts[i]
        txt = PAGE_SEP.join([pages[k][1] for k in range(i, j)])
        matches = analyze_text(tool, txt)
        local_rows: List[Dict[str, Any]] = []
        for m in matches:
            try:
                off = int(getattr(m, "offset", -1) or -1)
            except Exception:
                off = -1

            global_off = group_start + (off if off >= 0 else 0)
            page_no = page_for_offset(bounds, global_off)

            sentence = safe_str(getattr(m, "sentence", ""))
            context = safe_str(getattr(m, "context", ""))

            if URL_RE.search(sentence) or URL_RE.search(context):
                continue

            if excluir_bibliografia:
                if page_no in skip_pages:
                    continue
                if is_reference_fragment(sentence) or is_reference_fragment(context):
                    continue

            if is_code_fragment(sentence) or is_code_fragment(context):
                continue

            if lang_code.startswith("es"):
                if is_english_fragment(sentence) or is_english_fragment(context):
                    continue
                if is_latin_fragment(sentence) or is_latin_fragment(context):
                    continue

            local_rows.append({
                "Archivo": file_name,
                "Página/Diapositiva": page_no,
                "BloqueTipo": unit_label,
                "Mensaje": safe_str(getattr(m, "message", "")),
                "Sugerencias": ", ".join(
                    getattr(m, "replacements", [])[:5]
                ) if isinstance(getattr(m, "replacements", []), list) else "",
                "Oración": sentence,
                "Contexto": context,
                "Regla": safe_str(getattr(m, "ruleId", "")),
                "Categoría": safe_str(getattr(m, "category", "")),
            })
        return local_rows

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = [ex.submit(worker, r) for r in ranges]
        for fut in as_completed(futures):
            part = fut.result()
            with lock:
                rows.extend(part)

    if not rows:
        df_lt = pd.DataFrame([])
    else:
        df_lt = pd.DataFrame.from_records(rows)

    if not df_lt.empty:
        if excluir_bibliografia:
            mask_ref = (
                df_lt["Oración"].fillna("").apply(is_reference_fragment)
                | df_lt["Contexto"].fillna("").apply(is_reference_fragment)
            )
            df_lt = df_lt.loc[~mask_ref].copy()

        mask_code = (
            df_lt["Oración"].fillna("").apply(is_code_fragment)
            | df_lt["Contexto"].fillna("").apply(is_code_fragment)
        )
        df_lt = df_lt.loc[~mask_code].copy()

        if lang_code.startswith("es"):
            mask_eng = (
                df_lt["Oración"].fillna("").apply(is_english_fragment)
                | df_lt["Contexto"].fillna("").apply(is_english_fragment)
            )
            df_lt = df_lt.loc[~mask_eng].copy()

            mask_lat = (
                df_lt["Oración"].fillna("").apply(is_latin_fragment)
                | df_lt["Contexto"].fillna("").apply(is_latin_fragment)
            )
            df_lt = df_lt.loc[~mask_lat].copy()

        mask_url = (
            df_lt["Oración"].fillna("").str.contains(URL_RE)
            | df_lt["Contexto"].fillna("").str.contains(URL_RE)
        )
        df_lt = df_lt.loc[~mask_url].copy()

    if analizar_modismos and lang_code.startswith("es") and modismos_patterns:
        if excluir_bibliografia and ext in (".txt", ".docx", ".pptx"):
            skip_for_modismos = None
        else:
            skip_for_modismos = skip_pages

        df_mod = detect_modismos_in_pages(
            file_name=file_name,
            pages=pages,
            unit_label=unit_label,
            patterns=modismos_patterns,
            skip_pages=skip_for_modismos,
        )
    else:
        df_mod = pd.DataFrame([])

    if df_lt is None or df_lt.empty:
        final_df = df_mod
    elif df_mod is None or df_mod.empty:
        final_df = df_lt
    else:
        final_df = pd.concat([df_lt, df_mod], ignore_index=True)

    if final_df is None or final_df.empty:
        return pd.DataFrame([])

    final_df.sort_values(["Archivo", "Página/Diapositiva"], inplace=True)
    final_df.drop_duplicates(
        subset=["Archivo", "Página/Diapositiva", "Mensaje", "Oración", "Contexto", "Regla"],
        keep="first",
        inplace=True,
    )
    final_df.reset_index(drop=True, inplace=True)
    return final_df

# Patrones para limpiar caracteres no válidos en XML/Excel
INVALID_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

def _sanitize_excel_df(df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    """
    Elimina de un DataFrame los caracteres de control no permitidos por
    el formato XLSX (XML 1.0) para evitar que Excel muestre errores
    de reparación al abrir el archivo.
    """
    if df is None or df.empty:
        return df

    df = df.copy()

    # Limpiar nombres de columnas
    df.columns = [INVALID_XML_RE.sub("", str(c)) for c in df.columns]

    # Limpiar el contenido de las celdas de tipo string
    def _clean_value(v: Any) -> Any:
        if isinstance(v, str):
            return INVALID_XML_RE.sub("", v)
        return v

    return df.applymap(_clean_value)
def to_excel_bytes(resultados_df: pd.DataFrame, resumen_completo_df: pd.DataFrame) -> bytes:
    # Sanitizar dataframes de entrada para evitar caracteres ilegales en XML
    resultados_df = _sanitize_excel_df(resultados_df)
    resumen_completo_df = _sanitize_excel_df(resumen_completo_df)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        # -----------------------------
        # Hoja 1: Resultados
        # -----------------------------
        if resultados_df is None or resultados_df.empty:
            tmp = pd.DataFrame(columns=[
                "Archivo", "Página/Diapositiva", "BloqueTipo", "Mensaje",
                "Sugerencias", "Oración", "Contexto", "Regla", "Categoría"
            ])
            tmp = _sanitize_excel_df(tmp)
            tmp.to_excel(w, index=False, sheet_name="Resultados")
            ws = w.sheets["Resultados"]
            ws.set_column(0, 0, 40)
        else:
            resultados_df.to_excel(w, index=False, sheet_name="Resultados")
            ws = w.sheets["Resultados"]
            for i, col in enumerate(resultados_df.columns):
                try:
                    width = min(
                        60,
                        max(
                            12,
                            int(resultados_df[col].astype(str).str.len().quantile(0.9)) + 2
                        )
                    )
                except Exception:
                    width = 22
                ws.set_column(i, i, width)

        # -----------------------------
        # Hoja 2: ResumenIncidencias
        # -----------------------------
        if resultados_df is None or resultados_df.empty:
            resumen_inc = pd.DataFrame(columns=["Archivo", "TotalIncidencias"])
        else:
            resumen_inc = (
                resultados_df.groupby("Archivo")
                .size()
                .reset_index(name="TotalIncidencias")
                .sort_values("TotalIncidencias", ascending=False)
            )

        resumen_inc = _sanitize_excel_df(resumen_inc)
        resumen_inc.to_excel(w, index=False, sheet_name="ResumenIncidencias")
        ws2 = w.sheets["ResumenIncidencias"]
        ws2.set_column(0, 0, 40)
        ws2.set_column(1, 1, 22)

        # -----------------------------
        # Hoja 3: ResumenCompleto
        # -----------------------------
        if resumen_completo_df is None:
            resumen_completo_df = pd.DataFrame()

        resumen_completo_df = _sanitize_excel_df(resumen_completo_df)
        if resumen_completo_df is None:
            resumen_completo_df = pd.DataFrame()

        resumen_completo_df.to_excel(w, index=False, sheet_name="ResumenCompleto")
        ws3 = w.sheets["ResumenCompleto"]
        for i, col in enumerate(resumen_completo_df.columns):
            ws3.set_column(i, i, 28 if i > 0 else 50)

    # ⚠️ Importante: después del with YA NO se vuelve a escribir nada en `w`
    out.seek(0)
    return out.getvalue()


# ======================================================
# FUNCIONES BROKEN LINK CHECKER (simplificadas)
# ======================================================
def _strip_invisible(s: str) -> str:
    return s.replace("\u200b", "").replace("\ufeff", "").strip()

def _looks_like_url(s: str) -> bool:
    s = s.lower().strip()
    return s.startswith(("http://", "https://")) or "." in s

def validate_url_structure(url: str) -> Tuple[bool, str]:
    if "\\" in url:
        return False, "URL contiene backslash (\\) - carácter inválido"
    if " " in url and "%20" not in url:
        return False, "URL contiene espacios sin encodear"
    if not url.startswith(("http://", "https://")):
        return False, "URL debe comenzar con http:// o https://"
    try:
        parsed = urlparse(url)
        if not parsed.netloc:
            return False, "URL sin dominio válido"
        if "." not in parsed.netloc and "localhost" not in parsed.netloc.lower():
            return False, "Dominio inválido (falta extensión)"
        return True, ""
    except Exception as e:
        return False, f"Error de estructura: {str(e)}"

def _normalize_one_url(
    raw: str,
    default_scheme: str = "https",
    allow_mailto: bool = False,
    allow_tel: bool = False,
    allow_anchors_only: bool = False,
) -> Tuple[Optional[str], str]:
    if raw is None:
        return None, "Vacío"
    s = _strip_invisible(str(raw))
    if not s:
        return None, "Vacío"
    s = s.replace("\n", " ").replace("\r", " ").strip()
    if s.startswith("#"):
        return (s, "") if allow_anchors_only else (None, "Anchor (#)")
    low = s.lower()
    if low.startswith("mailto:"):
        return (s, "") if allow_mailto else (None, "mailto")
    if low.startswith("tel:"):
        return (s, "") if allow_tel else (None, "tel")
    if not low.startswith(("http://", "https://")):
        if _looks_like_url(s):
            s = f"{default_scheme}://{s}"
        else:
            return None, "No parece URL"
    struct_valid, struct_reason = validate_url_structure(s)
    if not struct_valid:
        return None, struct_reason
    try:
        p = urlparse(s)
    except Exception:
        return None, "Parseo inválido"
    if not p.netloc:
        return None, "Sin dominio"
    netloc_raw = p.netloc.strip()
    userinfo = ""
    hostport = netloc_raw
    if "@" in netloc_raw:
        userinfo, hostport = netloc_raw.rsplit("@", 1)
    host = hostport
    port: Optional[str] = None
    if ":" in hostport:
        host, port = hostport.rsplit(":", 1)
    try:
        host_idn = host.encode("idna").decode("ascii")
    except Exception:
        host_idn = host
    scheme = p.scheme.lower()
    if port and ((scheme == "http" and port == "80") or (scheme == "https" and port == "443")):
        port = None
    if port:
        netloc_clean = f"{host_idn}:{port}"
    else:
        netloc_clean = host_idn
    if userinfo:
        netloc_clean = f"{userinfo}@{netloc_clean}"
    path = quote(p.path, safe="/%:@-._~!$&'()*+,;=")
    query = quote(p.query, safe="=&%:@-._~!$&'()*+,;/?")
    norm = urlunparse((scheme, netloc_clean, path, p.params, query, ""))
    return norm, ""

def _normalize_links(
    series: pd.Series,
    allow_mailto: bool,
    allow_tel: bool,
    allow_anchors_only: bool,
    default_scheme: str,
) -> Tuple[List[Tuple[int, str]], pd.DataFrame]:
    out: List[Tuple[int, str]] = []
    invalid_rows: List[Dict[str, Any]] = []
    for excel_row, v in enumerate(series.tolist(), start=2):
        url, reason = _normalize_one_url(
            v,
            default_scheme=default_scheme,
            allow_mailto=allow_mailto,
            allow_tel=allow_tel,
            allow_anchors_only=allow_anchors_only,
        )
        if url is None:
            invalid_rows.append({
                "Fila_Excel": excel_row,
                "Valor": "" if v is None else str(v),
                "Motivo": reason,
            })
            continue
        out.append((excel_row, url))
    return out, pd.DataFrame(invalid_rows)

def _httpx_available_or_warn() -> bool:
    if httpx is None:
        st.error("Falta la librería `httpx`. Instala con: `pip install httpx`")
        return False
    return True

def _requests_available_or_warn() -> bool:
    if requests is None:
        st.error("Falta la librería `requests`. Instala con: `pip install requests`")
        return False
    return True

def _get_random_user_agent() -> str:
    return random.choice(USER_AGENTS)

def _build_headers_for_domain(url: str) -> Dict[str, str]:
    headers = {
        "user-agent": _get_random_user_agent(),
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "accept-language": "es-PE,es;q=0.9,en;q=0.8",
        "accept-encoding": "gzip, deflate, br",
        "dnt": "1",
        "connection": "keep-alive",
        "upgrade-insecure-requests": "1",
    }
    return headers

def _host_key(url: str) -> str:
    try:
        p = urlparse(url)
        return p.netloc.lower()
    except Exception:
        return "unknown"

def _is_html_like(content_type: Optional[str]) -> bool:
    if not content_type:
        return False
    ct = content_type.lower()
    return "text/html" in ct or "application/xhtml" in ct

def _is_binary_candidate(url: str) -> bool:
    try:
        p = urlparse(url)
        path = p.path.lower()
    except Exception:
        return False
    return any(path.endswith(ext) for ext in BINARY_EXTS)

def _compute_retry_delay(retry_after_header: Optional[str], attempt: int) -> float:
    if retry_after_header:
        try:
            return float(retry_after_header)
        except Exception:
            pass
    return min(30, 1.0 * (2 ** (attempt - 1))) + random.random()

async def _fetch_limited_text_v5(
    client: "httpx.AsyncClient",
    url: str,
    timeout_s: float,
    max_bytes: int,
    range_bytes: int,
) -> Tuple[Optional[int], Dict[str, str], str, bool, str, List[str]]:
    headers = {"Range": f"bytes=0-{range_bytes-1}"}
    try:
        async with client.stream(
            "GET",
            url,
            timeout=timeout_s,
            follow_redirects=True,
            headers=headers,
        ) as r:
            final_url = str(r.url)
            history_urls = [str(resp.url) for resp in r.history]
            redirect_chain = (
                history_urls + [final_url]
                if history_urls or final_url != url
                else [final_url]
            )
            redirected = final_url != url or bool(history_urls)
            status = r.status_code
            h = {k.lower(): v for k, v in r.headers.items()}
            buf = bytearray()
            async for chunk in r.aiter_bytes():
                if not chunk:
                    continue
                take = min(len(chunk), max_bytes - len(buf))
                buf.extend(chunk[:take])
                if len(buf) >= max_bytes:
                    break
            encoding = r.encoding or "utf-8"
            try:
                text = buf.decode(encoding, errors="replace")
            except Exception:
                text = buf.decode("utf-8", errors="replace")
            return status, h, text, redirected, final_url, redirect_chain
    except Exception as e:
        return None, {}, f"{e.__class__.__name__}: {str(e)[:200]}", False, url, [url]

def _classify_v5(
    url: str,
    status_code: Optional[int],
    detail: str,
    redirected: bool,
) -> str:
    if status_code is None:
        return "ERROR"
    if status_code in (404, 410):
        return "ROTO"
    if 500 <= status_code <= 599:
        return "ERROR"
    if 400 <= status_code <= 499:
        return "ERROR"
    return "REDIRECT" if redirected else "ACTIVO"

async def _check_one_url_robust_v5(
    client: "httpx.AsyncClient",
    url: str,
    timeout_s: float,
    max_bytes: int,
    range_bytes: int,
    detect_soft_404: bool,
    retries: int,
) -> Dict[str, Any]:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if _is_binary_candidate(url):
        attempt = 0
        while attempt <= max(0, retries):
            attempt += 1
            try:
                r = await client.head(
                    url,
                    timeout=timeout_s,
                    follow_redirects=True,
                )
                final_url = str(r.url)
                history_urls = [str(resp.url) for resp in r.history]
                redirect_chain = (
                    history_urls + [final_url]
                    if history_urls or final_url != url
                    else [final_url]
                )
                redirected = final_url != url or bool(history_urls)
                status = r.status_code
                headers = {k.lower(): v for k, v in r.headers.items()}
                ct = headers.get("content-type", "")
                if status in (405, 501):
                    break
                if status in (408, 425, 429) or (500 <= status <= 599):
                    delay = _compute_retry_delay(headers.get("retry-after"), attempt)
                    if attempt <= retries:
                        await asyncio.sleep(delay)
                        continue
                detail = "OK" if status < 400 else f"HTTP {status}"
                return {
                    "Link": url,
                    "Status": _classify_v5(url, status, detail, redirected),
                    "HTTP_Code": status,
                    "Detalle": detail,
                    "Content_Type": ct,
                    "Redirected": "Sí" if redirected else "No",
                    "Timestamp": now_str,
                    "Final_URL": final_url,
                    "Redirect_Chain": " -> ".join(redirect_chain),
                    "Soft_404": "No",
                    "Score": 100,
                }
            except Exception as e:
                last_detail = f"{e.__class__.__name__}: {str(e)[:200]}"
                if attempt <= retries:
                    delay = _compute_retry_delay(None, attempt)
                    await asyncio.sleep(delay)
                    continue
                return {
                    "Link": url,
                    "Status": "ERROR",
                    "HTTP_Code": None,
                    "Detalle": last_detail,
                    "Content_Type": "",
                    "Redirected": "No",
                    "Timestamp": now_str,
                    "Final_URL": url,
                    "Redirect_Chain": url,
                    "Soft_404": "No",
                    "Score": -100,
                }
    attempt = 0
    last_detail = ""
    last_status: Optional[int] = None
    last_redirected = False
    last_ct = ""
    last_final_url = url
    last_chain: List[str] = [url]
    soft_flag = False
    content_score = 0
    while attempt <= max(0, retries):
        attempt += 1
        status, headers, text, redirected, final_url, chain = await _fetch_limited_text_v5(
            client,
            url,
            timeout_s=timeout_s,
            max_bytes=max_bytes,
            range_bytes=range_bytes,
        )
        last_status = status
        last_redirected = redirected
        last_final_url = final_url
        last_chain = chain
        last_ct = headers.get("content-type", "")
        if status is None:
            last_detail = text
            if attempt <= retries:
                delay = _compute_retry_delay(None, attempt)
                await asyncio.sleep(delay)
                continue
            break
        if status in (408, 425, 429) or (500 <= status <= 599):
            last_detail = f"HTTP {status} (transitorio)"
            if attempt <= retries:
                delay = _compute_retry_delay(headers.get("retry-after"), attempt)
                await asyncio.sleep(delay)
                continue
            break
        if status >= 400:
            last_detail = f"HTTP {status}"
            break
        detail = "OK"
        if detect_soft_404 and _is_html_like(last_ct):
            content_score = 0
            if SOFT_404_RE.search(text[:8000]):
                soft_flag = True
                last_detail = "Soft-404 detectado"
                return {
                    "Link": url,
                    "Status": "ROTO",
                    "HTTP_Code": status,
                    "Detalle": last_detail,
                    "Content_Type": last_ct,
                    "Redirected": "Sí" if redirected else "No",
                    "Timestamp": now_str,
                    "Final_URL": final_url,
                    "Redirect_Chain": " -> ".join(chain),
                    "Soft_404": "Sí",
                    "Score": content_score,
                }
        return {
            "Link": url,
            "Status": _classify_v5(url, status, detail, redirected),
            "HTTP_Code": status,
            "Detalle": detail,
            "Content_Type": last_ct,
            "Redirected": "Sí" if redirected else "No",
            "Timestamp": now_str,
            "Final_URL": final_url,
            "Redirect_Chain": " -> ".join(chain),
            "Soft_404": "No",
            "Score": content_score if content_score else 100,
        }
    return {
        "Link": url,
        "Status": _classify_v5(url, last_status, last_detail or "Error", last_redirected),
        "HTTP_Code": last_status,
        "Detalle": last_detail or "Error",
        "Content_Type": last_ct,
        "Redirected": "Sí" if last_redirected else "No",
        "Timestamp": now_str,
        "Final_URL": last_final_url,
        "Redirect_Chain": " -> ".join(last_chain),
        "Soft_404": "Sí" if soft_flag else "No",
        "Score": content_score,
    }

async def _run_link_check_ultra_v5(
    links_with_rows: List[Tuple[int, str]],
    timeout_s: float,
    concurrency_global: int,
    concurrency_per_host: int,
    detect_soft_404: bool,
    retries: int,
    verify_ssl: bool,
    max_bytes: int,
    range_bytes: int,
    progress_callback,
) -> List[Dict[str, Any]]:
    sem_global = asyncio.Semaphore(max(1, int(concurrency_global)))
    host_sems: Dict[str, asyncio.Semaphore] = {}
    def get_host_sem(url: str) -> asyncio.Semaphore:
        hk = _host_key(url)
        if hk not in host_sems:
            host_sems[hk] = asyncio.Semaphore(max(1, int(concurrency_per_host)))
        return host_sems[hk]
    limits = httpx.Limits(
        max_connections=max(10, int(concurrency_global) + 10),
        max_keepalive_connections=max(10, int(concurrency_global)),
        keepalive_expiry=30.0,
    )
    timeout = httpx.Timeout(timeout_s)
    base_headers = _build_headers_for_domain("https://example.com")
    async with httpx.AsyncClient(
        headers=base_headers,
        limits=limits,
        timeout=timeout,
        http2=False,
        verify=verify_ssl,
        follow_redirects=True,
    ) as client_ssl, httpx.AsyncClient(
        headers=base_headers,
        limits=limits,
        timeout=timeout,
        http2=False,
        verify=False,
        follow_redirects=True,
    ) as client_nossl:
        total = len(links_with_rows)
        done = 0
        results: List[Dict[str, Any]] = []
        async def worker(fila_excel: int, u: str):
            nonlocal done
            host_sem = get_host_sem(u)
            client = client_nossl if not verify_ssl else client_ssl
            client.headers.update(_build_headers_for_domain(u))
            async with sem_global:
                async with host_sem:
                    base = await _check_one_url_robust_v5(
                        client,
                        u,
                        timeout_s=timeout_s,
                        max_bytes=max_bytes,
                        range_bytes=range_bytes,
                        detect_soft_404=detect_soft_404,
                        retries=retries,
                    )
            row = dict(base)
            row["Fila_Excel"] = fila_excel
            done += 1
            progress_callback(done, total, u, row.get("Status", ""))
            return row
        tasks = [worker(fila, url) for (fila, url) in links_with_rows]
        for coro in asyncio.as_completed(tasks):
            results.append(await coro)
        return results

def run_async(coro):
    try:
        return asyncio.run(coro)
    except RuntimeError as e:
        msg = str(e)
        if "asyncio.run() cannot be called from a running event loop" in msg:
            loop = asyncio.get_event_loop()
            return loop.run_until_complete(coro)
        raise

def _infer_tipo_problema(row: pd.Series) -> str:
    status = str(row.get("Status", "") or "").upper()
    code = row.get("HTTP_Code")
    detalle = str(row.get("Detalle", "") or "")
    soft_404_flag = str(row.get("Soft_404", "") or "").strip().lower() == "sí"
    if status == "INVALIDO":
        return "FORMATO_INVALIDO"
    if status in ("ACTIVO", "REDIRECT"):
        return "SIN_PROBLEMA"
    if status == "ROTO":
        if soft_404_flag or "soft-404" in detalle.lower():
            return "SOFT_404"
        if code in (404, 410):
            return "ROTO_REAL"
        return "ROTO_REAL"
    if status == "ERROR":
        if code in (401, 403, 429):
            return "ACCESO_RESTRINGIDO"
        if code is None:
            return "ERROR_DESCONOCIDO"
        try:
            c_int = int(code)
        except Exception:
            return "ERROR_DESCONOCIDO"
        if 500 <= c_int <= 599:
            return "ERROR_SERVIDOR"
        if 400 <= c_int <= 499:
            return "ERROR_CLIENTE"
        return "ERROR_DESCONOCIDO"
    return ""

def _standardize_status_column(df: pd.DataFrame) -> pd.DataFrame:
    if "Status" not in df.columns:
        return df
    df = df.copy()
    status_upper = df["Status"].astype(str).str.upper()
    df.loc[status_upper.str.contains("REDIRECT"), "Status"] = "ACTIVO"
    df.loc[
        status_upper.str.contains("ERROR") | status_upper.str.contains("INVALIDO"),
        "Status"
    ] = "ROTO"
    return df

# ======================================================
# PDF to Word Transform
# ======================================================
BIB_HEAD_PATTERNS = [
    "fuentes bibliograficas",
    "referencias bibliograficas",
    "bibliografia",
    "referencias",
    "obras citadas",
]

class PDFBatchProcessor:
    def __init__(self, max_workers: int = None):
        self.max_workers = max_workers or min(4, os.cpu_count() or 1)
        self.cancel_event = threading.Event()
        self.progress_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()

    def process_single_pdf(self, pdf_path: str, output_dir: str, options: Dict) -> Dict:
        if self.cancel_event.is_set():
            return {"status": "cancelled", "file": pdf_path}
        if fitz is None or DocxDocument is None:
            return {
                "status": "error",
                "file": pdf_path,
                "error": "Faltan dependencias `pymupdf` o `python-docx`.",
                "elapsed_time": 0,
                "success": False,
            }
        try:
            start_time = datetime.now()
            pdf_name = Path(pdf_path).name
            self.progress_queue.put({"type": "file_start", "file": pdf_name, "total_pages": 0})
            success, output_path, stats = self._process_pdf_internal(pdf_path, output_dir, options)
            elapsed = (datetime.now() - start_time).total_seconds()
            result: Dict[str, Any] = {
                "status": "success" if success else "error",
                "file": pdf_path,
                "output": output_path,
                "stats": stats,
                "elapsed_time": elapsed,
                "success": success,
            }
            if not success:
                result["error"] = stats.get("error", "Error desconocido")
            return result
        except Exception as e:
            logger.error(f"Error procesando {pdf_path}: {e}")
            return {
                "status": "error",
                "file": pdf_path,
                "error": str(e),
                "elapsed_time": 0,
                "success": False,
            }

    def _process_pdf_internal(self, pdf_path: str, output_dir: str, options: Dict) -> Tuple[bool, str, Dict]:
        use_multithread = options.get("usar_multihilo", True)
        try:
            with fitz.open(pdf_path) as doc:
                num_pages = len(doc)
            self.progress_queue.put({"type": "file_pages", "file": Path(pdf_path).name, "total_pages": num_pages})
            output_path = Path(output_dir) / f"{Path(pdf_path).stem}.docx"
            doc_word = DocxDocument()
            if use_multithread and num_pages > 3:
                from concurrent.futures import ThreadPoolExecutor
                futures: List[Tuple[int, Any]] = []
                results: List[Tuple[int, str]] = []
                with ThreadPoolExecutor(max_workers=options.get("max_workers", 4)) as executor:
                    for page_num in range(num_pages):
                        if self.cancel_event.is_set():
                            break
                        future = executor.submit(self._extract_and_process_page, pdf_path, page_num, options)
                        futures.append((page_num, future))
                    for page_num, future in futures:
                        if self.cancel_event.is_set():
                            break
                        page_idx, text = future.result()
                        results.append((page_idx, text))
                        self.progress_queue.put({
                            "type": "page_progress",
                            "file": Path(pdf_path).name,
                            "page": page_num + 1,
                            "total": num_pages,
                        })
                results.sort(key=lambda x: x[0])
                page_texts = [text for _, text in results]
            else:
                page_texts: List[str] = []
                for page_num in range(num_pages):
                    if self.cancel_event.is_set():
                        break
                    _, text = self._extract_and_process_page(pdf_path, page_num, options)
                    page_texts.append(text)
                    self.progress_queue.put({
                        "type": "page_progress",
                        "file": Path(pdf_path).name,
                        "page": page_num + 1,
                        "total": num_pages,
                    })
            if self.cancel_event.is_set():
                return False, str(output_path), {"status": "cancelled"}
            for idx, text in enumerate(page_texts):
                if text.strip():
                    for line in text.split("\n"):
                        if line.strip():
                            doc_word.add_paragraph(line)
                if idx < len(page_texts) - 1:
                    doc_word.add_page_break()
            doc_word.save(str(output_path))
            stats = {
                "archivo": pdf_path,
                "nombre_archivo": Path(pdf_path).stem,
                "paginas_procesadas": num_pages,
                "archivo_salida": str(output_path),
                "tamano_salida": os.path.getsize(output_path) if os.path.exists(output_path) else 0,
                "errores": sum(1 for text in page_texts if "ERROR:" in text),
                "timestamp": datetime.now().isoformat(),
            }
            return True, str(output_path), stats
        except Exception as e:
            logger.error(f"Error interno procesando {pdf_path}: {e}")
            return False, str(Path(output_dir) / f"{Path(pdf_path).stem}.docx"), {"error": str(e)}

    def _extract_and_process_page(self, pdf_path: str, page_num: int, options: Dict) -> Tuple[int, str]:
        try:
            with fitz.open(pdf_path) as doc:
                if page_num >= len(doc):
                    return page_num, f"=== PÁGINA {page_num + 1} ===\nERROR: Página no existe"
                page = doc[page_num]
                raw_text = page.get_text()
                if len(raw_text.strip()) < 50:
                    raw_text = page.get_text("text")
                cleaned_text = self._clean_text(raw_text)
                if options.get("filtrar_bibliografia", False):
                    text_base = self._filter_references(cleaned_text)
                else:
                    text_base = cleaned_text
                text_no_formulas = self._filter_formulas(text_base)
                reformatted_text = self._reformat_sentences(text_no_formulas)
                result = f"=== PÁGINA {page_num + 1} ===\n{reformatted_text.strip()}"
                return page_num, result
        except Exception as e:
            error_msg = f"ERROR procesando página {page_num + 1}: {str(e)}"
            logger.error(f"{error_msg} en {pdf_path}")
            return page_num, f"=== PÁGINA {page_num + 1} ===\n{error_msg}"

    def _clean_text(self, text: str) -> str:
        replacements = {
            "\x00": "", "\x0c": "\n", "\uf0b7": "•", "\uf0a7": "§",
            "\uf0d8": "°", "\xad": "", "\t": "    ",
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        lines = [line.strip() for line in text.split("\n")]
        lines = [line for line in lines if line]
        final_lines: List[str] = []
        buffer: List[str] = []
        def flush_buffer():
            nonlocal buffer, final_lines
            if buffer:
                final_lines.append(" ".join(buffer))
                buffer = []
        for line in lines:
            if "=" in line:
                flush_buffer()
                final_lines.append(line)
                continue
            if len(line) < 80 and not line.endswith((".", "!", "?", ":", ";", ",", ")")):
                buffer.append(line)
            else:
                if buffer:
                    final_lines.append(" ".join(buffer + [line]))
                    buffer = []
                else:
                    final_lines.append(line)
        flush_buffer()
        return "\n".join(final_lines).strip()

    def _is_reference_line(self, line: str) -> bool:
        if not line or len(line.strip()) < 5:
            return False
        text = line.strip()
        if re.search(r"\(\d{4}[a-z]?\)", text) or re.search(r"\b\d{4}\.", text):
            return True
        if re.search(r"https?://\S+", text, re.IGNORECASE):
            return True
        numbers = sum(ch.isdigit() for ch in text)
        if numbers >= 4 and ("," in text or "." in text):
            if "pp." in text.lower() or "p." in text.lower():
                return True
        return False

    def _filter_references(self, text: str) -> str:
        if not text.strip():
            return text
        lines = text.split("\n")
        result: List[str] = []
        in_ref_block = False
        for i, line in enumerate(lines):
            norm = self._normalize_text(line)
            if in_ref_block:
                if not line.strip():
                    continue
                if self._is_reference_line(line):
                    continue
                in_ref_block = False
            is_bib_header = any(pattern in norm for pattern in BIB_HEAD_PATTERNS)
            if is_bib_header and self._is_reference_header(lines, i):
                in_ref_block = True
                continue
            result.append(line)
        return "\n".join(result).strip()

    def _filter_formulas(self, text: str) -> str:
        return text

    def _reformat_sentences(self, text: str) -> str:
        if not text.strip():
            return text
        text = re.sub(r"\s*\n\s*", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        text = re.sub(r"(\d+)\.\s+(\d{1,3})", r"\1.\2", text)
        text = re.sub(r"\b(\d+)\.\s+(?=[A-ZÁÉÍÓÚÑ])", r"\1§ ", text)
        text = re.sub(r"(?<!\d)\.\s+(?!\d)", ".\n", text)
        text = re.sub(r"\)\s+(?=[A-ZÁÉÍÓÚÑ¿])", ")\n", text)
        text = text.replace("§", ".")
        lines = [line.strip() for line in text.split("\n")]
        lines = [line for line in lines if line]
        return "\n".join(lines)

    def _normalize_text(self, text: str) -> str:
        if not text:
            return ""
        text = text.lower()
        text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
        text = re.sub(r"\s+", " ", text).strip()
        return text

    def _is_reference_header(self, lines: List[str], idx: int) -> bool:
        if idx < 0 or idx >= len(lines):
            return False
        start = idx + 1
        end = min(len(lines), idx + 11)
        window = [line for line in lines[start:end] if line.strip()]
        if not window:
            return False
        total = len(window)
        ref_like = sum(1 for line in window if self._is_reference_line(line))
        return ref_like >= 2 or ref_like / total >= 0.4

# ======================================================
# DESCARGA MASIVA
# ======================================================
def _format_hms(seconds: float) -> str:
    seconds_int = int(max(0, seconds))
    m, s = divmod(seconds_int, 60)
    h, m = divmod(m, 60)
    if h > 0:
        return f"{h:02d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"

def nombre_archivo_seguro(url: str, carpeta_destino: str, max_ruta: int = 240) -> str:
    nombre = url.split('/')[-1]
    nombre = nombre.split('?')[0]
    from urllib.parse import unquote as _unq
    nombre = _unq(nombre)
    caracteres_invalidos = '<>:"/\\|?*'
    for c in caracteres_invalidos:
        nombre = nombre.replace(c, '_')
    if not nombre.strip():
        nombre = "archivo_descargado"
    ruta_base = os.path.join(carpeta_destino, "")
    espacio_disponible = max_ruta - len(ruta_base)
    if espacio_disponible < 50:
        espacio_disponible = 50
    if len(nombre) > espacio_disponible:
        base, ext = os.path.splitext(nombre)
        if len(ext) > 10:
            ext = ext[:10]
        max_base = espacio_disponible - len(ext)
        if max_base < 1:
            max_base = 1
        base = base[:max_base]
        nombre = base + ext
    return nombre

def _ensure_unique_path(base_dir: Path, filename: str) -> Path:
    base_dir = Path(base_dir)
    base_dir.mkdir(parents=True, exist_ok=True)
    candidate = base_dir / filename
    if not candidate.exists():
        return candidate
    stem = Path(filename).stem
    ext = Path(filename).suffix
    counter = 1
    while True:
        alt = base_dir / f"{stem}_{counter}{ext}"
        if not alt.exists():
            return alt
        counter += 1

def _group_files_by_size(file_paths: List[str], max_mb: int = MAX_ZIP_BLOCK_MB) -> List[List[str]]:
    groups: List[List[str]] = []
    current_group: List[str] = []
    current_size = 0
    max_bytes = max_mb * 1024 * 1024
    for path in file_paths:
        try:
            size = os.path.getsize(path)
        except OSError:
            continue
        if size >= max_bytes:
            if current_group:
                groups.append(current_group)
                current_group = []
                current_size = 0
            groups.append([path])
            continue
        if current_group and current_size + size > max_bytes:
            groups.append(current_group)
            current_group = [path]
            current_size = size
        else:
            current_group.append(path)
            current_size += size
    if current_group:
        groups.append(current_group)
    return groups

def _zip_paths_to_bytes(file_paths: List[str]) -> bytes:
    buf = BytesIO()   # ✅ usar la clase ya importada
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in file_paths:
            if not os.path.exists(p):
                continue
            arcname = Path(p).name
            zf.write(p, arcname=arcname)
    buf.seek(0)
    return buf.getvalue()


def _run_descarga_masiva_streamlit(
    urls_archivos: List[str],
    progress_placeholder,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], str, Optional[str]]:

    if not urls_archivos:
        tmp_dir = Path(tempfile.mkdtemp(prefix="utp_descarga_masiva_"))
        return [], [], str(tmp_dir), None
    tmp_dir = Path(tempfile.mkdtemp(prefix="utp_descarga_masiva_"))
    session = requests.Session()
    resultados: List[Dict[str, Any]] = []
    fallidos: List[Dict[str, Any]] = []
    total = len(urls_archivos)
    start_time = time.time()

    if progress_placeholder is not None and total > 0:
        render_task_progress(
            progress_placeholder,
            "Descargando archivos",
            0.0,
            f"0/{total} archivos",
        )

    for idx, url in enumerate(urls_archivos, start=1):
        url = str(url).strip()
        descargado_ok = False
        ultimo_error = ""
        nombre_archivo = ""
        ruta_archivo = None
        try:
            nombre_archivo = nombre_archivo_seguro(url, str(tmp_dir))
            ruta_archivo = tmp_dir / nombre_archivo
            if ruta_archivo.exists():
                base, ext = os.path.splitext(nombre_archivo)
                contador = 1
                while ruta_archivo.exists():
                    sufijo = f"_{contador}"
                    espacio_disponible = 240 - len(os.path.join(str(tmp_dir), ""))
                    max_base = espacio_disponible - len(ext) - len(sufijo)
                    if max_base < 1:
                        max_base = 1
                    base_trunc = base[:max_base]
                    nombre_archivo_alt = f"{base_trunc}{sufijo}{ext}"
                    ruta_archivo = tmp_dir / nombre_archivo_alt
                    contador += 1
            for intento in range(1, MAX_INTENTOS_DESCARGA + 1):
                try:
                    resp = session.get(url, stream=True, timeout=30, headers=REQUEST_HEADERS, allow_redirects=True)
                    if resp.status_code == 200:
                        tam_header = resp.headers.get("content-length")
                        if tam_header:
                            try:
                                tam_esperado = int(tam_header)
                            except ValueError:
                                tam_esperado = None
                        else:
                            tam_esperado = None
                        with open(ruta_archivo, "wb") as f:
                            for chunk in resp.iter_content(chunk_size=CHUNK_SIZE):
                                if chunk:
                                    f.write(chunk)
                        tam_real = os.path.getsize(ruta_archivo)
                        if (tam_esperado is not None and tam_esperado == 0) or tam_real == 0:
                            try:
                                os.remove(ruta_archivo)
                            except OSError:
                                pass
                            raise Exception(f"Archivo descargado con tamaño 0 (tam_esperado={tam_esperado}, tam_real={tam_real}).")
                        descargado_ok = True
                        resultados.append({
                            "url": url,
                            "nombre_archivo": ruta_archivo.name,
                            "ruta_archivo": str(ruta_archivo),
                            "status": "OK",
                        })
                        break
                    else:
                        raise Exception(f"Código HTTP: {resp.status_code}")
                except Exception as e:
                    ultimo_error = str(e)
                    logger.warning(f"Error en intento {intento} para {nombre_archivo}: {e}")
                    if ruta_archivo and ruta_archivo.exists():
                        try:
                            os.remove(ruta_archivo)
                        except OSError:
                            pass
                    if intento < MAX_INTENTOS_DESCARGA:
                        espera = min(60, 2 ** intento)
                        logger.info(f"Reintentando en {espera} segundos para {url}...")
                        time.sleep(espera)
            if not descargado_ok:
                fallidos.append({
                    "url": url,
                    "nombre_archivo": nombre_archivo,
                    "error": ultimo_error or "Error desconocido",
                })
        except Exception as e:
            logger.error(f"Error al procesar {url}: {e}")
            fallidos.append({
                "url": url,
                "nombre_archivo": nombre_archivo,
                "error": str(e),
            })
        elapsed = time.time() - start_time
        processed = idx
        pct = processed / total
        speed = processed / elapsed if elapsed > 0 else 0.0
        eta = (total - processed) / speed if speed > 0 else 0.0
        if progress_placeholder is not None and total > 0:
            detail = (
                f"{pct*100:.1f}% | {processed}/{total} archivos "
                f"[{_format_hms(elapsed)}<{_format_hms(eta)}, {speed:.2f} archivo/s]"
            )
            render_task_progress(
                progress_placeholder,
                "Descargando archivos",
                pct,
                detail,
            )

    csv_fallidos_path: Optional[Path] = None
    if fallidos:
        csv_fallidos_path = tmp_dir / "descargas_fallidas.csv"
        pd.DataFrame(fallidos).to_csv(csv_fallidos_path, index=False, encoding="utf-8-sig")
    return resultados, fallidos, str(tmp_dir), str(csv_fallidos_path) if csv_fallidos_path else None

# ======================================================
# UI HELPERS
# ======================================================
def apply_global_styles():
    st.markdown(
        """
        <style>
        /* Reducir tamaño base de fuente para que todo se vea más compacto */
        html {
            font-size: 11px;  /* en lugar de 16px */
        }

        [data-testid="stAppViewContainer"] { background: #f3f4f6; }
        [data-testid="stSidebar"] { background: #f9fafb; }

        .utp-hero {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 18px;
            padding: 1.8rem 2.4rem;
            color: #ffffff;
            margin-bottom: 1.8rem;
            box-shadow: 0 18px 40px rgba(76, 81, 191, 0.35);
            display: flex;
            align-items: center;
            gap: 1.0rem;
        }

        .utp-hero-icon {
            width: 3.1rem;
            height: 3.1rem;
            border-radius: 999px;
            background: rgba(255,255,255,0.18);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 2.0rem;
        }
        .utp-hero-title { font-weight: 800; font-size: 1.8rem; margin-bottom: 0.15rem; }
        .utp-hero-sub { font-size: 0.92rem; opacity: 0.96; line-height: 1.4; }
        .utp-sidebar-brand {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 18px;
            padding: 1.0rem 1.1rem;
            color: #ffffff;
            box-shadow: 0 14px 32px rgba(76, 81, 191, 0.35);
            margin-bottom: 1.1rem;
        }
        .utp-sidebar-brand-title {
            font-weight: 800;
            font-size: 1.05rem;
            margin-bottom: 0.2rem;
            display: flex;
            align-items: center;
            gap: 0.4rem;
        }
        .utp-sidebar-brand-subtitle { font-size: 0.82rem; opacity: 0.92; }
        .utp-card {
            border-radius: 14px;
            border: 1px solid #e5e7eb;
            padding: 1.1rem 1.3rem 1.15rem 1.3rem;
            margin-bottom: 1.0rem;
            background: #ffffff;
            box-shadow: 0 10px 25px rgba(15,23,42,0.05);
        }
        .utp-step-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 0.7rem;
        }
        .utp-step-main {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            font-size: 1.0rem;
            font-weight: 700;
            color: #111827;
        }
        .utp-step-number {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 26px;
            height: 26px;
            border-radius: 999px;
            background: #4f46e5;
            color: #ffffff;
            font-size: 0.9rem;
            font-weight: 700;
            box-shadow: 0 3px 8px rgba(79,70,229,0.45);
        }
        .utp-step-status {
            padding: 0.18rem 0.7rem;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 600;
            border: 1px solid transparent;
            white-space: nowrap;
        }
        .utp-step-status-ok { background-color: #dcfce7; color: #166534; border-color: #bbf7d0; }
        .utp-step-status-warn { background-color: #ffedd5; color: #9a3412; border-color: #fed7aa; }
        .utp-step-status-error { background-color: #fee2e2; color: #b91c1c; border-color: #fecaca; }
        .utp-step-header-simple {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            font-size: 1.0rem;
            font-weight: 700;
            color: #111827;
            margin-bottom: 0.7rem;
        }
        .utp-step-header-simple .utp-step-number {
            width: 26px;
            height: 26px;
            border-radius: 999px;
            background: #4f46e5;
            color: #ffffff;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 0.9rem;
            font-weight: 700;
        }
        .stButton>button {
            border-radius: 999px;
            font-weight: 700;
            padding: 0.6rem 1.3rem;
            border: none;
            transition: all 0.2s ease;
        }
        .stButton>button:hover {
            transform: translateY(-1px);
            box-shadow: 0 10px 25px rgba(79,70,229,0.45);
        }
        .stDataFrame { border-radius: 10px; border: 1px solid #e5e7eb; }
        .hero-reset-anchor + div[data-testid="stButton"] {
            margin-top: -3.0rem;
            margin-bottom: 0.6rem;
            display: flex;
            justify-content: flex-end;
            padding-right: 2.4rem;
        }
        .hero-reset-anchor + div[data-testid="stButton"] > button {
            position: relative;
            width: 44px;
            height: 44px;
            border-radius: 999px;
            padding: 0;
            background-color: #ff1654;
            color: transparent;
            box-shadow: 0 12px 28px rgba(15,23,42,0.45);
        }
        .hero-reset-anchor + div[data-testid="stButton"] > button::before {
            content: "↻";
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 22px;
            color: #ffffff;
        }
        .hero-reset-anchor + div[data-testid="stButton"] > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 16px 32px rgba(15,23,42,0.55);
        }
        .utp-success-chip {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 0.5rem 1.1rem;
            border-radius: 999px;
            background-color: #22c55e;
            color: #ffffff;
            font-weight: 700;
            font-size: 0.9rem;
            box-shadow: 0 10px 24px rgba(22,163,74,0.45);
            border: none;
            margin-top: 0.45rem;
        }

        .progress-bar-ui-task {
            margin-top: 0.35rem;
            margin-bottom: 0.35rem;
        }

        .progress-bar-ui-task-header {
            font-size: 0.85rem;
            font-weight: 600;
            color: #111827;
            margin-bottom: 0.25rem;
        }

        /* Contenedor de la barra: pill ancho tipo screenshot */
        .progress-bar-ui-task-track {
            position: relative;
            width: 100%;
            height: 20px;                      /* ⬅️ antes 18px: barra mucho más gruesa */
            border-radius: 999px;
            background: #e0f2ff;               /* azul muy claro como el fondo de la captura */
            overflow: hidden;
        }

        /* Barra de progreso azul */
        .progress-bar-ui-task-bar {
            position: absolute;
            top: 0;
            left: 0;
            height: 100%;                      /* ocupa todo el alto del track */
            display: flex;
            align-items: center;               /* centra el texto verticalmente */
            justify-content: flex-start;       /* texto pegado a la izquierda */
            padding: 0 12px;
            border-radius: 999px;
            background: linear-gradient(90deg, #1d8cf8, #1554d1);
            color: #ffffff;
            font-size: 0.80rem;
            font-weight: 700;
            text-shadow: 0 1px 2px rgba(15,23,42,0.4);
            min-width: 40px;                   /* asegura una “pastilla” azul visible */
            max-width: 100%;
            box-sizing: border-box;
        }

        .progress-bar-ui-task-subtext {
            margin-top: 0.25rem;
            font-size: 0.80rem;
            color: #4b5563;
        }


        </style>
    """, unsafe_allow_html=True)

def render_sidebar_header():
    st.markdown("""
        <div class="utp-sidebar-brand">
            <div class="utp-sidebar-brand-title">
                <span>📚</span><span>Plataforma - GrammarScan</span>
            </div>
            <div class="utp-sidebar-brand-subtitle">
                Revisión automatizada de ortografía y gramática de diversos documentos académicos.
            </div>
        </div>
    """, unsafe_allow_html=True)


def render_hero(title: str, subtitle: str, icon: str = "📚"):
    st.markdown(f"""
        <div class="utp-hero">
            <div class="utp-hero-icon">{icon}</div>
            <div>
                <div class="utp-hero-title">{title}</div>
                <div class="utp-hero-sub">{subtitle}</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

def render_step_header_html(step_label: str, title: str, status: str) -> str:
    map_text = {"ok": "Listo", "warn": "Pendiente", "error": "Falta"}
    map_class = {"ok": "utp-step-status-ok", "warn": "utp-step-status-warn", "error": "utp-step-status-error"}
    status_text = map_text.get(status, "Pendiente")
    status_class = map_class.get(status, "utp-step-status-warn")
    return f"""
    <div class="utp-step-row">
        <div class="utp-step-main">
            <span class="utp-step-number">{step_label}</span>
            <span>{title}</span>
        </div>
        <div class="utp-step-status {status_class}">{status_text}</div>
    </div>
    """

def render_simple_step_header(step_label: str, title: str):
    st.markdown(f"""
        <div class="utp-step-header-simple">
            <span class="utp-step-number">{step_label}</span>
            <span>{title}</span>
        </div>
    """, unsafe_allow_html=True)

def ui_card_open():
    st.markdown('<div class="utp-card">', unsafe_allow_html=True)

def ui_card_close():
    st.markdown("</div>", unsafe_allow_html=True)

def render_success_chip(text: str):
    st.markdown(f'<div class="utp-success-chip">{text}</div>', unsafe_allow_html=True)

def render_task_progress(
    placeholder,
    title: str,
    pct: float,
    detail: str = "",
) -> None:
    """Barra de progreso unificada tipo progress-bar-ui-task."""
    pct = max(0.0, min(1.0, float(pct or 0.0)))
    pct_label = f"{pct * 100:.1f}%"
    detail_html = f'<div class="progress-bar-ui-task-subtext">{detail}</div>' if detail else ""
    html = f"""
    <div class="progress-bar-ui-task">
        <div class="progress-bar-ui-task-header">{title}</div>
        <div class="progress-bar-ui-task-track">
            <div class="progress-bar-ui-task-bar" style="width: {pct * 100:.1f}%;">
                {pct_label}
            </div>
        </div>
        {detail_html}
    </div>
    """
    placeholder.markdown(html, unsafe_allow_html=True)


# ======================================================
# ESTADO DE SESIÓN
# ======================================================
def init_session_state():
    if "module" not in st.session_state:
        st.session_state["module"] = "Home"
    if "module_radio" not in st.session_state:
        st.session_state["module_radio"] = st.session_state["module"]
    
    # Broken Link Checker states
    st.session_state.setdefault("output_dir", str(Path.cwd() / "SALIDA_LINK_CHECKER"))
    st.session_state.setdefault("status_input_filename", None)
    st.session_state.setdefault("status_input_df", None)
    st.session_state.setdefault("status_links_list", None)
    st.session_state.setdefault("status_cache", {})
    st.session_state.setdefault("status_result_df", None)
    st.session_state.setdefault("status_invalid_df", None)
    st.session_state.setdefault("status_export_df", None)
    
    # Descarga Masiva states
    st.session_state.setdefault("descarga_zip_bytes", None)
    st.session_state.setdefault("descarga_resultados", None)
    st.session_state.setdefault("descarga_fallidos", None)
    st.session_state.setdefault("descarga_download_dir", None)
    st.session_state.setdefault("descarga_fallidos_csv", None)
    
    # PDF to Word states
    st.session_state.setdefault("extraccion_zip_bytes", None)
    st.session_state.setdefault("extraccion_resultados", None)
    st.session_state.setdefault("extraccion_errores", None)
    
    # Pipeline states
    st.session_state.setdefault("pipeline_pdf_signature", None)
    st.session_state.setdefault("pipeline_pdf_done", False)
    st.session_state.setdefault("pipeline_pdf_results", None)
    st.session_state.setdefault("pipeline_pdf_errors", None)
    st.session_state.setdefault("pipeline_word_docs", None)
    st.session_state.setdefault("pipeline_ppt_docs", None)
    st.session_state.setdefault("pipeline_word_done", False)
    st.session_state.setdefault("pipeline_df_links", None)
    st.session_state.setdefault("pipeline_word_errors", None)
    st.session_state.setdefault("pipeline_word_inputs_count", 0)
    st.session_state.setdefault("pipeline_ppt_inputs_count", 0)
    st.session_state.setdefault("pipeline_status_done", False)
    st.session_state.setdefault("pipeline_reset_token", 0)
    st.session_state.setdefault("pipeline_docx_paths", [])
    st.session_state.setdefault("pipeline_docx_meta", {})
    st.session_state.setdefault("pipeline_pptx_paths", [])
    st.session_state.setdefault("pipeline_pptx_meta", {})
    
    # Bulk download states
    st.session_state.setdefault("bulk_has_valid_urls", False)
    st.session_state.setdefault("bulk_urls_archivos", None)
    st.session_state.setdefault("bulk_excel_df", None)
    st.session_state.setdefault("bulk_url_mapping", None)
    st.session_state.setdefault("pipeline_bulk_signature", None)
    st.session_state.setdefault("pipeline_bulk_done", False)
    
    # Manual upload dir
    st.session_state.setdefault("pipeline_manual_dir", None)
    
    # GrammarScan states
    st.session_state.setdefault("gs_uploader_key", 0)
    st.session_state.setdefault("gs_lang", "es")
    st.session_state.setdefault("gs_max_chars", 30000)
    st.session_state.setdefault("gs_workers", 4)
    st.session_state.setdefault("gs_excluir_biblio", True)
    st.session_state.setdefault("gs_modismos", True)
    st.session_state.setdefault("gs_final_df", None)
    st.session_state.setdefault("gs_resumen_completo_df", None)
    st.session_state.setdefault("gs_metrics", None)
    st.session_state.setdefault("gs_elapsed", 0.0)
    st.session_state.setdefault("gs_last_files_signature", None)

    st.session_state.setdefault("gs_excel_bytes", None)
    st.session_state.setdefault("gs_excel_autotrigger_done", False)

def reset_report_broken_pipeline():
    keys_to_clear = [
        "pipeline_bulk_signature", "pipeline_bulk_done", "bulk_has_valid_urls",
        "bulk_urls_archivos", "descarga_resultados", "descarga_fallidos",
        "descarga_zip_bytes", "descarga_download_dir", "descarga_fallidos_csv",
        "pipeline_pdf_signature", "pipeline_pdf_done", "pipeline_pdf_results",
        "pipeline_pdf_errors", "extraccion_resultados", "extraccion_errores",
        "extraccion_zip_bytes", "extr_usar_multihilo", "extr_max_workers",
        "pipeline_word_docs", "pipeline_ppt_docs", "pipeline_word_done",
        "pipeline_df_links", "pipeline_word_errors", "reporte_links_df",
        "pipeline_word_inputs_count", "pipeline_ppt_inputs_count",
        "pipeline_docx_paths", "pipeline_docx_meta", "pipeline_pptx_paths",
        "pipeline_pptx_meta", "pipeline_status_done", "status_input_filename",
        "status_input_df", "status_links_list", "status_cache", "status_result_df",
        "status_invalid_df", "status_export_df", "bulk_excel_df", "bulk_url_mapping",
        "pipeline_manual_dir"
    ]
    for k in keys_to_clear:
        st.session_state.pop(k, None)
    st.session_state["pipeline_reset_token"] = st.session_state.get("pipeline_reset_token", 0) + 1

def reset_grammarscan_state():
    keys_to_clear = [
        "gs_lang", "gs_max_chars", "gs_workers", "gs_excluir_biblio",
        "gs_modismos", "gs_final_df", "gs_resumen_completo_df", "gs_metrics",
        "gs_elapsed", "gs_last_files_signature",
        "gs_excel_bytes", "gs_excel_autotrigger_done",   # 👈 IMPORTANTE
    ]
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]

    st.session_state["gs_uploader_key"] += 1
    if "gs_uploader" in st.session_state:
        del st.session_state["gs_uploader"]

def reset_full_pipeline():
    reset_report_broken_pipeline()
    reset_grammarscan_state()
    st.rerun()

def on_change_module():
    st.session_state["module"] = st.session_state["module_radio"]

# ======================================================
# FUNCIONES AUXILIARES
# ======================================================
def _read_excel_safe(uploaded_file) -> pd.DataFrame:
    try:
        return pd.read_excel(uploaded_file)
    except Exception as e:
        raise RuntimeError(f"No se pudo leer el Excel: {e}") from e

def build_files_signature(uploaded_files) -> str:
    parts = []
    for uf in uploaded_files:
        size = getattr(uf, "size", None)
        parts.append(f"{uf.name}:{size}")
    parts.sort()
    return "|".join(parts)

def expand_uploaded_files(ups) -> List[LogicalFileSource]:
    logical_files: List[LogicalFileSource] = []

    for up in ups:
        name = up.name
        ext = os.path.splitext(name)[1].lower()

        if ext == ".zip":
            try:
                zip_bytes = up.getvalue()
                with zipfile.ZipFile(BytesIO(zip_bytes)) as zf:
                    for info in zf.infolist():
                        if info.is_dir():
                            continue
                        inner_ext = os.path.splitext(info.filename)[1].lower()
                        if inner_ext not in ALLOWED_DOC_EXTS:
                            continue

                        display_name = f"{Path(name).stem}/{Path(info.filename).name}"

                        def _make_reader(zbytes: bytes, inner_name: str) -> Callable[[], bytes]:
                            def _reader() -> bytes:
                                with zipfile.ZipFile(BytesIO(zbytes)) as _zf:
                                    return _zf.read(inner_name)
                            return _reader

                        logical_files.append(
                            LogicalFileSource(
                                display_name=display_name,
                                ext=inner_ext,
                                read_bytes=_make_reader(zip_bytes, info.filename),
                            )
                        )
            except Exception as e:
                st.error(f"Error leyendo ZIP '{name}': {e}")
                continue

        elif ext in ALLOWED_DOC_EXTS:
            logical_files.append(
                LogicalFileSource(
                    display_name=name,
                    ext=ext,
                    read_bytes=up.getvalue,
                )
            )
        else:
            st.warning(f"Archivo omitido por extensión no soportada: {name}")

    return logical_files


def process_grammarscan_files(
    ups,
    lang_code: str,
    max_chars_call: int,
    workers: int,
    excluir_biblio: bool,
    analizar_modismos: bool,
):
    try:
        _ = get_language_tool(lang_code)
    except Exception as e:
        st.error(f"No se pudo iniciar LanguageTool local: {e}")
        return pd.DataFrame([]), pd.DataFrame([]), {"total": 0, "n_inc": 0, "n_zero": 0, "n_err": 0}, 0.0

    modismos_patterns: List[ModismoPattern] = []
    if analizar_modismos and lang_code.startswith("es"):
        script_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
        modismos_path = os.path.join(script_dir, "modismos_ar.xlsx")
        try:
            modismos_patterns = get_modismos_patterns(modismos_path)
            st.success(f"Diccionario de modismos cargado: {len(modismos_patterns)} entradas.")
        except Exception as e:
            st.error(f"No se pudieron cargar los modismos desde '{modismos_path}': {e}")
            modismos_patterns = []

    logical_files = expand_uploaded_files(ups)
    total_seleccionados = len(logical_files)

    if total_seleccionados == 0:
        return pd.DataFrame([]), pd.DataFrame([]), {"total": 0, "n_inc": 0, "n_zero": 0, "n_err": 0}, 0.0

    all_dfs: List[pd.DataFrame] = []
    resumen_rows: List[Dict[str, Any]] = []

    progress_task = st.empty()
    if total_seleccionados > 0:
        render_task_progress(
            progress_task,
            "Analizando documentos",
            0.0,
            f"0/{total_seleccionados} archivos",
        )

    t0 = time.time()

    for i, lf in enumerate(logical_files, start=1):
        pct = i / total_seleccionados if total_seleccionados > 0 else 1.0
        detail = f"{i}/{total_seleccionados} archivos procesados" if total_seleccionados > 0 else ""
        render_task_progress(
            progress_task,
            "Analizando documentos",
            pct,
            detail,
        )
        try:
            data = lf.read_bytes()
            ext = lf.ext
            df = analyze_file(
                lf.display_name,
                data,
                lang_code,
                max_chars_call,
                workers,
                excluir_bibliografia=excluir_biblio,
                modismos_patterns=modismos_patterns,
                analizar_modismos=analizar_modismos,
            )
            if not df.empty:
                all_dfs.append(df)
                resumen_rows.append({
                    "Archivo": lf.display_name,
                    "Extension": ext,
                    "Estado": "Con incidencias",
                    "TotalIncidencias": int(df.shape[0]),
                    "Detalle": ""
                })
            else:
                resumen_rows.append({
                    "Archivo": lf.display_name,
                    "Extension": ext,
                    "Estado": "Sin incidencias o sin texto",
                    "TotalIncidencias": 0,
                    "Detalle": ""
                })
        except Exception as e:
            resumen_rows.append({
                "Archivo": lf.display_name,
                "Extension": lf.ext,
                "Estado": "Error",
                "TotalIncidencias": None,
                "Detalle": safe_str(e)
            })
            st.error(f"Error procesando {lf.display_name}: {e}")

    resumen_completo_df = pd.DataFrame(resumen_rows)
    n_inc = int(resumen_completo_df.query("Estado == 'Con incidencias'")["Archivo"].nunique()) if not resumen_completo_df.empty else 0
    n_zero = int(resumen_completo_df.query("Estado == 'Sin incidencias o sin texto'")["Archivo"].nunique()) if not resumen_completo_df.empty else 0
    n_err = int(resumen_completo_df.query("Estado == 'Error'")["Archivo"].nunique()) if not resumen_completo_df.empty else 0

    if any(len(df) for df in all_dfs):
        final_df = pd.concat(all_dfs, ignore_index=True)
    else:
        final_df = pd.DataFrame([])

    elapsed = time.time() - t0
    metrics = {
        "total": total_seleccionados,
        "n_inc": n_inc,
        "n_zero": n_zero,
        "n_err": n_err,
    }
    return final_df, resumen_completo_df, metrics, elapsed

# ======================================================
# PÁGINAS / MÓDULOS
# ======================================================

def page_home():
    # Hero principal
    render_hero(
        title=APP_TITLE,
        subtitle=(
            "Revisión automatizada y validación inteligente de enlaces contenidos "
            "en documentos académicos y administrativos."
        ),
        icon="🔗",
    )

    # Card principal de contenido
    ui_card_open()

    # Sección: Home + propósito general
    st.markdown(
        """
        ### 🏠 Home

        UTP GrammarScan es una herramienta inteligente desarrollada para el análisis automatizado de información
        contenida en diversos tipos de archivos académicos y administrativos.
        Su objetivo es optimizar los procesos de revisión documental y reducir el tiempo de revisión manual, 
        asegurando la calidad lingüística y la consistencia de los textos académicos producidos dentro de la institución.

        ### 🎯 Propósito de la Plataforma

        Automatizar el proceso completo de revisión documental desde la recolección de materiales hasta el análisis 
        lingüístico avanzado, ofreciendo una solución integral que:

        - Unifica múltiples fuentes de documentos (descarga masiva, carga manual, procesamiento automático).
        - Estandariza formatos (PDF, Word, PPT → Texto estructurado).
        - Analiza contenido con motores de gramática y ortografía profesional.
        - Genera reportes detallados y exportables listos para revisión.
        """,
        unsafe_allow_html=False,
    )

    # Sección: Funcionalidades principales
    st.markdown(
        """
        ---
        ### ✨ Funcionalidades Principales

        - **Descarga Masiva Inteligente (Extracción Automática)**. Identifica y descarga automáticamente documentos **PDF, Word y PPT** desde listados de URLs en Excel.
        - **Transformación Avanzada de Documentos PDF a Word**. Convierte documentos PDF a formato Word manteniendo la estructura y contenido textual.  
        - **Análisis Lingüístico Avanzado**. Revisión ortográfica y gramatical con LanguageTool. Detección de modismos argentinos con diccionario personalizado. 
        - **Reporte Status Excel**. Genera reportes en Excel detallados (Archivo, página, error, sugerencia, contexto).  
        """,
        unsafe_allow_html=False,
    )

    # Sección: Flujo de trabajo (NUEVA VERSIÓN)
    st.markdown("---")
    st.markdown("### ✨ Flujo de Trabajo")

    # 1) Resumen en tarjetas horizontales CON FLECHAS
    pasos_resumen = [
        ("1", "Carga de URLs", "Excel con enlaces"),
        ("2", "Descarga masiva", "Documentos origen"),
        ("3", "Transformación", "PDF → Word"),
        ("4", "Extracción", "Análisis de Documentos"),
        ("5", "Validación", "Ortográfica y Gramatical"),
        ("6", "Reporte final", "Excel Status"),
    ]

    # [card, arrow, card, arrow, ..., card]
    pesos = []
    for i in range(len(pasos_resumen)):
        pesos.append(4)  # columna de tarjeta
        if i < len(pasos_resumen) - 1:
            pesos.append(1)  # columna de flecha

    cols = st.columns(pesos)

    idx = 0
    for i, (numero, titulo, subtitulo) in enumerate(pasos_resumen):
        # tarjeta
        with cols[idx]:
            st.markdown(
                f"""
                <div style="
                    background-color: #ffffff;
                    border-radius: 14px;
                    padding: 16px 18px;
                    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.12);
                    text-align: center;
                    font-size: 0.85rem;
                ">
                    <div style="
                        display:inline-flex;
                        align-items:center;
                        justify-content:center;
                        width:32px;
                        height:32px;
                        border-radius:8px;
                        background:#3b82f6;
                        color:#ffffff;
                        font-weight:700;
                        margin-bottom:6px;
                    ">
                        {numero}
                    </div>
                    <div style="font-weight: 600;">{titulo}</div>
                    <div style="color: #6b7280; font-size: 0.75rem;">{subtitulo}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        idx += 1

        # flecha (entre tarjetas, excepto después de la última)
        if i < len(pasos_resumen) - 1:
            with cols[idx]:
                st.markdown(
                    """
                    <div style="
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        min-height:120px;
                    ">
                        <span style="font-size:1.8rem; color:#9ca3af;">➜</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            idx += 1

    st.markdown("")  # pequeño espacio debajo del timeline

 # 2) Detalle por paso en expanders (6 pasos completos)
    with st.expander("🔹 Paso 1: Carga de Excel de URLs", expanded=False):
        st.markdown(
            """
            - **Formato flexible**: soporta múltiples estructuras de columnas en Excel.  
            - **Normalización automática**: corrige y estandariza formatos de URLs.  
            - **Validación preliminar**: detecta problemas estructurales antes del procesamiento.
            """
        )

    with st.expander("🔹 Paso 2: Descarga Masiva", expanded=False):
        st.markdown(
            """
            - **Procesamiento automático**: descarga simultánea de múltiples documentos.  
            - **Gestión de errores**: registro detallado de fallos con motivos específicos.  
            - **Organización automática**: archivos descargados listos para procesamiento posterior.
            """
        )

    with st.expander("🔹 Paso 3: Transformación de Documentos", expanded=False):
        st.markdown(
            """
            - **Conversión** **PDF → Word**: extracción textual manteniendo referencias.  
            - **Procesamiento paralelo**: uso eficiente de recursos para documentos grandes.  
            - **Preservación de metadatos**: mantenimiento de información de origen.
            """
        )

    with st.expander("🔹 Paso 4: Análisis de Documentos", expanded=False):
        st.markdown(
            """
            - **Análisis exhaustivo**: procesamiento completo de documentos convertidos (PDF, Word, PPT) para extracción de contenido.  
            - **Contexto completo**: preparación optimizada del contenido para la validación lingüística.
            """
        )

    with st.expander("🔹 Paso 5: Ortografía y Gramática", expanded=False):
        st.markdown(
            """
            - **Verificación en tiempo real**: Análisis con LanguageTool.  
            - **Clasificación modismos**: detección de modismos argentinos específicos.  
            """
        )

    with st.expander("🔹 Paso 6: Reporte Final", expanded=False):
        st.markdown(
            """
            - **Excel estructurado**: exportación de reporte Excel.  
            - **Descarga Automatizada**: Descarga automática al completar el análisis.
            """
        )
    # Sección: Seguridad
    st.markdown(
        """
        ---
        ### 🔒 Seguridad y Privacidad

        - **Procesamiento local**: no se almacenan documentos en servidores externos.  
        - **Metadatos anónimos**: solo se registra información necesaria para el análisis.  
        - **Sin persistencia**: los archivos temporales se eliminan después del procesamiento.  
        - **Control total**: el usuario mantiene control completo sobre sus documentos.
        """,
        unsafe_allow_html=False,
    )

    ui_card_close()

def render_report_grammarscan():
    # ======================================================
    # 1. Bulk Document (PDF, WORD and PPT) Download
    # ======================================================
    render_hero(
        "Descarga Masiva Automática de Documentos (PDF's - Word y PPT) desde un Reporte Excel",
        "Sube tu Excel con URLs y descarga automáticamente todos los documentos y el reporte en Excel.",
        "⬇️",
    )
    
    # Botón Reiniciar
    st.markdown('<div class="hero-reset-anchor"></div>', unsafe_allow_html=True)
    if st.button("Reiniciar", key="btn_reset_report_broken"):
        reset_full_pipeline()

    # ------------------------------------------------------------------
    # VENTANA COLAPSABLE: Descarga masiva de documentos (PDF, Word, PPT)
    #   - Colapsada por defecto.
    #   - Al pulsar Reiniciar se incrementa pipeline_reset_token y se
    #     genera un label interno distinto añadiendo caracteres invisibles
    #     (\u200B), lo que fuerza a Streamlit a recrear el expander
    #     completamente colapsado, sin usar key.
    # ------------------------------------------------------------------
    exp_base_label = "Descarga masiva de documentos (PDF, Word, PPT)"
    reset_counter = st.session_state.get("pipeline_reset_token", 0)
    exp_label_internal = exp_base_label + ("\u200B" * reset_counter)

    with st.expander(exp_label_internal, expanded=False):

        # ---------- TARJETA: Selección de Excel (Paso 1) ----------
        ui_card_open()
        step1_ph = st.empty()
        
        bulk_uploader_key = f"pipeline_bulk_excel_uploader_{st.session_state.get('pipeline_reset_token', 0)}"
        uploaded_excel = st.file_uploader(
            "Seleccione el archivo Excel que contiene las URLs de los documentos a descargar",
            type=["xlsx", "xls"],
            key=bulk_uploader_key,
        )
        
        file_ok = uploaded_excel is not None
        step1_ph.markdown(
            render_step_header_html(
                "1",
                "Seleccione el archivo Excel que contiene las URLs de los documentos a descargar",
                "ok" if file_ok else "warn",
            ),
            unsafe_allow_html=True,
        )
        
        bulk_urls_archivos: List[str] = []
        df_in_bulk: Optional[pd.DataFrame] = None
        
        if file_ok:
            try:
                df_in_bulk = _read_excel_safe(uploaded_excel)
            except Exception as e:
                st.error(str(e))
            else:
                st.session_state["bulk_excel_df"] = df_in_bulk
                if "url" not in df_in_bulk.columns:
                    st.error("El Excel no contiene la columna requerida: **url**.")
                    st.caption(f"Columnas detectadas: {', '.join(map(str, df_in_bulk.columns.tolist()))}")
                else:
                    # Ya está normalizada, solo descartamos NaN
                    df_urls = df_in_bulk["url"].dropna()

                    # Aseguramos que lo que se manda a descarga también vaya limpio
                    bulk_urls_archivos = [
                        str(u).strip()
                        for u in df_urls
                        if str(u).strip().lower().endswith(DESC_EXT_PERMITIDAS)
                    ]
                    total_urls = len(df_urls)
                    total_permitidas = len(bulk_urls_archivos)

                    # 🔹 Límite duro de seguridad cuando se ejecuta en Streamlit Cloud
                    if IS_STREAMLIT_CLOUD and total_permitidas > MAX_BULK_URLS_CLOUD:
                        st.error(
                            f"El Excel contiene {total_permitidas} URLs descargables, "
                            f"lo cual supera el límite de {MAX_BULK_URLS_CLOUD} URLs "
                            "por ejecución en Streamlit Cloud. "
                            "Para más de 700 URLs, divide el Excel o ejecuta la herramienta en local."
                        )
                        st.info(
                            "📌 Documento demasiado grande para **Streamlit Cloud**; "
                            "divide el Excel en varios archivos más pequeños "
                            "o ejecuta la herramienta en tu equipo local."
                        )
                        # No habilitar la descarga masiva automática
                        st.session_state["bulk_has_valid_urls"] = False
                        st.session_state["bulk_urls_archivos"] = None
                    else:
                        if total_permitidas == 0:
                            st.warning(
                                "No se encontraron URLs que terminen en .ppt, .pptx, .pdf, .doc o .docx."
                            )
                        else:
                            st.session_state["bulk_has_valid_urls"] = True
                            st.session_state["bulk_urls_archivos"] = bulk_urls_archivos
                        
                        try:
                            excel_bytes = uploaded_excel.getbuffer()
                            bulk_signature = (uploaded_excel.name, len(excel_bytes))
                        except Exception:
                            bulk_signature = (uploaded_excel.name, 0)
                        
                        prev_bulk_sig = st.session_state.get("pipeline_bulk_signature")
                        if prev_bulk_sig != bulk_signature:
                            st.session_state["pipeline_bulk_signature"] = bulk_signature
                            st.session_state["pipeline_bulk_done"] = False
        else:
            st.caption("Carga un archivo Excel para continuar con la descarga masiva.")
        
        ui_card_close()
        
        # ---------- TARJETA: Procesar Descarga Masiva (Paso 2) ----------
        ui_card_open()
        render_simple_step_header("2", "Procesar Descarga Masiva")
        
        if not _requests_available_or_warn():
            ui_card_close()
            # salimos sólo del bloque del expander
        else:
            progress_task_bulk = st.empty()
            
            urls_archivos_state = st.session_state.get("bulk_urls_archivos") or []
            auto_trigger_bulk = bool(urls_archivos_state) and not st.session_state.get("pipeline_bulk_done", False)
            
            if urls_archivos_state and auto_trigger_bulk:
                try:
                    render_task_progress(
                        progress_task_bulk,
                        "Descargando archivos",
                        0.0,
                        "Preparando descarga masiva...",
                    )
                    with st.spinner("Descargando archivos..."):
                        resultados, fallidos, download_dir, csv_fallidos_path = _run_descarga_masiva_streamlit(
                            urls_archivos_state,
                            progress_placeholder=progress_task_bulk,
                        )
    
                    st.session_state["descarga_resultados"] = resultados
                    st.session_state["descarga_fallidos"] = fallidos
                    st.session_state["descarga_download_dir"] = download_dir
                    st.session_state["descarga_fallidos_csv"] = csv_fallidos_path
                    st.session_state["pipeline_bulk_done"] = True
    
                except Exception as e:
                    progress_task_bulk.empty()
                    st.error(f"Ocurrió un error durante la descarga masiva: {e}")
            
            resultados_ready = st.session_state.get("descarga_resultados") or []
            fallidos_ready = st.session_state.get("descarga_fallidos") or []
            
            if not (resultados_ready or fallidos_ready):
                if not urls_archivos_state:
                    st.caption("Primero carga un Excel válido con URLs para poder procesar la descarga.")
            
            if st.session_state.get("pipeline_bulk_done", False):
                render_success_chip("Descarga masiva completada")
            
            ui_card_close()
        
        # ---------- TARJETA: Descargar ZIP (Paso 3) ----------
        ui_card_open()
        render_simple_step_header("3", "Descargar todos los archivos (PDF, Word, PPT) (ZIP)")
        
        resultados_ready = st.session_state.get("descarga_resultados") or []
        download_dir = st.session_state.get("descarga_download_dir")
        csv_fallidos_path = st.session_state.get("descarga_fallidos_csv")
        
        file_paths: List[str] = []
        for r in resultados_ready:
            p = r.get("ruta_archivo")
            if p and os.path.exists(p):
                file_paths.append(str(p))
        
        if not resultados_ready:
            st.warning("Primero ejecuta el paso 2 para generar las descargas.")
        elif not file_paths:
            st.info(
                "No hay archivos descargados correctamente para comprimir en ZIP "
                "(es posible que todas las descargas hayan fallado)."
            )
        else:
            groups = _group_files_by_size(file_paths, max_mb=MAX_ZIP_BLOCK_MB)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            if len(groups) == 1:
                zip_bytes = _zip_paths_to_bytes(groups[0])
                zip_name = f"Descarga_Masiva_Documentos_{ts}.zip"
                st.download_button(
                    "⬇️ Descargar todos los archivos (ZIP)",
                    data=zip_bytes,
                    file_name=zip_name,
                    mime="application/zip",
                )
            else:
                st.info(
                    f"Los archivos descargados se han dividido en **{len(groups)} ZIPs** "
                    f"para que cada uno tenga como máximo ~{MAX_ZIP_BLOCK_MB} MB."
                )
                for idx, group in enumerate(groups, start=1):
                    zip_bytes = _zip_paths_to_bytes(group)
                    zip_name = f"Descarga_Masiva_Documentos_{ts}_parte{idx}.zip"
                    st.download_button(
                        f"⬇️ Descargar ZIP parte {idx}",
                        data=zip_bytes,
                        file_name=zip_name,
                        mime="application/zip",
                        key=f"btn_zip_part_{idx}",
                    )
        
        if csv_fallidos_path and os.path.exists(csv_fallidos_path):
            try:
                with open(csv_fallidos_path, "rb") as fh:
                    csv_bytes = fh.read()
                st.download_button(
                    "⬇️ Descargar CSV de descargas fallidas",
                    data=csv_bytes,
                    file_name="descargas_fallidas.csv",
                    mime="text/csv",
                )
            except Exception as e:
                st.warning(f"No se pudo leer el CSV de descargas fallidas: {e}")
        
        ui_card_close()
    
    # ======================================================
    # 2. PDF, WORD and PPT to Word Transformation (ZIP)
    # ======================================================
    render_hero(
        "Carga Directa de Documentos (PDF's - Word y PPT) y ZIP",
        "Arrastra tus PDFs, Word, PPT o ZIP y el sistema los procesa automáticamente.",
        "🧲",
    )
    # ... (resto de la función sigue igual)

    
    if fitz is None or DocxDocument is None:
        ui_card_open()
        st.error(
            "Faltan dependencias para este módulo.\n\n"
            "- Instala `pymupdf` (fitz)\n"
            "- Instala `python-docx`\n\n"
            "Luego vuelve a desplegar la aplicación."
        )
        ui_card_close()
        return
    
    ui_card_open()
    step_pdf1 = st.empty()
    
    pdf_uploader_key = f"pipeline_pdf_uploader_ultra_{st.session_state.get('pipeline_reset_token', 0)}"
    uploaded_files = st.file_uploader(
        "Selecciona uno o más archivos PDF, Word, PPT o ZIP (ZIP con PDFs/Word/PPT en su interior)",
        type=["pdf", "docx", "pptx", "doc", "ppt", "zip"],
        accept_multiple_files=True,
        key=pdf_uploader_key,
        help="Si ya ejecutaste la Descarga Masiva, aquí llegarán automáticamente los documentos.",
    )
    
    # --- NUEVO: listas basadas en rutas + meta ---
    pdf_input_paths: List[str] = []
    docx_input_paths: List[str] = []
    pptx_input_paths: List[str] = []
    unsupported_office: List[Dict[str, str]] = []
    
    docx_meta: Dict[str, Dict[str, Any]] = {}
    pptx_meta: Dict[str, Dict[str, Any]] = {}
    pdf_meta: Dict[str, Dict[str, Any]] = {}
    
    # Directorio base para guardar archivos subidos manualmente
    manual_dir = st.session_state.get("pipeline_manual_dir")
    if manual_dir is None:
        manual_dir = tempfile.mkdtemp(prefix="utp_pipeline_manual_")
        st.session_state["pipeline_manual_dir"] = manual_dir
    manual_dir_path = Path(manual_dir)
    
    # --- 2.1 Documentos provenientes de Descarga Masiva ---
    resultados_desc = st.session_state.get("descarga_resultados") or []
    for r in resultados_desc:
        ruta = r.get("ruta_archivo")
        if not ruta:
            continue
        ruta_str = str(ruta)
        ext = Path(ruta_str).suffix.lower()
        url_or_source = r.get("url")
        
        if ext == ".pdf":
            pdf_input_paths.append(ruta_str)
            pdf_meta[ruta_str] = {
                "display_name": Path(ruta_str).name,
                "source_url": url_or_source,
                "origin": "bulk_download",
            }
        elif ext == ".docx":
            docx_input_paths.append(ruta_str)
            docx_meta[ruta_str] = {
                "display_name": Path(ruta_str).name,
                "source_url": url_or_source,
                "origin": "bulk_download",
            }
        elif ext == ".pptx":
            pptx_input_paths.append(ruta_str)
            pptx_meta[ruta_str] = {
                "display_name": Path(ruta_str).name,
                "source_url": url_or_source,
                "origin": "bulk_download",
            }
        elif ext in (".doc", ".ppt"):
            unsupported_office.append({
                "Archivo": Path(ruta_str).name,
                "Motivo": f"Formato {ext} no soportado. Convierte a .docx o .pptx para poder analizar los links.",
            })
    
    # --- 2.2 Documentos / ZIP subidos manualmente ---
    if uploaded_files:
        for f in uploaded_files:
            fname = f.name
            ext = Path(fname).suffix.lower()
            dest_path = manual_dir_path / Path(fname).name
            
            def _write_if_needed(target: Path, data_bytes: bytes):
                try:
                    with open(target, "wb") as fw:
                        fw.write(data_bytes)
                except Exception as exc:
                    st.warning(f"No se pudo guardar el archivo '{target.name}' en disco: {exc}")
            
            if ext in (".pdf", ".docx", ".pptx", ".doc", ".ppt"):
                data_bytes = f.getbuffer()
                
                if ext == ".pdf":
                    _write_if_needed(dest_path, data_bytes)
                    pdf_input_paths.append(str(dest_path))
                    pdf_meta[str(dest_path)] = {
                        "display_name": Path(fname).name,
                        "source_url": fname,
                        "origin": "upload",
                    }
                elif ext == ".docx":
                    _write_if_needed(dest_path, data_bytes)
                    docx_input_paths.append(str(dest_path))
                    docx_meta[str(dest_path)] = {
                        "display_name": Path(fname).name,
                        "source_url": fname,
                        "origin": "upload",
                    }
                elif ext == ".pptx":
                    _write_if_needed(dest_path, data_bytes)
                    pptx_input_paths.append(str(dest_path))
                    pptx_meta[str(dest_path)] = {
                        "display_name": Path(fname).name,
                        "source_url": fname,
                        "origin": "upload",
                    }
                elif ext in (".doc", ".ppt"):
                    unsupported_office.append({
                        "Archivo": fname,
                        "Motivo": f"Formato {ext} no soportado. Convierte a .docx o .pptx antes de cargarlo.",
                    })
            
            elif ext == ".zip":
                try:
                    zdata = BytesIO(f.getbuffer())
                    with zipfile.ZipFile(zdata, "r") as zf:
                        for info in zf.infolist():
                            if info.is_dir():
                                continue
                            inner_name = Path(info.filename).name
                            inner_ext = Path(inner_name).suffix.lower()
                            file_bytes = zf.read(info)
                            inner_dest = manual_dir_path / inner_name
                            _write_if_needed(inner_dest, file_bytes)
                            
                            if inner_ext == ".pdf":
                                pdf_input_paths.append(str(inner_dest))
                                pdf_meta[str(inner_dest)] = {
                                    "display_name": inner_name,
                                    "source_url": fname,
                                    "origin": "upload_zip",
                                }
                            elif inner_ext == ".docx":
                                docx_input_paths.append(str(inner_dest))
                                docx_meta[str(inner_dest)] = {
                                    "display_name": inner_name,
                                    "source_url": fname,
                                    "origin": "upload_zip",
                                }
                            elif inner_ext == ".pptx":
                                pptx_input_paths.append(str(inner_dest))
                                pptx_meta[str(inner_dest)] = {
                                    "display_name": inner_name,
                                    "source_url": fname,
                                    "origin": "upload_zip",
                                }
                            elif inner_ext in (".doc", ".ppt"):
                                unsupported_office.append({
                                    "Archivo": inner_name,
                                    "Motivo": f"Formato {inner_ext} no soportado. Convierte a .docx o .pptx para poder analizar los links.",
                                })
                except Exception as e:
                    st.warning(f"No se pudo leer el ZIP `{fname}`: {e}")
    
    has_docs = bool(pdf_input_paths or docx_input_paths or pptx_input_paths)
    
    step_pdf1.markdown(
        render_step_header_html(
            "4",
            "Agregar documentos (PDF, Word, PPT) directos o desde ZIP",
            "ok" if has_docs else "warn",
        ),
        unsafe_allow_html=True,
    )
    
    if not has_docs:
        st.caption("Agrega documentos manualmente o ejecuta primero la Descarga Masiva para que lleguen aquí automáticamente.")
        ui_card_close()
        return
    
    # --- Firma de documentos para controlar re-ejecuciones ---
    all_paths_for_signature: List[str] = []
    all_paths_for_signature.extend(pdf_input_paths)
    all_paths_for_signature.extend(docx_input_paths)
    all_paths_for_signature.extend(pptx_input_paths)
    
    signature: List[Tuple[str, int]] = []
    for p in all_paths_for_signature:
        try:
            size = os.path.getsize(p)
        except OSError:
            size = 0
        signature.append((str(Path(p).name), int(size)))
    signature.sort()
    
    prev_sig = st.session_state.get("pipeline_pdf_signature")
    if prev_sig != signature:
        st.session_state["pipeline_pdf_signature"] = signature
        st.session_state["pipeline_pdf_done"] = False
        st.session_state["pipeline_pdf_results"] = None
        st.session_state["pipeline_pdf_errors"] = None
        
        # RESET de rutas generadas y reporte de links / status
        st.session_state["pipeline_docx_paths"] = None
        st.session_state["pipeline_docx_meta"] = None
        st.session_state["pipeline_pptx_paths"] = None
        st.session_state["pipeline_pptx_meta"] = None
        st.session_state["pipeline_word_done"] = False
        st.session_state["pipeline_df_links"] = None
        st.session_state["pipeline_word_errors"] = None
        
        st.session_state["pipeline_status_done"] = False
        st.session_state["status_result_df"] = None
        st.session_state["status_export_df"] = None
        st.session_state["status_invalid_df"] = None
    
    # --- Listado de archivos seleccionados ---
    def _build_pdf_table_from_paths(paths: List[str]) -> pd.DataFrame:
        rows = []
        for p in paths:
            path_obj = Path(p)
            try:
                size_mb = path_obj.stat().st_size / (1024 * 1024)
            except OSError:
                size_mb = 0.0
            pages = "?"
            if fitz is not None and path_obj.exists():
                try:
                    with fitz.open(path_obj) as doc_pdf:
                        pages = len(doc_pdf)
                except Exception:
                    pages = "?"
            rows.append({
                "Nombre": path_obj.name,
                "Tamaño_MB": round(size_mb, 2),
                "Páginas": pages,
            })
        return pd.DataFrame(rows)
    
    def _build_office_table_from_paths(paths: List[str], meta: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
        rows = []
        for p in paths:
            path_obj = Path(p)
            try:
                size_mb = path_obj.stat().st_size / (1024 * 1024)
            except OSError:
                size_mb = 0.0
            m = meta.get(p, {})
            display = m.get("display_name", path_obj.name)
            origen = m.get("origin", "")
            rows.append({
                "Nombre": display,
                "Tamaño_MB": round(size_mb, 2),
                "Origen": origen,
            })
        return pd.DataFrame(rows)
    
    if pdf_input_paths:
        with st.expander("Archivos PDF seleccionados", expanded=False):
            df_files = _build_pdf_table_from_paths(pdf_input_paths)
            st.dataframe(df_files, use_container_width=True, height=260)
    
    if docx_input_paths:
        with st.expander("Archivos Word (DOCX) seleccionados", expanded=False):
            df_files_word = _build_office_table_from_paths(docx_input_paths, docx_meta)
            st.dataframe(df_files_word, use_container_width=True, height=260)
    
    if pptx_input_paths:
        with st.expander("Archivos PPTX seleccionados", expanded=False):
            df_files_ppt = _build_office_table_from_paths(pptx_input_paths, pptx_meta)
            st.dataframe(df_files_ppt, use_container_width=True, height=260)
    
    if unsupported_office:
        with st.expander("⚠️ Archivos Office no soportados", expanded=False):
            df_unsup = pd.DataFrame(unsupported_office)
            st.dataframe(df_unsup, use_container_width=True, height=260)
    
    # Opciones de procesamiento
    with st.expander("Opciones de procesamiento de PDFs", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            usar_multihilo = st.toggle(
                "Usar procesamiento paralelo por páginas (solo PDFs)",
                value=True,
                help="Recomendado cuando los PDFs tienen muchas páginas.",
                key="extr_usar_multihilo",
            )
        with col2:
            max_workers = st.number_input(
                "Número máximo de workers para PDFs",
                min_value=1,
                max_value=16,
                value=4,
                step=1,
                key="extr_max_workers",
            )
    
    # 4. Procesar todos los documentos (Paso 5)
    render_simple_step_header("5", "Procesar todos los documentos (PDF, Word, PPT)")
    
    progress_pdf_task = st.empty()

    
    auto_trigger_pdf = not st.session_state.get("pipeline_pdf_done", False)
    
    if auto_trigger_pdf:
        try:
            render_task_progress(
                progress_pdf_task,
                "Procesando documentos (PDF → Word)",
                0.0,
                "Iniciando procesamiento de documentos...",
            )

            resultados_pdf: List[Dict[str, Any]] = []
            errores_pdf: List[Dict[str, Any]] = []

            if pdf_input_paths:
                total_pdf_files = len(pdf_input_paths)
                with st.spinner("Extrayendo texto y generando archivos Word desde PDFs..."):
                    for idx, pdf_path in enumerate(pdf_input_paths, start=1):
                        try:
                            processor = PDFBatchProcessor(max_workers=int(max_workers))
                            options = {
                                "usar_multihilo": bool(usar_multihilo),
                                "max_workers": int(max_workers),
                                "filtrar_bibliografia": False,
                            }
                            tmp_dir = tempfile.mkdtemp(prefix="utp_pdf_extr_")
                            result = processor.process_single_pdf(pdf_path, tmp_dir, options)
                            resultados_pdf.append(result)
                        except Exception as e:
                            errores_pdf.append({
                                "status": "error",
                                "file": pdf_path,
                                "error": str(e),
                                "success": False,
                            })

                        pct = idx / total_pdf_files if total_pdf_files > 0 else 1.0
                        detail = f"{idx}/{total_pdf_files} PDFs procesados"
                        render_task_progress(
                            progress_pdf_task,
                            "Procesando documentos (PDF → Word)",
                            pct,
                            detail,
                        )
            else:
                # No hay PDFs; marcamos la tarea como completada
                render_task_progress(
                    progress_pdf_task,
                    "Procesando documentos (PDF → Word)",
                    1.0,
                    "No hay PDFs pendientes; se continuará con Word/PPTX.",
                )

            # ... (resto del código que construye combined_docx_paths, etc.)

            
            # 2) DOCX generados desde PDFs + DOCX originales (rutas)
            generated_docx_paths: List[str] = []
            generated_docx_meta: Dict[str, Dict[str, Any]] = {}
            
            for r in resultados_pdf or []:
                if r.get("status") == "success":
                    out_path = r.get("output")
                    if out_path and os.path.exists(out_path):
                        out_path_str = str(out_path)
                        original_pdf_name = r.get("file", "")
                        source_url = pdf_meta.get(original_pdf_name, {}).get("source_url")
                        display_stem = Path(original_pdf_name).stem
                        display_name = f"{display_stem}.docx"
                        
                        generated_docx_paths.append(out_path_str)
                        generated_docx_meta[out_path_str] = {
                            "display_name": display_name,
                            "source_url": source_url,
                            "origin": "pdf_to_word",
                        }
            
            combined_docx_paths = list(dict.fromkeys(docx_input_paths + generated_docx_paths))
            combined_docx_meta = dict(docx_meta)
            combined_docx_meta.update(generated_docx_meta)
            
            st.session_state["pipeline_docx_paths"] = combined_docx_paths
            st.session_state["pipeline_docx_meta"] = combined_docx_meta
            st.session_state["pipeline_pptx_paths"] = pptx_input_paths
            st.session_state["pipeline_pptx_meta"] = pptx_meta
            
            st.session_state["pipeline_pdf_results"] = resultados_pdf
            st.session_state["pipeline_pdf_errors"] = errores_pdf
            st.session_state["extraccion_resultados"] = resultados_pdf
            st.session_state["extraccion_errores"] = errores_pdf
            st.session_state["pipeline_pdf_done"] = True
            st.session_state["pipeline_word_inputs_count"] = len(combined_docx_paths)
            st.session_state["pipeline_ppt_inputs_count"] = len(pptx_input_paths)
            
            # progress_bar_pdf.empty()
            # status_text_pdf.empty()
        except Exception as e:
            # progress_bar_pdf.empty()
            # status_text_pdf.empty()
            st.error(f"Ocurrió un error durante el procesamiento de documentos: {e}")
    else:
        resultados_pdf = st.session_state.get("pipeline_pdf_results") or []
        errores_pdf = st.session_state.get("pipeline_pdf_errors") or []
        word_count = st.session_state.get("pipeline_word_inputs_count", 0)
        ppt_count = st.session_state.get("pipeline_ppt_inputs_count", 0)
        
        if resultados_pdf or errores_pdf or word_count or ppt_count:
            total_ok = sum(1 for r in resultados_pdf if r.get("status") == "success")
            total_err = len(errores_pdf)
            total_pdf_files = total_ok + total_err
            
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("PDF procesados", total_pdf_files)
            m2.metric("PDF OK", total_ok)
            m3.metric("PDF con error", total_err)
            m4.metric("Word (DOCX) detectados", word_count)
            m5.metric("PPTX detectados", ppt_count)
    
    if st.session_state.get("pipeline_pdf_done", False):
        render_success_chip("Procesamiento de documentos completado")
    
    ui_card_close()
    
    # ======================================================
    # 3. GrammarScan - Ortografía y Gramática
    # ======================================================
    render_hero(
        "UTP GrammarScan — Ortografía y Gramática (PDF, DOCX, PPTX, TXT)",
        "Herramienta inteligente desarrollada para el análisis y revisión automatizada de ortografía y gramática de diversos tipos de archivos académicos.",
        "📂",
    )
    
    # Parámetros (expander colapsado por defecto)
    with st.expander("Parámetros", expanded=False):
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            lang_code = st.selectbox("Idioma", ["es", "en-US", "pt-BR", "fr", "de"], index=0, key="gs_lang")
        with c2:
            max_chars_call = st.number_input(
                "Máx. caracteres por llamada (LOCAL)",
                5000, 80000, 30000,
                help="Se agrupan páginas/diapos hasta este límite para mantener contexto.",
                key="gs_max_chars",
            )
        with c3:
            workers = st.slider(
                "Trabajadores (hilos)",
                1, max(2, os.cpu_count() or 4), min(4, (os.cpu_count() or 4)),
                help="Paraleliza el troceo por páginas. Las llamadas a LT se serializan para estabilidad.",
                key="gs_workers",
            )
    
    excluir_biblio = st.checkbox(
        "Excluir secciones/entradas de bibliografía (APA, MLA, IEEE, Vancouver)",
        value=True,
        key="gs_excluir_biblio",
    )
    
    analizar_modismos = st.checkbox(
        "Detectar modismos argentinos (modismos_ar.xlsx)",
        value=True if st.session_state.get("gs_lang", "es").startswith("es") else False,
        key="gs_modismos",
    )
    
    with st.expander("Estado del motor", expanded=False):
        st.write(f"Java detectado: **{find_java()}**")
        st.write("Backend LanguageTool: **local** (una sola instancia por sesión).")
        if not find_java():
            st.error("No se puede continuar sin Java.")
            st.stop()
    
    # Paso 6: subir archivos
    ui_card_open()
    step6_ph = st.empty()
    
    # Combinar archivos de diferentes fuentes
    all_uploaded_files = []
    
    # 1. Archivos DOCX generados del pipeline
    docx_paths = st.session_state.get("pipeline_docx_paths") or []
    for docx_path in docx_paths:
        # Crear un objeto similar a UploadedFile
        class FakeUploadedFile:
            def __init__(self, path, display_name):
                self.name = display_name
                self.path = path
                
            def getvalue(self):
                with open(self.path, 'rb') as f:
                    return f.read()
                
            @property
            def size(self):
                return os.path.getsize(self.path)
        
        meta = st.session_state.get("pipeline_docx_meta", {}).get(docx_path, {})
        display_name = meta.get("display_name", Path(docx_path).name)
        fake_file = FakeUploadedFile(docx_path, display_name)
        all_uploaded_files.append(fake_file)
    
    # 2. Archivos PPTX del pipeline
    pptx_paths = st.session_state.get("pipeline_pptx_paths") or []
    for pptx_path in pptx_paths:
        class FakeUploadedFile:
            def __init__(self, path, display_name):
                self.name = display_name
                self.path = path
                
            def getvalue(self):
                with open(self.path, 'rb') as f:
                    return f.read()
                
            @property
            def size(self):
                return os.path.getsize(self.path)
        
        meta = st.session_state.get("pipeline_pptx_meta", {}).get(pptx_path, {})
        display_name = meta.get("display_name", Path(pptx_path).name)
        fake_file = FakeUploadedFile(pptx_path, display_name)
        all_uploaded_files.append(fake_file)
    
    # 3. Archivos PDF originales (opcional, si se quieren analizar directamente)
    # Podríamos permitir analizar PDFs directamente también
    
    # Uploader adicional para archivos manuales
    uploader_key = f"gs_uploader_{st.session_state['gs_uploader_key']}"
    manual_ups = st.file_uploader(
        "6 Sube uno o varios archivos (.pdf, .docx, .pptx, .txt, .zip)",
        type=["pdf", "docx", "pptx", "txt", "zip"],
        accept_multiple_files=True,
        key=uploader_key,
        help="También puedes subir archivos adicionales manualmente.",
    )
    
    if manual_ups:
        all_uploaded_files.extend(manual_ups)
    
    have_files = bool(all_uploaded_files)
    step6_ph.markdown(
        render_step_header_html("6", "Sube uno o varios archivos (.pdf, .docx, .pptx, .txt, .zip)", have_files),
        unsafe_allow_html=True,
    )
    
    if have_files:
        st.success(f"{len(all_uploaded_files)} archivo(s) seleccionado(s).")
        st.caption(
            "Documentos del pipeline automático + archivos subidos manualmente. "
            "Para lotes muy grandes es más estable subir archivos .zip."
        )
    else:
        st.info("Sube al menos un archivo para iniciar el análisis automático.")
    
    ui_card_close()
    
    # Recuperar estado previo
    final_df = st.session_state.get("gs_final_df")
    resumen_completo_df = st.session_state.get("gs_resumen_completo_df")
    metrics = st.session_state.get("gs_metrics")
    elapsed = st.session_state.get("gs_elapsed", 0.0)
    
    # Procesamiento automático cuando hay archivos
    if have_files:
        signature = build_files_signature(all_uploaded_files)
        last_signature = st.session_state.get("gs_last_files_signature")
        
        need_processing = (final_df is None) or (signature != last_signature)
        
        if need_processing:
            final_df, resumen_completo_df, metrics, elapsed = process_grammarscan_files(
                ups=all_uploaded_files,
                lang_code=st.session_state.get("gs_lang", "es"),
                max_chars_call=st.session_state.get("gs_max_chars", 30000),
                workers=st.session_state.get("gs_workers", 4),
                excluir_biblio=st.session_state.get("gs_excluir_biblio", True),
                analizar_modismos=st.session_state.get("gs_modismos", True),
            )
            st.session_state["gs_final_df"] = final_df
            st.session_state["gs_resumen_completo_df"] = resumen_completo_df
            st.session_state["gs_metrics"] = metrics
            st.session_state["gs_elapsed"] = elapsed
            st.session_state["gs_last_files_signature"] = signature
    else:
        st.session_state["gs_last_files_signature"] = None
    
    # Paso 7: Procesar documentos (automático) + resultados
    ui_card_open()
    render_simple_step_header("7", "Procesar documentos (análisis automático)")
    
    if not have_files:
        st.info("Aún no hay análisis. Sube archivos en el Paso 6 para iniciar el proceso automáticamente.")
    else:
        if metrics:
            cA, cB, cC, cD = st.columns(4)
            with cA:
                st.metric("Seleccionados", metrics.get("total", 0))
            with cB:
                st.metric("Con incidencias", metrics.get("n_inc", 0))
            with cC:
                st.metric("Sin incidencias / sin texto", metrics.get("n_zero", 0))
            with cD:
                st.metric("Errores", metrics.get("n_err", 0))
        
        if final_df is not None and not final_df.empty:
            st.subheader("📑 Resultados (detalle de incidencias)")
            st.dataframe(final_df, use_container_width=True, hide_index=True)
        else:
            st.info("No se encontraron incidencias en los archivos procesados.")
        
        if elapsed:
            st.caption(f"⏱️ Tiempo total: {elapsed:0.2f}s")
    
    ui_card_close()

    # Paso 8: Excel (Resultados + Resúmenes)
    ui_card_open()
    render_simple_step_header("8", "Excel (Resultados + Resúmenes)")
    
    if final_df is None or resumen_completo_df is None or resumen_completo_df.empty:
        st.info("Aún no hay resultados para exportar. Procesa documentos primero.")
        # Si no hay datos, limpiamos estado de Excel para futuros corridas
        st.session_state["gs_excel_bytes"] = None
        st.session_state["gs_excel_autotrigger_done"] = False
    else:
        # Generar bytes del Excel y guardarlos en sesión
        excel_bytes = to_excel_bytes(final_df, resumen_completo_df)
        st.session_state["gs_excel_bytes"] = excel_bytes

        # Botón manual (por si el usuario quiere descargar de nuevo)
        st.download_button(
            "⬇️ Excel (Resultados + Resúmenes)",
            data=excel_bytes,
            file_name="UTP_GrammarScan_Resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Auto-descarga solo la primera vez que se genera el Excel
        if not st.session_state.get("gs_excel_autotrigger_done", False):
            try:
                b64 = base64.b64encode(excel_bytes).decode("utf-8")
                file_name = "UTP_GrammarScan_Resultados.xlsx"

                # Componente HTML invisible que dispara la descarga al cargarse
                components.html(
                    f"""
                    <html>
                        <body>
                            <a id="auto_download_link"
                               download="{file_name}"
                               href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}">
                            </a>
                            <script>
                                document.getElementById('auto_download_link').click();
                            </script>
                        </body>
                    </html>
                    """,
                    height=0,
                    width=0,
                )

                st.session_state["gs_excel_autotrigger_done"] = True
            except Exception as e:
                st.warning(f"No se pudo iniciar la descarga automática del Excel: {e}")
    
    ui_card_close()

# ======================================================
# MAIN
# ======================================================
def main():
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon=APP_ICON,
        layout="wide",
        initial_sidebar_state="expanded",
    )
    
    apply_global_styles()
    init_session_state()
    
    with st.sidebar:
        render_sidebar_header()
        
        st.radio(
            "Módulos",
            MODULES,
            index=MODULES.index(st.session_state["module"]),
            key="module_radio",
            on_change=on_change_module,
        )
        module = st.session_state["module"]
        
        st.markdown("---")
        with st.expander("Recomendaciones"):
            st.markdown(
                """
                - Se recomienda en el proceso "1" descargar de forma masiva entre **500 - 700** registros como máximo.  
                - Para más de **700** registros, lo recomendable es:  
                    Ejecutar el app en local o dividir el Excel de Url en varios archivos (por ejemplo bloques de 500 o 700 registros) y procesarlos por partes.  
                - Esto para evitar que el contenedor de **Streamlit Cloud** se quede sin memoria (~1 GB de RAM) 
                """
            )
    
    module = st.session_state["module"]
    
    if module == "Home":
        page_home()
    elif module == "Report GrammarScan":
        render_report_grammarscan()
    else:
        render_hero(title=module, subtitle="Módulo no encontrado.", icon="⚠️")
        st.error("Módulo seleccionado no existe.")

if __name__ == "__main__":
    main()




























