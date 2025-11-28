import os
import re
import time
import shutil
import threading
import zipfile
from xml.etree import ElementTree as ET
from bisect import bisect_right
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from typing import Any, Dict, List, Tuple

import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from pptx import Presentation
from dataclasses import dataclass  # <<< NUEVO


# =========================
# Parámetros y utilidades
# =========================

LINES_PER_TXT_PAGE = 50
PAGE_SEP = "\n\f\n"  # separador único entre páginas/diapositivas (no visible)

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


# --- NUEVO: util para tildes en modismos ---
ACCENTED_VOWELS = "áéíóúÁÉÍÓÚ"


def _has_accented_vowel(s: str) -> bool:
    return any(ch in ACCENTED_VOWELS for ch in s or "")


# =========================
# MODELOS Y CARGA DE MODISMOS
# =========================

@dataclass
class ModismoPattern:
    modismo: str
    tipo: str           # 'literal' o 'regex'
    patron: str         # patrón regex en texto
    sugerencia: str
    comentario: str
    regex: re.Pattern   # patrón compilado


def _normalize_regex_pattern(patron: str) -> str:
    """
    Normaliza el patrón leído desde Excel.
    Convierte '\\\\' en '\\' para que secuencias como '\\b' se interpreten bien.
    """
    if not isinstance(patron, str):
        patron = str(patron or "")
    patron = patron.replace("\\\\", "\\")
    return patron.strip()


def load_modismos_from_excel(path: str) -> List[ModismoPattern]:
    """
    Lee 'modismos_ar.xlsx' y devuelve una lista de patrones compilados.
    Columnas esperadas:
      - modismo (str)
      - tipo ('literal' o 'regex')
      - patron (opcional si tipo = 'literal')
      - sugerencia (str)
      - comentario (opcional)
    """
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
            # patrón literal más robusto: frontera de palabra basada en no-\w
            patron = r"(?<!\w)" + re.escape(modismo) + r"(?!\w)"
        else:
            base = patron_cfg if patron_cfg else modismo
            patron = _normalize_regex_pattern(base)

        try:
            rx = re.compile(patron, flags=re.IGNORECASE | re.UNICODE)
        except re.error as e:
            print(f"[AVISO] Regex inválida en modismos_ar.xlsx para '{modismo}': {e}")
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


# =========================
# Detección de bibliografía
# =========================

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


# Patrones de líneas de referencia
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
    re.compile(
        r"^[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+,\s[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+\.\s.+\.\s\d{4}\.",
        re.UNICODE
    ),
    re.compile(r"^\[\d+\]\s.+"),
    re.compile(
        r"^\d+\.\s[A-ZÁÉÍÓÚÑ][\w'\-]+.*\b(19|20)\d{2}\b.*\d+(?:\(\d+\))?:\d+(-\d+)?\.?$"
    ),
    re.compile(
        r"^[A-ZÁÉÍÓÚÑ][\w'\-]+.*\.\s+\S+\.\s+\(?\b(19|20)\d{2}\)?\.?$"
    )
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


# =========================
# Detección de inglés, latín y código
# =========================

EN_COMMON_WORDS = {
    "the", "and", "or", "but", "if", "then", "else", "when", "where", "who", "what",
    "which", "while", "for", "from", "to", "in", "on", "at", "by", "of", "with",
    "this", "that", "these", "those",
    "is", "are", "was", "were", "be", "been", "being",
    "have", "has", "had", "do", "does", "did",
    "can", "could", "should", "would", "will", "may", "might", "must",
    "not", "no", "yes", "a", "an", "as",
    "we", "you", "they", "he", "she", "it", "his", "her", "their", "our", "us", "your",
    "about", "into", "after", "before", "over", "under", "again", "further", "then",
    "once", "here", "there", "because", "very", "also", "just", "such", "only",
    "own", "same", "so", "than", "too", "more", "most", "less",
    "any", "each", "few", "some", "other", "another",
    "new", "use", "used", "using", "example",
    "important", "information", "data", "result", "results",
    "conclusion", "paper", "study", "work", "research",
    "introduction", "method", "methods", "discussion", "however", "therefore",
    "between", "within", "without"
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
    "mutatis", "mutandis", "bonae", "fidei", "sui", "generis", "erga", "omnes",
    "habeas"
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
    "boolean ", "bool ", "null", "true", "false",
    "system.out", "console.log", "printf", "println",
    "#include", "using ", "namespace ",
    "def ", "import ", "from ", "return ",
    "try:", "catch", "finally", "except ",
    "=>", "lambda", "function "
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


# =========================
# Lectores por página/diapo
# =========================

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


def read_pptx_slides(bio: BytesIO) -> List[Tuple[int, str]]:
    try:
        prs = Presentation(bio)
        slides: List[Tuple[int, str]] = []
        slide_h = float(prs.slide_height) if hasattr(prs, "slide_height") else None

        for s_idx, slide in enumerate(prs.slides, start=1):
            chunk: List[str] = []
            for sh in slide.shapes:
                if (
                    hasattr(sh, "has_text_frame")
                    and sh.has_text_frame
                    and sh.text_frame
                    and sh.text_frame.text
                ):
                    raw = sh.text_frame.text or ""
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


# =========================
# Concatenación y límites
# =========================

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


# =========================
# LanguageTool local
# =========================

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


# =========================
# Capa de modismos argentinos
# =========================

def detect_modismos_in_pages(
    file_name: str,
    pages: List[Tuple[int, str]],
    unit_label: str,
    patterns: List[ModismoPattern],
    skip_pages: set | None = None,
) -> pd.DataFrame:
    """
    Recorre cada página/diapositiva y detecta modismos argentinos usando los patrones.
    Devuelve un DataFrame con las mismas columnas que LanguageTool.
    """
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
                # contexto alrededor del match
                ctx_start = max(0, start - 60)
                ctx_end = min(len(text), end + 60)
                contexto = text[ctx_start:ctx_end]

                # filtro extra: no marcar si el contexto parece bibliografía
                if is_reference_fragment(contexto):
                    continue

                # el texto coincidente con la misma posición en el original
                match_text = text[start:end]

                # --- NUEVO FILTRO: modismos con tilde ---
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


# =========================
# Análisis de archivo completo
# =========================

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

    # Para TXT no excluimos páginas completas como bibliografía;
    # solo se filtran fragmentos con is_reference_fragment.
    skip_pages = detect_bibliography_pages(pages) if excluir_bibliografia else set()
    if excluir_bibliografia and ext == ".txt":
        skip_pages = set()

    st.session_state["_lang_code"] = lang_code
    _, starts, bounds = build_global_text(pages)
    ranges = chunk_by_pages(pages, max_chars_call)
    tool = get_language_tool(lang_code)

    rows: List[Dict[str, Any]] = []
    lock = threading.Lock()
    prog = st.progress(0.0, text=f"{file_name}: 0/{len(ranges)} grupos")
    done = 0

    def worker(rng: Tuple[int, int]):
        nonlocal done
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
                done += 1
                prog.progress(done / len(ranges), text=f"{file_name}: {done}/{len(ranges)} grupos")

    prog.progress(1.0, text=f"{file_name}: completado")

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

    # --- Modismos argentinos (capa extra) ---
    if analizar_modismos and lang_code.startswith("es") and modismos_patterns:
        # Para TXT no pasamos páginas a excluir (ya filtramos por fragmento)
        skip_for_modismos = skip_pages if (excluir_bibliografia and ext != ".txt") else None
        df_mod = detect_modismos_in_pages(
            file_name=file_name,
            pages=pages,
            unit_label=unit_label,
            patterns=modismos_patterns,
            skip_pages=skip_for_modismos,
        )
    else:
        df_mod = pd.DataFrame([])

    # Unificar resultados
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


# =========================
# Exportes
# =========================

def to_excel_bytes(resultados_df: pd.DataFrame, resumen_completo_df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:

        if resultados_df is None or resultados_df.empty:
            tmp = pd.DataFrame(columns=[
                "Archivo", "Página/Diapositiva", "BloqueTipo", "Mensaje",
                "Sugerencias", "Oración", "Contexto", "Regla", "Categoría"
            ])
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

        if resultados_df is None or resultados_df.empty:
            resumen_inc = pd.DataFrame(columns=["Archivo", "TotalIncidencias"])
        else:
            resumen_inc = (
                resultados_df.groupby("Archivo")
                .size()
                .reset_index(name="TotalIncidencias")
                .sort_values("TotalIncidencias", ascending=False)
            )
        resumen_inc.to_excel(w, index=False, sheet_name="ResumenIncidencias")
        ws2 = w.sheets["ResumenIncidencias"]
        ws2.set_column(0, 0, 40)
        ws2.set_column(1, 1, 22)

        resumen_completo_df.to_excel(w, index=False, sheet_name="ResumenCompleto")
        ws3 = w.sheets["ResumenCompleto"]
        for i, col in enumerate(resumen_completo_df.columns):
            ws3.set_column(i, i, 28 if i > 0 else 50)

    out.seek(0)
    return out.read()


# =========================
# UI — solo LOCAL
# =========================

def main():
    st.set_page_config(page_title="UTP GrammarScan — Local", page_icon="📂", layout="wide")
    st.title("📂 UTP GrammarScan — Ortografía y Gramática (PDF, DOCX, PPTX, TXT)")

    # <<< NUEVO: se agrupan los parámetros en un expander >>>
    with st.expander("Parámetros", expanded=False):
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            lang_code = st.selectbox("Idioma", ["es", "en-US", "pt-BR", "fr", "de"], index=0)
        with c2:
            max_chars_call = st.number_input(
                "Máx. caracteres por llamada (LOCAL)",
                5000, 80000, 30000,
                help="Se agrupan páginas/diapos hasta este límite para mantener contexto."
            )
        with c3:
            workers = st.slider(
                "Trabajadores (hilos)",
                1, max(2, os.cpu_count() or 4), min(4, (os.cpu_count() or 4)),
                help="Paraleliza el troceo por páginas. Las llamadas a LT se serializan para estabilidad."
            )
    # <<< FIN CAMBIO >>>

    excluir_biblio = st.checkbox(
        "Excluir secciones/entradas de bibliografía (APA, MLA, IEEE, Vancouver)",
        value=True
    )

    analizar_modismos = st.checkbox(
        "Detectar modismos argentinos (modismos_ar.xlsx)",
        value=True if lang_code.startswith("es") else False
    )

    with st.expander("Estado del motor"):
        st.write(f"Java detectado: **{find_java()}**")
        st.write("Backend LanguageTool: **local** (una sola instancia por sesión).")
        if not find_java():
            st.error("No se puede continuar sin Java.")
            st.stop()

    ups = st.file_uploader(
        "Sube uno o varios archivos (.pdf, .docx, .pptx, .txt)",
        type=["pdf", "docx", "pptx", "txt"],
        accept_multiple_files=True
    )

    if st.button("Procesar documentos", use_container_width=True):
        if not ups:
            st.warning("Sube al menos un archivo.")
            st.stop()

        try:
            _ = get_language_tool(lang_code)
        except Exception as e:
            st.error(f"No se pudo iniciar LanguageTool local: {e}")
            st.stop()

        # Carga de modismos (si aplica)
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

        total_seleccionados = len(ups)
        all_dfs: List[pd.DataFrame] = []
        resumen_rows: List[Dict[str, Any]] = []

        overall = st.progress(0.0, text="Preparando…")
        t0 = time.time()

        for i, up in enumerate(ups, start=1):
            overall.progress(
                (i - 1) / total_seleccionados,
                text=f"Procesando {up.name} ({i}/{total_seleccionados})"
            )

            try:
                data = up.read()
                ext = os.path.splitext(up.name)[1].lower()

                df = analyze_file(
                    up.name,
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
                        "Archivo": up.name,
                        "Extension": ext,
                        "Estado": "Con incidencias",
                        "TotalIncidencias": int(df.shape[0]),
                        "Detalle": ""
                    })
                else:
                    resumen_rows.append({
                        "Archivo": up.name,
                        "Extension": ext,
                        "Estado": "Sin incidencias o sin texto",
                        "TotalIncidencias": 0,
                        "Detalle": ""
                    })

            except Exception as e:
                resumen_rows.append({
                    "Archivo": up.name,
                    "Extension": os.path.splitext(up.name)[1].lower(),
                    "Estado": "Error",
                    "TotalIncidencias": None,
                    "Detalle": safe_str(e)
                })
                st.error(f"Error procesando {up.name}: {e}")

        overall.progress(1.0, text="Análisis finalizado")

        resumen_completo_df = pd.DataFrame(resumen_rows)

        n_inc = int(
            resumen_completo_df.query("Estado == 'Con incidencias'")["Archivo"].nunique()
        )
        n_zero = int(
            resumen_completo_df.query(
                "Estado == 'Sin incidencias o sin texto'"
            )["Archivo"].nunique()
        )
        n_err = int(
            resumen_completo_df.query("Estado == 'Error'")["Archivo"].nunique()
        )

        cA, cB, cC, cD = st.columns(4)
        with cA:
            st.metric("Seleccionados", total_seleccionados)
        with cB:
            st.metric("Con incidencias", n_inc)
        with cC:
            st.metric("Sin incidencias / sin texto", n_zero)
        with cD:
            st.metric("Errores", n_err)

        if any(len(df) for df in all_dfs):
            final_df = pd.concat(all_dfs, ignore_index=True)
            st.subheader("📑 Resultados (detalle de incidencias)")
            st.dataframe(final_df, use_container_width=True, hide_index=True)

            st.markdown("**Resumen por archivo (sólo con incidencias)**")
            resumen_inc = (
                final_df.groupby("Archivo")
                .size()
                .reset_index(name="TotalIncidencias")
                .sort_values("TotalIncidencias", ascending=False)
            )
            st.dataframe(resumen_inc, use_container_width=True, hide_index=True)
        else:
            final_df = pd.DataFrame([])
            st.info("No se encontraron incidencias en los archivos procesados.")

        st.markdown("**Resumen completo de archivos** (incluye 0 incidencias y errores)")
        st.dataframe(resumen_completo_df, use_container_width=True, hide_index=True)

        cD1, cD2 = st.columns(2)
        with cD1:
            st.download_button(
                "⬇️ Excel (Resultados + Resúmenes)",
                data=to_excel_bytes(final_df, resumen_completo_df),
                file_name="UTP_GrammarScan_Resultados.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml."
                    "sheet"
                ),
                use_container_width=True
            )
        with cD2:
            st.download_button(
                "⬇️ CSV (Resultados — sólo incidencias)",
                data=final_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                file_name="UTP_GrammarScan_Resultados.csv",
                mime="text/csv",
                use_container_width=True
            )

        st.caption(f"⏱️ Tiempo total: {time.time() - t0:0.2f}s")


if __name__ == "__main__":
    main()


















