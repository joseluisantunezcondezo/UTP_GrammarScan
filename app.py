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


# =========================
# Detección de bibliografía (mejorada y con editoriales)
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
JOURNAL_HINT_RE = re.compile(r"\b(vol\.?|no\.?|nº|n\.\s?o\.?|pp\.?|ed\.|edición|issn|isbn|issue|pages)\b", re.IGNORECASE)
PUBLISHER_HINT_RE = re.compile(
    r"\b(pearson|mcgraw[- ]?hill|elsevier|springer|wiley|cengage|prentice\s*hall|sage|oxford|cambridge|harvard\s*press|"
    r"editorial(?:es)?|ediciones?|universidad(?:\s+de)?\s+\w+)\b",
    re.IGNORECASE
)
ETAL_RE = re.compile(r"\bet\s+al\.?\b", re.IGNORECASE)
BULLET_PREFIX_RE = re.compile(r"^\s*(?:[\u2022\u2023\u25E6\u2043\u2219\u25CF\u25AA\u25AB\u25A0\u25A1]|[-–—·•▪●◦‣])\s*")

def _strip_bullet(line: str) -> str:
    return BULLET_PREFIX_RE.sub("", line).strip()

# Patrones de líneas de referencia (APA, MLA, IEEE, Vancouver y variantes comunes en PPT)
BIB_RE_LIST = [
    re.compile(
        r"^[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+,\s(?:[A-Z]\.\s?){1,4}(?:,\s(?:[A-Z]\.\s?)){0,3}"
        r"(?:\s(?:&|y|and)\s[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+,\s(?:[A-Z]\.\s?){1,4})*"
        r"\s\(\d{4}[a-z]?\)\.\s", re.UNICODE
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
    re.compile(r"^\d+\.\s[A-ZÁÉÍÓÚÑ][\w'\-]+.*\b(19|20)\d{2}\b.*\d+(?:\(\d+\))?:\d+(-\d+)?\.?$"),
    re.compile(r"^[A-ZÁÉÍÓÚÑ][\w'\-]+.*\.\s+\S+\.\s+\(?\b(19|20)\d{2}\)?\.?$")
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
    if YEAR_RE.search(s) and (DOI_URL_HINT_RE.search(s) or JOURNAL_HINT_RE.search(s) or PUBLISHER_HINT_RE.search(s)):
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
    if YEAR_RE.search(flat) and (DOI_URL_HINT_RE.search(flat) or JOURNAL_HINT_RE.search(flat) or PUBLISHER_HINT_RE.search(flat)):
        if flat.count(",") >= 2 or ETAL_RE.search(flat) or URL_RE.search(flat):
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
    """Divide un DOCX en 'páginas' aproximadas usando saltos de página y tablas."""
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
        else:  # Table
            for row in b.rows:
                row_text = " | ".join(
                    (c.text or "").strip() for c in row.cells if (c.text or "").strip()
                )
                if row_text:
                    buff.append(row_text)

    flush()

    if not pages:  # fallback
        flat = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        if flat.strip():
            pages = [(1, normalize_ws(flat))]

    return pages


def _clean_reference_lines_block(text: str) -> str:
    """Elimina renglones que parezcan citas dentro de un bloque de texto."""
    kept = []
    for ln in (text or "").splitlines():
        s = normalize_ws(ln)
        if not s:
            continue
        if is_bibliography_heading(s) or is_reference_line(s) or (PUBLISHER_HINT_RE.search(s) and YEAR_RE.search(s)):
            continue
        kept.append(s)
    return "\n".join(kept).strip()


def _read_pptx_via_zip(bio: BytesIO) -> List[Tuple[int, str]]:
    """
    Fallback tolerante a corrupción:
    lee directamente los XML de 'ppt/slides/slideN.xml' y extrae <a:t>.
    Evita tocar 'ppt/fonts/*' (donde suele estar la corrupción).
    """
    slides: List[Tuple[int, str]] = []
    bio.seek(0)
    with zipfile.ZipFile(bio) as z:
        names = [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
        # Ordenar por número de slide
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
                # Si una slide puntual está dañada, la omitimos y seguimos
                continue
    return slides


def read_pptx_slides(bio: BytesIO) -> List[Tuple[int, str]]:
    """
    Lector PPTX robusto:
    1) intenta con python-pptx
    2) si falla (ej. BadZipFile/CRC), cae a lectura directa del ZIP
    """
    # --- intento 1: python-pptx ---
    try:
        prs = Presentation(bio)
        slides: List[Tuple[int, str]] = []
        slide_h = float(prs.slide_height) if hasattr(prs, "slide_height") else None

        for s_idx, slide in enumerate(prs.slides, start=1):
            chunk: List[str] = []
            for sh in slide.shapes:
                if hasattr(sh, "has_text_frame") and sh.has_text_frame and sh.text_frame and sh.text_frame.text:
                    raw = sh.text_frame.text or ""
                    raw_norm = normalize_ws(raw)

                    # 1) Si el shape está en el 25% inferior y parece referencia -> descartar shape completo
                    try:
                        if slide_h and float(getattr(sh, "top", 0)) >= 0.75 * slide_h:
                            if is_reference_fragment(raw_norm) or URL_RE.search(raw_norm) or (PUBLISHER_HINT_RE.search(raw_norm) and YEAR_RE.search(raw_norm)):
                                continue
                    except Exception:
                        pass

                    # 2) Si no, limpiar renglones de referencia dentro del shape
                    cleaned = _clean_reference_lines_block(raw_norm)
                    if cleaned:
                        chunk.append(cleaned)

            txt = normalize_ws("\n".join(chunk))
            if txt:
                slides.append((s_idx, txt))
        return slides

    except Exception:
        # --- intento 2: fallback vía ZIP (tolerante a CRC roto) ---
        try:
            return _read_pptx_via_zip(bio)
        except Exception as e2:
            # Si también falla, propagamos para que lo capture la capa superior por archivo.
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

def build_global_text(pages: List[Tuple[int, str]]) -> Tuple[str, List[int], List[Tuple[int, int, int]]]:
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
# LanguageTool local (única instancia)
# =========================

@st.cache_resource(show_spinner=False)
def get_language_tool(lang_code: str):
    if not find_java():
        raise RuntimeError("Java no detectado. Activa tu JRE/JDK para usar LanguageTool local.")
    import language_tool_python as lt
    return lt.LanguageTool(lang_code)


LT_LOCK = threading.Lock()  # serializa check()


# =========================
# Análisis con reintentos
# =========================

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


def analyze_file(
    file_name: str,
    file_bytes: bytes,
    lang_code: str,
    max_chars_call: int,
    workers: int,
    excluir_bibliografia: bool = True
) -> pd.DataFrame:
    pages, unit_label = extract_pages(file_bytes, file_name)
    if not pages:
        return pd.DataFrame([])

    skip_pages = detect_bibliography_pages(pages) if excluir_bibliografia else set()

    st.session_state["_lang_code"] = lang_code
    _, starts, bounds = build_global_text(pages)
    ranges = chunk_by_pages(pages, max_chars_call)
    tool = get_language_tool(lang_code)

    rows: List[Dict[str, Any]] = []
    lock = threading.Lock()
    prog = st.progress(0.0, text=f"{file_name}: 0/{len(ranges)} grupos")
    done = 0

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

            if excluir_bibliografia:
                if page_no in skip_pages:
                    continue
                if is_reference_fragment(sentence) or is_reference_fragment(context):
                    continue

            local_rows.append({
                "Archivo": file_name,
                "Página/Diapositiva": page_no,
                "BloqueTipo": unit_label,
                "Mensaje": safe_str(getattr(m, "message", "")),
                "Sugerencias": ", ".join(getattr(m, "replacements", [])[:5])
                               if isinstance(getattr(m, "replacements", []), list) else "",
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
        return pd.DataFrame([])

    df = pd.DataFrame.from_records(rows)

    # Post-filtro extra por si algo quedó
    if excluir_bibliografia and not df.empty:
        mask_ref = (
            df["Oración"].fillna("").apply(is_reference_fragment) |
            df["Contexto"].fillna("").apply(is_reference_fragment)
        )
        df = df.loc[~mask_ref].copy()

    df.sort_values(["Archivo", "Página/Diapositiva"], inplace=True)
    df.drop_duplicates(
        subset=["Archivo", "Página/Diapositiva", "Mensaje", "Oración", "Contexto", "Regla"],
        keep="first",
        inplace=True,
    )
    df.reset_index(drop=True, inplace=True)
    return df


# =========================
# Exportes
# =========================

def to_excel_bytes(resultados_df: pd.DataFrame, resumen_completo_df: pd.DataFrame) -> bytes:
    """
    Genera un Excel con:
      - Resultados          (detalle de incidencias)
      - ResumenIncidencias  (sólo arch. con incidencias)
      - ResumenCompleto     (todos los archivos procesados, incl. 0 y errores)
    """
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:

        # --- Resultados (detalle) ---
        if resultados_df is None or resultados_df.empty:
            tmp = pd.DataFrame(columns=[
                "Archivo","Página/Diapositiva","BloqueTipo","Mensaje",
                "Sugerencias","Oración","Contexto","Regla","Categoría"
            ])
            tmp.to_excel(w, index=False, sheet_name="Resultados")
            ws = w.sheets["Resultados"]
            ws.set_column(0, 0, 40)
        else:
            resultados_df.to_excel(w, index=False, sheet_name="Resultados")
            ws = w.sheets["Resultados"]
            for i, col in enumerate(resultados_df.columns):
                try:
                    width = min(60, max(12, int(resultados_df[col].astype(str).str.len().quantile(0.9)) + 2))
                except Exception:
                    width = 22
                ws.set_column(i, i, width)

        # --- Resumen incidencias (como antes) ---
        if resultados_df is None or resultados_df.empty:
            resumen_inc = pd.DataFrame(columns=["Archivo","TotalIncidencias"])
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

        # --- Resumen completo (incluye 0 y errores) ---
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

    excluir_biblio = st.checkbox(
        "Excluir secciones/entradas de bibliografía (APA, MLA, IEEE, Vancouver)",
        value=True
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
            _ = get_language_tool(lang_code)  # warm-up
        except Exception as e:
            st.error(f"No se pudo iniciar LanguageTool local: {e}")
            st.stop()

        total_seleccionados = len(ups)
        all_dfs: List[pd.DataFrame] = []
        resumen_rows: List[Dict[str, Any]] = []

        overall = st.progress(0.0, text="Preparando…")
        t0 = time.time()

        for i, up in enumerate(ups, start=1):
            overall.progress((i - 1) / total_seleccionados, text=f"Procesando {up.name} ({i}/{total_seleccionados})")

            try:
                data = up.read()
                ext = os.path.splitext(up.name)[1].lower()

                df = analyze_file(up.name, data, lang_code, max_chars_call, workers, excluir_bibliografia=excluir_biblio)

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
                # Captura cualquier error (p. ej., PPTX con CRC roto + fallback fallido) y continúa
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

        # Métricas
        n_inc = int(resumen_completo_df.query("Estado == 'Con incidencias'")["Archivo"].nunique())
        n_zero = int(resumen_completo_df.query("Estado == 'Sin incidencias o sin texto'")["Archivo"].nunique())
        n_err = int(resumen_completo_df.query("Estado == 'Error'")["Archivo"].nunique())

        cA, cB, cC, cD = st.columns(4)
        with cA: st.metric("Seleccionados", total_seleccionados)
        with cB: st.metric("Con incidencias", n_inc)
        with cC: st.metric("Sin incidencias / sin texto", n_zero)
        with cD: st.metric("Errores", n_err)

        # Resultados detallados
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

        # Descargas
        cD1, cD2 = st.columns(2)
        with cD1:
            st.download_button(
                "⬇️ Excel (Resultados + Resúmenes)",
                data=to_excel_bytes(final_df, resumen_completo_df),
                file_name="UTP_GrammarScan_Resultados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
















