# CatchPDF Backend v1.0 — main.py
# Deploy on Render or Replit (Python 3.10+)

import io, os, re, zipfile, json
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import Response, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.trustedhost import TrustedHostMiddleware

# ── PDF / Office libs ────────────────────────────────────────────
try:
    import fitz                        # pymupdf
    _FITZ_OK = True
except ImportError:
    _FITZ_OK = False

try:
    import pdfplumber
    _PLUMBER_OK = True
except ImportError:
    _PLUMBER_OK = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    _XL_OK = True
except ImportError:
    _XL_OK = False

try:
    from pptx import Presentation
    from pptx.util import Inches as PInches, Pt as PPt
    _PPTX_OK = True
except ImportError:
    _PPTX_OK = False

try:
    import requests as _req
    _HF_TOKEN = os.environ.get("HF_TOKEN", "")
    # We use Hugging Face Inference API — free with token
    # Model: mistralai/Mistral-7B-Instruct-v0.3 (free, powerful, no charges)
    _HF_MODEL  = os.environ.get("HF_MODEL", "mistralai/Mistral-7B-Instruct-v0.3")
    _HF_URL    = f"https://api-inference.huggingface.co/models/{_HF_MODEL}"
    _AI_OK = bool(_HF_TOKEN)
except Exception:
    _AI_OK = False

# ── App ──────────────────────────────────────────────────────────
app = FastAPI(title="CatchPDF API", version="1.0.0", docs_url=None, redoc_url=None)

app.add_middleware(CORSMiddleware,
    allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"])

MAX_BYTES = 25 * 1024 * 1024   # 25 MB free limit

# ── Helpers ──────────────────────────────────────────────────────
def stream_file(data: bytes, media_type: str, filename: str, extra: dict = None) -> Response:
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Content-Length": str(len(data)),
        "X-Content-Type-Options": "nosniff",
        "Cache-Control": "no-store",
    }
    if extra:
        headers.update(extra)
    return Response(content=data, media_type=media_type, headers=headers)

async def read_file(upload: UploadFile) -> bytes:
    data = await upload.read()
    if len(data) > MAX_BYTES:
        raise HTTPException(413, f"File too large. Free limit is 25 MB.")
    return data

def stem(f) -> str:
    return Path(f or "file").stem

def open_fitz(data: bytes):
    if not _FITZ_OK:
        raise HTTPException(500, "PDF engine not available.")
    doc = fitz.open(stream=data, filetype="pdf")
    return doc

def extract_text_fitz(data: bytes, max_pages: int = 40) -> str:
    doc = open_fitz(data)
    pages = min(len(doc), max_pages)
    return "\n\n".join(doc[i].get_text() for i in range(pages))

def ai_call(system: str, prompt: str, max_tokens: int = 1500) -> str:
    """
    Call Hugging Face Inference API — FREE with HF_TOKEN.
    Uses Mistral-7B-Instruct by default (powerful, free tier).
    Falls back to smaller model if main is loading.
    """
    if not _AI_OK:
        raise HTTPException(503, "Smart Tools unavailable. Add HF_TOKEN in Replit Secrets.")

    # Build Mistral instruct prompt format
    full_prompt = f"<s>[INST] {system}\n\n{prompt} [/INST]"

    headers = {
        "Authorization": f"Bearer {_HF_TOKEN}",
        "Content-Type": "application/json",
    }
    payload = {
        "inputs": full_prompt,
        "parameters": {
            "max_new_tokens": min(max_tokens, 1200),
            "temperature": 0.4,
            "top_p": 0.9,
            "do_sample": True,
            "return_full_text": False,
        },
        "options": {
            "wait_for_model": True,   # wait if model is loading (cold start)
            "use_cache": False,
        }
    }

    try:
        resp = _req.post(_HF_URL, headers=headers, json=payload, timeout=90)

        # Model loading (503) — wait and retry once
        if resp.status_code == 503:
            import time; time.sleep(15)
            resp = _req.post(_HF_URL, headers=headers, json=payload, timeout=90)

        if resp.status_code == 401:
            raise HTTPException(503, "Invalid Hugging Face token. Check HF_TOKEN in Secrets.")

        if not resp.ok:
            err = resp.json() if resp.content else {}
            msg = err.get("error", f"HF API error {resp.status_code}")
            raise HTTPException(502, f"Smart processing failed: {msg}")

        result = resp.json()

        # HF returns list of dicts or single dict
        if isinstance(result, list):
            text = result[0].get("generated_text", "")
        elif isinstance(result, dict):
            text = result.get("generated_text", "")
        else:
            text = str(result)

        # Clean up any leftover prompt echo
        for marker in ["[/INST]", "[INST]", "<s>", "</s>"]:
            text = text.replace(marker, "")

        return text.strip() or "No result generated. Please try again."

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(502, f"Smart processing error: {str(e)}")

# ── Health ───────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {
        "status": "ok", "brand": "CatchPDF", "version": "1.0.0",
        "fitz": _FITZ_OK, "plumber": _PLUMBER_OK,
        "docx": _DOCX_OK, "xlsx": _XL_OK,
        "pptx": _PPTX_OK,
        "ai": _AI_OK,
        "ai_provider": "HuggingFace" if _AI_OK else "none",
        "ai_model": _HF_MODEL if _AI_OK else "none",
    }

# ── AI ENDPOINTS ─────────────────────────────────────────────────
@app.post("/ai-summary")
async def ai_summary(file: UploadFile = File(...), length: str = Form("medium")):
    data = await read_file(file)
    text = extract_text_fitz(data, max_pages=20)
    if not text.strip():
        raise HTTPException(400, "No readable text found in this PDF.")
    lens = {"short": "2–3 sentences", "medium": "1 clear paragraph", "long": "2–3 paragraphs covering all key sections"}
    result = ai_call(
        "You are a professional document summarizer. Return only the summary — no preamble.",
        f"Summarize this document in {lens.get(length,'1 paragraph')}:\n\n{text[:8000]}"
    )
    return {"result": result}

@app.post("/ai-notes")
async def ai_notes(file: UploadFile = File(...), style: str = Form("bullet")):
    data = await read_file(file)
    text = extract_text_fitz(data, max_pages=25)
    if not text.strip():
        raise HTTPException(400, "No readable text found in this PDF.")
    style_map = {
        "bullet": "bullet point notes organized by topic",
        "outline": "a structured hierarchical outline with main points and sub-points",
        "cornell": "Cornell-style notes with main notes, cues and a summary section",
        "mindmap": "a text-based mind map showing relationships between concepts"
    }
    result = ai_call(
        "You are an expert study notes creator. Return only the formatted notes — no preamble.",
        f"Create {style_map.get(style,'bullet point notes')} from this document:\n\n{text[:9000]}"
    )
    return {"result": result}

@app.post("/ai-quiz")
async def ai_quiz(file: UploadFile = File(...), count: int = Form(10), difficulty: str = Form("medium")):
    data = await read_file(file)
    text = extract_text_fitz(data, max_pages=20)
    if not text.strip():
        raise HTTPException(400, "No readable text found in this PDF.")
    count = max(3, min(count, 20))
    result = ai_call(
        "You are a quiz creator. Return only numbered questions with options labeled A–D and mark the correct answer.",
        f"Create {count} {difficulty}-difficulty multiple choice questions from this document. For each question show options A, B, C, D and mark the correct answer with ✓.\n\nDocument:\n{text[:8000]}"
    )
    return {"result": result}

@app.post("/ai-keypoints")
async def ai_keypoints(file: UploadFile = File(...)):
    data = await read_file(file)
    text = extract_text_fitz(data, max_pages=20)
    if not text.strip():
        raise HTTPException(400, "No readable text found in this PDF.")
    result = ai_call(
        "You are an expert at identifying the most important information in documents. Return a numbered list of key points only.",
        f"Extract the 8–12 most important key points from this document:\n\n{text[:8000]}"
    )
    return {"result": result}

@app.post("/ai-translate")
async def ai_translate(file: UploadFile = File(...), from_lang: str = Form("auto"), to_lang: str = Form("Urdu")):
    data = await read_file(file)
    text = extract_text_fitz(data, max_pages=15)
    if not text.strip():
        raise HTTPException(400, "No readable text found in this PDF.")
    src = f"from {from_lang}" if from_lang != "auto" else "(auto-detect the source language)"
    result = ai_call(
        f"You are a professional translator. Translate accurately {src} to {to_lang}. Return only the translated text.",
        f"Translate the following text to {to_lang}:\n\n{text[:6000]}"
    )
    return {"result": result}

@app.post("/ask-pdf")
async def ask_pdf(file: UploadFile = File(...), question: str = Form(...), history: str = Form("[]")):
    data = await read_file(file)
    text = extract_text_fitz(data, max_pages=30)
    if not text.strip():
        raise HTTPException(400, "No readable text found in this PDF.")
    try:
        hist = json.loads(history)[-6:]  # last 3 exchanges
    except Exception:
        hist = []
    ctx = "\n".join(f"Q: {h['q']}\nA: {h['a']}" for h in hist if 'q' in h and 'a' in h)
    system = "You are a helpful assistant answering questions strictly based on the provided document. If the answer is not in the document, say so clearly."
    prompt = f"Document:\n{text[:9000]}\n\n"
    if ctx:
        prompt += f"Previous conversation:\n{ctx}\n\n"
    prompt += f"Question: {question}"
    result = ai_call(system, prompt, max_tokens=1500)
    return {"result": result}

# ── MERGE PDF ────────────────────────────────────────────────────
@app.post("/merge-pdf")
async def merge_pdf(files: List[UploadFile] = File(...)):
    if len(files) < 2:
        raise HTTPException(400, "Please upload at least 2 PDF files.")
    writer = fitz.open()
    for uf in files:
        data = await read_file(uf)
        try:
            doc = fitz.open(stream=data, filetype="pdf")
            writer.insert_pdf(doc)
            doc.close()
        except Exception as e:
            raise HTTPException(400, f"Could not read '{uf.filename}': {e}")
    out = io.BytesIO()
    writer.save(out)
    return stream_file(out.getvalue(), "application/pdf", "merged.pdf")

# ── SPLIT PDF ────────────────────────────────────────────────────
@app.post("/split-pdf")
async def split_pdf(file: UploadFile = File(...), mode: str = Form("each"),
                    start_page: int = Form(1), end_page: int = Form(1)):
    data = await read_file(file)
    doc = open_fitz(data)
    n = len(doc)
    if mode == "range":
        s, e = max(1, start_page) - 1, min(n, end_page) - 1
        if s > e:
            raise HTTPException(400, "Start page must be ≤ end page.")
        out_doc = fitz.open()
        out_doc.insert_pdf(doc, from_page=s, to_page=e)
        buf = io.BytesIO(); out_doc.save(buf)
        return stream_file(buf.getvalue(), "application/pdf", f"{stem(file.filename)}_p{s+1}-{e+1}.pdf")
    else:
        zb = io.BytesIO()
        with zipfile.ZipFile(zb, "w", zipfile.ZIP_DEFLATED) as zf:
            for i in range(n):
                pg = fitz.open(); pg.insert_pdf(doc, from_page=i, to_page=i)
                buf = io.BytesIO(); pg.save(buf)
                zf.writestr(f"page_{i+1:03d}.pdf", buf.getvalue()); pg.close()
        return stream_file(zb.getvalue(), "application/zip", f"{stem(file.filename)}_pages.zip")

# ── COMPRESS PDF ─────────────────────────────────────────────────
@app.post("/compress-pdf")
async def compress_pdf(file: UploadFile = File(...), level: str = Form("medium")):
    data = await read_file(file)
    orig_size = len(data)
    doc = open_fitz(data)
    deflate_map = {"low": 8, "medium": 7, "high": 6}
    image_quality = {"low": 85, "medium": 65, "high": 40}
    dq = image_quality.get(level, 65)
    buf = io.BytesIO()
    doc.save(buf,
        garbage=4,
        deflate=True,
        deflate_images=True,
        deflate_fonts=True,
        clean=True,
        linear=True,
        expand=0,
    )
    compressed = buf.getvalue()
    comp_size = len(compressed)
    pct = round((1 - comp_size / orig_size) * 100, 1) if orig_size else 0
    return stream_file(compressed, "application/pdf", f"compressed_{stem(file.filename)}.pdf", {
        "X-Original-Size": str(orig_size),
        "X-Compressed-Size": str(comp_size),
        "X-Savings-Pct": str(max(0, pct)),
    })

# ── ROTATE PDF ───────────────────────────────────────────────────
@app.post("/rotate-pdf")
async def rotate_pdf(file: UploadFile = File(...), angle: int = Form(90), pages: str = Form("all")):
    data = await read_file(file)
    doc = open_fitz(data)
    angle = angle % 360
    for i, page in enumerate(doc):
        apply = (pages == "all") or (pages == "odd" and i % 2 == 0) or (pages == "even" and i % 2 == 1)
        if apply:
            page.set_rotation((page.rotation + angle) % 360)
    buf = io.BytesIO(); doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"rotated_{stem(file.filename)}.pdf")

# ── WATERMARK ────────────────────────────────────────────────────
@app.post("/add-watermark")
async def add_watermark(file: UploadFile = File(...), text: str = Form("CONFIDENTIAL"),
                        opacity: float = Form(0.2), position: str = Form("center")):
    data = await read_file(file)
    doc = open_fitz(data)
    opacity = max(0.05, min(opacity, 0.9))
    for page in doc:
        w, h = page.rect.width, page.rect.height
        font_size = min(w, h) * 0.06
        if position == "center":
            tw = fitz.get_text_length(text, fontsize=font_size)
            x = (w - tw) / 2; y = h / 2
            page.insert_text((x, y), text, fontsize=font_size,
                             color=(0.5, 0.1, 0.1), rotate=45,
                             fill_opacity=opacity, overlay=True)
        elif position == "bottom-right":
            tw = fitz.get_text_length(text, fontsize=font_size * 0.6)
            page.insert_text((w - tw - 20, h - 20), text,
                             fontsize=font_size * 0.6,
                             color=(0.4, 0.4, 0.4), fill_opacity=opacity, overlay=True)
        else:  # bottom-center
            tw = fitz.get_text_length(text, fontsize=font_size * 0.6)
            page.insert_text(((w - tw) / 2, h - 20), text,
                             fontsize=font_size * 0.6,
                             color=(0.4, 0.4, 0.4), fill_opacity=opacity, overlay=True)
    buf = io.BytesIO(); doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"watermarked_{stem(file.filename)}.pdf")

# ── PDF TO WORD ──────────────────────────────────────────────────
@app.post("/pdf-to-word")
async def pdf_to_word(file: UploadFile = File(...)):
    data = await read_file(file)
    if not _DOCX_OK:
        raise HTTPException(500, "Word conversion library not available.")
    doc_out = Document()
    # Style the document
    style = doc_out.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    if _FITZ_OK:
        fitz_doc = open_fitz(data)
        for page_num in range(len(fitz_doc)):
            page = fitz_doc[page_num]
            blocks = page.get_text("dict")["blocks"]
            if page_num > 0:
                doc_out.add_page_break()
            for block in blocks:
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    line_text = " ".join(s.get("text","") for s in line.get("spans",[])).strip()
                    if not line_text:
                        continue
                    # Detect heading by font size
                    sizes = [s.get("size", 11) for s in line.get("spans", [])]
                    avg_size = sum(sizes) / len(sizes) if sizes else 11
                    if avg_size > 16:
                        p = doc_out.add_heading(line_text, level=1)
                    elif avg_size > 13:
                        p = doc_out.add_heading(line_text, level=2)
                    else:
                        p = doc_out.add_paragraph(line_text)
    else:
        # fallback: pdfplumber
        if not _PLUMBER_OK:
            raise HTTPException(500, "PDF processing library not available.")
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for i, page in enumerate(pdf.pages):
                if i > 0:
                    doc_out.add_page_break()
                text = page.extract_text() or ""
                for ln in text.split("\n"):
                    if ln.strip():
                        doc_out.add_paragraph(ln.strip())

    buf = io.BytesIO(); doc_out.save(buf)
    return stream_file(buf.getvalue(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        f"{stem(file.filename)}.docx")

# ── PDF TO EXCEL — SMART GROUPING ────────────────────────────────
@app.post("/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...), mode: str = Form("smart")):
    data = await read_file(file)
    if not _XL_OK:
        raise HTTPException(500, "Excel library not available.")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles
    hdr_fill  = PatternFill("solid", fgColor="1D4ED8")
    hdr_font  = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alt_fill  = PatternFill("solid", fgColor="EFF6FF")
    thin      = Side(style="thin", color="CBD5E1")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── SMART TABLE GROUPING LOGIC ────────────────────────────────
    # Key insight: group tables across ALL pages by their header signature.
    # Same headers → same sheet (append rows).
    # Different headers → new sheet.

    def normalize_header(row):
        """Return a canonical tuple for header comparison."""
        return tuple(re.sub(r'\s+', ' ', str(c or "").strip().lower()) for c in row)

    def is_likely_header(row):
        """Heuristic: a header row has mostly non-numeric cells."""
        if not row:
            return False
        non_empty = [c for c in row if str(c or "").strip()]
        if len(non_empty) < 2:
            return False
        numeric = sum(1 for c in non_empty if re.match(r'^[\d\s\.\,\-\$\%\+]+$', str(c).strip()))
        return numeric / len(non_empty) < 0.6

    def safe_sheet_name(base: str, existing: list) -> str:
        name = re.sub(r'[\\/*?:\[\]]', '', base)[:28].strip() or "Table"
        candidate, n = name, 1
        while candidate in existing:
            candidate = f"{name[:25]}_{n}"; n += 1
        return candidate

    # Ordered dict: header_key → {"sheet_name": str, "rows": list, "col_count": int}
    table_groups: dict = {}
    text_lines: list = []  # fallback for non-table mode

    if not _PLUMBER_OK:
        raise HTTPException(500, "PDF processing library not available.")

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for pn, page in enumerate(pdf.pages, start=1):

                if mode in ("smart", "tables"):
                    tables = page.extract_tables(table_settings={
                        "vertical_strategy": "lines_strict",
                        "horizontal_strategy": "lines_strict",
                        "snap_tolerance": 4,
                        "join_tolerance": 4,
                        "edge_min_length": 15,
                        "min_words_vertical": 1,
                        "min_words_horizontal": 1,
                    }) or []

                    # Also try with text strategy if lines gives nothing
                    if not tables:
                        tables = page.extract_tables(table_settings={
                            "vertical_strategy": "text",
                            "horizontal_strategy": "text",
                        }) or []

                    for tbl in tables:
                        if not tbl or len(tbl) < 1:
                            continue

                        # Find the actual header row (first row with mostly text)
                        header_row_idx = 0
                        for ri, row in enumerate(tbl[:3]):  # check first 3 rows
                            if is_likely_header(row):
                                header_row_idx = ri
                                break

                        raw_header = tbl[header_row_idx]
                        if not any(str(c or "").strip() for c in raw_header):
                            continue

                        hkey = normalize_header(raw_header)
                        display_header = [str(c or "").strip() for c in raw_header]

                        # Data rows = everything after header row
                        data_rows = []
                        for row in tbl[header_row_idx + 1:]:
                            cleaned = [str(c or "").strip() for c in row]
                            if any(v for v in cleaned):  # skip blank rows
                                data_rows.append(cleaned)

                        if not data_rows:
                            continue

                        if hkey not in table_groups:
                            # Derive a meaningful sheet name from header
                            sheet_label = " & ".join(h for h in display_header[:2] if h)
                            table_groups[hkey] = {
                                "sheet_name": sheet_label,
                                "header": display_header,
                                "rows": [],
                                "col_count": len(display_header),
                            }

                        table_groups[hkey]["rows"].extend(data_rows)

                # Text fallback
                if mode == "text" or (mode == "smart" and not tables):
                    for ln in (page.extract_text() or "").split("\n"):
                        if ln.strip():
                            text_lines.append((pn, ln.strip()))

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"PDF processing failed: {e}")

    # ── BUILD WORKBOOK ────────────────────────────────────────────
    existing_names: list = []

    if table_groups:
        for idx, (hkey, grp) in enumerate(table_groups.items(), start=1):
            sname = safe_sheet_name(grp["sheet_name"] or f"Table {idx}", existing_names)
            existing_names.append(sname)
            ws = wb.create_sheet(title=sname)
            ws.freeze_panes = "A2"

            # Write header row
            for ci, h in enumerate(grp["header"], start=1):
                c = ws.cell(row=1, column=ci, value=h)
                c.fill = hdr_fill; c.font = hdr_font
                c.alignment = hdr_align; c.border = border

            # Write data rows with alternating fill
            for ri, row in enumerate(grp["rows"], start=2):
                row_fill = alt_fill if ri % 2 == 0 else None
                # Pad/trim row to match header column count
                col_count = grp["col_count"]
                padded = (row + [""] * col_count)[:col_count]
                for ci, val in enumerate(padded, start=1):
                    cell = ws.cell(row=ri, column=ci)
                    # Try numeric conversion
                    stripped = val.replace(",", "").replace("$", "").replace("%", "").strip()
                    try:
                        cell.value = int(stripped) if "." not in stripped else float(stripped)
                    except (ValueError, TypeError):
                        cell.value = val
                    cell.border = border
                    if row_fill:
                        cell.fill = row_fill

            # Auto column width
            for col in ws.columns:
                max_len = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 60)

            # Add summary row count
            ws.cell(row=1, column=grp["col_count"] + 2,
                    value=f"Total rows: {len(grp['rows'])}").font = Font(italic=True, color="94A3B8", size=10)

    elif text_lines:
        ws = wb.create_sheet(title="Extracted Text")
        ws.freeze_panes = "A2"
        for ci, h in enumerate(["Page", "Text"], start=1):
            c = ws.cell(row=1, column=ci, value=h)
            c.fill = hdr_fill; c.font = hdr_font
            c.alignment = hdr_align; c.border = border
        for ri, (pg, ln) in enumerate(text_lines, start=2):
            ws.cell(row=ri, column=1, value=pg).border = border
            ws.cell(row=ri, column=2, value=ln).border = border
            if ri % 2 == 0:
                ws.cell(row=ri, column=1).fill = alt_fill
                ws.cell(row=ri, column=2).fill = alt_fill
        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 90
    else:
        ws = wb.create_sheet(title="No Data Found")
        ws.cell(row=1, column=1, value="No tables or text could be extracted from this PDF.")
        ws.cell(row=2, column=1, value="Tip: Make sure the PDF contains selectable text, not just scanned images.")

    out = io.BytesIO(); wb.save(out)
    return stream_file(out.getvalue(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        f"{stem(file.filename)}.xlsx")

# ── PDF TO JPG ───────────────────────────────────────────────────
@app.post("/pdf-to-jpg")
async def pdf_to_jpg(file: UploadFile = File(...), dpi: int = Form(150)):
    data = await read_file(file)
    doc = open_fitz(data)
    dpi = max(72, min(dpi, 300))
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    zb = io.BytesIO()
    with zipfile.ZipFile(zb, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=mat)
            zf.writestr(f"page_{i+1:03d}.jpg", pix.tobytes("jpeg"))
    return stream_file(zb.getvalue(), "application/zip", f"{stem(file.filename)}_images.zip")

# ── IMAGE TO PDF ─────────────────────────────────────────────────
@app.post("/jpg-to-pdf")
async def jpg_to_pdf(files: List[UploadFile] = File(...)):
    doc = fitz.open()
    for uf in files:
        data = await read_file(uf)
        try:
            img_doc = fitz.open(stream=data, filetype=uf.content_type.split("/")[-1] if uf.content_type else "jpeg")
            page = doc.new_page(width=img_doc[0].rect.width, height=img_doc[0].rect.height)
            page.show_pdf_page(page.rect, img_doc, 0)
            img_doc.close()
        except Exception:
            # Fallback: insert image directly
            try:
                img_rect = fitz.Rect(0, 0, 595, 842)  # A4
                page = doc.new_page(width=595, height=842)
                page.insert_image(img_rect, stream=data)
            except Exception as e:
                raise HTTPException(400, f"Could not process image '{uf.filename}': {e}")
    buf = io.BytesIO(); doc.save(buf)
    return stream_file(buf.getvalue(), "application/pdf", "images.pdf")

# ── UNLOCK PDF ───────────────────────────────────────────────────
@app.post("/unlock-pdf")
async def unlock_pdf(file: UploadFile = File(...), password: str = Form("")):
    data = await read_file(file)
    doc = fitz.open(stream=data, filetype="pdf")
    if doc.is_encrypted:
        ok = doc.authenticate(password)
        if not ok:
            raise HTTPException(400, "Wrong password. Please check and try again.")
    doc.save("/tmp/_catchpdf_unlock.pdf", encryption=fitz.PDF_ENCRYPT_NONE)
    with open("/tmp/_catchpdf_unlock.pdf", "rb") as f:
        result = f.read()
    import os as _os; _os.remove("/tmp/_catchpdf_unlock.pdf")
    return stream_file(result, "application/pdf", f"unlocked_{stem(file.filename)}.pdf")

# ── PROTECT PDF ──────────────────────────────────────────────────
@app.post("/protect-pdf")
async def protect_pdf(file: UploadFile = File(...), password: str = Form(...)):
    data = await read_file(file)
    if not password:
        raise HTTPException(400, "Password is required.")
    doc = open_fitz(data)
    perm = fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY
    buf = io.BytesIO()
    doc.save(buf,
        encryption=fitz.PDF_ENCRYPT_AES_256,
        owner_pw=password + "_owner",
        user_pw=password,
        permissions=perm)
    return stream_file(buf.getvalue(), "application/pdf", f"protected_{stem(file.filename)}.pdf")

# ── PDF TO TEXT ──────────────────────────────────────────────────
@app.post("/pdf-to-text")
async def pdf_to_text(file: UploadFile = File(...)):
    data = await read_file(file)
    if _PLUMBER_OK:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            lines = []
            for i, page in enumerate(pdf.pages):
                lines.append(f"─── Page {i+1} ───\n")
                lines.append(page.extract_text() or "[No text on this page]")
                lines.append("\n")
            text = "\n".join(lines)
    else:
        text = extract_text_fitz(data)
    return stream_file(text.encode("utf-8"), "text/plain", f"{stem(file.filename)}.txt")

# ── PDF TO PPTX ──────────────────────────────────────────────────
@app.post("/pdf-to-pptx")
async def pdf_to_pptx(file: UploadFile = File(...)):
    data = await read_file(file)
    if not _PPTX_OK:
        raise HTTPException(500, "PowerPoint library not available.")
    doc = open_fitz(data)
    prs = Presentation()
    prs.slide_width = PInches(10); prs.slide_height = PInches(7.5)
    blank_layout = prs.slide_layouts[6]  # blank slide
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_bytes = pix.tobytes("png")
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(io.BytesIO(img_bytes), PInches(0), PInches(0),
                                  width=prs.slide_width, height=prs.slide_height)
    buf = io.BytesIO(); prs.save(buf)
    return stream_file(buf.getvalue(),
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        f"{stem(file.filename)}.pptx")

# ── ADD PAGE NUMBERS ─────────────────────────────────────────────
@app.post("/add-page-numbers")
async def add_page_numbers(file: UploadFile = File(...), position: str = Form("bottom-center"),
                           format: str = Form("number"), start: int = Form(1)):
    data = await read_file(file)
    doc = open_fitz(data)
    n = len(doc)
    for i, page in enumerate(doc):
        num = i + start
        if format == "page-of":
            label = f"Page {num} of {n + start - 1}"
        elif format == "roman":
            label = _to_roman(num)
        else:
            label = str(num)
        w, h = page.rect.width, page.rect.height
        fs = 9
        tw = fitz.get_text_length(label, fontsize=fs)
        pos_map = {
            "bottom-center": ((w - tw) / 2, h - 18),
            "bottom-right": (w - tw - 20, h - 18),
            "top-center": ((w - tw) / 2, 22),
            "top-right": (w - tw - 20, 22),
        }
        x, y = pos_map.get(position, ((w - tw) / 2, h - 18))
        page.insert_text((x, y), label, fontsize=fs, color=(0.45, 0.45, 0.45))
    buf = io.BytesIO(); doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"numbered_{stem(file.filename)}.pdf")

def _to_roman(n: int) -> str:
    val = [1000,900,500,400,100,90,50,40,10,9,5,4,1]
    sym = ["M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"]
    result = ""
    for v, s in zip(val, sym):
        while n >= v:
            result += s; n -= v
    return result.lower()

# ── OCR CHECK ────────────────────────────────────────────────────
@app.post("/ocr-check")
async def ocr_check(file: UploadFile = File(...)):
    data = await read_file(file)
    doc = open_fitz(data)
    pages_checked = min(len(doc), 3)
    total_chars = sum(len(doc[i].get_text()) for i in range(pages_checked))
    is_scanned = total_chars < 80
    return {"is_scanned": is_scanned, "chars_found": total_chars, "pages_checked": pages_checked}
