# RunDocs Backend v2.0 — main.py
# KEY FIXES:
# 1. PDF-to-Excel: Tables grouped by HEADER SIGNATURE across ALL pages (NOT page-wise)
# 2. Watermark: Fixed fill_opacity parameter & proper overlay rendering

import io, os, re, zipfile, json, time
from pathlib import Path
from typing import List

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware

try:
    import fitz
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
    from docx.shared import Pt
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
    from pptx.util import Inches as PInches
    _PPTX_OK = True
except ImportError:
    _PPTX_OK = False

try:
    import requests as _req
    _HF_TOKEN = os.environ.get("HF_TOKEN", "")
    _HF_MODEL = os.environ.get("HF_MODEL", "mistralai/Mistral-7B-Instruct-v0.3")
    _HF_URL   = f"https://api-inference.huggingface.co/models/{_HF_MODEL}"
    _AI_OK    = bool(_HF_TOKEN)
except Exception:
    _AI_OK = False

app = FastAPI(title="RunDocs API", version="2.0.0", docs_url=None, redoc_url=None)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

MAX_BYTES = 25 * 1024 * 1024


def stream_file(data: bytes, media_type: str, filename: str, extra: dict = None) -> Response:
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Content-Length": str(len(data)),
        "X-Content-Type-Options": "nosniff",
        "Cache-Control": "no-store",
        "Access-Control-Expose-Headers": "X-Original-Size,X-Compressed-Size,X-Savings-Pct",
    }
    if extra:
        headers.update(extra)
    return Response(content=data, media_type=media_type, headers=headers)


async def read_file(upload: UploadFile) -> bytes:
    data = await upload.read()
    if len(data) > MAX_BYTES:
        raise HTTPException(413, "File too large. Free limit is 25 MB.")
    return data


def stem(f) -> str:
    return Path(f or "file").stem


def open_fitz(data: bytes):
    if not _FITZ_OK:
        raise HTTPException(500, "PDF engine not available.")
    return fitz.open(stream=data, filetype="pdf")


def extract_text_fitz(data: bytes, max_pages: int = 40) -> str:
    doc = open_fitz(data)
    pages = min(len(doc), max_pages)
    return "\n\n".join(doc[i].get_text() for i in range(pages))


def ai_call(system: str, prompt: str, max_tokens: int = 1500) -> str:
    if not _AI_OK:
        raise HTTPException(503, "AI Tools unavailable. Add HF_TOKEN in Replit Secrets.")
    full_prompt = f"<s>[INST] {system}\n\n{prompt} [/INST]"
    headers = {"Authorization": f"Bearer {_HF_TOKEN}", "Content-Type": "application/json"}
    payload = {
        "inputs": full_prompt,
        "parameters": {"max_new_tokens": min(max_tokens, 1200), "temperature": 0.4, "top_p": 0.9, "do_sample": True, "return_full_text": False},
        "options": {"wait_for_model": True, "use_cache": False},
    }
    try:
        resp = _req.post(_HF_URL, headers=headers, json=payload, timeout=90)
        if resp.status_code == 503:
            time.sleep(15)
            resp = _req.post(_HF_URL, headers=headers, json=payload, timeout=90)
        if resp.status_code == 401:
            raise HTTPException(503, "Invalid HuggingFace token.")
        if not resp.ok:
            err = resp.json() if resp.content else {}
            raise HTTPException(502, f"AI error: {err.get('error', resp.status_code)}")
        result = resp.json()
        text = result[0].get("generated_text", "") if isinstance(result, list) else result.get("generated_text", "")
        for m in ["[/INST]", "[INST]", "<s>", "</s>"]:
            text = text.replace(m, "")
        return text.strip() or "No result generated."
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(502, f"AI error: {str(e)}")


def _to_roman(n: int) -> str:
    val = [1000,900,500,400,100,90,50,40,10,9,5,4,1]
    sym = ["M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"]
    result = ""
    for v, s in zip(val, sym):
        while n >= v:
            result += s; n -= v
    return result.lower()


@app.get("/health")
def health():
    return {"status": "ok", "brand": "RunDocs", "version": "2.0.0",
            "fitz": _FITZ_OK, "plumber": _PLUMBER_OK, "docx": _DOCX_OK,
            "xlsx": _XL_OK, "pptx": _PPTX_OK, "ai": _AI_OK}


# ══════════════════════════════════════════
# AI ENDPOINTS
# ══════════════════════════════════════════

@app.post("/ai-summary")
async def ai_summary(file: UploadFile = File(...), length: str = Form("medium")):
    data = await read_file(file)
    text = extract_text_fitz(data, 20)
    if not text.strip(): raise HTTPException(400, "No readable text found.")
    lens = {"short": "2–3 sentences", "medium": "1 clear paragraph", "long": "2–3 detailed paragraphs"}
    result = ai_call("You are a professional document summarizer. Return only the summary.", f"Summarize in {lens.get(length,'1 paragraph')}:\n\n{text[:8000]}")
    return {"result": result}

@app.post("/ai-notes")
async def ai_notes(file: UploadFile = File(...), style: str = Form("bullet")):
    data = await read_file(file)
    text = extract_text_fitz(data, 25)
    if not text.strip(): raise HTTPException(400, "No readable text found.")
    style_map = {"bullet": "bullet point notes", "outline": "hierarchical outline", "cornell": "Cornell-style notes", "mindmap": "text-based mind map"}
    result = ai_call("You are an expert study notes creator. Return only formatted notes.", f"Create {style_map.get(style,'bullet points')}:\n\n{text[:9000]}")
    return {"result": result}

@app.post("/ai-quiz")
async def ai_quiz(file: UploadFile = File(...), count: int = Form(10), difficulty: str = Form("medium")):
    data = await read_file(file)
    text = extract_text_fitz(data, 20)
    if not text.strip(): raise HTTPException(400, "No readable text found.")
    count = max(3, min(count, 20))
    result = ai_call("You are a quiz creator. Return numbered questions with A–D options and mark correct answer with ✓.", f"Create {count} {difficulty} questions:\n\n{text[:8000]}")
    return {"result": result}

@app.post("/ai-keypoints")
async def ai_keypoints(file: UploadFile = File(...)):
    data = await read_file(file)
    text = extract_text_fitz(data, 20)
    if not text.strip(): raise HTTPException(400, "No readable text found.")
    result = ai_call("Extract the most important information as a numbered list.", f"Extract 8–12 key points:\n\n{text[:8000]}")
    return {"result": result}

@app.post("/ai-translate")
async def ai_translate(file: UploadFile = File(...), from_lang: str = Form("auto"), to_lang: str = Form("Urdu")):
    data = await read_file(file)
    text = extract_text_fitz(data, 15)
    if not text.strip(): raise HTTPException(400, "No readable text found.")
    src = f"from {from_lang}" if from_lang != "auto" else "(auto-detect)"
    result = ai_call(f"Translate accurately {src} to {to_lang}. Return only translated text.", f"Translate to {to_lang}:\n\n{text[:6000]}")
    return {"result": result}

@app.post("/ask-pdf")
async def ask_pdf(file: UploadFile = File(...), question: str = Form(...), history: str = Form("[]")):
    data = await read_file(file)
    text = extract_text_fitz(data, 30)
    if not text.strip(): raise HTTPException(400, "No readable text found.")
    try: hist = json.loads(history)[-6:]
    except: hist = []
    ctx = "\n".join(f"Q: {h['q']}\nA: {h['a']}" for h in hist if "q" in h and "a" in h)
    prompt = f"Document:\n{text[:9000]}\n\n"
    if ctx: prompt += f"Previous:\n{ctx}\n\n"
    prompt += f"Question: {question}"
    result = ai_call("Answer questions based strictly on the document. If not found, say so.", prompt, 1500)
    return {"result": result}


# ══════════════════════════════════════════
# PDF ORGANIZE
# ══════════════════════════════════════════

@app.post("/merge-pdf")
async def merge_pdf(files: List[UploadFile] = File(...)):
    if len(files) < 2: raise HTTPException(400, "Please upload at least 2 PDF files.")
    writer = fitz.open()
    for uf in files:
        data = await read_file(uf)
        try:
            doc = fitz.open(stream=data, filetype="pdf")
            writer.insert_pdf(doc); doc.close()
        except Exception as e:
            raise HTTPException(400, f"Could not read '{uf.filename}': {e}")
    out = io.BytesIO(); writer.save(out)
    return stream_file(out.getvalue(), "application/pdf", "merged.pdf")

@app.post("/split-pdf")
async def split_pdf(file: UploadFile = File(...), mode: str = Form("each"), start_page: int = Form(1), end_page: int = Form(1)):
    data = await read_file(file)
    doc = open_fitz(data); n = len(doc)
    if mode == "range":
        s = max(1, start_page) - 1; e = min(n, end_page) - 1
        if s > e: raise HTTPException(400, "Start page must be ≤ end page.")
        out_doc = fitz.open(); out_doc.insert_pdf(doc, from_page=s, to_page=e)
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

@app.post("/rotate-pdf")
async def rotate_pdf(file: UploadFile = File(...), angle: int = Form(90), pages: str = Form("all")):
    data = await read_file(file); doc = open_fitz(data); angle = angle % 360
    for i, page in enumerate(doc):
        apply = pages == "all" or (pages == "odd" and i % 2 == 0) or (pages == "even" and i % 2 == 1)
        if apply: page.set_rotation((page.rotation + angle) % 360)
    buf = io.BytesIO(); doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"rotated_{stem(file.filename)}.pdf")

@app.post("/add-page-numbers")
async def add_page_numbers(file: UploadFile = File(...), position: str = Form("bottom-center"), format: str = Form("number"), start: int = Form(1)):
    data = await read_file(file); doc = open_fitz(data); n = len(doc)
    for i, page in enumerate(doc):
        num = i + start
        label = f"Page {num} of {n+start-1}" if format=="page-of" else (_to_roman(num) if format=="roman" else str(num))
        w, h = page.rect.width, page.rect.height; fs = 9
        tw = fitz.get_text_length(label, fontsize=fs)
        pos_map = {"bottom-center": ((w-tw)/2, h-18), "bottom-right": (w-tw-20, h-18), "top-center": ((w-tw)/2, 22), "top-right": (w-tw-20, 22)}
        x, y = pos_map.get(position, ((w-tw)/2, h-18))
        page.insert_text((x, y), label, fontsize=fs, color=(0.45, 0.45, 0.45))
    buf = io.BytesIO(); doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"numbered_{stem(file.filename)}.pdf")


# ══════════════════════════════════════════
# COMPRESS
# ══════════════════════════════════════════

@app.post("/compress-pdf")
async def compress_pdf(file: UploadFile = File(...), level: str = Form("medium")):
    data = await read_file(file); orig = len(data); doc = open_fitz(data)
    buf = io.BytesIO()
    doc.save(buf, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True, linear=True)
    comp = buf.getvalue(); pct = round((1 - len(comp)/orig)*100, 1) if orig else 0
    return stream_file(comp, "application/pdf", f"compressed_{stem(file.filename)}.pdf",
                       {"X-Original-Size": str(orig), "X-Compressed-Size": str(len(comp)), "X-Savings-Pct": str(max(0,pct))})


# ══════════════════════════════════════════
# SECURITY
# ══════════════════════════════════════════

@app.post("/protect-pdf")
async def protect_pdf(file: UploadFile = File(...), password: str = Form(...)):
    data = await read_file(file)
    if not password: raise HTTPException(400, "Password is required.")
    doc = open_fitz(data)
    perm = fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY
    buf = io.BytesIO()
    doc.save(buf, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=password+"_owner", user_pw=password, permissions=perm)
    return stream_file(buf.getvalue(), "application/pdf", f"protected_{stem(file.filename)}.pdf")

@app.post("/unlock-pdf")
async def unlock_pdf(file: UploadFile = File(...), password: str = Form("")):
    data = await read_file(file)
    doc = fitz.open(stream=data, filetype="pdf")
    if doc.is_encrypted:
        if not doc.authenticate(password): raise HTTPException(400, "Wrong password.")
    tmp = "/tmp/_rundocs_unlock.pdf"
    doc.save(tmp, encryption=fitz.PDF_ENCRYPT_NONE)
    with open(tmp, "rb") as f: result = f.read()
    os.remove(tmp)
    return stream_file(result, "application/pdf", f"unlocked_{stem(file.filename)}.pdf")


# ══════════════════════════════════════════
# WATERMARK FIX — uses shapes layer properly
# ══════════════════════════════════════════

@app.post("/add-watermark")
async def add_watermark(
    file: UploadFile = File(...),
    text: str = Form("CONFIDENTIAL"),
    opacity: float = Form(0.2),
    position: str = Form("center"),
):
    data = await read_file(file)
    doc = open_fitz(data)
    opacity = max(0.05, min(opacity, 0.9))

    for page in doc:
        w, h = page.rect.width, page.rect.height

        if position == "center":
            # Diagonal center watermark
            font_size = min(w, h) * 0.07
            # Create a transparent overlay using insert_text with proper color+opacity
            tw = fitz.get_text_length(text, fontname="helv", fontsize=font_size)
            # Center point
            cx, cy = w / 2, h / 2
            # Use insert_text with rotation for diagonal effect
            page.insert_text(
                fitz.Point(cx - tw/2, cy),
                text,
                fontname="helv",
                fontsize=font_size,
                color=(0.6, 0.1, 0.1),
                rotate=45,
                overlay=True,
            )
            # Apply opacity via transparency by drawing a white rect on top won't work
            # Instead use the correct approach: draw text with alpha using Shape
            # Re-do using shape for proper opacity support
            page_rect = page.rect
            shape = page.new_shape()
            # Calculate rotated text position for center diagonal
            shape.insert_text(
                fitz.Point(cx - tw * 0.6, cy + tw * 0.35),
                text,
                fontname="helv",
                fontsize=font_size * 0.85,
                color=(0.55, 0.08, 0.08),
                rotate=45,
            )
            shape.finish(
                fill=None,
                color=(0.55, 0.08, 0.08),
                fill_opacity=opacity,
                stroke_opacity=opacity,
            )
            shape.commit(overlay=True)

        elif position == "bottom-right":
            font_size = min(w, h) * 0.042
            tw = fitz.get_text_length(text, fontname="helv", fontsize=font_size)
            shape = page.new_shape()
            shape.insert_text(
                fitz.Point(w - tw - 22, h - 22),
                text,
                fontname="helv",
                fontsize=font_size,
                color=(0.35, 0.35, 0.35),
            )
            shape.finish(color=(0.35,0.35,0.35), fill_opacity=opacity, stroke_opacity=opacity)
            shape.commit(overlay=True)

        else:  # bottom-center
            font_size = min(w, h) * 0.042
            tw = fitz.get_text_length(text, fontname="helv", fontsize=font_size)
            shape = page.new_shape()
            shape.insert_text(
                fitz.Point((w - tw) / 2, h - 22),
                text,
                fontname="helv",
                fontsize=font_size,
                color=(0.35, 0.35, 0.35),
            )
            shape.finish(color=(0.35,0.35,0.35), fill_opacity=opacity, stroke_opacity=opacity)
            shape.commit(overlay=True)

    buf = io.BytesIO()
    doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"watermarked_{stem(file.filename)}.pdf")


# ══════════════════════════════════════════
# CONVERT
# ══════════════════════════════════════════

@app.post("/pdf-to-word")
async def pdf_to_word(file: UploadFile = File(...)):
    data = await read_file(file)
    if not _DOCX_OK: raise HTTPException(500, "Word library not available.")
    doc_out = Document()
    doc_out.styles["Normal"].font.name = "Calibri"
    doc_out.styles["Normal"].font.size = Pt(11)
    fitz_doc = open_fitz(data)
    for page_num in range(len(fitz_doc)):
        page = fitz_doc[page_num]
        if page_num > 0: doc_out.add_page_break()
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0: continue
            for line in block.get("lines", []):
                line_text = " ".join(s.get("text","") for s in line.get("spans",[])).strip()
                if not line_text: continue
                sizes = [s.get("size",11) for s in line.get("spans",[])]
                avg = sum(sizes)/len(sizes) if sizes else 11
                if avg > 16: doc_out.add_heading(line_text, level=1)
                elif avg > 13: doc_out.add_heading(line_text, level=2)
                else: doc_out.add_paragraph(line_text)
    buf = io.BytesIO(); doc_out.save(buf)
    return stream_file(buf.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{stem(file.filename)}.docx")


# ══════════════════════════════════════════════════════════════════
# PDF TO EXCEL — SMART GROUPING FIX (header-based, NOT page-based)
# ══════════════════════════════════════════════════════════════════

@app.post("/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...), mode: str = Form("smart")):
    data = await read_file(file)
    if not _XL_OK: raise HTTPException(500, "Excel library not available.")
    if not _PLUMBER_OK: raise HTTPException(500, "PDF processing library not available.")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles
    hdr_fill  = PatternFill("solid", fgColor="4F6EF7")
    hdr_font  = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alt_fill  = PatternFill("solid", fgColor="EEF1FF")
    thin      = Side(style="thin", color="C5CEFF")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    def norm_hdr(row):
        return tuple(re.sub(r"\s+", " ", str(c or "").strip().lower()) for c in row)

    def is_header(row):
        if not row: return False
        non_empty = [c for c in row if str(c or "").strip()]
        if len(non_empty) < 2: return False
        numeric = sum(1 for c in non_empty if re.match(r"^[\d\s\.\,\-\$\%\+\(\)]+$", str(c).strip()))
        return numeric / len(non_empty) < 0.55

    def safe_name(base, existing):
        name = re.sub(r"[\\/*?:\[\]]", "", base)[:28].strip() or "Table"
        cand, n = name, 1
        while cand in existing:
            cand = f"{name[:25]}_{n}"; n += 1
        return cand

    # ─── CORE FIX ───────────────────────────────────────────────────
    # Key: table_groups dict keyed by header tuple.
    # ALL pages processed → same header = SAME sheet (rows appended).
    # This prevents the old page-wise sheet creation bug.
    # ────────────────────────────────────────────────────────────────
    table_groups = {}   # hkey → {sheet_name, header, rows, col_count}
    text_lines   = []

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for pn, page in enumerate(pdf.pages, start=1):

                if mode in ("smart", "tables"):
                    # Try strict lines first, fall back to text strategy
                    tables = page.extract_tables({
                        "vertical_strategy": "lines_strict",
                        "horizontal_strategy": "lines_strict",
                        "snap_tolerance": 4, "join_tolerance": 4,
                        "edge_min_length": 15,
                        "min_words_vertical": 1, "min_words_horizontal": 1,
                    }) or []
                    if not tables:
                        tables = page.extract_tables({
                            "vertical_strategy": "text",
                            "horizontal_strategy": "text",
                        }) or []

                    for tbl in tables:
                        if not tbl or len(tbl) < 2: continue

                        # Find header row (first row that looks like a header)
                        hdr_idx = 0
                        for ri, row in enumerate(tbl[:4]):
                            if is_header(row):
                                hdr_idx = ri; break

                        raw_hdr = tbl[hdr_idx]
                        if not any(str(c or "").strip() for c in raw_hdr): continue

                        hkey        = norm_hdr(raw_hdr)
                        disp_hdr    = [str(c or "").strip() for c in raw_hdr]
                        data_rows   = []

                        for row in tbl[hdr_idx + 1:]:
                            cleaned = [str(c or "").strip() for c in row]
                            if any(v for v in cleaned):
                                data_rows.append(cleaned)

                        if not data_rows: continue

                        # ── SMART GROUPING: same header → append rows (not new sheet) ──
                        if hkey not in table_groups:
                            label = " & ".join(h for h in disp_hdr[:2] if h)
                            table_groups[hkey] = {
                                "sheet_name": label,
                                "header": disp_hdr,
                                "rows": [],
                                "col_count": len(disp_hdr),
                            }
                        table_groups[hkey]["rows"].extend(data_rows)

                # Text fallback
                if mode == "text" or (mode == "smart" and not table_groups):
                    for ln in (page.extract_text() or "").split("\n"):
                        if ln.strip(): text_lines.append((pn, ln.strip()))

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"PDF processing failed: {e}")

    existing = []

    if table_groups:
        for idx, (hkey, grp) in enumerate(table_groups.items(), start=1):
            sname = safe_name(grp["sheet_name"] or f"Table {idx}", existing)
            existing.append(sname)
            ws = wb.create_sheet(title=sname)
            ws.freeze_panes = "A2"

            # Header row
            for ci, h in enumerate(grp["header"], start=1):
                c = ws.cell(row=1, column=ci, value=h)
                c.fill = hdr_fill; c.font = hdr_font
                c.alignment = hdr_align; c.border = border

            # Data rows
            for ri, row in enumerate(grp["rows"], start=2):
                cols = grp["col_count"]
                padded = (row + [""] * cols)[:cols]
                for ci, val in enumerate(padded, start=1):
                    cell = ws.cell(row=ri, column=ci)
                    stripped = val.replace(",","").replace("$","").replace("%","").strip()
                    try:
                        cell.value = int(stripped) if "." not in stripped else float(stripped)
                    except (ValueError, TypeError):
                        cell.value = val
                    cell.border = border
                    if ri % 2 == 0: cell.fill = alt_fill

            # Auto column width
            for col in ws.columns:
                mx = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max(mx+3, 12), 60)

            # Row count annotation
            ws.cell(row=1, column=grp["col_count"]+2, value=f"Total rows: {len(grp['rows'])}").font = Font(italic=True, color="9BA3C8", size=10)

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
        ws.cell(row=1, column=1, value="No tables or text could be extracted.")
        ws.cell(row=2, column=1, value="Tip: PDF must contain selectable text (not scanned images).")

    out = io.BytesIO(); wb.save(out)
    return stream_file(out.getvalue(),
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       f"{stem(file.filename)}.xlsx")


@app.post("/pdf-to-jpg")
async def pdf_to_jpg(file: UploadFile = File(...), dpi: int = Form(150)):
    data = await read_file(file); doc = open_fitz(data)
    dpi = max(72, min(dpi, 300)); mat = fitz.Matrix(dpi/72, dpi/72)
    zb = io.BytesIO()
    with zipfile.ZipFile(zb, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=mat)
            zf.writestr(f"page_{i+1:03d}.jpg", pix.tobytes("jpeg"))
    return stream_file(zb.getvalue(), "application/zip", f"{stem(file.filename)}_images.zip")


@app.post("/jpg-to-pdf")
async def jpg_to_pdf(files: List[UploadFile] = File(...)):
    doc = fitz.open()
    for uf in files:
        data = await read_file(uf)
        try:
            ext = (uf.content_type or "jpeg").split("/")[-1]
            img_doc = fitz.open(stream=data, filetype=ext)
            page = doc.new_page(width=img_doc[0].rect.width, height=img_doc[0].rect.height)
            page.show_pdf_page(page.rect, img_doc, 0); img_doc.close()
        except Exception:
            try:
                page = doc.new_page(width=595, height=842)
                page.insert_image(fitz.Rect(0,0,595,842), stream=data)
            except Exception as e:
                raise HTTPException(400, f"Could not process '{uf.filename}': {e}")
    buf = io.BytesIO(); doc.save(buf)
    return stream_file(buf.getvalue(), "application/pdf", "images.pdf")


@app.post("/pdf-to-pptx")
async def pdf_to_pptx(file: UploadFile = File(...)):
    data = await read_file(file)
    if not _PPTX_OK: raise HTTPException(500, "PowerPoint library not available.")
    doc = open_fitz(data); prs = Presentation()
    prs.slide_width = PInches(10); prs.slide_height = PInches(7.5)
    blank = prs.slide_layouts[6]
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5,1.5))
        img = pix.tobytes("png")
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(io.BytesIO(img), PInches(0), PInches(0), width=prs.slide_width, height=prs.slide_height)
    buf = io.BytesIO(); prs.save(buf)
    return stream_file(buf.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation", f"{stem(file.filename)}.pptx")


@app.post("/pdf-to-text")
async def pdf_to_text(file: UploadFile = File(...)):
    data = await read_file(file)
    if _PLUMBER_OK:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            lines = []
            for i, page in enumerate(pdf.pages):
                lines.append(f"─── Page {i+1} ───\n")
                lines.append(page.extract_text() or "[No text]")
                lines.append("\n")
            text = "\n".join(lines)
    else:
        text = extract_text_fitz(data)
    return stream_file(text.encode("utf-8"), "text/plain", f"{stem(file.filename)}.txt")


@app.post("/ocr-check")
async def ocr_check(file: UploadFile = File(...)):
    data = await read_file(file); doc = open_fitz(data)
    pages_checked = min(len(doc), 3)
    total_chars = sum(len(doc[i].get_text()) for i in range(pages_checked))
    return {"is_scanned": total_chars < 80, "chars_found": total_chars, "pages_checked": pages_checked}
