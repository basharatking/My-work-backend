# ══════════════════════════════════════════════════════════════════
# ADDITIONAL ENDPOINTS — paste these into main.py
# These power: word-to-pdf.html, excel-to-pdf.html, delete-pages.html,
# reorder-pages.html, sign-pdf.html (previously these pages called the
# WRONG/wrong-purpose endpoints — e.g. word-to-pdf called /jpg-to-pdf).
#
# HOW TO INSTALL:
# 1. Add the imports below to the top of main.py (only the ones you
#    don't already have)
# 2. Paste each @app.post(...) function anywhere after `app = FastAPI(...)`
# 3. Restart your Replit backend
# ══════════════════════════════════════════════════════════════════

# ---- Extra imports needed (add to top of main.py if missing) ----
# import re   <- already imported in main.py
# from fastapi import HTTPException, UploadFile, File, Form  <- already imported
# import fitz  <- already imported as _FITZ_OK guarded import
#
# For Word -> PDF and Excel -> PDF we use LibreOffice via subprocess,
# which is the most reliable free way to get real .docx/.xlsx -> PDF
# conversion (font/layout accurate). Install on Replit with:
#   nix: add `pkgs.libreoffice` to replit.nix, OR
#   apt: `apt-get install -y libreoffice` in a Nix/Debian environment
#
# If LibreOffice is not available, these endpoints fall back to a
# "best effort" text-based PDF so the tool never hard-fails.

import subprocess
import tempfile
import shutil


def _libreoffice_available() -> bool:
    return shutil.which("libreoffice") is not None or shutil.which("soffice") is not None


def _convert_with_libreoffice(input_bytes: bytes, input_ext: str, out_ext: str = "pdf") -> bytes:
    """Convert a file to PDF (or other format) using headless LibreOffice."""
    binary = shutil.which("soffice") or shutil.which("libreoffice")
    if not binary:
        raise HTTPException(500, "Conversion engine not available on server.")

    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, f"input.{input_ext}")
        with open(in_path, "wb") as f:
            f.write(input_bytes)

        cmd = [
            binary, "--headless", "--norestore", "--invisible",
            "--convert-to", out_ext, "--outdir", tmp, in_path,
        ]
        try:
            subprocess.run(cmd, check=True, timeout=60, capture_output=True)
        except subprocess.TimeoutExpired:
            raise HTTPException(504, "Conversion timed out. Try a smaller file.")
        except subprocess.CalledProcessError as e:
            raise HTTPException(500, f"Conversion failed: {e.stderr.decode(errors='ignore')[:200]}")

        out_path = os.path.join(tmp, f"input.{out_ext}")
        if not os.path.exists(out_path):
            raise HTTPException(500, "Conversion produced no output file.")
        with open(out_path, "rb") as f:
            return f.read()


# ══════════════════════════════════════════
# WORD TO PDF
# ══════════════════════════════════════════

@app.post("/word-to-pdf")
async def word_to_pdf(file: UploadFile = File(...)):
    data = await read_file(file)
    ext = "docx" if file.filename.lower().endswith(".docx") else "doc"

    if _libreoffice_available():
        pdf_bytes = _convert_with_libreoffice(data, ext, "pdf")
        return stream_file(pdf_bytes, "application/pdf", f"{stem(file.filename)}.pdf")

    # Fallback: extract text with python-docx and render a simple PDF with fitz
    if not _DOCX_OK or not _FITZ_OK:
        raise HTTPException(500, "Conversion engine not available. Contact support.")
    try:
        word_doc = Document(io.BytesIO(data))
    except Exception as e:
        raise HTTPException(400, f"Could not read Word file: {e}")

    pdf_doc = fitz.open()
    page = pdf_doc.new_page()
    y = 50
    for para in word_doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        if y > 780:
            page = pdf_doc.new_page()
            y = 50
        page.insert_text((50, y), text, fontsize=11)
        y += 16
    buf = io.BytesIO()
    pdf_doc.save(buf)
    return stream_file(buf.getvalue(), "application/pdf", f"{stem(file.filename)}.pdf")


# ══════════════════════════════════════════
# EXCEL TO PDF
# ══════════════════════════════════════════

@app.post("/excel-to-pdf")
async def excel_to_pdf(file: UploadFile = File(...)):
    data = await read_file(file)
    ext = "xlsx" if file.filename.lower().endswith(".xlsx") else "xls"

    if _libreoffice_available():
        pdf_bytes = _convert_with_libreoffice(data, ext, "pdf")
        return stream_file(pdf_bytes, "application/pdf", f"{stem(file.filename)}.pdf")

    # Fallback: render Excel cell values as a simple table in a PDF
    if not _XL_OK or not _FITZ_OK:
        raise HTTPException(500, "Conversion engine not available. Contact support.")
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        raise HTTPException(400, f"Could not read Excel file: {e}")

    pdf_doc = fitz.open()
    for ws in wb.worksheets:
        page = pdf_doc.new_page()
        y = 50
        page.insert_text((40, 35), f"Sheet: {ws.title}", fontsize=13)
        for row in ws.iter_rows(values_only=True):
            if y > 780:
                page = pdf_doc.new_page()
                y = 50
            line = "  |  ".join(str(c) if c is not None else "" for c in row)
            page.insert_text((40, y), line[:140], fontsize=9)
            y += 14
    buf = io.BytesIO()
    pdf_doc.save(buf)
    return stream_file(buf.getvalue(), "application/pdf", f"{stem(file.filename)}.pdf")


# ══════════════════════════════════════════
# DELETE PAGES  (e.g. "1,3,5" or "2-4" or "1,3-5,8")
# ══════════════════════════════════════════

def _parse_page_spec(spec: str, total_pages: int) -> set:
    """Parses '1,3,5' or '2-4' or '1,3-5,8' into a 0-indexed set of page numbers."""
    pages = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            try:
                a, b = part.split("-")
                a, b = int(a.strip()), int(b.strip())
                for p in range(min(a, b), max(a, b) + 1):
                    if 1 <= p <= total_pages:
                        pages.add(p - 1)
            except ValueError:
                continue
        else:
            try:
                p = int(part)
                if 1 <= p <= total_pages:
                    pages.add(p - 1)
            except ValueError:
                continue
    return pages


@app.post("/delete-pages")
async def delete_pages(file: UploadFile = File(...), pages: str = Form(...)):
    data = await read_file(file)
    doc = open_fitz(data)
    total = len(doc)

    to_delete = _parse_page_spec(pages, total)
    if not to_delete:
        raise HTTPException(400, "No valid page numbers found. Use format like 1,3,5 or 2-4.")
    if len(to_delete) >= total:
        raise HTTPException(400, "Cannot delete all pages from the document.")

    doc.delete_pages(sorted(to_delete))
    buf = io.BytesIO()
    doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"edited_{stem(file.filename)}.pdf")


# ══════════════════════════════════════════
# REORDER PAGES  (e.g. "3,1,2,4")
# ══════════════════════════════════════════

@app.post("/reorder-pages")
async def reorder_pages(file: UploadFile = File(...), order: str = Form(...)):
    data = await read_file(file)
    doc = open_fitz(data)
    total = len(doc)

    try:
        new_order = [int(x.strip()) - 1 for x in order.split(",") if x.strip()]
    except ValueError:
        raise HTTPException(400, "Invalid page order format. Use e.g. 3,1,2,4")

    if sorted(new_order) != list(range(total)):
        raise HTTPException(
            400,
            f"Page order must include every page exactly once (1 to {total}).",
        )

    new_doc = fitz.open()
    for idx in new_order:
        new_doc.insert_pdf(doc, from_page=idx, to_page=idx)

    buf = io.BytesIO()
    new_doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"reordered_{stem(file.filename)}.pdf")


# ══════════════════════════════════════════
# SIGN PDF  (text-based signature stamp)
# ══════════════════════════════════════════

@app.post("/sign-pdf")
async def sign_pdf(
    file: UploadFile = File(...),
    text: str = Form(...),
    position: str = Form("bottom-right"),
    pages: str = Form("last"),
):
    data = await read_file(file)
    doc = open_fitz(data)
    total = len(doc)

    if pages == "all":
        target_pages = range(total)
    elif pages == "first":
        target_pages = [0]
    else:  # "last"
        target_pages = [total - 1]

    for i in target_pages:
        page = doc[i]
        w, h = page.rect.width, page.rect.height
        font_size = 14
        tw = fitz.get_text_length(text, fontname="helv", fontsize=font_size)

        pos_map = {
            "bottom-right": (w - tw - 40, h - 45),
            "bottom-center": ((w - tw) / 2, h - 45),
            "bottom-left": (40, h - 45),
        }
        x, y = pos_map.get(position, (w - tw - 40, h - 45))

        # Draw a subtle line above the signature for a more "signed" look
        page.draw_line(
            fitz.Point(x, y - 6), fitz.Point(x + tw, y - 6),
            color=(0.2, 0.2, 0.2), width=0.6,
        )
        page.insert_text(
            (x, y), text,
            fontname="helv", fontsize=font_size,
            color=(0.1, 0.1, 0.5),
        )
        page.insert_text(
            (x, y + 14), "Digitally signed via RunDocs",
            fontname="helv", fontsize=7,
            color=(0.5, 0.5, 0.5),
        )

    buf = io.BytesIO()
    doc.save(buf, garbage=3, deflate=True)
    return stream_file(buf.getvalue(), "application/pdf", f"signed_{stem(file.filename)}.pdf")
