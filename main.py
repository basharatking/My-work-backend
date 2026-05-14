"""CatchPDF v9 - Premium SaaS Backend"""
import io, os, hashlib
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill

app = FastAPI(title="CatchPDF")

# CORS for Netlify
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def stem(f): return Path(f or "file").stem

@app.post("/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    d = await file.read()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    table_groups = {} 
    header_fill = PatternFill("solid", fgColor="4F46E5")
    header_font = Font(bold=True, color="FFFFFF")

    try:
        with pdfplumber.open(io.BytesIO(d)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for tbl in tables:
                    if not tbl or not any(tbl[0]): continue
                    
                    # Logic: Group by Header Hash
                    headers = [str(c or "").strip() for c in tbl[0]]
                    header_hash = hashlib.md5("".join(headers).lower().encode()).hexdigest()
                    
                    if header_hash not in table_groups:
                        table_groups[header_hash] = {
                            "name": f"Sheet_{len(table_groups)+1}",
                            "headers": headers,
                            "rows": []
                        }
                    
                    for row in tbl[1:]:
                        if any(row):
                            table_groups[header_hash]["rows"].append([str(v or "").strip() for v in row])

        for group in table_groups.values():
            ws = wb.create_sheet(title=group["name"])
            ws.append(group["headers"])
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for row in group["rows"]:
                ws.append(row)

    except Exception as e:
        raise HTTPException(500, f"Error: {str(e)}")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return StreamingResponse(out, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename=CatchPDF_{stem(file.filename)}.xlsx"})

# NOTE: Keep all your other existing @app.post routes below this...
