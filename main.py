import os
import io
import requests
import pandas as pd
import pdfplumber
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pypdf import PdfReader, PdfWriter
from googletrans import Translator

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- FREE AI CONFIG ---
HF_TOKEN = "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" 
HF_API_URL = "https://api-inference.huggingface.co/models/mistralai/Mistral-7B-Instruct-v0.2"
headers = {"Authorization": f"Bearer {HF_TOKEN}"}

translator = Translator()

def get_free_ai_response(prompt: str):
    payload = {
        "inputs": f"<s>[INST] {prompt} [/INST]",
        "parameters": {"max_new_tokens": 500, "temperature": 0.7}
    }
    try:
        response = requests.post(HF_API_URL, headers=headers, json=payload)
        result = response.json()
        return result[0]['generated_text'].split("[/INST]")[-1].strip()
    except:
        return "AI is busy, please try again."

# --- UPDATED EXCEL TOOL (GROUP BY HEADER TABS) ---

@app.post("/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    try:
        content = await file.read()
        # Dictionary taake hum headers ke mutabiq data group kar saken
        grouped_data = {}
        
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table and len(table) > 1:
                    # Header ko string bana kar key banayenge (e.g. "Name-Age-Salary")
                    header_tuple = tuple(table[0])
                    header_name = "-".join([str(h) for h in header_tuple if h])[:30] # Excel tab limit
                    
                    if header_name not in grouped_data:
                        grouped_data[header_name] = {"columns": table[0], "rows": []}
                    
                    # Data rows add karein
                    grouped_data[header_name]["rows"].extend(table[1:])
        
        if not grouped_data:
            raise HTTPException(status_code=400, detail="No tables found.")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_title, content in grouped_data.items():
                # Har unique header ke liye ek alag sheet tab
                df = pd.DataFrame(content["rows"], columns=content["columns"])
                df.to_excel(writer, sheet_name=sheet_title, index=False)
            
        output.seek(0)
        return StreamingResponse(
            output, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=jasonpdf_header_tabs.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- OTHER ENDPOINTS ---

@app.post("/ai-summary")
async def ai_summary(file: UploadFile = File(...)):
    pdf_content = await file.read()
    reader = PdfReader(io.BytesIO(pdf_content))
    text = "".join([page.extract_text() for page in reader.pages[:5]])
    return {"result": get_free_ai_response(f"Summarize:\n{text[:3000]}")}

@app.post("/ai-translate")
async def ai_translate(file: UploadFile = File(...), target_lang: str = Form("ur")):
    pdf_content = await file.read()
    reader = PdfReader(io.BytesIO(pdf_content))
    text = reader.pages[0].extract_text()
    return {"result": translator.translate(text[:2500], dest=target_lang).text}

@app.post("/merge-pdf")
async def merge_pdf(files: list[UploadFile] = File(...)):
    merger = PdfWriter()
    for file in files:
        merger.append(io.BytesIO(await file.read()))
    out = io.BytesIO()
    merger.write(out)
    out.seek(0)
    return StreamingResponse(out, media_type="application/pdf")

@app.get("/health")
def health():
    return {"status": "online"}
