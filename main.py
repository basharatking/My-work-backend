import os
import io
import requests
import pandas as pd
import pdfplumber
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
from googletrans import Translator

app = FastAPI()

# CORS taake Netlify frontend connect ho sake
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
        if response.status_code != 200:
            return "AI model is loading... Please wait 30 seconds."
        result = response.json()
        return result[0]['generated_text'].split("[/INST]")[-1].strip()
    except Exception:
        return "AI service temporarily unavailable."

# --- AI ENDPOINTS ---

@app.post("/ai-summary")
async def ai_summary(file: UploadFile = File(...)):
    try:
        pdf_content = await file.read()
        reader = PdfReader(io.BytesIO(pdf_content))
        text = "".join([page.extract_text() for page in reader.pages[:5]])
        prompt = f"Summarize this in bullet points:\n\n{text[:3000]}"
        return {"result": get_free_ai_response(prompt)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/ai-translate")
async def ai_translate(file: UploadFile = File(...), target_lang: str = Form("ur")):
    try:
        pdf_content = await file.read()
        reader = PdfReader(io.BytesIO(pdf_content))
        text = reader.pages[0].extract_text()
        translated = translator.translate(text[:2500], dest=target_lang)
        return {"result": translated.text}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- EXCEL TOOL (MERGED TABLES LOGIC) ---

@app.post("/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    try:
        content = await file.read()
        all_tables_data = []
        
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    # Table ko DataFrame mein convert karna
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables_data.append(df)
        
        if not all_tables_data:
            raise HTTPException(status_code=400, detail="No tables found.")

        # Saare pages ke tables ko aik hi sheet mein jorna
        final_df = pd.concat(all_tables_data, ignore_index=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name="MasterData", index=False)
            
        output.seek(0)
        return StreamingResponse(
            output, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=jasonpdf_excel.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- STANDARD PDF TOOLS ---

@app.post("/merge-pdf")
async def merge_pdf(files: list[UploadFile] = File(...)):
    merger = PdfWriter()
    for file in files:
        content = await file.read()
        merger.append(io.BytesIO(content))
    out = io.BytesIO()
    merger.write(out)
    out.seek(0)
    return StreamingResponse(out, media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=merged.pdf"})

@app.get("/health")
def health():
    return {"status": "online", "mode": "free-tier"}
