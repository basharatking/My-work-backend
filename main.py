import os
import io
import requests
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
from googletrans import Translator

app = FastAPI()

# CORS taake Netlify frontend se backend connect ho sake
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- FREE AI CONFIG (Hugging Face) ---
# Bhai aapka token maine yahan add kar diya hai
HF_TOKEN = "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" 
HF_API_URL = "https://api-inference.huggingface.co/models/mistralai/Mistral-7B-Instruct-v0.2"
headers = {"Authorization": f"Bearer {HF_TOKEN}"}

translator = Translator()

def get_free_ai_response(prompt: str):
    """Hugging Face API for free text generation"""
    payload = {
        "inputs": f"<s>[INST] {prompt} [/INST]",
        "parameters": {"max_new_tokens": 500, "temperature": 0.7}
    }
    try:
        response = requests.post(HF_API_URL, headers=headers, json=payload)
        if response.status_code != 200:
            return "AI model is loading... Please wait 30 seconds and try again."
        
        result = response.json()
        if isinstance(result, list) and len(result) > 0:
            return result[0]['generated_text'].split("[/INST]")[-1].strip()
        return "AI response format error."
    except Exception as e:
        return f"Error: {str(e)}"

# --- AI ENDPOINTS ---

@app.post("/ai-summary")
async def ai_summary(file: UploadFile = File(...)):
    """Free Summary without Claude (Zero Cost)"""
    try:
        pdf_content = await file.read()
        reader = PdfReader(io.BytesIO(pdf_content))
        text = ""
        # Sirf pehle 5 pages summarize karein taake fast rahe
        for page in reader.pages[:5]: 
            extracted = page.extract_text()
            if extracted:
                text += extracted
        
        if not text.strip():
            return {"result": "Could not extract text from this PDF."}

        prompt = f"Summarize the following text in clear bullet points:\n\n{text[:3000]}"
        summary = get_free_ai_response(prompt)
        return {"result": summary}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/ai-translate")
async def ai_translate(file: UploadFile = File(...), target_lang: str = Form("ur")):
    """Free Translation using Googletrans"""
    try:
        pdf_content = await file.read()
        reader = PdfReader(io.BytesIO(pdf_content))
        text = reader.pages[0].extract_text() if reader.pages else ""
        
        if not text.strip():
            return {"result": "No text found to translate."}
            
        translated = translator.translate(text[:2500], dest=target_lang)
        return {"result": translated.text}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- STANDARD PDF TOOLS ---

@app.post("/merge-pdf")
async def merge_pdf(files: list[UploadFile] = File(...)):
    """Standard Merge (100% Free)"""
    try:
        merger = PdfWriter()
        for file in files:
            content = await file.read()
            merger.append(io.BytesIO(content))
        out = io.BytesIO()
        merger.write(out)
        out.seek(0)
        return StreamingResponse(out, media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=merged_jasonpdf.pdf"})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
def health():
    return {"status": "online", "mode": "free-tier-activated"}
