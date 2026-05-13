import os
import io
from fastapi import FastAPI, UploadFile, File, HTTPException, Header
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from PyPDF2 import PdfReader, PdfWriter
from dotenv import load_dotenv

# Environment variables (.env file) load karne ke liye
load_dotenv()

app = FastAPI(title="ITSPDF Utility API")

# Netlify frontend ko connect karne ke liye CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- SECURE TOKEN SYSTEM ---
# Ye 'os' library system ki settings se token uthayegi
# Aapko code mein MyAi... likhne ki zaroorat nahi
API_TOKEN = os.getenv("MYAI_TOKEN")

@app.get("/")
async def health_check():
    return {
        "status": "online",
        "project": "ITSPDF",
        "auth_configured": API_TOKEN is not None
    }

# 1. PDF MERGE FUNCTION
@app.post("/merge")
async def merge_pdfs(files: list[UploadFile] = File(...)):
    if not API_TOKEN:
        raise HTTPException(status_code=500, detail="Server Error: Token missing in environment variables.")
    
    pdf_writer = PdfWriter()
    try:
        for file in files:
            content = await file.read()
            pdf_reader = PdfReader(io.BytesIO(content))
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)

        output = io.BytesIO()
        pdf_writer.write(output)
        output.seek(0)
        
        return StreamingResponse(
            output,
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=itspdf_merged.pdf"}
        )
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})

# 2. PDF INFO/METADATA FUNCTION
@app.post("/metadata")
async def get_metadata(file: UploadFile = File(...)):
    try:
        pdf_reader = PdfReader(io.BytesIO(await file.read()))
        info = pdf_reader.metadata
        return {
            "filename": file.filename,
            "pages": len(pdf_reader.pages),
            "author": info.author if info.author else "Unknown",
            "creator": info.creator if info.creator else "Unknown"
        }
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})

# 3. SECURITY CHECK ROUTE (Testing ke liye)
@app.get("/verify-token")
async def verify():
    if API_TOKEN:
        # Security ke liye poora token nahi dikhayenge, sirf check karenge
        return {"message": "Token is active and hidden", "prefix": API_TOKEN[:7]}
    return {"message": "Token not found!"}
