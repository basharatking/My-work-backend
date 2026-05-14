import os
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from mangum import Mangum
import pandas as pd
import pdfplumber
from dotenv import load_dotenv

# Environment variables load karne ke liye
load_dotenv()

app = FastAPI()

# Netlify settings se token uthane ke liye
# Yaad rahe Netlify mein Key: MYAI_TOKEN honi chahiye
API_TOKEN = os.getenv("MYAI_TOKEN")

@app.get("/")
def home():
    return {"message": "Welcome to ITSPDF API. System is Live!"}

@app.get("/verify-token")
def verify():
    if API_TOKEN:
        # Security ki wajah se pura token nahi dikhayenge
        return {"status": "Active", "prefix": API_TOKEN[:7]}
    return {"status": "Error", "message": "Token not found in Netlify settings"}

@app.post("/convert-to-excel")
async def convert_pdf_to_excel(file: UploadFile = File(...)):
    # 1. Token Check
    if not API_TOKEN:
        raise HTTPException(status_code=500, detail="API Token missing in server settings")

    # 2. PDF processing
    try:
        all_data = []
        with pdfplumber.open(file.file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    # Table data ko list mein add karna
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_data.append(df)

        if not all_data:
            return JSONResponse(status_code=400, content={"message": "No tables found in this PDF"})

        # 3. Excel file banana
        final_df = pd.concat(all_data, ignore_index=True)
        output_path = f"converted_{file.filename}.xlsx"
        final_df.to_excel(output_path, index=False)

        return FileResponse(
            path=output_path, 
            filename=f"converted_{file.filename}.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

# --- NETLIFY HANDLER (Ye sab se zaroori hai) ---
handler = Mangum(app)
