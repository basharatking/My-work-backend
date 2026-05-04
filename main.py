"""JasonPDF v8 - FastAPI Backend with AI Tools"""
import io, os, zipfile, json
from pathlib import Path
from typing import List
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import pypdf, pdfplumber
from PIL import Image
import img2pdf
from docx import Document
from docx.shared import Pt, RGBColor
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
try:
    import fitz; _FITZ_OK=True
except: _FITZ_OK=False
try:
    from pptx import Presentation
    from pptx.util import Inches,Pt as PptPt
    from pptx.dml.color import RGBColor as PptxRGB
    _PPTX_OK=True
except: _PPTX_OK=False
try:
    import anthropic
    _AI=anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY",""))
    _AI_OK=True
except: _AI=None; _AI_OK=False

app=FastAPI(title="JasonPDF",version="8.0",docs_url="/api/docs")
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_credentials=True,allow_methods=["*"],allow_headers=["*"],
    expose_headers=["X-Original-Size","X-Compressed-Size","X-Savings-Pct","Content-Disposition"])

FREE_LIMIT_BYTES=int(os.environ.get("FREE_LIMIT_MB","25"))*1024*1024
STATIC_DIR=Path(__file__).parent/"static"

def stream_file(data,media_type,filename,extra=None):
    h={"Content-Disposition":f'attachment; filename="{filename}"'}
    if extra:h.update(extra)
    return StreamingResponse(io.BytesIO(data),media_type=media_type,headers=h)
async def read_file(u):
    d=await u.read()
    if len(d)>FREE_LIMIT_BYTES:raise HTTPException(413,"File too large. Limit is 25 MB.")
    if not d:raise HTTPException(400,"File is empty.")
    return d
def stem(f):return Path(f or "file").stem
def open_fitz(data):
    if not _FITZ_OK:raise HTTPException(500,"PDF engine not available.")
    try:return fitz.open(stream=data,filetype="pdf")
    except Exception as e:raise HTTPException(400,f"Cannot open PDF: {e}")

def extract_text(data:bytes,max_pages:int=30)->str:
    pages=[]
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for i,page in enumerate(pdf.pages[:max_pages]):
                t=page.extract_text() or ""
                if t.strip():pages.append(f"[Page {i+1}]\n{t.strip()}")
    except Exception as e:raise HTTPException(400,f"Cannot extract text: {e}")
    if not pages:raise HTTPException(400,"No extractable text. PDF may be scanned/image-based.")
    return "\n\n".join(pages)

def ai_call(system:str,prompt:str,max_tokens:int=2000)->str:
    if not _AI_OK:raise HTTPException(500,"AI not configured. Set ANTHROPIC_API_KEY in Replit Secrets.")
    try:
        r=_AI.messages.create(model="claude-sonnet-4-20250514",max_tokens=max_tokens,
            system=system,messages=[{"role":"user","content":prompt}])
        return r.content[0].text
    except anthropic.AuthenticationError:raise HTTPException(500,"AI API key invalid.")
    except anthropic.RateLimitError:raise HTTPException(429,"AI rate limit. Try again.")
    except Exception as e:raise HTTPException(500,f"AI error: {e}")

@app.get("/health")
def health():return{"status":"ok","version":"8.0","fitz":_FITZ_OK,"pptx":_PPTX_OK,"ai":_AI_OK}

# ── AI ENDPOINTS ──────────────────────────────────────────────────
@app.post("/ai-summary")
async def ai_summary(file:UploadFile=File(...),length:str=Form("medium")):
    data=await read_file(file);text=extract_text(data)
    lmap={"short":"Write a short 2-3 paragraph summary.","medium":"Write a detailed summary covering main topics, arguments and conclusions.","long":"Write a comprehensive summary with all major points, details and conclusions."}
    result=ai_call("You are an expert document analyst. Summarize documents clearly and accurately.",f"{lmap.get(length,lmap['medium'])}\n\nDocument:\n{text[:8000]}")
    return JSONResponse({"result":result})

@app.post("/ai-notes")
async def ai_notes(file:UploadFile=File(...),style:str=Form("bullet")):
    data=await read_file(file);text=extract_text(data)
    smap={"bullet":"Create clean bullet-point notes with clear headings for each topic.","cornell":"Create Cornell-style notes: Cue column left, notes right, summary at bottom.","outline":"Create hierarchical outline notes: I. Main topic A. Subtopic 1. Detail"}
    result=ai_call("You are an expert study notes creator. Create clear, student-friendly notes.",f"{smap.get(style,smap['bullet'])}\n\nDocument:\n{text[:8000]}")
    return JSONResponse({"result":result})

@app.post("/ai-quiz")
async def ai_quiz(file:UploadFile=File(...),count:int=Form(10),difficulty:str=Form("medium")):
    data=await read_file(file);text=extract_text(data);count=max(3,min(count,20))
    prompt=f"Create exactly {count} MCQ questions at {difficulty} difficulty.\n\nFormat each:\nQ1: [Question]\nA) [Option]\nB) [Option]\nC) [Option]\nD) [Option]\nAnswer: [Letter]\n\nDocument:\n{text[:7000]}"
    result=ai_call("You are an expert quiz creator. Generate MCQ questions from document content.",prompt)
    return JSONResponse({"result":result})

@app.post("/ai-keypoints")
async def ai_keypoints(file:UploadFile=File(...)):
    data=await read_file(file);text=extract_text(data)
    prompt=f"Extract the most important key points.\n\nFormat:\n📌 KEY POINTS\n• Point 1\n• Point 2\n\n📊 IMPORTANT FACTS & FIGURES\n[Numbers, dates, stats]\n\n💡 MAIN CONCLUSIONS\n[Key takeaways]\n\nDocument:\n{text[:8000]}"
    result=ai_call("You are an expert at identifying the most important information in documents.",prompt)
    return JSONResponse({"result":result})

@app.post("/ai-translate")
async def ai_translate(file:UploadFile=File(...),from_lang:str=Form("auto"),to_lang:str=Form("Urdu")):
    data=await read_file(file);text=extract_text(data,max_pages=15)
    frm="the detected language" if from_lang=="auto" else from_lang
    prompt=f"Translate from {frm} to {to_lang}. Preserve structure and formatting. Only translate, no explanations.\n\nText:\n{text[:6000]}"
    result=ai_call(f"You are an expert translator. Translate documents accurately.",prompt)
    return JSONResponse({"result":result,"from":from_lang,"to":to_lang})

@app.post("/ask-pdf")
async def ask_pdf(file:UploadFile=File(...),question:str=Form(...),history:str=Form("[]")):
    if not question.strip():raise HTTPException(400,"Question cannot be empty.")
    data=await read_file(file);text=extract_text(data,max_pages=25)
    try:hist=json.loads(history)[:6]
    except:hist=[]
    if not _AI_OK:raise HTTPException(500,"AI not configured. Set ANTHROPIC_API_KEY in Replit Secrets.")
    msgs=[]
    for h in hist:
        if h.get("role") in ("user","assistant"):msgs.append({"role":h["role"],"content":h["content"]})
    msgs.append({"role":"user","content":f"Document:\n{text[:6000]}\n\nQuestion: {question}"})
    try:
        r=_AI.messages.create(model="claude-sonnet-4-20250514",max_tokens=800,
            system="You are an intelligent PDF assistant. Answer ONLY based on the document. Be concise and helpful.",
            messages=msgs)
        return JSONResponse({"answer":r.content[0].text})
    except Exception as e:raise HTTPException(500,f"AI error: {e}")

# ── PDF TOOL ENDPOINTS ────────────────────────────────────────────
@app.post("/merge-pdf")
async def merge_pdf(files:List[UploadFile]=File(...)):
    if len(files)<2:raise HTTPException(400,"Select at least 2 PDFs.")
    w=pypdf.PdfWriter()
    for u in files:
        d=await read_file(u)
        try:r=pypdf.PdfReader(io.BytesIO(d));[w.add_page(p) for p in r.pages]
        except Exception as e:raise HTTPException(400,f"Cannot read '{u.filename}': {e}")
    out=io.BytesIO();w.write(out)
    return stream_file(out.getvalue(),"application/pdf","merged.pdf")

@app.post("/split-pdf")
async def split_pdf(file:UploadFile=File(...),mode:str=Form("each"),start_page:int=Form(1),end_page:int=Form(1)):
    data=await read_file(file)
    try:reader=pypdf.PdfReader(io.BytesIO(data));total=len(reader.pages)
    except Exception as e:raise HTTPException(400,f"Cannot read PDF: {e}")
    if mode=="range":
        s,e=max(1,start_page)-1,min(total,end_page)
        if s>=e:raise HTTPException(400,f"Invalid range. PDF has {total} pages.")
        w=pypdf.PdfWriter();[w.add_page(reader.pages[i]) for i in range(s,e)]
        out=io.BytesIO();w.write(out)
        return stream_file(out.getvalue(),"application/pdf",f"{stem(file.filename)}_p{s+1}-{e}.pdf")
    zb=io.BytesIO()
    with zipfile.ZipFile(zb,"w",zipfile.ZIP_DEFLATED) as zf:
        for i in range(total):
            w=pypdf.PdfWriter();w.add_page(reader.pages[i]);pb=io.BytesIO();w.write(pb)
            zf.writestr(f"page_{str(i+1).zfill(3)}.pdf",pb.getvalue())
    return stream_file(zb.getvalue(),"application/zip","split_pages.zip")

@app.post("/compress-pdf")
async def compress_pdf(file:UploadFile=File(...),level:str=Form("medium")):
    data=await read_file(file);orig=len(data)
    q={"low":85,"medium":60,"high":35}.get(level,60);md={"low":1600,"medium":1200,"high":900}.get(level,1200)
    try:
        doc=open_fitz(data)
        for page in doc:
            for img in page.get_images(full=True):
                xref=img[0]
                try:
                    base=doc.extract_image(xref);pil=Image.open(io.BytesIO(base["image"])).convert("RGB")
                    w,h=pil.size
                    if max(w,h)>md:scale=md/max(w,h);pil=pil.resize((int(w*scale),int(h*scale)),Image.LANCZOS)
                    buf=io.BytesIO();pil.save(buf,format="JPEG",quality=q,optimize=True);doc.update_stream(xref,buf.getvalue())
                except:pass
        out=io.BytesIO();doc.save(out,garbage=4,deflate=True,clean=True);doc.close()
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Compression failed: {e}")
    comp=out.getvalue();cs=len(comp);sav=round((1-cs/orig)*100,1) if orig else 0
    return stream_file(comp,"application/pdf",f"compressed_{file.filename}",{"X-Original-Size":str(orig),"X-Compressed-Size":str(cs),"X-Savings-Pct":str(sav)})

@app.post("/rotate-pdf")
async def rotate_pdf(file:UploadFile=File(...),angle:int=Form(90),pages:str=Form("all")):
    data=await read_file(file)
    if angle not in(90,180,270):raise HTTPException(400,"Angle must be 90,180 or 270.")
    try:
        doc=open_fitz(data)
        for i,page in enumerate(doc):
            pn=i+1
            if pages=="odd" and pn%2==0:continue
            if pages=="even" and pn%2!=0:continue
            page.set_rotation((page.rotation+angle)%360)
        out=io.BytesIO();doc.save(out);doc.close()
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Rotation failed: {e}")
    return stream_file(out.getvalue(),"application/pdf",f"rotated_{file.filename}")

@app.post("/add-watermark")
async def add_watermark(file:UploadFile=File(...),text:str=Form("CONFIDENTIAL"),opacity:float=Form(0.2),position:str=Form("center")):
    if not text.strip():raise HTTPException(400,"Watermark text empty.")
    opacity=max(0.05,min(opacity,0.95));data=await read_file(file)
    try:
        doc=open_fitz(data)
        for page in doc:
            w,h=page.rect.width,page.rect.height;fs=min(w,h)*0.08;color=(0.55,0.55,0.55)
            if position=="center":page.insert_text(fitz.Point(w*0.15,h*0.55),text,fontsize=fs,rotate=45,color=color,fill_opacity=opacity,overlay=True)
            elif position=="top":page.insert_text(fitz.Point(w*0.5-len(text)*fs*0.25,h-fs-20),text,fontsize=fs,color=color,fill_opacity=opacity,overlay=True)
            else:page.insert_text(fitz.Point(w*0.5-len(text)*fs*0.25,fs+20),text,fontsize=fs,color=color,fill_opacity=opacity,overlay=True)
        out=io.BytesIO();doc.save(out);doc.close()
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Watermark failed: {e}")
    return stream_file(out.getvalue(),"application/pdf",f"watermarked_{file.filename}")

@app.post("/pdf-to-word")
async def pdf_to_word(file:UploadFile=File(...)):
    data=await read_file(file);doc=Document()
    doc.styles["Normal"].font.name="Calibri";doc.styles["Normal"].font.size=Pt(11)
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            total=len(pdf.pages)
            if total==0:raise HTTPException(400,"PDF has no pages.")
            for i,page in enumerate(pdf.pages):
                h=doc.add_heading(f"Page {i+1}",level=1);h.runs[0].font.color.rgb=RGBColor(0xFF,0x38,0x14)
                text=page.extract_text() or ""
                if text.strip():[doc.add_paragraph(l).paragraph_format.__setattr__("space_after",Pt(2)) for l in text.split("\n")]
                else:p=doc.add_paragraph("[No extractable text]");p.runs[0].italic=True
                if i<total-1:doc.add_page_break()
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Conversion failed: {e}")
    out=io.BytesIO();doc.save(out)
    return stream_file(out.getvalue(),"application/vnd.openxmlformats-officedocument.wordprocessingml.document",f"{stem(file.filename)}.docx")

@app.post("/pdf-to-excel")
async def pdf_to_excel(file:UploadFile=File(...),mode:str=Form("smart")):
    data=await read_file(file);wb=openpyxl.Workbook();wb.remove(wb.active)
    hf=PatternFill("solid",fgColor="FF3814");font=Font(bold=True,color="FFFFFF",size=11)
    ha=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ts=Side(style="thin",color="DDDDDD");bdr=Border(left=ts,right=ts,top=ts,bottom=ts)
    af=PatternFill("solid",fgColor="FFF4F2");tgroups={};tlines=[]
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for pn,page in enumerate(pdf.pages,start=1):
                if mode in("smart","tables"):
                    for tbl in(page.extract_tables() or []):
                        if not tbl or len(tbl)<1:continue
                        hdrs=tuple(str(c or "").strip() for c in tbl[0])
                        if not any(hdrs):continue
                        if hdrs not in tgroups:tgroups[hdrs]=[]
                        for row in tbl[1:]:
                            if any(v for v in row):tgroups[hdrs].append([str(v or "").strip() for v in row])
                if mode=="text" or(mode=="smart" and not tgroups):
                    for ln in(page.extract_text() or "").split("\n"):
                        if ln.strip():tlines.append((pn,ln.strip()))
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Excel failed: {e}")
    if tgroups:
        for idx,(hdrs,rows) in enumerate(tgroups.items(),start=1):
            sn=" & ".join(h for h in hdrs[:2] if h)[:28] or f"Table {idx}"
            ex=[ws.title for ws in wb.worksheets];base,cnt=sn,1
            while sn in ex:sn=f"{base[:25]}_{cnt}";cnt+=1
            ws=wb.create_sheet(title=sn);ws.freeze_panes="A2"
            for ci,h in enumerate(hdrs,start=1):
                c=ws.cell(row=1,column=ci,value=h);c.fill=hf;c.font=font;c.alignment=ha;c.border=bdr
            for ri,row in enumerate(rows,start=2):
                fill=af if ri%2==0 else None
                for ci,val in enumerate(row,start=1):
                    c=ws.cell(row=ri,column=ci,value=val);c.border=bdr
                    if fill:c.fill=fill
            for col in ws.columns:
                mx=max((len(str(c.value or "")) for c in col),default=10)
                ws.column_dimensions[col[0].column_letter].width=min(max(mx+3,12),55)
    elif tlines:
        ws=wb.create_sheet(title="Extracted Text");ws.freeze_panes="A2"
        for ci,h in enumerate(["Page","Text"],start=1):
            c=ws.cell(row=1,column=ci,value=h);c.fill=hf;c.font=font;c.alignment=ha;c.border=bdr
        for ri,(pg,ln) in enumerate(tlines,start=2):
            ws.cell(row=ri,column=1,value=pg).border=bdr;ws.cell(row=ri,column=2,value=ln).border=bdr
        ws.column_dimensions["A"].width=8;ws.column_dimensions["B"].width=80
    else:ws=wb.create_sheet(title="No Data");ws.cell(row=1,column=1,value="No tables found.")
    out=io.BytesIO();wb.save(out)
    return stream_file(out.getvalue(),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",f"{stem(file.filename)}.xlsx")

@app.post("/pdf-to-jpg")
async def pdf_to_jpg(file:UploadFile=File(...),dpi:int=Form(150)):
    data=await read_file(file);dpi=max(72,min(dpi,300))
    try:
        doc=open_fitz(data);mat=fitz.Matrix(dpi/72.0,dpi/72.0);zb=io.BytesIO()
        with zipfile.ZipFile(zb,"w",zipfile.ZIP_DEFLATED) as zf:
            for i,page in enumerate(doc):
                pix=page.get_pixmap(matrix=mat,alpha=False);zf.writestr(f"page_{str(i+1).zfill(3)}.jpg",pix.tobytes("jpeg"))
        doc.close()
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Conversion failed: {e}")
    return stream_file(zb.getvalue(),"application/zip",f"{stem(file.filename)}_images.zip")

@app.post("/jpg-to-pdf")
async def jpg_to_pdf(files:List[UploadFile]=File(...)):
    if not files:raise HTTPException(400,"No files provided.")
    images=[]
    for u in files:
        raw=await read_file(u);img=Image.open(io.BytesIO(raw)).convert("RGB")
        buf=io.BytesIO();img.save(buf,format="JPEG",quality=92);images.append(buf.getvalue())
    out=io.BytesIO()
    try:out.write(img2pdf.convert(images))
    except Exception as e:raise HTTPException(500,f"PDF creation failed: {e}")
    return stream_file(out.getvalue(),"application/pdf","converted.pdf")

@app.post("/unlock-pdf")
async def unlock_pdf(file:UploadFile=File(...),password:str=Form("")):
    data=await read_file(file)
    try:
        doc=open_fitz(data)
        if doc.is_encrypted and not doc.authenticate(password):raise HTTPException(400,"Incorrect password.")
        out=io.BytesIO();doc.save(out,encryption=fitz.PDF_ENCRYPT_NONE);doc.close()
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Unlock failed: {e}")
    return stream_file(out.getvalue(),"application/pdf",f"unlocked_{file.filename}")

@app.post("/protect-pdf")
async def protect_pdf(file:UploadFile=File(...),password:str=Form(...)):
    if not password:raise HTTPException(400,"Password cannot be empty.")
    data=await read_file(file)
    try:
        doc=open_fitz(data);out=io.BytesIO()
        doc.save(out,encryption=fitz.PDF_ENCRYPT_AES_256,user_pw=password,owner_pw=password+"_owner");doc.close()
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Encryption failed: {e}")
    return stream_file(out.getvalue(),"application/pdf",f"protected_{file.filename}")

@app.post("/pdf-to-text")
async def pdf_to_text(file:UploadFile=File(...)):
    data=await read_file(file);lines=[]
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for i,page in enumerate(pdf.pages):
                text=page.extract_text() or ""
                lines.append(f"=== Page {i+1} ===\n{text.strip() or '[No text found]'}")
    except Exception as e:raise HTTPException(500,f"Extraction failed: {e}")
    return stream_file("\n\n".join(lines).encode("utf-8"),"text/plain",f"{stem(file.filename)}.txt")

@app.post("/pdf-to-pptx")
async def pdf_to_pptx(file:UploadFile=File(...)):
    if not _PPTX_OK:raise HTTPException(500,"python-pptx not installed.")
    data=await read_file(file);prs=Presentation()
    prs.slide_width=Inches(13.33);prs.slide_height=Inches(7.5);blank=prs.slide_layouts[6]
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for i,page in enumerate(pdf.pages):
                slide=prs.slides.add_slide(blank);text=page.extract_text() or ""
                lines=[l.strip() for l in text.split("\n") if l.strip()]
                tt=lines[0][:100] if lines else f"Page {i+1}";bt="\n".join(lines[1:]) if len(lines)>1 else ""
                tb=slide.shapes.add_textbox(Inches(0.5),Inches(0.4),Inches(12.3),Inches(1.0))
                tf=tb.text_frame;tf.word_wrap=True;p=tf.paragraphs[0];p.text=tt
                p.runs[0].font.size=PptPt(24);p.runs[0].font.bold=True;p.runs[0].font.color.rgb=PptxRGB(0xFF,0x38,0x14)
                if bt:
                    bb=slide.shapes.add_textbox(Inches(0.5),Inches(1.6),Inches(12.3),Inches(5.5))
                    btf=bb.text_frame;btf.word_wrap=True;bp=btf.paragraphs[0];bp.text=bt[:2000];bp.runs[0].font.size=PptPt(14)
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"PPTX failed: {e}")
    out=io.BytesIO();prs.save(out)
    return stream_file(out.getvalue(),"application/vnd.openxmlformats-officedocument.presentationml.presentation",f"{stem(file.filename)}.pptx")

@app.post("/add-page-numbers")
async def add_page_numbers(file:UploadFile=File(...),position:str=Form("bottom-center"),format:str=Form("number"),start:int=Form(1)):
    data=await read_file(file)
    try:
        doc=open_fitz(data);total=len(doc)
        for i,page in enumerate(doc):
            num=start+i;w,h=page.rect.width,page.rect.height;fs=10
            label=str(num) if format=="number" else f"Page {num} of {total}" if format=="page_of" else f"— {num} —"
            tw=len(label)*fs*0.55;m=20
            pm={"bottom-center":fitz.Point(w/2-tw/2,m),"bottom-right":fitz.Point(w-tw-m,m),
                "bottom-left":fitz.Point(m,m),"top-center":fitz.Point(w/2-tw/2,h-m-fs),
                "top-right":fitz.Point(w-tw-m,h-m-fs),"top-left":fitz.Point(m,h-m-fs)}
            page.insert_text(pm.get(position,fitz.Point(w/2-tw/2,m)),label,fontsize=fs,color=(0.3,0.3,0.3),overlay=True)
        out=io.BytesIO();doc.save(out);doc.close()
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f"Page numbering failed: {e}")
    return stream_file(out.getvalue(),"application/pdf",f"numbered_{file.filename}")

@app.post("/ocr-check")
async def ocr_check(file:UploadFile=File(...)):
    data=await read_file(file);tc=0;tp=0
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            tp=len(pdf.pages)
            for page in pdf.pages[:5]:tc+=len((page.extract_text() or "").strip())
    except Exception as e:raise HTTPException(400,f"Cannot read PDF: {e}")
    avg=tc/max(tp,1)
    return JSONResponse({"is_scanned":avg<50,"avg_chars":round(avg,1),"total_pages":tp})

# SERVE FRONTEND — MUST BE LAST
if STATIC_DIR.exists():
    app.mount("/",StaticFiles(directory=str(STATIC_DIR),html=True),name="static")
