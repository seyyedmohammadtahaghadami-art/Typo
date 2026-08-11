import os,re,shutil,uuid
from pathlib import Path
from fastapi import FastAPI,File,HTTPException,UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from .processor import convert_document

ROOT=Path("/app/data"); ROOT.mkdir(parents=True,exist_ok=True)
MAX_MB=int(os.getenv("MAX_UPLOAD_MB","75"))
app=FastAPI(title="TYPO Real Document Engine",version="2.0.0")
app.add_middleware(CORSMiddleware,allow_origins=[x.strip() for x in os.getenv("CORS_ORIGINS","*").split(",")],allow_methods=["GET","POST"],allow_headers=["*"])

@app.get("/health")
def health(): return {"ok":True,"service":"TYPO Real","version":app.version}

@app.post("/api/convert")
async def convert(file:UploadFile=File(...)):
    original=Path(file.filename or "document").name
    ext=Path(original).suffix.lower()
    if ext not in {".pdf",".png",".jpg",".jpeg"}: raise HTTPException(400,"Supported: PDF, PNG, JPG, JPEG")
    job=uuid.uuid4().hex; work=ROOT/job; work.mkdir()
    source=work/("input"+ext)
    try:
        total=0
        with source.open("wb") as out:
            while True:
                chunk=await file.read(1024*1024)
                if not chunk: break
                total+=len(chunk)
                if total>MAX_MB*1024*1024: raise HTTPException(413,f"Maximum file size is {MAX_MB} MB")
                out.write(chunk)
        result=convert_document(source,work)
        return {"ok":True,"job_id":job,"filename":result.output.name,"download_url":f"/api/download/{job}/{result.output.name}",**result.stats}
    except HTTPException:
        shutil.rmtree(work,ignore_errors=True); raise
    except Exception as e:
        shutil.rmtree(work,ignore_errors=True)
        raise HTTPException(500,f"Conversion failed: {type(e).__name__}: {e}")

@app.get("/api/download/{job}/{filename}")
def download(job:str,filename:str):
    if not re.fullmatch(r"[a-f0-9]{32}",job): raise HTTPException(400,"Invalid job")
    safe=Path(filename).name
    path=ROOT/job/safe
    if safe!=filename or not safe.endswith(".docx") or not path.exists(): raise HTTPException(404,"Output not found")
    return FileResponse(path,media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",filename=safe)
