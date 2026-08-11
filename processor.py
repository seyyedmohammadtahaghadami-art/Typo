from pathlib import Path
import re,statistics,math
import fitz,pytesseract
from PIL import Image,ImageEnhance,ImageOps
from pytesseract import Output
from docx import Document
from docx.shared import Pt,Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT

def norm(s):
    s=str(s or "").replace("ي","ی").replace("ى","ی").replace("ك","ک").replace("ـ","").replace("\u200b","")
    s=re.sub(r"\s+"," ",s).strip()
    return re.sub(r"\s+([،؛؟,.!:%٪)\]}])",r"\1",s)

def rtl(s):
    p=len(re.findall(r"[\u0600-\u06ff]",s)); l=len(re.findall(r"[A-Za-z]",s))
    return p>=max(1,l)

def native_words(page):
    out=[]
    for x0,y0,x1,y1,t,*_ in page.get_text("words",sort=False):
        t=norm(t)
        if t: out.append(dict(text=t,x=x0,y=y0,x1=x1,y1=y1,width=x1-x0,height=y1-y0,conf=100,ocr=False))
    return out

def ocr(img):
    img=ImageOps.grayscale(img); img=ImageEnhance.Contrast(img).enhance(1.25)
    d=pytesseract.image_to_data(img,lang="fas+eng",config="--oem 3 --psm 3",output_type=Output.DICT)
    out=[]
    for i,t in enumerate(d["text"]):
        t=norm(t)
        try:c=float(d["conf"][i])
        except:c=0
        if not t or c<20: continue
        x,y,w,h=[int(d[k][i]) for k in ("left","top","width","height")]
        out.append(dict(text=t,x=x,y=y,x1=x+w,y1=y+h,width=w,height=h,conf=c,ocr=True))
    return out

def lines(words):
    words=sorted(words,key=lambda w:(w["y"],w["x"])); ls=[]
    for w in words:
        cy=w["y"]+w["height"]/2; hit=None
        for l in ls:
            if abs(cy-(l["y"]+l["height"]/2))<=max(4,min(w["height"],l["height"])*.65): hit=l; break
        if not hit:
            hit={"x":w["x"],"y":w["y"],"x1":w["x1"],"y1":w["y1"],"height":w["height"],"words":[]}; ls.append(hit)
        hit["words"].append(w); hit["x"]=min(hit["x"],w["x"]);hit["y"]=min(hit["y"],w["y"]);hit["x1"]=max(hit["x1"],w["x1"]);hit["y1"]=max(hit["y1"],w["y1"]);hit["height"]=hit["y1"]-hit["y"]
    out=[]
    for l in sorted(ls,key=lambda z:z["y"]):
        text=norm(" ".join(w["text"] for w in l["words"])); r=rtl(text)
        ws=sorted(l["words"],key=lambda z:z["x"],reverse=r)
        out.append({**l,"words":ws,"text":norm(" ".join(w["text"] for w in ws)),"rtl":r})
    return out

def paragraphs(ls):
    out=[]
    for l in ls:
        if not out: out.append({"lines":[l]}); continue
        p=out[-1]["lines"][-1]; gap=l["y"]-p["y1"]
        if l["rtl"]==p["rtl"] and gap<=max(12,p["height"]*1.55): out[-1]["lines"].append(l)
        else: out.append({"lines":[l]})
    return out

def table_detect(ls):
    rows=[l for l in ls if len(l["words"])>=2]
    if len(rows)<4:return None
    centers=[]
    for l in rows:
        for w in l["words"]:
            c=(w["x"]+w["x1"])/2; q=next((q for q in centers if abs(q[0]-c)<max(12,w["width"]*.45)),None)
            if q:q[1]+=1;q[0]=(q[0]*(q[1]-1)+c)/q[1]
            else:centers.append([c,1])
    stable=[c for c,n in centers if n>=max(3,math.ceil(len(rows)*.6))]
    if len(stable)<2:return None
    data=[]
    for l in rows:
        row=[]
        for c in stable:
            row.append(norm(" ".join(w["text"] for w in l["words"] if abs((w["x"]+w["x1"])/2-c)<max(14,w["width"]*.65))))
        if sum(bool(x) for x in row)>=2:data.append(row)
    fill=sum(bool(x) for r in data for x in r)/max(1,len(data)*len(stable))
    return data if len(data)>=4 and fill>=.58 else None

def cell(cell,text):
    cell.text="";p=cell.paragraphs[0];r=rtl(text);p.alignment=WD_ALIGN_PARAGRAPH.RIGHT if r else WD_ALIGN_PARAGRAPH.LEFT
    run=p.add_run(text);run.font.name="Tahoma" if r else "Arial";run.font.size=Pt(10);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def build(pages,out):
    doc=Document();s=doc.sections[0];s.top_margin=s.bottom_margin=s.left_margin=s.right_margin=Cm(1.6);tables=0
    for pi,page in enumerate(pages):
        if page["table"]:
            t=doc.add_table(rows=len(page["table"]),cols=max(map(len,page["table"])));t.style="Table Grid";t.alignment=WD_TABLE_ALIGNMENT.CENTER
            for i,row in enumerate(page["table"]):
                for j,v in enumerate(row):cell(t.cell(i,j),v)
            doc.add_paragraph();tables+=1
        for para in page["paragraphs"]:
            for l in para["lines"]:
                p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.RIGHT if l["rtl"] else WD_ALIGN_PARAGRAPH.LEFT
                for i,w in enumerate(l["words"]):
                    run=p.add_run(w["text"]+(" " if i<len(l["words"])-1 else ""));run.font.name="Tahoma" if l["rtl"] else "Arial";run.font.size=Pt(max(8,min(36,w["height"]*.72)))
        if pi<len(pages)-1:doc.add_page_break()
    doc.save(out);return tables

def convert_document(src,work):
    pages=[]
    if src.suffix.lower()==".pdf":
        pdf=fitz.open(src)
        for page in pdf:
            ws=native_words(page); used=len(ws)<6 or sum(len(w["text"]) for w in ws)<18
            if used:
                pix=page.get_pixmap(matrix=fitz.Matrix(2.25,2.25),alpha=False)
                ws=ocr(Image.frombytes("RGB",[pix.width,pix.height],pix.samples))
            ls=lines(ws);pages.append({"lines":ls,"paragraphs":paragraphs(ls),"table":table_detect(ls),"ocr":used,"words":len(ws),"conf":statistics.mean([w["conf"] for w in ws]) if ws else 0})
        pdf.close()
    else:
        im=Image.open(src).convert("RGB");ws=ocr(im);ls=lines(ws);pages.append({"lines":ls,"paragraphs":paragraphs(ls),"table":table_detect(ls),"ocr":True,"words":len(ws),"conf":statistics.mean([w["conf"] for w in ws]) if ws else 0})
    out=work/(src.stem+"-TYPO-Real.docx");tc=build(pages,out)
    return type("R",(),{"output":out,"stats":{"pages":len(pages),"words":sum(p["words"] for p in pages),"ocr_pages":sum(p["ocr"] for p in pages),"tables":tc,"confidence":round(statistics.mean([p["conf"] for p in pages]),1) if pages else 0}})()
