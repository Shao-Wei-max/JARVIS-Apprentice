"""Review-first lexical triage, with source provenance; no medical inference."""
from pathlib import Path
import re, unicodedata, hashlib
from datetime import datetime, timezone
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from docx.shared import Pt

def clean(s):
    return re.sub(r'\s+', ' ', s or '').strip()

def norm(s):
    return ''.join(c for c in unicodedata.normalize('NFKC',s).casefold() if c.isalnum())

# Conservative metadata and layout checks; preserve excluded units for audit.
def metadata(t):
    return bool(re.search(r'^[\w.+-]+@[\w.-]+\.[a-z]{2,}$|^(?:tel|fax|email|e-mail)\s*[:：]|^授課資料來源|^copy from\b|^\d{4}\s+.{0,80}\b(?:mosby|ltd|press|publisher)\b|^.{0,35}(?:教授|講師)$',t,re.I))

def pdf_units(page):
    # Preserve visual lines rather than flattening arbitrary PDF blocks.
    lines=[]
    for b in page.get_text('dict')['blocks']:
        if b['type']!=0: continue
        for line in b['lines']:
            t=clean(''.join(s['text'] for s in line['spans']))
            if t: lines.append((line['bbox'],t))
    lines.sort(key=lambda x:(x[0][1],x[0][0]))
    out=[]; i=0
    while i<len(lines):
        box,t=lines[i]; parts=[t]; boxes=[box]; i+=1
        # Only join clear English continuations with close aligned geometry.
        while i<len(lines) and not metadata(t):
            nb,nt=lines[i]; gap=nb[1]-boxes[-1][3]
            if not (0<=gap<=max(20,(boxes[-1][3]-boxes[-1][1])*.8) and abs(nb[0]-boxes[-1][0])<=30 and re.search(r'\b(?:of|or|and|by|into|the|to|for|with|blood)\s*$',parts[-1],re.I) and re.match(r'[a-z]',nt) and not metadata(nt)): break
            parts.append(nt);boxes.append(nb);i+=1
        text=' '.join(parts)
        # Numeric-leading lines are never paired with a following label by inference.
        risk=bool(re.match(r'^[~≈<>≤≥+−-]?\s*\d',text) or re.search(r'\b(?:mL|mg|mmHg|µg|μg)\s*/|\b(?:of|or|and|by|into|the|to|for|with)\s*$',text,re.I))
        out.append((text,boxes,'metadata_filtered' if metadata(text) else 'layout_review' if risk else ''))
    return out

def extract_slides(path, source_name=None):
    path=Path(path); source=source_name or path.name; units=[]
    if path.suffix.lower()=='.pdf':
        import fitz
        with fitz.open(path) as d:
            pages=[(i,pdf_units(p)) for i,p in enumerate(d,1)]
    elif path.suffix.lower()=='.pptx':
        from pptx import Presentation
        pages=[(i,[p.text for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs]+[c.text for sh in s.shapes if sh.has_table for r in sh.table.rows for c in r.cells]) for i,s in enumerate(Presentation(path).slides,1)]
    else: raise ValueError('PDF or PPTX required')
    for page,blocks in pages:
        for j,b in enumerate(blocks,1):
            t,boxes,flag=b if isinstance(b,tuple) else (clean(b),[], 'metadata_filtered' if metadata(clean(b)) else '')
            if not norm(t): continue
            uid=hashlib.sha256(f'{source}|{page}|{j}|{t}'.encode()).hexdigest()[:24]
            units.append(dict(id=uid,slide=page,block=j,text=t,source=source,bboxes=boxes,layout_flag=flag))
    return units

def extract_docx_paragraphs(path):
    d=Document(path)
    return [clean(p.text) for p in d.paragraphs if clean(p.text)] + [clean(p.text) for t in d.tables for r in t.rows for c in r.cells for p in c.paragraphs if clean(p.text)]

def compare_content(slides, previous_paragraphs, threshold=.36):
    if not 0<=threshold<=1: raise ValueError('threshold out of range')
    prev=[p for p in previous_paragraphs if norm(p)]; pn=[norm(p) for p in prev]; out=[]
    if not slides: return out
    sn=[norm(s['text']) for s in slides]
    vectors=None
    if prev:
        try: vectors=TfidfVectorizer(analyzer='char',ngram_range=(2,4)).fit_transform(sn+pn)
        except ValueError: pass
    seen=set()
    for i,s in enumerate(slides):
        n=sn[i]; idx=None; sim=0.; coverage=0.
        if vectors is not None:
            scores=cosine_similarity(vectors[i],vectors[len(slides):])[0]; idx=int(scores.argmax()); sim=float(scores[idx])
        exact=next((j for j,p in enumerate(pn) if len(n)>=12 and n in p),None)
        if exact is not None: idx=exact; coverage=1.
        elif idx is not None:
            grams={n[k:k+3] for k in range(len(n)-2)}; p=pn[idx]; coverage=sum(g in p for g in grams)/max(1,len(grams))
        if s.get('layout_flag'): status=s['layout_flag']
        elif n in seen: status='repeated_slide_text'
        elif len(n)<12: status='short_review'
        elif exact is not None: status='existing_exact'
        elif sim>=threshold or coverage>=.65: status='review_overlap_or_change'
        else: status='potential_new'
        if not s.get('layout_flag'): seen.add(n)
        out.append(dict(s,similarity=round(sim,4),coverage=round(coverage,4),closest_existing=prev[idx] if idx is not None else '',status=status))
    return out

def integrate_selected(base_docx,candidates,selected_ids,output_path,source_name=None):
    selected=set(selected_ids); known={c['id'] for c in candidates}
    if selected-known: raise ValueError('Unknown candidate IDs')
    if Path(base_docx).resolve()==Path(output_path).resolve(): raise ValueError('Cannot overwrite base')
    if any(c['id'] in selected and c['status'] in ('metadata_filtered','layout_review') for c in candidates): raise ValueError('Excluded or layout-risk units cannot be appended directly; verify source manually.')
    doc=Document(base_docx); approved=[c for c in candidates if c['id'] in selected]
    def para(t,bold=False):
        p=doc.add_paragraph(style='Normal'); p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_after=Pt(6); r=p.add_run(t); r.bold=bold; r.font.size=Pt(12)
    if approved:
        doc.add_page_break(); para('今年簡報補充內容（人工確認）',True)
        for c in approved:
            para(c['text']); para(f"來源：{c.get('source',source_name)}，頁碼 {c['slide']}，區塊 {c.get('block',1)}")
    output=Path(output_path); output.parent.mkdir(parents=True,exist_ok=True); doc.save(output)
    return dict(output=str(output),approved_units=len(approved),approved=approved,timestamp_utc=datetime.now(timezone.utc).isoformat(),method='human-approved append-only addendum',limitations=['Lexical triage only','No OCR or image interpretation','No automatic replacement or chapter placement'])
