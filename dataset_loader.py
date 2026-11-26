import json
from pathlib import Path
import fitz
from docx import Document

def load_tasks(manifest, data_root):
    root=Path(data_root); data=json.loads(Path(manifest).read_text(encoding='utf-8')); seen=set(); topics={}
    for t in data['tasks']:
        assert t['task_id'] not in seen; seen.add(t['task_id'])
        assert topics.setdefault(t['topic'],t['split'])==t['split']
        t=dict(t)
        for k in ('previous_pdf','slides_pdf','reference_docx','template_docx'):
            p=root/t[k]
            if not p.exists():
                matches=list(root.rglob(t[k]));
                if len(matches)!=1: raise FileNotFoundError(t[k])
                p=matches[0]
            t[k]=str(p.resolve())
        for key,field in [('previous_pdf','previous_page_ranges'),('slides_pdf','slide_page_ranges')]:
            with fitz.open(t[key]) as d:
                for lo,hi in t[field]:
                    if not 1<=lo<=hi<=len(d): raise ValueError((t['task_id'],field,lo,hi))
        n=len(Document(t['reference_docx']).paragraphs)
        for lo,hi in t['reference_paragraph_ranges']:
            if not 0<=lo<hi<=n: raise ValueError((t['task_id'],lo,hi,n))
        yield t
