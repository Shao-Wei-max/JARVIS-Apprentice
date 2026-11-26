"""Conservative, reproducible DOCX structure metrics; not a semantic quality judge."""
import json,re
from pathlib import Path
from difflib import SequenceMatcher
from docx import Document
from lxml import etree
from zipfile import ZipFile
NS={'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
def stats(path):
    d=Document(path)
    paragraphs=[(p.text.strip(),p.style.name) for p in d.paragraphs if p.text.strip()]
    with ZipFile(path) as z:
        xml=etree.fromstring(z.read('word/document.xml'))
        anchors=xml.xpath('//wp:anchor',namespaces=NS)
        tops=sum(bool(a.xpath('./wp:wrapTopAndBottom',namespaces=NS)) for a in anchors)
        images=len(xml.xpath('//wp:inline | //wp:anchor',namespaces=NS))
    margins=[round(getattr(d.sections[0],k).cm,2) for k in ('top_margin','bottom_margin','left_margin','right_margin')]
    return {'paragraphs':paragraphs,'images':images,'top_bottom_images':tops,'margins_cm':margins}
def evaluate(candidate,reference):
    a,b=stats(candidate),stats(reference)
    # Matching is by normalized text, and may undercount when paragraph segmentation differs.
    norm=lambda s:re.sub(r'\s+','',s)
    refs={}
    for text,style in b['paragraphs']:refs.setdefault(norm(text),[]).append(style)
    matched=correct=0
    for text,style in a['paragraphs']:
        key=norm(text)
        if refs.get(key):
            matched+=1;correct+=style==refs[key].pop(0)
    text_a=''.join(norm(x) for x,_ in a['paragraphs'])
    text_b=''.join(norm(x) for x,_ in b['paragraphs'])
    return {'candidate':str(candidate),'reference':str(reference),'image_count':a['images'],'top_bottom_image_rate':a['top_bottom_images']/a['images'] if a['images'] else None,'reference_top_bottom_image_rate':b['top_bottom_images']/b['images'] if b['images'] else None,'margins_cm':a['margins_cm'],'reference_margins_cm':b['margins_cm'],'margins_match':all(abs(x-y)<=0.02 for x,y in zip(a['margins_cm'],b['margins_cm'])),'exact_text_matched_paragraphs':matched,'exact_text_style_accuracy':correct/matched if matched else None,'text_similarity_proxy':SequenceMatcher(None,text_a,text_b,autojunk=False).ratio(),'warning':'Text similarity is NOT medical content preservation; paragraph style accuracy excludes unmatched paragraphs; inspect images/tables manually.'}
def compare_runs(off,on):
    a,b=stats(Path(off)/'output.docx'),stats(Path(on)/'output.docx')
    logs={k:json.loads((Path(p)/'execution_log.json').read_text(encoding='utf-8')) for k,p in [('off',off),('on',on)]}
    return {'memory_off':{'images':a['images'],'top_bottom_images':a['top_bottom_images'],'margins_cm':a['margins_cm']},'memory_on':{'images':b['images'],'top_bottom_images':b['top_bottom_images'],'margins_cm':b['margins_cm']},'applied_events_on':logs['on'].get('rule_events',[]),'note':'A/B shows memory-attributable structural differences, not overall accuracy.'}

# v0.7: coverage-aware metrics; reference ranges are half-open.
def evaluate_task(candidate, reference, ranges, source_pdf=None, pages=None):
    from collections import Counter, defaultdict, deque
    import unicodedata
    import fitz
    def norm(s):
        s=unicodedata.normalize('NFKC',s)
        s=re.sub(r'^(?:[◎●■]|[一二三四五六七八九十]+[、.]|\d+[.、]|\(\d+\)|[A-Za-z][.、])\s*','',s.strip())
        return re.sub(r'\s+','',s)
    def pars(doc, indices=None):
        p=doc.paragraphs
        return [(norm(p[i].text),p[i].style.name) for i in (range(len(p)) if indices is None else indices) if norm(p[i].text)]
    def ratio(n,d): return n/d if d else None
    a=pars(Document(candidate)); rd=Document(reference)
    b=pars(rd,sorted({i for lo,hi in ranges for i in range(lo,hi)}))
    refs=defaultdict(deque)
    for t,s in b: refs[t].append(s)
    matched=correct=0
    for t,s in a:
        if refs[t]:
            matched+=1; correct+=s==refs[t].popleft()
    d=Document(candidate); margins=[getattr(sec,k).cm for sec in d.sections for k in ('top_margin','bottom_margin','left_margin','right_margin')]
    st=stats(candidate)
    result=dict(reference_paragraphs=len(b),candidate_paragraphs=len(a),matched=matched,style_correct=correct,matched_style_accuracy=ratio(correct,matched),matching_coverage=ratio(matched,len(b)),coverage_adjusted_style_score=ratio(correct,len(b)),margin_compliance=ratio(sum(abs(x-1.27)<=.02 for x in margins),len(margins)),image_wrap_compliance=ratio(st['top_bottom_images'],st['images']),candidate_images=st['images'])
    if source_pdf is not None:
        from jarvis import extract
        blocks=[]; links=[]; images=0
        for lo,hi in pages:
            for page in extract(source_pdf,lo,hi):
                for block in page['blocks']:
                    if block['kind']=='text': blocks.append(norm(block['text']))
                    elif block['bbox'][2]-block['bbox'][0]>=5 and block['bbox'][3]-block['bbox'][1]>=5: images+=1
                links.extend(x['uri'] for x in page['links'])
        src=Counter(''.join(blocks)); dst=Counter(''.join(t for t,s in a))
        actual=Counter(rel.target_ref for rel in d.part.rels.values() if rel.reltype.endswith('/hyperlink'))
        expected=Counter(links)
        result.update(source_character_recall_proxy=ratio(sum((src&dst).values()),sum(src.values())),source_image_regions=images,image_count_recall_proxy=ratio(min(images,st['images']),images),source_unique_links=len(expected),unique_link_recall=ratio(len(set(expected)&set(actual)),len(expected)))
    result.update(correction_count=None,correction_seconds=None,integration_f1=None,warning='Exact normalized paragraph alignment is provisional; character and image count recall are proxies, not semantic or image identity verification. Missing manual labels are null.')
    return result
