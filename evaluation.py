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
