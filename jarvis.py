"""JARVIS Apprentice MVP: PDF to template DOCX, structural diff, reviewed rule memory.
Python 3.11. Prototype, not clinical or production document processing.
"""
import argparse, json, re, sqlite3, io, hashlib, copy
from pathlib import Path
from datetime import datetime, timezone
import fitz
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

STYLE_NAMES=['1. 主標題','2. 副標題','3. 第一內文','4. 第二內文','5. 第三內文','6. 第四內文','7. 第五內文']
DEFAULTS=[{'id':'G001','action':'set_margins','value':1.27,'scope':'medical_conote','source':'user_confirmed'}, {'id':'G002','action':'vertical_images','value':True,'scope':'medical_conote','source':'user_confirmed'}, {'id':'G003','action':'preserve_links','value':True,'scope':'medical_conote','source':'user_confirmed'}]

def db_connect(path):
    c=sqlite3.connect(str(path)); c.execute('CREATE TABLE IF NOT EXISTS rules(id INTEGER PRIMARY KEY, trigger TEXT NOT NULL, action TEXT NOT NULL, value TEXT NOT NULL, scope TEXT NOT NULL, evidence TEXT, approved INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL)');c.commit();return c

def rules_for(c,task):
    rows=c.execute('SELECT id,trigger,action,value,scope,evidence FROM rules WHERE approved=1 AND (scope=? OR scope=?)', (task,'all')).fetchall()
    return [{'id':x[0],'trigger':x[1],'action':x[2],'value':json.loads(x[3]),'scope':x[4],'evidence':x[5]} for x in rows]

def choose_style(t):
    s=t.strip()
    if re.match(r'^(?:◎|●|■)',s):return STYLE_NAMES[0]
    if re.match(r'^[一二三四五六七八九十]+[、.]',s):return STYLE_NAMES[1]
    if re.match(r'^\d+[.、]',s):return STYLE_NAMES[2]
    if re.match(r'^\(\d+\)',s):return STYLE_NAMES[3]
    if re.match(r'^[A-Z][.、]',s):return STYLE_NAMES[4]
    if re.match(r'^[a-z][.、]',s):return STYLE_NAMES[5]
    if re.match(r'^(?:i|ii|iii|iv|v)[.)、\s]',s):return STYLE_NAMES[6]
    return STYLE_NAMES[2]

def hyperlink(p, text, url):
    rel=p.part.relate_to(url,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',is_external=True)
    h=OxmlElement('w:hyperlink');h.set(qn('r:id'),rel)
    run=OxmlElement('w:r');pr=OxmlElement('w:rPr');color=OxmlElement('w:color');color.set(qn('w:val'),'0563C1');pr.append(color);run.append(pr)
    tx=OxmlElement('w:t');tx.text=text;run.append(tx);h.append(run);p._p.append(h)

def set_top_bottom_wrap(inline):
    """Create a schema-ordered floating drawing preserving image relationship IDs."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    anchor=parse_xml('<wp:anchor %s distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0"/>' % nsdecls('wp'))
    simple=OxmlElement('wp:simplePos');simple.set('x','0');simple.set('y','0');anchor.append(simple)
    h=OxmlElement('wp:positionH');h.set('relativeFrom','column');align=OxmlElement('wp:align');align.text='center';h.append(align);anchor.append(h)
    v=OxmlElement('wp:positionV');v.set('relativeFrom','paragraph');off=OxmlElement('wp:posOffset');off.text='0';v.append(off);anchor.append(v)
    # Required order: simplePos, positionH, positionV, extent, effectExtent?, wrap*, docPr, cNvGraphicFramePr?, graphic.
    for tag in ('wp:extent','wp:effectExtent'):
        el=inline.find(qn(tag))
        if el is not None:anchor.append(copy.deepcopy(el))
    anchor.append(OxmlElement('wp:wrapTopAndBottom'))
    for tag in ('wp:docPr','wp:cNvGraphicFramePr','a:graphic'):
        el=inline.find(qn(tag))
        if el is not None:anchor.append(copy.deepcopy(el))
    inline.getparent().replace(inline,anchor)

def validate_docx(path,expected_images=None,expect_top_bottom=False):
    from zipfile import ZipFile
    from lxml import etree
    errors=[]
    with ZipFile(path) as z:
        bad=z.testzip()
        if bad:errors.append('Corrupt zip member: '+bad)
        for name in z.namelist():
            if name.endswith(('.xml','.rels')):
                try:etree.fromstring(z.read(name))
                except Exception as e:errors.append(f'Invalid XML {name}: {e}')
        names=set(z.namelist())
        image_files=[n for n in names if n.startswith('word/media/')]
        root=etree.fromstring(z.read('word/document.xml'))
        ns={'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        drawings=root.xpath('//wp:inline|//wp:anchor',namespaces=ns)
        if expected_images is not None and len(drawings)!=expected_images:errors.append(f'Drawing count {len(drawings)} != expected {expected_images}')
        if expect_top_bottom and any(x.tag.endswith('inline') or not x.xpath('./wp:wrapTopAndBottom',namespaces=ns) for x in drawings):errors.append('Missing Top and Bottom on one or more drawings')
        for x in root.xpath('//wp:anchor',namespaces=ns):
            children=[etree.QName(y).localname for y in x]
            required=['simplePos','positionH','positionV','extent','wrapTopAndBottom','docPr','graphic']
            pos=[children.index(n) for n in required if n in children]
            if len(pos)!=len(required) or pos!=sorted(pos):errors.append('Invalid anchor element order: '+str(children))
        relroot=etree.fromstring(z.read('word/_rels/document.xml.rels'))
        rels={x.get('Id'):x.get('Target') for x in relroot}
        for rid in root.xpath('//a:blip/@r:embed',namespaces=ns):
            target=rels.get(rid)
            if not target:errors.append('Missing image relationship '+rid)
            elif 'word/'+target.lstrip('/') not in names and target.lstrip('/') not in names:errors.append('Missing image target '+target)
    return {'passed':not errors,'errors':errors,'drawing_count':len(drawings),'embedded_media_count':len(image_files),'image_layout':image_layout_counts(path)}

def image_layout_counts(path):
    from zipfile import ZipFile
    from lxml import etree
    with ZipFile(path) as z:
        root=etree.fromstring(z.read('word/document.xml'))
    ns={'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
    return {'inline':len(root.xpath('//wp:inline',namespaces=ns)),
            'top_bottom':len(root.xpath('//wp:anchor/wp:wrapTopAndBottom',namespaces=ns)),
            'other_anchor':len(root.xpath('//wp:anchor[not(wp:wrapTopAndBottom)]',namespaces=ns))}

def explicit_level(text):
    t=text.strip()
    if re.match(r'^(?:◎|●|■)',t):return 0
    if re.match(r'^[一二三四五六七八九十]+[、.]',t):return 1
    if re.match(r'^\d+[.、]',t):return 2
    if re.match(r'^\(\d+\)',t):return 3
    if re.match(r'^[A-Z][.、]',t):return 4
    if re.match(r'^[a-z][.、]',t):return 5
    if re.match(r'^(?:i|ii|iii|iv|v)[.)、\s]',t):return 6
    return None

def extract(pdf_path,start,end):
    pdf=fitz.open(pdf_path)
    if start<1 or end>len(pdf) or end<start:raise ValueError(f'Page range must be within 1..{len(pdf)}')
    pages=[]
    for page_no in range(start,end+1):
        page=pdf[page_no-1];blocks=[]
        for b in page.get_text('dict')['blocks']:
            if b['type']==0:
                lines=[''.join(span['text'] for span in line['spans']) for line in b.get('lines',[])]
                content=''.join(lines).strip()
                if content:blocks.append({'kind':'text','text':content,'bbox':list(b['bbox'])})
            elif b['type']==1:blocks.append({'kind':'image','bbox':list(b['bbox'])})
        blocks.sort(key=lambda x:(round(x['bbox'][1]/6),x['bbox'][0]))
        links=[{'uri':x.get('uri'),'bbox':list(x['from'])} for x in page.get_links() if x.get('uri')]
        pages.append({'page':page_no,'blocks':blocks,'links':links})
    pdf.close();return pages

def is_legacy_header(text, bbox, page_height):
    """Conservative: only a standalone prior-year conote label near page top."""
    t=' '.join(text.split())
    return bool(re.fullmatch(r'Conote for KMU Medicine\s+\d{2,3}(?:\s+\d{1,3})?',t,re.I) and bbox[1] < page_height*0.23)

def likely_continuation(text, bbox, previous_bbox, previous_style):
    """Spatial heuristic; never joins or rewrites source content."""
    if previous_bbox is None or previous_style is None or explicit_level(text) is not None:
        return False
    gap=bbox[1]-previous_bbox[3]
    return -3 <= gap <= 18 and abs(bbox[0]-previous_bbox[0]) <= 45

def add_candidate_rule(db_path,trigger,action,value,evidence):
    c=db_connect(db_path); encoded=json.dumps(value)
    row=c.execute('SELECT id FROM rules WHERE trigger=? AND action=? AND value=? AND scope=?',
                  (trigger,action,encoded,'medical_conote')).fetchone()
    if row: result=row[0]
    else:
        result=c.execute('INSERT INTO rules(trigger,action,value,scope,evidence,approved,created_at) VALUES (?,?,?,?,?,0,?)',
            (trigger,action,encoded,'medical_conote',evidence,datetime.now(timezone.utc).isoformat())).lastrowid
    c.commit();c.close();return result

def propose_v05(diff_path,db_path):
    """Only propose; user must approve in Memory tab."""
    data=json.loads(Path(diff_path).read_text(encoding='utf-8'))
    events=data.get('events',[]); ids=[]
    removed=' '.join(' '.join(e.get('before',[])) for e in events if e.get('type')=='content_or_structure_change')
    if re.search(r'Conote for KMU Medicine\s+\d{2,3}',removed,re.I):
        ids.append(add_candidate_rule(db_path,'legacy_pdf_header','remove_legacy_header',True,
                     f"{data.get('after','corrected.docx')}: legacy header removed; verify PDF location"))
    if any(e.get('type')=='style_change' for e in events):
        ids.append(add_candidate_rule(db_path,'probable_pdf_line_continuation','inherit_continuation_style_v2',True,
                     f"{data.get('after','corrected.docx')}: style changes; spatial heuristic requires review"))
    return ids

def generate(pdf_path,template,start,end,out_dir,db_path=None,memory_enabled=True):
    out_dir=Path(out_dir);out_dir.mkdir(parents=True,exist_ok=True)
    pages=extract(pdf_path,start,end)
    d=Document(template)
    # Preserve template section, header/footer, styles, numbering; clear body only.
    for el in list(d.element.body):
        if el.tag!=qn('w:sectPr'):d.element.body.remove(el)
    memory=[]
    if db_path and memory_enabled:
        c=db_connect(db_path);memory=rules_for(c,'medical_conote');c.close()
    margin=1.27
    for r in memory:
        if r['action']=='set_margins_cm' and isinstance(r['value'],(int,float)) and 0.8<=r['value']<=3:margin=r['value']
    for sec in d.sections:sec.top_margin=sec.bottom_margin=sec.left_margin=sec.right_margin=Cm(margin)
    styles=set(x.name for x in d.styles)
    learned_wrap=any(r['action']=='set_image_wrap_top_bottom' and r['value'] is True for r in memory)
    learned_continuation=any(r['action']=='inherit_continuation_style' and r['value'] is True for r in memory)
    learned_header=any(r['action']=='remove_legacy_header' and r['value'] is True for r in memory)
    learned_continuation_v2=any(r['action']=='inherit_continuation_style_v2' and r['value'] is True for r in memory)
    trace=[];warnings=[];links_found=0;rule_events=[]
    if memory_enabled:
        for r in memory:
            if r['action']=='set_margins_cm':rule_events.append({'rule_id':r['id'],'action':r['action'],'target':'document','status':'applied' if margin==r['value'] else 'overridden','value':margin})
    pdf=fitz.open(pdf_path)
    for ix,page_data in enumerate(pages):
        if ix:d.add_page_break()
        page=pdf[page_data['page']-1]
        previous_style=None
        previous_bbox=None
        for bi,b in enumerate(page_data['blocks']):
            if b['kind']=='text':
                if learned_header and is_legacy_header(b['text'],b['bbox'],page.rect.height):
                    for r in memory:
                        if r['action']=='remove_legacy_header' and r['value'] is True:
                            rule_events.append({'rule_id':r['id'],'action':r['action'],'target':f"page:{page_data['page']}/block:{bi}",'status':'applied'})
                    trace.append({'page':page_data['page'],'index':bi,'kind':'filtered_legacy_header','source_text':b['text']})
                    continue
                style=choose_style(b['text'])
                if learned_continuation_v2 and likely_continuation(b['text'],b['bbox'],previous_bbox,previous_style):
                    style=previous_style
                    for r in memory:
                        if r['action']=='inherit_continuation_style_v2' and r['value'] is True:
                            rule_events.append({'rule_id':r['id'],'action':r['action'],'target':f"page:{page_data['page']}/block:{bi}",'status':'applied'})
                if learned_continuation and not learned_continuation_v2 and explicit_level(b['text']) is None and previous_style and previous_bbox:
                    # Conservative: only short vertically adjacent unnumbered blocks.
                    gap=b['bbox'][1]-previous_bbox[3]
                    if -2 <= gap <= 14 and abs(b['bbox'][0]-previous_bbox[0]) <= 18:
                        style=previous_style
                        for r in memory:
                            if r['action']=='inherit_continuation_style' and r['value'] is True:rule_events.append({'rule_id':r['id'],'action':r['action'],'target':f"page:{page_data['page']}/block:{bi}",'status':'applied'})
                p=d.add_paragraph(style=style if style in styles else 'Normal')
                display_text=b['text']
                # Word template supplies automatic list numbers; remove only a matching visible prefix.
                # Source text remains untouched in source_map.json.
                patterns={STYLE_NAMES[2]:r'^\d+[.、]\s*',STYLE_NAMES[3]:r'^\(\d+\)\s*',STYLE_NAMES[4]:r'^[A-Z][.、]\s*',STYLE_NAMES[5]:r'^[a-z][.、]\s*',STYLE_NAMES[6]:r'^(?:i|ii|iii|iv|v)[.)、]\s*'}
                if style in patterns and p.style.name==style:
                    display_text=re.sub(patterns[style],'',display_text,count=1)
                p.add_run(display_text)
                # Do not invent or rewrite text. Preserve PDF block text verbatim.
                previous_style=style
                previous_bbox=b['bbox']
                trace.append({'page':page_data['page'],'index':bi,'kind':'text','bbox':b['bbox'],'style':p.style.name,'source_text':b['text']})
            else:
                rect=fitz.Rect(b['bbox'])
                if rect.width<5 or rect.height<5:continue
                try:
                    pix=page.get_pixmap(matrix=fitz.Matrix(1.7,1.7),clip=rect,alpha=False)
                    p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                    shape=p.add_run().add_picture(io.BytesIO(pix.tobytes('png')),width=Cm(min(15.5,rect.width*2.54/72)))
                    if learned_wrap:
                        set_top_bottom_wrap(shape._inline)
                        for r in memory:
                            if r['action']=='set_image_wrap_top_bottom' and r['value'] is True:rule_events.append({'rule_id':r['id'],'action':r['action'],'target':f"page:{page_data['page']}/block:{bi}",'status':'applied'})
                    trace.append({'page':page_data['page'],'index':bi,'kind':'image','bbox':b['bbox'],'wrap':'top_bottom' if learned_wrap else 'inline'})
                except Exception as e:warnings.append(f'Image on page {page_data["page"]}: {e}')
        for link in page_data['links']:
            # PDF URI annotations are preserved as clickable appendix links, not silently dropped.
            p=d.add_paragraph(style='Normal');hyperlink(p,link['uri'],link['uri']);links_found+=1
            trace.append({'page':page_data['page'],'kind':'hyperlink','uri':link['uri'],'bbox':link['bbox']})
    pdf.close()
    docx_path=out_dir/'output.docx';d.save(docx_path)
    integrity=validate_docx(docx_path,expected_images=sum(x['kind']=='image' for x in trace),expect_top_bottom=learned_wrap)
    if not integrity['passed']:
        raise RuntimeError('DOCX validation failed: '+str(integrity['errors']))
    report={'docx_integrity':integrity,'status':'prototype','source_pdf':Path(pdf_path).name,'pages':[start,end],'margins_cm':margin,'source_text_blocks':sum(x['kind']=='text' for x in trace),'inserted_image_regions':sum(x['kind']=='image' for x in trace),'hyperlinks':links_found,'memory_enabled':memory_enabled,'memory_rules_applied':sorted(set(e['rule_id'] for e in rule_events if e['status']=='applied')),'warnings':warnings,'limitations':['PDF block order and paragraph boundaries are heuristic','Embedded image text and complex tables are not fully editable','Hyperlinks are placed in a per-page list, not at exact original text positions','Content fidelity requires human review']}
    for filename,obj in [('source_map.json',trace),('validation_report.json',report),('execution_log.json',{'fixed_rules':DEFAULTS,'memory_retrieved':memory,'rule_events':rule_events,'timestamp':datetime.now(timezone.utc).isoformat()})]:
        (out_dir/filename).write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding='utf-8')
    return report

def snapshot(path):
    d=Document(path);pars=[{'text':p.text,'style':p.style.name} for p in d.paragraphs if p.text.strip()]
    margins=[round(getattr(d.sections[0],v).cm,2) for v in ('top_margin','bottom_margin','left_margin','right_margin')]
    return {'paragraphs':pars,'margins':margins,'image_layout':image_layout_counts(path),'links':len(d.part.rels)}

def compare(before,after,out_path):
    import difflib
    a=snapshot(before);b=snapshot(after)
    matcher=difflib.SequenceMatcher(None,[x['text'] for x in a['paragraphs']],[x['text'] for x in b['paragraphs']],autojunk=False)
    events=[]
    for tag,i,j,k,l in matcher.get_opcodes():
        if tag=='equal':
            for offset in range(j-i):
                x=a['paragraphs'][i+offset];y=b['paragraphs'][k+offset]
                if x['style']!=y['style']:events.append({'type':'style_change','before':x['style'],'after':y['style'],'text':x['text'][:140]})
        else:events.append({'type':'content_or_structure_change','operation':tag,'before':[x['text'] for x in a['paragraphs'][i:j]],'after':[x['text'] for x in b['paragraphs'][k:l]],'requires_review':True})
    if a['margins']!=b['margins']:events.append({'type':'margin_change','before':a['margins'],'after':b['margins']})
    if a['image_layout']!=b['image_layout']:
        events.append({'type':'image_wrap_change','before':a['image_layout'],'after':b['image_layout'],'requires_review':True})
    result={'before':Path(before).name,'after':Path(after).name,'events':events,'note':'Structural differences are observations, not automatically approved preferences.'}
    Path(out_path).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');return result

def propose(diff_path,db_path):
    data=json.loads(Path(diff_path).read_text(encoding='utf-8'));c=db_connect(db_path);ids=[]
    events=data['events']
    proposals=[]
    for e in events:
        if e['type']=='margin_change' and len(set(e['after']))==1:
            val=e['after'][0]
            if 0.8<=val<=3:proposals.append(('new_document','set_margins_cm',val,'Margins corrected to uniform value'))
        if e['type']=='image_wrap_change' and e['after']['top_bottom']>e['before']['top_bottom'] and e['after']['top_bottom']>0:
            proposals.append(('new_image','set_image_wrap_top_bottom',True,'Human changed image wrapping to Word Top and Bottom'))
    style_events=[e for e in events if e['type']=='style_change']
    # Only suggest a conservative contextual rule, not a global style replacement.
    if len(style_events)>=2 and any(e['before']!=e['after'] for e in style_events):
        proposals.append(('unnumbered_adjacent_block','inherit_continuation_style',True,'Multiple paragraph hierarchy corrections; heuristic requires review'))
    for trigger,action,value,description in proposals:
        encoded=json.dumps(value)
        existing=c.execute('SELECT id FROM rules WHERE trigger=? AND action=? AND value=? AND scope=?',(trigger,action,encoded,'medical_conote')).fetchone()
        if existing:
            ids.append(existing[0]);continue
        cur=c.execute('INSERT INTO rules(trigger,action,value,scope,evidence,approved,created_at) VALUES (?,?,?,?,?,?,?)',
                      (trigger,action,encoded,'medical_conote',f'{data["after"]}: {description}',0,datetime.now(timezone.utc).isoformat()))
        ids.append(cur.lastrowid)
    c.commit();c.close();return ids

def approve(db_path,id):
    c=db_connect(db_path);c.execute('UPDATE rules SET approved=1 WHERE id=?',(id,));c.commit();c.close()

def main():
    ap=argparse.ArgumentParser();sub=ap.add_subparsers(dest='cmd',required=True)
    g=sub.add_parser('generate');g.add_argument('--pdf',required=True);g.add_argument('--template',required=True);g.add_argument('--start',type=int,required=True);g.add_argument('--end',type=int,required=True);g.add_argument('--out',required=True);g.add_argument('--db')
    x=sub.add_parser('compare');x.add_argument('--before',required=True);x.add_argument('--after',required=True);x.add_argument('--out',required=True)
    p=sub.add_parser('propose');p.add_argument('--diff',required=True);p.add_argument('--db',required=True)
    a=sub.add_parser('approve');a.add_argument('--db',required=True);a.add_argument('--id',type=int,required=True)
    args=ap.parse_args()
    if args.cmd=='generate':print(json.dumps(generate(args.pdf,args.template,args.start,args.end,args.out,args.db),ensure_ascii=False,indent=2))
    elif args.cmd=='compare':print(json.dumps(compare(args.before,args.after,args.out),ensure_ascii=False,indent=2)[:2000])
    elif args.cmd=='propose':print('Candidate rule IDs:',propose(args.diff,args.db))
    elif args.cmd=='approve':approve(args.db,args.id);print('Rule approved:',args.id)
if __name__=='__main__':main()
