"""JARVIS Apprentice v0.9.1: controlled memory ablation and document evaluation."""
import json
from pathlib import Path
import streamlit as st
from jarvis import generate,compare,propose,propose_v05,approve,db_connect
from evaluation import evaluate, compare_runs
from content_comparator import extract_slides, extract_docx_paragraphs, compare_content, integrate_selected
st.set_page_config(page_title='JARVIS Apprentice v0.9.1',layout='wide')
st.title('JARVIS Apprentice · Medical Conote v0.9.1')
st.caption('Execute → Correct → Learn → Reuse | Memory ablation and quantitative checks')
root=Path(__file__).resolve().parent/'workspace';root.mkdir(exist_ok=True)
db=root/'memory.db'
t1,t2,t3,t4,t5=st.tabs(['1. Execute & A/B','2. Correct & Learn','3. Memory','4. Evaluate','5. This Year Slides'])
with t1:
    pdf=st.file_uploader('Previous-year PDF',type='pdf');template=st.file_uploader('Word template',type='docx')
    a,b=st.columns(2);start=a.number_input('Start page',min_value=1,value=1);end=b.number_input('End page',min_value=1,value=2)
    mode=st.radio('Execution mode',['A/B: Memory OFF vs ON','Memory ON only','Memory OFF only'])
    if st.button('Generate DOCX',type='primary'):
        if not pdf or not template:st.error('Upload both files.')
        elif end<start:st.error('End page must be >= start page.')
        else:
            (root/'source.pdf').write_bytes(pdf.getvalue());(root/'template.docx').write_bytes(template.getvalue())
            try:
                modes=([False,True] if mode.startswith('A/B') else [mode=='Memory ON only'])
                for enabled in modes:
                    tag='memory_on' if enabled else 'memory_off'
                    out=root/tag
                    report=generate(root/'source.pdf',root/'template.docx',int(start),int(end),out,db,memory_enabled=enabled)
                    st.subheader(tag);st.json(report)
                    for name in ('output.docx','source_map.json','validation_report.json','execution_log.json'):
                        f=out/name;st.download_button(f'Download {tag}/{name}',f.read_bytes(),file_name=f'{tag}_{name}',key=f'{tag}_{name}')
                if len(modes)==2:
                    summary=compare_runs(root/'memory_off',root/'memory_on')
                    st.subheader('A/B structural comparison');st.json(summary)
                    st.download_button('Download A/B summary',json.dumps(summary,ensure_ascii=False,indent=2),file_name='ab_summary.json')
            except Exception as e:st.exception(e)
with t2:
    original=st.file_uploader('Agent output DOCX',type='docx',key='before');corrected=st.file_uploader('Human-corrected DOCX',type='docx',key='after')
    if st.button('Analyze differences'):
        if original and corrected:
            (root/'before.docx').write_bytes(original.getvalue());(root/'after.docx').write_bytes(corrected.getvalue())
            st.json(compare(root/'before.docx',root/'after.docx',root/'diff.json'))
        else:st.error('Upload both DOCX files.')
    if (root/'diff.json').exists() and st.button('Propose candidate rules'):
        st.write('Proposed rule IDs:',propose(root/'diff.json',db) + propose_v05(root/'diff.json',db))
with t3:
    c=db_connect(db);rows=c.execute('SELECT id,trigger,action,value,scope,evidence,approved FROM rules ORDER BY id').fetchall();c.close()
    if not rows:st.info('No rules yet. Use Correct & Learn.')
    for r in rows:
        st.json(dict(zip(('id','trigger','action','value','scope','evidence','approved'),(*r[:6],bool(r[6])))))
        if not r[6] and st.button(f'Approve rule {r[0]}',key=f'approve_{r[0]}'):
            approve(db,r[0]);st.rerun()
with t4:
    st.write('Upload a reference human-corrected DOCX to evaluate the latest generated output. Results are structural proxies, not full semantic accuracy.')
    reference=st.file_uploader('Reference DOCX',type='docx',key='reference')
    if reference and st.button('Evaluate latest A/B outputs'):
        path=root/'reference.docx';path.write_bytes(reference.getvalue())
        for tag in ('memory_off','memory_on'):
            candidate=root/tag/'output.docx'
            if candidate.exists():
                result=evaluate(candidate,path)
                st.subheader(tag);st.json(result)
                st.download_button(f'Download {tag} metrics',json.dumps(result,ensure_ascii=False,indent=2),file_name=f'{tag}_metrics.json',key=f'metric_{tag}')

with t5:
    import hashlib
    from collections import Counter
    st.subheader('今年簡報：逐區塊比對與人工確認')
    base=st.file_uploader('去年共筆生成 Word（Memory ON）',type='docx',key='v091base')
    uploads=st.file_uploader('今年簡報（可多選）',type=['pdf','pptx'],accept_multiple_files=True,key='v091slides')
    threshold=st.slider('文字相似度門檻',0.1,0.9,0.36,0.02)
    signature=hashlib.sha256(repr([(f.name,hashlib.sha256(f.getvalue()).hexdigest()) for f in ([base] if base else [])+list(uploads)]+[threshold]).encode()).hexdigest()
    if st.button('分析今年簡報'):
        if not base or not uploads: st.error('請上傳 Word 及簡報')
        else:
            try:
                work=root/'content_v091'/signature; work.mkdir(parents=True,exist_ok=True)
                bp=work/'base.docx'; bp.write_bytes(base.getvalue()); units=[]
                for i,f in enumerate(uploads):
                    path=work/(str(i)+Path(f.name).suffix.lower()); path.write_bytes(f.getvalue()); units.extend(extract_slides(path,f'{i+1}: {f.name}'))
                st.session_state['v091_result']=(signature,str(bp),compare_content(units,extract_docx_paragraphs(bp),threshold))
            except Exception as e: st.exception(e)
    result=st.session_state.get('v091_result')
    if result and result[0]==signature:
        _,bp,candidates=result
        st.write(dict(Counter(c['status'] for c in candidates)))
        st.caption('僅文字比對，不代表真正新增或醫學正確；所有選項預設不批准。')
        st.download_button('下載候選 JSON',json.dumps(candidates,ensure_ascii=False,indent=2),file_name='candidates.json')
        statuses=st.multiselect('顯示分類',sorted({c['status'] for c in candidates}),default=[x for x in ['potential_new','review_overlap_or_change'] if any(c['status']==x for c in candidates)])
        visible=[c for c in candidates if c['status'] in statuses]
        page=st.number_input('候選清單頁（每頁20筆）',min_value=1,max_value=max(1,(len(visible)+19)//20),value=1)
        for c in visible[(page-1)*20:page*20]:
            with st.expander(f"{c['source']} / p.{c['slide']} / {c['block']} / {c['status']}"):
                st.write(c['text']); st.caption(f"原始文字位置：{c.get('bboxes', [])}"); st.write('最相近舊內容：',c['closest_existing']); st.write('相似度',c['similarity'],'涵蓋率',c['coverage'])
                key='v091_'+signature+c['id']; store='v091_selected_'+signature
                if store not in st.session_state: st.session_state[store]=set()
                checked=st.checkbox('批准加入',value=c['id'] in st.session_state[store],key=key,disabled=c['status'] in ('metadata_filtered','layout_review'))
                if c['status'] in ('metadata_filtered','layout_review'): st.caption('不可直接加入；請核對原始投影片。')
                if checked: st.session_state[store].add(c['id'])
                else: st.session_state[store].discard(c['id'])
        selected=st.session_state.get('v091_selected_'+signature,set()); st.write('已批准區塊數',len(selected))
        if st.button('整合已批准內容',disabled=not selected):
            output=Path(bp).parent/'updated_conote.docx'; report=integrate_selected(bp,candidates,selected,output)
            st.session_state['v091_output']=(signature,output.read_bytes(),json.dumps(report,ensure_ascii=False,indent=2))
        saved=st.session_state.get('v091_output')
        if saved and saved[0]==signature:
            st.download_button('下載新版 Word',saved[1],file_name='updated_conote.docx'); st.download_button('下載整合紀錄',saved[2],file_name='integration_report.json')
