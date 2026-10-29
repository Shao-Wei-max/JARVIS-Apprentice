"""JARVIS Apprentice v0.5: controlled memory ablation and document evaluation."""
import json
from pathlib import Path
import streamlit as st
from jarvis import generate,compare,propose,propose_v05,approve,db_connect
from evaluation import evaluate, compare_runs
st.set_page_config(page_title='JARVIS Apprentice v0.5',layout='wide')
st.title('JARVIS Apprentice · Medical Conote v0.5')
st.caption('Execute → Correct → Learn → Reuse | Memory ablation and quantitative checks')
root=Path('workspace');root.mkdir(exist_ok=True)
db=root/'memory.db'
t1,t2,t3,t4=st.tabs(['1. Execute & A/B','2. Correct & Learn','3. Memory','4. Evaluate'])
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
