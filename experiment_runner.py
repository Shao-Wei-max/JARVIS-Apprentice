import argparse,csv,json,sqlite3,time,tempfile
from pathlib import Path
from contextlib import closing
from dataset_loader import load_tasks
from jarvis import generate,db_connect
from evaluation import evaluate_task
from docx import Document
from docx.oxml.ns import qn
import fitz

def run(manifest,data_root,out,split='validation',memory=None,limit=None):
    out=Path(out);out.mkdir(parents=True,exist_ok=True)
    tasks=[t for t in load_tasks(manifest,data_root) if t['split']==split]
    if limit: tasks=tasks[:limit]
    rows=[]
    with tempfile.TemporaryDirectory() as tmp:
        db=Path(tmp)/'frozen.db'
        if memory:
            with closing(sqlite3.connect(Path(memory).resolve().as_uri()+'?mode=ro',uri=True)) as src, closing(sqlite3.connect(db)) as dst: src.backup(dst)
        else: db_connect(db).close()
        with closing(sqlite3.connect(db)) as c: rule_count=c.execute('SELECT count(*) FROM rules WHERE approved=1').fetchone()[0]
        for t in tasks:
            for enabled in (False,True):
                folder=out/t['task_id']/('on' if enabled else 'off');folder.mkdir(parents=True,exist_ok=True)
                row={'task_id':t['task_id'],'split':split,'memory_on':enabled,'approved_rules_snapshot':rule_count}
                try:
                    start=time.perf_counter()
                    # Build a temporary PDF for disjoint ranges; preserve URI annotations.
                    with fitz.open(t['previous_pdf']) as source, fitz.open() as subset:
                        for lo,hi in t['previous_page_ranges']: subset.insert_pdf(source,from_page=lo-1,to_page=hi-1,links=True)
                        p=Path(tmp)/'subset.pdf';subset.save(p);n=len(subset)
                    report=generate(p,t['template_docx'],1,n,folder,db,memory_enabled=enabled)
                    row['execution_seconds']=time.perf_counter()-start
                    row['applied_rule_count']=len(report['memory_rules_applied'])
                    row.update(evaluate_task(folder/'output.docx',t['reference_docx'],t['reference_paragraph_ranges'],t['previous_pdf'],t['previous_page_ranges']))
                    row['status']='ok'
                except Exception as e: row.update(status='error',error=str(e))
                rows.append(row)
                (out/'results.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
    keys=list(dict.fromkeys(k for r in rows for k in r))
    with (out/'results.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=keys);w.writeheader();w.writerows(rows)
    return rows
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--manifest',default='dataset_manifest_v1.1.json');p.add_argument('--data-root',default='dataset');p.add_argument('--out',default='experiments');p.add_argument('--split',choices=['train','validation','test'],default='validation');p.add_argument('--memory');p.add_argument('--limit',type=int);a=p.parse_args();print(json.dumps(run(a.manifest,a.data_root,a.out,a.split,a.memory,a.limit),ensure_ascii=False,indent=2))
