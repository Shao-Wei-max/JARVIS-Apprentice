"""Reproduce T01 learning from paired human corrections; explicit approval only."""
import argparse,json,hashlib
from pathlib import Path
from contextlib import closing
from jarvis import compare,propose,propose_v05,db_connect
APPROVED={'set_image_wrap_top_bottom','remove_legacy_header','inherit_continuation_style_v2'}
def learn(before,after,out,seconds):
    out=Path(out);out.mkdir(parents=True,exist_ok=True);db=out/'stage3_memory.db'
    if db.exists(): raise FileExistsError('Use a new output directory; existing memory is preserved')
    diff=out/'correction_diff.json';data=compare(before,after,diff)
    propose(diff,db);propose_v05(diff,db)
    with closing(db_connect(db)) as c:
        rows=c.execute('SELECT id,action FROM rules').fetchall()
        for i,a in rows:
            if a in APPROVED:c.execute('UPDATE rules SET approved=1 WHERE id=?',(i,))
        c.commit();rules=c.execute('SELECT id,action,approved FROM rules').fetchall()
    record={'task_id':'T01','correction_seconds':seconds,'events':len(data['events']),'rules':rules,'approval':'explicit user approval of three named actions','source_sha256':{Path(p).name:hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in (before,after)}}
    (out/'learning_record.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8');return record
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--before',required=True);p.add_argument('--after',required=True);p.add_argument('--out',required=True);p.add_argument('--seconds',type=int,required=True);a=p.parse_args();print(json.dumps(learn(a.before,a.after,a.out,a.seconds),ensure_ascii=False,indent=2))
