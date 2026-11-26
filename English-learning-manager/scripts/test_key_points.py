"""Behavioral checks; all writes stay in a named isolated project root."""
from pathlib import Path
from copy import deepcopy
from datetime import datetime
from zipfile import ZipFile
import argparse,io,json,shutil
import manager as engine
from manager import Manager,TZ,jd,repack,package
from key_points import audit_points,build_document

def main(project):
    project=Path(project).resolve()
    sandbox=project/'work/learning-manager-qa'/('sandbox-key-points-'+datetime.now(TZ).strftime('%Y%m%dT%H%M%S%f'))
    assert sandbox.is_relative_to(project/'work/learning-manager-qa')
    (sandbox/'.learning-manager/transactions').mkdir(parents=True)
    original=Manager(project).load()
    state=deepcopy(original)
    for name in ['plans','lessons','submissions','knowledge','weekly']:state[name]={}
    state['baseline_errors']=[];state['adjustments']=[];state.pop('key_points',None);state.pop('vocabulary_summary',None)
    (sandbox/'.learning-manager/state.json').write_bytes(jd(state))
    m=Manager(sandbox);checks=[]
    def check(label,condition):
        assert condition,label
        checks.append(label)
    entry={'id':'collocation.test-consider-doing','category':['词汇与搭配','动词后接 doing'],
           'original':'SYNTHETIC consider doing','pattern':'consider doing sth','meaning':'考虑做某事',
           'examples':[{'english':'I am considering studying abroad.','chinese':'我正在考虑出国留学。'}],
           'notes':['SYNTHETIC TEST ONLY']}
    payload={'date':'2026-11-26','raw_input':entry['original'],'entries':[entry]}
    before=m.state_path.read_bytes();real_replace=engine.os.replace;calls=0
    def fail_once(source,dest):
        nonlocal calls
        calls+=1
        if calls==2:raise OSError('injected key-point write failure')
        return real_replace(source,dest)
    engine.os.replace=fail_once
    try:
        try:m.record_points(payload)
        except OSError:pass
        else:raise AssertionError('Injected failure must propagate')
    finally:engine.os.replace=real_replace
    check('Interrupted initial record rolls back Word and state',m.state_path.read_bytes()==before and not m.safe('06_Learning Notes/English Learning_Key Points.docx').exists())
    first=m.record_points(payload);after=m.load();doc=m.safe(first['file']);word=doc.read_bytes()
    check('First record creates one categorized Word entry',len(after['key_points']['entries'])==1 and not audit_points(m,after))
    check('Recording does not change graded learning evidence',all(after[k]==state[k] for k in ['plans','lessons','submissions','knowledge','weekly','baseline_errors','adjustments']))
    rev=after['revision'];state_bytes=m.state_path.read_bytes()
    check('Identical input is a byte-preserving no-op',m.record_points(payload)['action']=='unchanged' and m.state_path.read_bytes()==state_bytes and doc.read_bytes()==word)
    extra=deepcopy(payload);extra['date']='2026-11-27';extra['raw_input']='SYNTHETIC supplement consider doing'
    extra['entries'][0]['original']=extra['raw_input'];extra['entries'][0]['notes'].append('可用进行时表示正在考虑。')
    result=m.record_points(extra);merged=m.load()['key_points']['entries'][entry['id']]
    check('Supplement merges with original date and provenance preserved',result['total']==1 and merged['first_recorded_date']=='2026-11-26' and len(merged['originals'])==2 and len(merged['examples'])==1 and len(merged['notes'])==2)
    with ZipFile(io.BytesIO(word)) as a,ZipFile(doc) as b:
        check('Later Word update preserves every other package part',set(a.namelist())==set(b.namelist()) and all(a.read(n)==b.read(n) for n in a.namelist() if n!='word/document.xml'))
    slim=deepcopy(m.load()['key_points']);slim['display_fields']={'original':False,'translation':False,'first_recorded_date':False}
    visible=engine.text(package(build_document(slim,doc.read_bytes())))
    check('User-hidden fields stay hidden while original evidence is retained',entry['original'] not in visible and '首次收录：' not in visible and '翻译：' not in visible and slim['entries'][entry['id']]['originals'][0]['text']==entry['original'])
    bad=deepcopy(payload);bad['entries'][0]['id']='collocation.duplicate'
    saved=m.state_path.read_bytes()
    try:m.record_points(bad)
    except ValueError:pass
    else:raise AssertionError('Duplicate pattern must reject')
    check('Different IDs cannot duplicate the same pattern',m.state_path.read_bytes()==saved)
    saved_word=doc.read_bytes();root=package(saved_word)
    text_node=next(root.iter('{'+engine.W+'}t'));text_node.text='USER EXTERNAL EDIT'
    doc.write_bytes(repack(saved_word,engine.ET.tostring(root,encoding='utf-8',xml_declaration=True)))
    edited=doc.read_bytes()
    try:m.record_points(extra)
    except ValueError:pass
    else:raise AssertionError('External edit must block overwrite')
    check('External Word edits are protected even on repeat input',doc.read_bytes()==edited and m.state_path.read_bytes()==saved)
    report={'passed':True,'checks':checks,'sandbox':str(sandbox),'scope':'Synthetic records stay in this isolated root.'}
    (sandbox/'test-results.json').write_bytes(jd(report))
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--project',required=True)
    main(p.parse_args().project)
