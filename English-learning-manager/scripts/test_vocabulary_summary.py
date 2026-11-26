"""Vocabulary integration checks in a separate project; no formal data writes."""
from pathlib import Path
from copy import deepcopy
from datetime import datetime
from zipfile import ZipFile
import argparse,importlib.util,io,json,shutil
from manager import Manager,TZ,jd
from vocabulary_summary import DOCUMENT,prepare_summary,audit_summary

def main(project):
    project=Path(project).resolve();m0=Manager(project);initial=m0.load()
    sandbox=project/'work/learning-manager-qa'/('sandbox-vocabulary-'+datetime.now(TZ).strftime('%Y%m%dT%H%M%S%f'))
    assert sandbox.is_relative_to(project/'work/learning-manager-qa')
    (sandbox/'.learning-manager/transactions').mkdir(parents=True)
    state=deepcopy(initial);second=deepcopy(state['lessons']['2026-11-24'])
    state['lessons'].pop('2026-11-24');state.pop('key_points',None);state.pop('vocabulary_summary',None)
    state['submissions']={};state['weekly']={}
    for rel in [state['config']['profile'],state['config']['plan'],*[l['path'] for l in initial['lessons'].values()]]:
        dest=sandbox/rel;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(project/rel,dest)
    (sandbox/'.learning-manager/state.json').write_bytes(jd(state))
    m=Manager(sandbox);checks=[]
    def check(label,value):
        assert value,label
        checks.append(label)
    m.summarize_vocabulary()
    check('Backfill reads the registered 20 words without changing lessons',m.vocabulary()['total_unique_words']==20 and m.load()['lessons']==state['lessons'])
    helper=Path(state['config']['worksheet_skill'])/'scripts/normalize_table_borders.py'
    spec=importlib.util.spec_from_file_location('test_summary_borders',helper);border=importlib.util.module_from_spec(spec);spec.loader.exec_module(border)
    candidate=sandbox/'candidates'/Path(second['path']).name;candidate.parent.mkdir()
    candidate.write_bytes(border.normalize_docx_bytes((project/second['path']).read_bytes()))
    m.register({'date':'2026-11-24','day':2,'file':str(candidate),'blueprint':second['blueprint'],'word_mode':'new','qa':{'visual':'unavailable','reason':'Isolated integration test'}})
    check('Worksheet registration automatically updates the same summary to 40 words',m.vocabulary()['total_unique_words']==40 and len(m.vocabulary()['days'])==2)
    word=m.safe(DOCUMENT).read_bytes();saved=m.state_path.read_bytes()
    check('Repeat summary is a byte-preserving no-op',m.summarize_vocabulary()['action']=='unchanged' and m.state_path.read_bytes()==saved and m.safe(DOCUMENT).read_bytes()==word)
    import test_manager as fixtures
    fixtures.SANDBOX=sandbox
    review=fixtures.lesson(3,'2026-11-28')
    m.register(review)
    check('Weekend review adds zero new words and does not rewrite summary',m.vocabulary()['total_unique_words']==40 and m.safe(DOCUMENT).read_bytes()==word)
    duplicate=deepcopy(m.load());duplicate['lessons']['2026-11-25']=deepcopy(duplicate['lessons']['2026-11-23'])
    duplicate['lessons']['2026-11-25'].update(date='2026-11-25',day=4)
    store,_,_=prepare_summary(m,duplicate)
    check('Repeated headwords on another day do not create duplicate entries',store['total_unique_words']==40 and len(store['days'])==2)
    check('Source snapshots and actual table signature audit correctly',not audit_summary(m,m.load()))
    after=m.load()
    check('Summary and review registration do not invent grades or mastery',after['submissions']==state['submissions'] and after['knowledge']==state['knowledge'] and after['baseline_errors']==state['baseline_errors'])
    m.safe(DOCUMENT).write_bytes(word+b'USER EDIT')
    before=m.state_path.read_bytes()
    try:m.summarize_vocabulary()
    except ValueError:pass
    else:raise AssertionError('External Word edit must block overwrite')
    check('External summary edits are preserved instead of overwritten',m.safe(DOCUMENT).read_bytes()==word+b'USER EDIT' and m.state_path.read_bytes()==before)
    report={'passed':True,'checks':checks,'sandbox':str(sandbox)}
    (sandbox/'test-results.json').write_bytes(jd(report))
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--project',required=True);main(p.parse_args().project)
