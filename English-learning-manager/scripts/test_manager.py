"""Run behavioral integration checks in a disposable, clearly isolated root."""
from pathlib import Path
import base64
import copy
from datetime import date, datetime
import importlib.util
import json
import shutil
import sys
from zipfile import ZipFile
from lxml import etree as ET

from manager import Manager, NS, W, TZ, jd, digest, package, repack, cell_text

PROJECT=Path(__file__).resolve().parents[3]
OUT=PROJECT/'work/learning-manager-qa'
SANDBOX=OUT/('sandbox-'+datetime.now(TZ).strftime('%Y%m%dT%H%M%S%f'))
SKILL=Path('C:/Users/Kohye/.codex/skills/English-learning-template')
CHECKS=[]


def check(label,condition):
    if not condition:
        raise AssertionError(label)
    CHECKS.append(label)


def rejects(label,fn):
    try:
        fn()
    except ValueError:
        check(label,True)
    else:
        raise AssertionError(label+' should reject')


def lesson(day,dt):
    data=(SKILL/'assets/reference.docx').read_bytes();root=package(data);body=root.find('w:body',NS)
    # This is an internal review lesson authored by applying the existing
    # template workflow, not a second generator shipped for daily lessons.
    replacements={
        0:f'DAY {day}  ·  ENGLISH WORKSHEET',
        1:'Date: '+date.fromisoformat(dt).strftime('%d %B %Y').lstrip('0'),
        4:'Words 1–20  ·  Review previously studied vocabulary',
        17:'Grammar Focus: Hope and Want Followed by To',
        18:'一、表达愿望与目标',
        19:'hope和want后接to加动词原形，表达希望或想做某事。',
        20:'Examples: I hope to improve my English.（我希望提高英语。）',
        21:'She wants to develop her skills.（她想培养技能。）',
        22:'二、基本结构',
        23:'主语 + hope / want + to + 动词原形。',
        24:'第三人称单数：She hopes to achieve her goal.',
        25:'三、注意动词形式',
        26:'to后用原形：to improve，不用to improves。',
        27:'否定句：I do not want to reduce practice time.',
        28:'主语为she时：She does not want to stop learning.',
        29:'四、不要漏掉to',
        30:'错误：I hope improve. 正确：I hope to improve.',
        31:'错误：She wants develops. 正确：She wants to develop.',
        32:'五、迁移到自己的目标',
        33:'先确定希望做的事，再用hope/want to do写完整句。',
        34:'Examples: We want to communicate clearly.（我们想清楚交流。）',
        37:'1. I hope ___ my English.',38:'     A. improve       B. to improve       C. improving',
        39:'2. She wants ___ a new skill.',40:'     A. to develop       B. develops       C. developed',
        41:'3. They hope to ___ their goals.',42:'     A. achieves       B. achieving       C. achieve',
        43:'4. He ___ to communicate clearly.',44:'     A. want       B. wants       C. wanting',
        45:'5. We do not want ___ our study time.',46:'     A. reducing       B. reduce       C. to reduce',
        50:'Read this original teaching passage and answer the questions.',
        51:'Small Goals for a Design Student',
        52:'Mina studies design. She hopes to improve her English so that she can explain her ideas clearly.',
        53:'Every Saturday, she reviews old words. She does not add new words because she wants to remember useful expressions.',
        54:'She practises with a friend for ten minutes. They compare their sentences and provide support for each other.',
        55:'Mina wants to describe one design in English next week. A small goal helps her maintain a regular habit.',
        57:'review  —  to study something again',58:'goal  —  something you want to achieve',
        59:'expression  —  a word or group of words',60:'regular  —  happening often in a pattern',
        61:'clearly  —  in a way that is easy to understand',62:'support  —  help given to someone',
        64:'1. Why does Mina want to improve her English?',65:'     A. To explain her ideas.       B. To stop designing.       C. To add more homework.',
        66:'2. What does Mina do on Saturdays?',67:'     A. Adds new words.       B. Reviews old words.       C. Avoids English.',
        68:'3. How long does she practise with her friend?',69:'     A. Five minutes.       B. Ten minutes.       C. Thirty minutes.',
        70:'4. What is her goal for next week?',71:'     A. Describe one design.       B. Learn no words.       C. Take an official test.',
        77:'1. 我希望用英语清楚地描述我的设计。',80:'2. 她想培养新的技能。',83:'3. 我们希望下周实现这个小目标。',
        87:'Write 3–5 sentences about a small learning goal. Use hope to or want to and at least four words below:',
        88:'improve     describe     develop     achieve',
    }
    for index,value in replacements.items():
        if body[index].tag!=f'{{{W}}}p':raise AssertionError(index)
        cell_text(body[index],value)
    target=SANDBOX/'candidates'/f"{dt.replace('-','')}_English worksheet_Day{day}.docx"
    target.parent.mkdir(exist_ok=True)
    target.write_bytes(repack(data,ET.tostring(root,encoding='utf-8',xml_declaration=True)))
    blueprint=[]
    table=body.findall('w:tbl',NS)[0]
    for i,row in enumerate(table.findall('w:tr',NS)[1:],1):
        cells=row.findall('w:tc',NS)
        for direction in ['en','zh']:
            blueprint.append({'id':f'vocab.{direction}.{i:02}','prompt':f'双向默写 {direction} 第{i}词','points':1})
    for section,n,start in [('grammar',5,37),('reading',4,64),('translation',3,77)]:
        for i in range(n):
            stride=3 if section=='translation' else 2
            blueprint.append({'id':f'{section}.{i+1:02}','prompt':''.join(body[start+i*stride].itertext()),'points':1})
    blueprint.append({'id':'writing.01','prompt':replacements[87],'points':5})
    return {'date':dt,'day':day,'file':str(target),'word_mode':'review','blueprint':blueprint,'review_question_map':{'grammar.hope-to':['grammar.01','translation.01']},'qa':{'visual':'unavailable','reason':'隔离测试；运行时没有原生排版引擎，未宣称视觉通过'}}


def answers(reg,success=True,pending=False):
    items=[]
    for q in reg['blueprint']:
        target=q['id'] in {'grammar.01','translation.01'}
        ok=success or not target
        original='I hope to improve my English.' if ok else 'I hope improve my English.'
        item={'id':q['id'],'original_answer':original,'result':'correct' if ok else 'incorrect','score':q['points'] if ok else 0,'expected_answer':'I hope to improve my English.','explanation':'内部合成样例；用于检验记录联动，不是学习者成绩。','checks':[],'issues':[]}
        if target:
            item['checks']=[{'knowledge_id':'grammar.hope-to','success':ok,'independent':True,'mode':'review'}]
            if not ok:item['issues']=[{'knowledge_id':'grammar.hope-to','reason':'缺少to','explanation':'hope后接to加原形。','practice':'不看提示翻译：我希望下周实现目标。'}]
        if pending and q['id']=='reading.04':
            item.update({'original_answer':None,'result':'pending','score':None,'pending_reason':'手写图像第4题字迹无法确认','checks':[],'issues':[]})
        items.append(item)
    return {'date':reg['date'],'raw_text':'SYNTHETIC TEST DATA ONLY','items':items}


def same_parts(a,b):
    with ZipFile(a) as za,ZipFile(b) as zb:
        return set(za.namelist())==set(zb.namelist()) and all(za.read(n)==zb.read(n) for n in za.namelist() if n!='word/document.xml')


def main():
    SANDBOX.mkdir(parents=True)
    for folder in ['01_English Proficiency Profile','02_Weekly Study Plan','03_Eglish worksheet']:
        (SANDBOX/folder).mkdir()
        for source in (PROJECT/folder).iterdir():
            if source.is_file() and source.suffix in {'.docx','.md'}:
                shutil.copyfile(source,SANDBOX/folder/source.name)
    mgr=Manager(SANDBOX);mgr.bootstrap(SKILL);s=mgr.load()
    check('既有42天周表全部接入',len(s['plans'])==42)
    check('初始客观错题与Lesson01证据接入',len(s['baseline_errors'])==31)
    check('初始试卷不冒认已完成或已批改',mgr.progress()['graded_days']==0 and mgr.progress()['completed_days']==0)
    check('初始档案及原计划字节备份',digest((SANDBOX/'.learning-manager/baseline/weekly-plan.docx').read_bytes())==digest((PROJECT/'02_Weekly Study Plan/English Learning_Weekly Study Plan.docx').read_bytes()))
    ctx=mgr.context('2026-11-28',2)
    check('出卷上下文带原档案/周表/近期反馈/到期错题',bool(ctx['profile_text']) and bool(ctx['plan']) and 'recent_feedback' in ctx and bool(ctx['due_reviews']))
    reg=lesson(2,'2026-11-28');original=(SKILL/'assets/reference.docx').read_bytes();mgr.register(reg)
    check('原模板检查通过并以draft登记',mgr.load()['lessons'][reg['date']]['qa']['structural']['passed'] and mgr.load()['lessons'][reg['date']]['status']=='draft')
    check('真正替换教学内容并保留其他包部件',same_parts(SKILL/'assets/reference.docx',Path(reg['file'])) and 'Small Goals for a Design Student' in ''.join(package(Path(reg['file']).read_bytes()).itertext()))
    check('重复登记相同日期幂等',mgr.register(reg)['action']=='unchanged')
    rejects('同日期不同Day拒绝',lambda:mgr.context('2026-11-28',3))
    rejects('同Day不同日期拒绝',lambda:mgr.context('2026-11-29',2))
    invalid=copy.deepcopy(reg);invalid['word_mode']='new'
    rejects('周末新词及重复新词拒绝',lambda:mgr.register(invalid))
    base_count=mgr.bank(mgr.load())['grammar.hope-to']['error_count']
    grade=answers(reg,False,True)
    source=SANDBOX/'student-submission.txt';source.write_text('SYNTHETIC: I hope improve my English.',encoding='utf-8')
    grade['source_files']=[str(source),reg['file']]
    mgr.grade(grade);first=mgr.load()['submissions'][reg['date']]
    check('原答案和文件证据完整保留',first['items'][40]['original_answer']=='I hope improve my English.' and all(digest((SANDBOX/e['path']).read_bytes())==e['sha256'] for e in first['evidence']))
    check('同题知识点唯一，重复错误计真实次数',mgr.bank(mgr.load())['grammar.hope-to']['error_count']==base_count+2)
    check('模糊手写答案保持待确认不计分',first['items'][48]['result']=='pending' and first['items'][48]['score'] is None)
    revision=mgr.load()['revision'];mgr.grade(grade)
    check('相同批改重复提交不增版本/错误次数',mgr.load()['revision']==revision)
    malformed=copy.deepcopy(grade);malformed['items'][48]['score']=0
    rejects('待确认不得擅自计零分',lambda:mgr.grade(malformed))
    missing=copy.deepcopy(grade);missing['items']=missing['items'][:-1]
    rejects('漏题不能静默当作完成',lambda:mgr.grade(missing))
    duplicate=copy.deepcopy(grade);duplicate['items'].append(duplicate['items'][0])
    rejects('同次作答重复题号拒绝',lambda:mgr.grade(duplicate))
    mgr.weekly('2026-11-23');stats=mgr.load()['weekly']['2026-11-23']
    check('周统计排除待确认并单列覆盖',stats['pending_questions']==1 and stats['graded_questions']==52 and stats['submission_days']==1)
    check('初始摸底不计入周正确率',stats['total_questions']==53)
    check('听说无证据保持未评估',stats['knowledge']['listening.diagnostic']['state']=='未评估' and stats['knowledge']['speaking.diagnostic']['state']=='未评估')
    corrected=answers(reg,True,True);corrected['source_files']=grade['source_files'];mgr.grade(corrected)
    check('重批撤销误判，不累加旧错误',mgr.bank(mgr.load())['grammar.hope-to']['error_count']==base_count and mgr.load()['submissions'][reg['date']]['revision']==2)
    check('重批自动同步既有周评估',mgr.load()['weekly']['2026-11-23']['score']==mgr.load()['weekly']['2026-11-23']['possible_points'])
    check('初始档案全文保留并追加周评估链接',(SANDBOX/'01_English Proficiency Profile/English Learning_English Proficiency Profile.md').read_text(encoding='utf-8').startswith((PROJECT/'01_English Proficiency Profile/English Learning_English Proficiency Profile.md').read_text(encoding='utf-8').rstrip()))
    mgr.grade(grade)
    check('历史重批版本备份存在',any('grade_or_regrade'==json.loads(p.read_text(encoding='utf-8'))['action'] and any(e['path']=='.learning-manager/state.json' and e['before_sha256'] for e in json.loads(p.read_text(encoding='utf-8'))['files']) for p in (SANDBOX/'.learning-manager/transactions').glob('*/manifest.json')))
    before=SANDBOX/'before-adjust.docx';plan=SANDBOX/'02_Weekly Study Plan/English Learning_Weekly Study Plan.docx';shutil.copyfile(plan,before)
    oldslots=mgr.load()['plans']['2026-11-25']['original'][:]
    adjust={'date':'2026-11-25','slots':['词汇10′：复习旧词','语法10′：hope to do','听力5′：用户自备','写作5′：独立写句'],'minutes':30,'status':'partial','actual':{'minutes':20,'content':'完成词汇和语法','completed_tasks':2,'assigned_tasks':4},'reason':'模拟用户明确改为30分钟并只完成前两项'}
    mgr.adjust(adjust)
    check('时间变更同步同一Word且保留初始计划',mgr.load()['plans']['2026-11-25']['original']==oldslots and mgr.load()['plans']['2026-11-25']['minutes']==30 and 'hope to do' in ''.join(package(plan.read_bytes()).itertext()))
    check('Word样式/公式/页脚等包部件逐字节保留',same_parts(before,plan))
    a,b=package(before.read_bytes()),package(plan.read_bytes())
    check('周表合并格/列宽/可点击复选框未丢失',len(a.findall('.//w:tblGrid',NS))==len(b.findall('.//w:tblGrid',NS)) and [ET.tostring(x) for x in a.findall('.//w:tblGrid',NS)]==[ET.tostring(x) for x in b.findall('.//w:tblGrid',NS)] and len(a.findall('.//w:sdt',NS))==len(b.findall('.//w:sdt',NS)))
    bad=copy.deepcopy(adjust);bad['minutes']=60
    rejects('训练格时长与总时长冲突拒绝',lambda:mgr.adjust(bad))
    mgr.adjust({'date':'2026-11-26','status':'missed','actual':{'minutes':0,'content':'用户明确今日未学习','completed_tasks':0,'assigned_tasks':4},'reason':'模拟用户报告今天没有完成'})
    check('未完成事实入账且不改其他未知天',mgr.load()['plans']['2026-11-26']['status']=='missed' and mgr.load()['plans']['2026-11-24']['status']=='unconfirmed')
    complete={'date':'2026-11-28','status':'completed','actual':{'minutes':60,'content':'完成四项周测','completed_tasks':4,'assigned_tasks':4},'reason':'模拟用户报告完成'};mgr.adjust(complete)
    bad=copy.deepcopy(complete);bad['status']='missed';bad['actual']={'minutes':0,'content':'改写历史','completed_tasks':0,'assigned_tasks':4}
    rejects('已完成历史任务冻结',lambda:mgr.adjust(bad))
    for day,dt in [(3,'2026-11-29'),(4,'2026-12-01'),(5,'2026-12-02')]:
        r=lesson(day,dt);mgr.register(r);mgr.grade(answers(r,True))
        if day==3:
            e=mgr.bank(mgr.load())['grammar.hope-to'];check('同日多题复习最多升级一次至+3天',e['stage']==1 and e['next_due']=='2026-12-02')
        if day==4:check('未到期提前成功不延后复习',mgr.bank(mgr.load())['grammar.hope-to']['next_due']=='2026-12-02')
    check('到期复习成功升级至+7天',mgr.bank(mgr.load())['grammar.hope-to']['next_due']=='2026-12-09')
    r=lesson(6,'2026-12-09');mgr.register(r);mgr.grade(answers(r,False))
    check('复习失败重置+1天且掌握状态降级',mgr.bank(mgr.load())['grammar.hope-to']['next_due']=='2026-12-10' and mgr.bank(mgr.load())['grammar.hope-to']['mastery'] in {'不稳定','未掌握'})
    mgr.weekly('2026-11-23');rev=mgr.load()['revision'];mgr.weekly('2026-11-23')
    check('同一周重复评估幂等',mgr.load()['revision']==rev)
    profile=SANDBOX/'01_English Proficiency Profile/English Learning_English Proficiency Profile.md'
    with profile.open('a',encoding='utf-8') as f:f.write('\n用户在管理区之后追加的学习目标。\n')
    mgr.weekly('2026-11-30')
    check('管理区之后的用户追加内容不会丢失','用户在管理区之后追加的学习目标。' in profile.read_text(encoding='utf-8'))
    check('准确区分已确认任务完成率与未知天',mgr.load()['weekly']['2026-11-23']['confirmed_task_counts']=={'completed':6,'assigned':12})
    mgr.weekly('2026-12-28')
    days=[]
    for i in range(7):
        dt=date(2027,1,4).fromordinal(date(2027,1,4).toordinal()+i).isoformat()
        days.append({'date':dt,'slots':['词汇15′：旧词复习','语法20′：冠词与不定式','阅读10′：细节定位','写作15′：五句独立输出'],'minutes':60})
    mgr.plan_week({'week_start':'2027-01-04','days':days,'goal':'巩固已检测薄弱点','criteria':'独立造句并提供真实作答','reason':'隔离测试下一周延展'})
    check('超出现有周数时延展同一Word而非新建系统',len(mgr.load()['plans'])==49 and len(package(plan.read_bytes()).findall('w:body/w:tbl',NS))==7)
    root=package(plan.read_bytes());ids=[n.get(f'{{{W}}}val') for n in root.findall('.//w:sdtPr/w:id',NS)]
    check('新周交互复选框ID无重复',len(ids)==len(set(ids)))
    stored=plan.read_bytes();plan.write_bytes(stored+b'external edit')
    rejects('检测Word外部编辑并阻止覆盖',lambda:mgr.adjust({'date':'2026-12-10','slots':days[0]['slots'],'reason':'测试外部冲突'}))
    plan.write_bytes(stored)
    check('隔离全链路审计通过',mgr.audit()['passed'])
    check('原模板文件始终不变',(SKILL/'assets/reference.docx').read_bytes()==original)
    # Simulate a crash between two file replacements, then recover snapshots.
    tx=SANDBOX/'.learning-manager/transactions/simulated-interruption';tx.mkdir()
    victim=SANDBOX/'recovery-test.txt';victim.write_bytes(b'new')
    (tx/'0.before').write_bytes(b'old');(tx/'0.after').write_bytes(b'new')
    (tx/'manifest.json').write_bytes(jd({'transaction':'simulated-interruption','action':'test','status':'prepared','files':[{'path':'recovery-test.txt','index':0,'before_sha256':digest(b'old'),'after_sha256':digest(b'new')}]}))
    rejects('未完成事务阻止继续读取',mgr.load)
    mgr.recover();check('中断事务恢复旧版本',victim.read_bytes()==b'old' and mgr.audit()['passed'])
    # Inject one write failure after the first replacement. Before-images must
    # restore state and Word together rather than merely returning an error.
    import manager as engine
    before_state=mgr.state_path.read_bytes();before_word=plan.read_bytes()
    real_replace=engine.os.replace;calls=0
    def fail_once(source,destination):
        nonlocal calls
        calls+=1
        if calls==2:raise OSError('injected disk write failure')
        return real_replace(source,destination)
    engine.os.replace=fail_once
    try:
        try:mgr.adjust({'date':'2026-12-11','slots':days[0]['slots'],'reason':'注入写入失败检查回滚'})
        except OSError:pass
        else:raise AssertionError('injected failure must propagate')
    finally:engine.os.replace=real_replace
    check('多文件写入异常后状态和Word一起回滚',mgr.state_path.read_bytes()==before_state and plan.read_bytes()==before_word and mgr.audit()['passed'])
    # Exercise actual mastery decisions with independent dated evidence,
    # rather than checking that five status words merely appear in the files.
    facts=copy.deepcopy(mgr.load());facts['submissions']={};facts['baseline_errors']=[]
    facts['knowledge']={'grammar.hope-to':{'label':'hope to do','category':'grammar'}}
    def evidence(dt,success=True,independent=True):
        facts['submissions'][dt]={'items':[{'id':'translation.01','prompt':'独立句子','original_answer':'learner original','expected_answer':'reference','result':'correct' if success else 'incorrect','checks':[{'knowledge_id':'grammar.hope-to','success':success,'independent':independent,'mode':'review'}],'issues':[] if success else [{'knowledge_id':'grammar.hope-to','reason':'实际检测失败'}]}]}
    check('无检测证据状态为未评估',mgr.bank(facts)['grammar.hope-to']['mastery']=='未评估')
    evidence('2026-11-23',False)
    check('单次实际失败状态为未掌握',mgr.bank(facts)['grammar.hope-to']['mastery']=='未掌握')
    facts['submissions']={};evidence('2026-11-23')
    check('少量成功证据不宣称掌握',mgr.bank(facts)['grammar.hope-to']['mastery']=='不稳定')
    evidence('2026-11-24');evidence('2026-11-25')
    check('多日期稳定成功达到基本掌握',mgr.bank(facts)['grammar.hope-to']['mastery']=='基本掌握')
    evidence('2026-11-30');evidence('2026-12-01')
    check('跨周独立回忆和复习成功才达到已掌握',mgr.bank(facts)['grammar.hope-to']['mastery']=='已掌握')
    for submission in facts['submissions'].values():submission['items'][0]['checks'][0]['independent']=False
    check('有提示的成功不能冒认独立掌握',mgr.bank(facts)['grammar.hope-to']['mastery']=='不稳定')
    facts['submissions']={};evidence('2026-11-23',False);evidence('2026-11-24',True,False)
    check('提示后复习成功不推进间隔',mgr.bank(facts)['grammar.hope-to']['next_due']=='2026-11-24' and mgr.bank(facts)['grammar.hope-to']['stage']==0)
    report={'passed':True,'checks':CHECKS,'check_count':len(CHECKS),'sandbox':str(SANDBOX),'scope':'合成数据隔离测试，不计入真实学习进度；DOCX结构和包部件通过，原生逐页渲染尚不可用。'}
    (OUT/'integration-results.json').write_bytes(jd(report))
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__=='__main__':main()
