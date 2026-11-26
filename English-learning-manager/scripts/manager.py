"""English-learning-manager: local evidence, plans, review and assessment.

The language model grades and authors through the existing worksheet skill.
This module validates and persists those decisions; it never guesses answers.
Python 3.12 and bundled lxml for lossless namespace handling in Word XML.
"""
from __future__ import annotations
import argparse
import copy
from datetime import date, datetime, timedelta, timezone
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import re
import shutil
import sys
import uuid
from lxml import etree as ET
from zipfile import ZipFile

TZ = timezone(timedelta(hours=8))
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W, 'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'}
INTERVALS = [1, 3, 7, 14, 30]
STATES = ['已掌握', '基本掌握', '不稳定', '未掌握', '未评估']


def now():
    return datetime.now(TZ).isoformat(timespec='microseconds')


def today():
    return datetime.now(TZ).date().isoformat()


def iso(value):
    return date.fromisoformat(value).isoformat()


def jd(value):
    return (json.dumps(value, ensure_ascii=False, indent=2) + '\n').encode('utf-8')


def digest(value):
    return hashlib.sha256(value).hexdigest()


def text(el):
    return ''.join(n.text or '' for n in el.iter() if n.tag in {f'{{{W}}}t', '{http://schemas.openxmlformats.org/officeDocument/2006/math}t'})


def package(data):
    with ZipFile(io.BytesIO(data)) as z:
        if z.testzip():
            raise ValueError('Word 文件损坏')
        return ET.fromstring(z.read('word/document.xml'))


def repack(data, document):
    out = io.BytesIO()
    with ZipFile(io.BytesIO(data)) as src, ZipFile(out, 'w') as dst:
        for entry in src.infolist():
            dst.writestr(entry, document if entry.filename == 'word/document.xml' else src.read(entry.filename))
    return out.getvalue()


def cell_text(cell, value):
    """Change text slots only; preserve paragraph/run/cell properties."""
    nodes = cell.findall('.//w:t', NS)
    if nodes:
        nodes[0].text = value
        nodes[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        for node in nodes[1:]:
            node.text = ''
    else:
        p = cell.find('w:p', NS)
        if p is None:
            p = ET.SubElement(cell, f'{{{W}}}p')
        ET.SubElement(ET.SubElement(p, f'{{{W}}}r'), f'{{{W}}}t').text = value


def check_id(value):
    if not isinstance(value, str) or not re.fullmatch(r'[a-z0-9][a-z0-9_.-]{0,95}', value):
        raise ValueError(f'无效知识点/题号标识: {value!r}')
    return value


def validate_slots(slots, minutes):
    if len(slots)!=4 or not all(isinstance(s,str) and s.strip() for s in slots):
        raise ValueError('沿用周表四训练格')
    allocated=sum(int(n) for slot in slots for n in re.findall(r'(\d+)\s*(?:′|’|分钟|min\b)',slot))
    if allocated != minutes:
        raise ValueError(f'训练格内分钟数合计{allocated}，与计划{minutes}不符')


class Manager:
    def __init__(self, root):
        self.root = Path(root).resolve()
        self.home = self.root / '.learning-manager'
        self.state_path = self.home / 'state.json'

    def safe(self, rel):
        dest = (self.root / rel).resolve()
        if not dest.is_relative_to(self.root) or dest == self.root:
            raise ValueError('输出路径必须位于本项目内')
        return dest

    def load(self):
        if not self.state_path.exists():
            raise ValueError('系统未初始化')
        pending = [p for p in (self.home / 'transactions').glob('*/manifest.json') if json.loads(p.read_text(encoding='utf-8'))['status'] == 'prepared']
        if pending:
            raise ValueError('存在中断事务；先运行 recover，禁止覆盖未完成写入')
        return json.loads(self.state_path.read_text(encoding='utf-8'))

    def commit(self, old, state, action, files=None):
        """Durable before-images + staged after-images + recovery manifest.

        A lock prevents concurrent writers. Readers refuse prepared transactions.
        A crash is rolled back by recover; all snapshots remain for audit.
        """
        self.home.mkdir(exist_ok=True)
        lock = self.home / 'write.lock'
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        txn = None
        try:
            os.write(fd, jd({'pid': os.getpid(), 'created_at': now()}))
            current = json.loads(self.state_path.read_text(encoding='utf-8')) if self.state_path.exists() else None
            if current != old:
                raise ValueError('学习记录已被其他操作修改，请重新读取')
            if any(json.loads(p.read_text(encoding='utf-8'))['status'] == 'prepared' for p in (self.home/'transactions').glob('*/manifest.json')):
                raise ValueError('先恢复中断事务')
            state['revision'] = (old or {}).get('revision', 0) + 1
            state['updated_at'] = now()
            output = dict(files or {})
            output.update(self.views(state))
            output['.learning-manager/state.json'] = jd(state)
            output = {rel: value for rel, value in output.items() if not self.safe(rel).is_file() or self.safe(rel).read_bytes() != value}
            txid = datetime.now(TZ).strftime('%Y%m%dT%H%M%S%f') + '-' + uuid.uuid4().hex[:8]
            txn = self.home / 'transactions' / txid
            txn.mkdir(parents=True)
            entries = []
            for i, (rel, content) in enumerate(output.items()):
                dest = self.safe(rel)
                if dest.exists() and not dest.is_file():
                    raise ValueError(f'目标不是文件: {dest}')
                before = dest.read_bytes() if dest.exists() else None
                if before is not None:
                    (txn / f'{i}.before').write_bytes(before)
                (txn / f'{i}.after').write_bytes(content)
                entries.append({'path': rel, 'index': i, 'before_sha256': digest(before) if before is not None else None, 'after_sha256': digest(content)})
            manifest = {'transaction': txid, 'action': action, 'created_at': now(), 'revision': state['revision'], 'status': 'prepared', 'files': entries}
            mp = txn / 'manifest.json'
            mp.write_bytes(jd(manifest))
            try:
                for entry in entries:
                    dest = self.safe(entry['path'])
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    stage = dest.with_name(dest.name + '.' + txid + '.tmp')
                    shutil.copyfile(txn / f"{entry['index']}.after", stage)
                    os.replace(stage, dest)
                manifest['status'] = 'committed'
                tmp = txn / 'manifest.tmp'
                tmp.write_bytes(jd(manifest))
                os.replace(tmp, mp)
            except BaseException:
                self._rollback(txn, manifest)
                raise
            return {'action': action, 'revision': state['revision'], 'transaction': txid, 'files': list(output)}
        finally:
            os.close(fd)
            lock.unlink(missing_ok=True)

    def _rollback(self, txn, manifest):
        for entry in manifest['files']:
            dest = self.safe(entry['path'])
            if entry['before_sha256'] is None:
                dest.unlink(missing_ok=True)
            else:
                dest.parent.mkdir(parents=True, exist_ok=True)
                stage = dest.with_name(dest.name + '.recovery.tmp')
                shutil.copyfile(txn / f"{entry['index']}.before", stage)
                os.replace(stage, dest)
        manifest['status'] = 'rolled_back'
        (txn / 'manifest.json').write_bytes(jd(manifest))

    def recover(self):
        lock = self.home / 'write.lock'
        if lock.exists():
            pid = json.loads(lock.read_text(encoding='utf-8'))['pid']
            if os.name == 'nt':
                import ctypes
                handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
                if handle:
                    ctypes.windll.kernel32.CloseHandle(handle)
                    raise ValueError('写入进程仍存活，不能移除锁')
                if ctypes.windll.kernel32.GetLastError() != 87:
                    raise ValueError('无法确认写入进程已退出，保留锁')
            else:
                try:
                    os.kill(pid, 0)
                except ProcessLookupError:
                    pass
                else:
                    raise ValueError('写入进程仍存活')
            lock.unlink()
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        try:
            restored = []
            for p in sorted((self.home/'transactions').glob('*/manifest.json')):
                manifest = json.loads(p.read_text(encoding='utf-8'))
                if manifest['status'] == 'prepared':
                    for e in manifest['files']:
                        dest = self.safe(e['path'])
                        actual = digest(dest.read_bytes()) if dest.is_file() else None
                        if actual not in {e['before_sha256'], e['after_sha256']}:
                            raise ValueError('中断后文件有外部修改；保留现场，需人工合并')
                    self._rollback(p.parent, manifest)
                    restored.append(manifest['transaction'])
            return {'restored': restored}
        finally:
            os.close(fd)
            lock.unlink(missing_ok=True)

    def bootstrap(self, worksheet_skill):
        if self.state_path.exists():
            return {'action': 'already_initialized', 'revision': self.load()['revision']}
        profile = next((self.root/'01_English Proficiency Profile').glob('*.md'))
        plan = next((self.root/'02_Weekly Study Plan').glob('*.docx'))
        dirs = [p for p in self.root.iterdir() if p.is_dir() and p.name.startswith('03_')]
        if len(dirs) != 1:
            raise ValueError('应只有一个已有的03试卷目录，请在配置中确认')
        worksheet_skill = str(Path(worksheet_skill).resolve())
        spec = json.loads((Path(worksheet_skill)/'references/design-spec.json').read_text(encoding='utf-8'))
        ref = Path(worksheet_skill)/spec['reference_file']
        if digest(ref.read_bytes()) != spec['reference_sha256']:
            raise ValueError('Worksheet 保留参考的哈希与规范不一致')
        state = {'schema_version': 1, 'revision': 0, 'timezone': 'Asia/Shanghai', 'config': {'profile': profile.relative_to(self.root).as_posix(), 'plan': plan.relative_to(self.root).as_posix(), 'worksheet_dir': dirs[0].name, 'worksheet_skill': worksheet_skill, 'reference_sha256': spec['reference_sha256'], 'default_minutes': 60, 'week_start': '2026-11-23'}, 'sources': {}, 'plans': {}, 'lessons': {}, 'submissions': {}, 'baseline_errors': [], 'knowledge': {}, 'weekly': {}, 'adjustments': []}
        for p in [profile, plan, *sorted(dirs[0].glob('*.docx'))]:
            state['sources'][p.relative_to(self.root).as_posix()] = digest(p.read_bytes())
        tables = package(plan.read_bytes()).findall('w:body/w:tbl', NS)
        for ti, table in enumerate(tables):
            rows = table.findall('w:tr', NS)
            if len(rows) != 15:
                raise ValueError('周计划不是每天两行的既有结构')
            for day in range(7):
                r1, r2 = rows[1+day*2:3+day*2]
                c1, c2 = r1.findall('w:tc', NS), r2.findall('w:tc', NS)
                slots = [text(c1[1]), text(c1[2]), text(c2[1]), text(c2[2])]
                dt = date(2026,11,23) + timedelta(days=ti*7+day)
                checked = c1[-1].find('.//w14:checked', NS)
                done = checked is not None and checked.get(f'{{{NS["w14"]}}}val') == '1'
                state['plans'][str(dt)] = {'date': str(dt), 'table': ti, 'day_index': day, 'original': slots[:], 'current': slots[:], 'minutes': 60, 'status': 'completed' if done else 'unconfirmed' if str(dt) < today() else 'planned', 'actual': None, 'reason': None, 'history': []}
        for p in sorted(dirs[0].glob('*.docx')):
            match = re.fullmatch(r'(\d{8})_English worksheet_Day([1-9]\d*)\.docx',p.name)
            if not match:
                continue
            dt = datetime.strptime(match[1],'%Y%m%d').date().isoformat()
            day = int(match[2])
            self.identity(state, dt, day)
            root = package(p.read_bytes())
            vocab = self.words(root)
            state['lessons'][dt] = {'date': dt, 'day': day, 'path': p.relative_to(self.root).as_posix(), 'sha256': digest(p.read_bytes()), 'status': 'existing_ungraded', 'blueprint': [], 'words': vocab, 'word_mode': 'new', 'qa': {'structural': 'not_rechecked', 'visual': 'not_verified'}}
        self.import_baseline(state, profile.read_text(encoding='utf-8'))
        files = {'.learning-manager/baseline/profile.md': profile.read_bytes(), '.learning-manager/baseline/weekly-plan.docx': plan.read_bytes(), '.learning-manager/baseline/source-inventory.json': jd(state['sources'])}
        return self.commit(None, state, 'bootstrap', files)

    @staticmethod
    def knowledge(state, kid, label=None, category=None):
        check_id(kid)
        if kid not in state['knowledge']:
            if not label or category not in {'vocabulary','grammar','reading','writing','listening','speaking','spelling','collocation'}:
                raise ValueError(f'新知识点 {kid} 必须提供名称和有效分类')
            state['knowledge'][kid] = {'label': label, 'category': category}

    def import_baseline(self, state, source):
        v1 = {1:'vocabulary.consensus',2:'vocabulary.enhance',3:'vocabulary.lucid',7:'vocabulary.contradictory',9:'vocabulary.identify',10:'collocation.come-into-effect',11:'vocabulary.attain',13:'vocabulary.caution',17:'grammar.inversion',18:'grammar.agreement',19:'grammar.future-perfect',20:'grammar.participle',21:'grammar.subjunctive',22:'grammar.inversion',23:'grammar.comparative',24:'grammar.gerund-passive',25:'grammar.gerund',26:'grammar.inversion',27:'grammar.present-perfect',28:'grammar.agreement',29:'grammar.subjunctive',39:'reading.paraphrase'}
        v2 = {7:'vocabulary.affect-attract',8:'vocabulary.achieve-reach',9:'vocabulary.experience-manager',11:'grammar.agreement',15:'grammar.conditional',20:'reading.number-detail'}
        section = ''
        for line in source.splitlines():
            if line.startswith('## '):
                section = line
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            if len(cells) != 5 or not cells[0].isdigit():
                continue
            q = int(cells[0])
            mapping = v1 if section.startswith('## 3.') else v2 if section.startswith('## 4.') else {}
            if q not in mapping or not ('×' in cells[4] or '错误' in cells[4] or '严格计分不算' in cells[4]):
                continue
            kid = mapping[q]
            self.knowledge(state,kid,kid.split('.',1)[1],kid.split('.',1)[0])
            state['baseline_errors'].append({'id':f'baseline.{"v1" if mapping is v1 else "v2"}.{q}', 'knowledge_id':kid, 'question':cells[1], 'original_answer':cells[2], 'expected_answer':cells[3], 'reason':cells[4], 'source':'初始档案 '+section, 'detected_at':None, 'review_from':today()})
        for i,kid,original,answer,reason in [(1,'grammar.article','I am a undergraduate.','I am an undergraduate.','undergraduate以元音音素开头，用an'),(2,'spelling.study','I sutdy stage design.','I study stage design.','study拼写错误'),(3,'grammar.hope-to','I hope improve my English ability.','I hope to improve my English.','hope后接to加动词原形')]:
            self.knowledge(state,kid,{'grammar.article':'a/an冠词','spelling.study':'study拼写','grammar.hope-to':'hope/want to do'}[kid],kid.split('.')[0])
            state['baseline_errors'].append({'id':f'baseline.lesson01.{i}','knowledge_id':kid,'question':'Lesson 01 自主造句','original_answer':original,'expected_answer':answer,'reason':reason,'source':'初始档案 第5节','detected_at':None,'review_from':today()})
        for kid,label,cat in [('listening.diagnostic','听力诊断','listening'),('speaking.diagnostic','口语诊断','speaking')]:
            self.knowledge(state,kid,label,cat)

    @staticmethod
    def identity(state, dt, day):
        iso(dt)
        if not isinstance(day,int) or isinstance(day,bool) or day < 1:
            raise ValueError('Day必须为正整数')
        for existing in state['lessons'].values():
            if (existing['date'] == dt and existing['day'] != day) or (existing['day'] == day and existing['date'] != dt):
                raise ValueError('日期与Day冲突；不能重复统计或自动重编号')

    @staticmethod
    def words(root):
        for table in root.findall('w:body/w:tbl', NS):
            rows = table.findall('w:tr', NS)
            if rows and 'Chinese definition' in text(rows[0]):
                return [text(row.findall('w:tc', NS)[1].find('w:p', NS)).strip().lower() for row in rows[1:]]
        return []

    def observations(self, state, until=None):
        events=[]
        for dt,submission in sorted(state['submissions'].items()):
            if until and dt > until:
                continue
            for item in submission['items']:
                if item['result'] in {'pending','unanswered'}:
                    continue
                issues={x['knowledge_id']:x for x in item.get('issues',[])}
                for check in item.get('checks',[]):
                    events.append({'date':dt,'qid':item['id'],'knowledge_id':check['knowledge_id'],'success':check['success'],'independent':check['independent'],'mode':check['mode'],'issue':issues.get(check['knowledge_id']),'original_answer':item['original_answer'],'expected_answer':item['expected_answer'],'question':item['prompt']})
        return events

    def bank(self, state, until=None):
        entries={}
        for kid,meta in state['knowledge'].items():
            entries[kid]={'knowledge_id':kid,**meta,'error_count':0,'last_detected':None,'last_reviewed':None,'next_due':None,'stage':0,'mastery':'未评估','evidence':[]}
        for item in state['baseline_errors']:
            entry=entries[item['knowledge_id']]
            entry['error_count']+=1
            entry['next_due']=item['review_from']
            entry['mastery']='不稳定'
            entry['evidence'].append(item)
        events=self.observations(state,until)
        groups={kid:[] for kid in entries}
        for event in events:
            entry=entries[event['knowledge_id']]
            groups[event['knowledge_id']].append(event)
            entry['last_detected']=event['date']
            if event['mode']=='review':
                entry['last_reviewed']=event['date']
            if not event['success']:
                entry['error_count']+=1
                entry['stage']=0
                entry['next_due']=str(date.fromisoformat(event['date'])+timedelta(days=1))
                entry['evidence'].append(event)
        for kid,group in groups.items():
            entry=entries[kid]
            # Schedule once per knowledge point per day. A failure overrides
            # successes on that day; prompted success does not advance a stage.
            for dt in sorted({e['date'] for e in group}):
                daily=[e for e in group if e['date']==dt]
                if any(not e['success'] for e in daily):
                    entry['stage']=0
                    entry['next_due']=str(date.fromisoformat(dt)+timedelta(days=1))
                elif entry['next_due'] and dt >= entry['next_due'] and any(e['mode']=='review' and e['independent'] for e in daily):
                    entry['stage']=min(entry['stage']+1,len(INTERVALS)-1)
                    entry['next_due']=str(date.fromisoformat(dt)+timedelta(days=INTERVALS[entry['stage']]))
            if group:
                recent=group[-10:]
                rate=sum(e['success'] for e in recent)/len(recent)
                independent={e['date'] for e in recent if e['independent'] and e['success']}
                span=(date.fromisoformat(max(independent))-date.fromisoformat(min(independent))).days if independent else 0
                last_day_failed=any(not e['success'] for e in recent if e['date']==recent[-1]['date'])
                if last_day_failed:
                    entry['mastery']='未掌握' if rate<.5 else '不稳定'
                elif rate>=.9 and len(recent)>=5 and len(independent)>=3 and span>=7 and len({e['date'] for e in recent if e['mode']=='review' and e['success'] and e['independent']})>=2:
                    entry['mastery']='已掌握'
                elif rate>=.8 and len(recent)>=3 and len(independent)>=2:
                    entry['mastery']='基本掌握'
                elif rate<.5:
                    entry['mastery']='未掌握'
                else:
                    entry['mastery']='不稳定'
        return entries

    def context(self, dt=None, day=None):
        state=self.load()
        dt=iso(dt or today())
        existing=state['lessons'].get(dt)
        day=day or (existing['day'] if existing else max([x['day'] for x in state['lessons'].values()],default=0)+1)
        self.identity(state,dt,day)
        if dt not in state['plans']:
            raise ValueError('该日期无已确认周计划；先制定该周计划')
        baseline=Path(self.root/state['config']['profile']).read_text(encoding='utf-8')
        due=sorted((e for e in self.bank(state,dt).values() if e['next_due'] and e['next_due']<=dt),key=lambda e:(e['next_due'],-e['error_count']))
        return {'date':dt,'day':day,'already_registered':existing,'latest_profile_path':state['config']['profile'],'profile_text':baseline,'latest_weekly_assessment':state['weekly'][max(state['weekly'])] if state['weekly'] else None,'plan':state['plans'][dt], 'recent_feedback':[state['submissions'][d] for d in sorted(state['submissions'])[-7:] if d<=dt], 'due_reviews':due,'worksheet_skill':state['config']['worksheet_skill'],'reference_sha256':state['config']['reference_sha256'],'word_mode':'review' if date.fromisoformat(dt).weekday()>=5 else 'new','output_path':f"{state['config']['worksheet_dir']}/{dt.replace('-','')}_English worksheet_Day{day}.docx",'existing_words':sorted({w for lesson in state['lessons'].values() for w in lesson['words']})}

    def register(self,payload):
        old=self.load(); state=copy.deepcopy(old)
        dt=iso(payload['date']); day=payload['day']; self.identity(state,dt,day)
        if dt not in state['plans']:
            raise ValueError('无对应学习计划')
        candidate=Path(payload['file']).resolve(); data=candidate.read_bytes()
        expected=f"{dt.replace('-','')}_English worksheet_Day{day}.docx"
        if candidate.name!=expected:
            raise ValueError('Worksheet文件名与日期/Day不一致')
        skill=Path(state['config']['worksheet_skill'])
        spec=json.loads((skill/'references/design-spec.json').read_text(encoding='utf-8'))
        if digest((skill/spec['reference_file']).read_bytes())!=state['config']['reference_sha256']:
            raise ValueError('保留参考已变更，需重新检查规范')
        module_spec=importlib.util.spec_from_file_location('worksheet_validator',skill/'scripts/validate_lesson.py')
        module=importlib.util.module_from_spec(module_spec);module_spec.loader.exec_module(module)
        validation=module.check_lesson(candidate)
        if not validation['passed']:
            raise ValueError('现有Worksheet Skill结构校验失败: '+str(validation['errors']))
        qa=payload.get('qa',{})
        if qa.get('visual') not in {'passed','unavailable'}:
            raise ValueError('必须明确逐页视觉检查结果')
        if qa['visual']=='passed' and (not qa.get('pages') or not qa.get('render_path')):
            raise ValueError('视觉通过必须注明页数及实际渲染目录')
        blueprint=payload['blueprint']
        if not blueprint or len({q['id'] for q in blueprint})!=len(blueprint):
            raise ValueError('必须给出完整且无重复的可批改题目清单')
        for q in blueprint:
            check_id(q['id'])
            if not isinstance(q['points'],(int,float)) or isinstance(q['points'],bool) or not 0<q['points']<=100:
                raise ValueError('题目满分必须为正数')
            if not q.get('prompt'):
                raise ValueError('题目清单缺少题干')
        words=self.words(package(data))
        if len(set(words))!=20:
            raise ValueError('20主词必须互不重复')
        mode=payload.get('word_mode','review' if date.fromisoformat(dt).weekday()>=5 else 'new')
        if mode not in {'new','review'}:
            raise ValueError('word_mode应为new/review')
        prior={w for d,lesson in state['lessons'].items() if d!=dt for w in lesson['words']}
        if mode=='new' and prior.intersection(words):
            raise ValueError('新词与既有词重复；复习词不可重复记作新词')
        if date.fromisoformat(dt).weekday()>=5 and (mode!='review' or not set(words)<=prior):
            raise ValueError('周末只复习已有词，不添加新词')
        reviews=payload.get('review_question_map',{})
        for kid,qids in reviews.items():
            if kid not in state['knowledge'] or not qids or not set(qids)<={q['id'] for q in blueprint}:
                raise ValueError('错题复习必须映射到实际题号')
        existing=state['lessons'].get(dt)
        if existing and digest(data)!=existing['sha256'] and dt in state['submissions']:
            raise ValueError('已有作答的试卷不能替换题目；可更新批改，不改历史题干')
        record={'date':dt,'day':day,'path':f"{state['config']['worksheet_dir']}/{expected}",'sha256':digest(data),'status':'ready' if qa['visual']=='passed' else 'draft','blueprint':blueprint,'words':words,'word_mode':mode,'review_question_map':reviews,'qa':{**qa,'structural':validation}}
        state['lessons'][dt]=record
        from vocabulary_summary import prepare_summary
        summary,summary_files,changed=prepare_summary(self,state,{record['path']:data})
        if existing==record and not changed:
            return {'action':'unchanged','date':dt}
        state['vocabulary_summary']=summary
        return self.commit(old,state,'register_worksheet',{record['path']:data,**summary_files})

    def grade(self,payload):
        old=self.load();state=copy.deepcopy(old);dt=iso(payload['date'])
        lesson=state['lessons'].get(dt)
        if not lesson or not lesson['blueprint']:
            raise ValueError('先登记实际试卷和完整题目清单，不能凭空批改')
        if digest(self.safe(lesson['path']).read_bytes())!=lesson['sha256']:
            raise ValueError('试卷已外部修改，先核对题目清单')
        for k in payload.get('knowledge',[]):
            self.knowledge(state,k['id'],k['label'],k['category'])
        blueprint={q['id']:q for q in lesson['blueprint']}
        items=payload['items']
        if len({q['id'] for q in items})!=len(items) or set(blueprint)!={q['id'] for q in items}:
            raise ValueError('需逐题覆盖题目清单；未收到/不清楚的答案标pending，明确空答标unanswered')
        for item in items:
            q=blueprint[item['id']]
            item['prompt']=q['prompt'];item['points']=q['points']
            result=item['result']
            if result not in {'correct','incorrect','partial','pending','unanswered'}:
                raise ValueError('未知批改状态')
            if 'original_answer' not in item or not isinstance(item['original_answer'],(str,type(None))):
                raise ValueError('必须保留原始答案文本或null，不可改写后存储')
            if result=='pending':
                if not item.get('pending_reason') or item.get('score') is not None or item.get('issues') or item.get('checks'):
                    raise ValueError('待确认题须注明原因，不给分、不进入知识点/错题统计')
                continue
            score=item.get('score')
            if not isinstance(score,(float,int)) or isinstance(score,bool) or not 0<=score<=q['points']:
                raise ValueError('分数不在有效范围')
            if (result=='correct' and score!=q['points']) or (result in {'incorrect','unanswered'} and score!=0) or (result=='partial' and not 0<score<q['points']):
                raise ValueError('判定与分数矛盾')
            if result not in {'unanswered'} and not item.get('original_answer'):
                raise ValueError('可判定题必须保留可辨认原答案')
            if not item.get('expected_answer') or not item.get('explanation'):
                raise ValueError('每题需正确答案和讲解')
            checks=item.get('checks',[]);issues=item.get('issues',[])
            if len({x['knowledge_id'] for x in checks})!=len(checks) or len({x['knowledge_id'] for x in issues})!=len(issues):
                raise ValueError('同题同知识点不可重复计数')
            for check in checks:
                if check['knowledge_id'] not in state['knowledge'] or type(check['success']) is not bool or type(check['independent']) is not bool or check['mode'] not in {'practice','review'}:
                    raise ValueError('无效知识点检测证据')
            if result in {'incorrect','partial'} and not issues:
                raise ValueError('错误题必须归类并给出针对练习')
            if result in {'incorrect','partial'} and not any(not c['success'] for c in checks):
                raise ValueError('错误题缺少失败知识点检测')
            if result=='correct' and (issues or any(not c['success'] for c in checks)):
                raise ValueError('正确题不能同时记录错误')
            for issue in issues:
                if issue['knowledge_id'] not in {c['knowledge_id'] for c in checks if not c['success']} or not all(issue.get(k) for k in ['reason','explanation','practice']):
                    raise ValueError('每个错因需要失败检测、原因、讲解和针对性练习')
            if {i['knowledge_id'] for i in issues}!={c['knowledge_id'] for c in checks if not c['success']}:
                raise ValueError('失败知识点与错因列表必须一一对应')
            if result=='unanswered' and (checks or issues or item['original_answer'] not in {'',None}):
                raise ValueError('确认空答计零分，但不能推断知识点未掌握')
        canonical=copy.deepcopy(payload)
        canonical.pop('source_files',None)
        evidence={};evidence_refs=[]
        for name in payload.get('source_files',[]):
            source=Path(name).resolve();data=source.read_bytes()
            rel=f".learning-manager/evidence/{dt}/{digest(data)[:16]}-{source.name}"
            evidence[rel]=data;evidence_refs.append({'path':rel,'sha256':digest(data)})
        canonical['evidence']=evidence_refs
        canonical['items']=items
        fingerprint=digest(jd(canonical))
        previous=state['submissions'].get(dt)
        if previous and previous['fingerprint']==fingerprint:
            return {'action':'unchanged','date':dt}
        state['submissions'][dt]={**canonical,'fingerprint':fingerprint,'revision':previous['revision']+1 if previous else 1,'graded_at':now()}
        # All aggregates are rebuilt from latest submissions. Old revisions
        # survive in transaction snapshots without contributing to totals.
        state['weekly']={week:self.week_stats(state,week) for week in state['weekly']}
        return self.commit(old,state,'grade_or_regrade',evidence)

    def patch_plan(self,state,dates):
        rel=state['config']['plan'];data=self.safe(rel).read_bytes()
        if digest(data)!=state['sources'][rel]:
            raise ValueError('周计划Word有外部修改，请先人工合并再同步，禁止覆盖')
        root=package(data);tables=root.findall('w:body/w:tbl',NS)
        for dt in dates:
            plan=state['plans'][dt];rows=tables[plan['table']].findall('w:tr',NS)
            r1,r2=rows[1+plan['day_index']*2:3+plan['day_index']*2]
            c1,c2=r1.findall('w:tc',NS),r2.findall('w:tc',NS)
            for cell,value in zip([c1[1],c1[2],c2[1],c2[2]],plan['current']):
                if text(cell)!=value:
                    cell_text(cell,value)
            for node in c1[-1].findall('.//w14:checked',NS):
                node.set(f'{{{NS["w14"]}}}val','1' if plan['status']=='completed' else '0')
            for node in c1[-1].findall('.//w:sdtContent//w:t',NS):
                node.text='✔' if plan['status']=='completed' else '☐'
        result=repack(data,ET.tostring(root,encoding='utf-8',xml_declaration=True))
        state['sources'][rel]=digest(result)
        return {rel:result}

    def adjust(self,payload):
        old=self.load();state=copy.deepcopy(old);changes=payload.get('changes',[payload]);dates=[]
        for change in changes:
            dt=iso(change['date']);plan=state['plans'].get(dt)
            if not plan:
                raise ValueError('调整日期无对应周表')
            status=change.get('status',plan['status'])
            slots=change.get('slots',plan['current']);minutes=change.get('minutes',plan['minutes'])
            actual=change.get('actual',plan['actual']);reason=change.get('reason','')
            if not reason:
                raise ValueError('调整或执行记录必须有原因')
            if plan['status']=='completed':
                if slots!=plan['current'] or status!='completed' or actual!=plan['actual'] or minutes!=plan['minutes']:
                    raise ValueError('已完成的历史任务冻结；不得改写原任务或实际记录')
                continue
            if len(slots)!=4 or not all(isinstance(s,str) and s.strip() for s in slots) or not isinstance(minutes,int) or isinstance(minutes,bool) or minutes<=0:
                raise ValueError('沿用周表四训练格与正整数分钟数')
            validate_slots(slots,minutes)
            if status not in {'planned','unconfirmed','in_progress','completed','partial','missed'}:
                raise ValueError('无效完成情况')
            if status in {'completed','partial','missed'} and not isinstance(actual,dict):
                raise ValueError('执行情况须记录actual: minutes/content/completed_tasks/assigned_tasks')
            if isinstance(actual,dict):
                if not isinstance(actual.get('minutes'),int) or actual['minutes']<0 or not actual.get('content') or not isinstance(actual.get('assigned_tasks'),int) or not isinstance(actual.get('completed_tasks'),int) or not 0<=actual['completed_tasks']<=actual['assigned_tasks'] or actual['assigned_tasks']<=0:
                    raise ValueError('实际执行时间/任务数/内容无效')
                if status=='completed' and actual['completed_tasks']!=actual['assigned_tasks']:
                    raise ValueError('全部完成状态与任务数矛盾')
                if status=='missed' and (actual['minutes'] or actual['completed_tasks']):
                    raise ValueError('未学习状态应为0分钟、0完成任务')
            if (slots,minutes,status,actual)==(plan['current'],plan['minutes'],plan['status'],plan['actual']):
                continue
            event={'date':dt,'at':now(),'reason':reason,'before':{k:copy.deepcopy(plan[k]) for k in ['current','minutes','status','actual']},'after':{'current':slots,'minutes':minutes,'status':status,'actual':actual}}
            plan['current']=slots;plan['minutes']=minutes;plan['status']=status;plan['actual']=actual;plan['reason']=reason;plan['history'].append(event);state['adjustments'].append(event);dates.append(dt)
        if not dates:
            return {'action':'unchanged'}
        files=self.patch_plan(state,dates)
        state['weekly']={week:self.week_stats(state,week) for week in state['weekly']}
        return self.commit(old,state,'adjust_plan_or_execution',files)

    def plan_week(self,payload):
        """Apply an assessed next week; extend the same Word only if absent."""
        old=self.load();state=copy.deepcopy(old);start=iso(payload['week_start'])
        if date.fromisoformat(start).weekday()!=0 or len(payload['days'])!=7 or not payload.get('reason'):
            raise ValueError('提供周一起点、7天计划及调整依据')
        dates=[str(date.fromisoformat(start)+timedelta(days=i)) for i in range(7)]
        if {d['date'] for d in payload['days']}!=set(dates):
            raise ValueError('7天日期须完整且不重复')
        present=[d in state['plans'] for d in dates]
        if all(present):
            return self.adjust({'changes':[{**d,'reason':payload['reason']} for d in payload['days']]})
        if any(present):
            raise ValueError('已有部分周计划，请人工核对结构，禁止重复建周')
        if start<=max(state['plans']):
            raise ValueError('只能在既有计划之后延展，不补写历史计划')
        previous=str(date.fromisoformat(start)-timedelta(days=7))
        if previous not in state['weekly']:
            raise ValueError('先汇总前一周实际证据，再制定下一周')
        rel=state['config']['plan'];data=self.safe(rel).read_bytes()
        if digest(data)!=state['sources'][rel]:
            raise ValueError('周计划Word已外部修改，禁止覆盖')
        root=package(data);body=root.find('w:body',NS);table=copy.deepcopy(body.findall('w:tbl',NS)[0]);rows=table.findall('w:tr',NS)
        title=copy.deepcopy(body.findall('w:p',NS)[0]);cell_text(title,f"第{(date.fromisoformat(start)-date.fromisoformat(state['config']['week_start'])).days//7+1}周  |  {start}—{dates[-1]}")
        ppr=title.find('w:pPr',NS)
        if ppr is None:ppr=ET.SubElement(title,f'{{{W}}}pPr')
        if ppr.find('w:pageBreakBefore',NS) is None:ET.SubElement(ppr,f'{{{W}}}pageBreakBefore')
        goal=copy.deepcopy(body.findall('w:p',NS)[1]);cell_text(goal,'周目标｜'+payload['goal'])
        criteria=copy.deepcopy(body.findall('w:p',NS)[2]);cell_text(criteria,'验收标准｜'+payload['criteria'])
        ti=len(body.findall('w:tbl',NS))
        existing_ids=[int(n.get(f'{{{W}}}val')) for n in root.findall('.//w:sdtPr/w:id',NS) if n.get(f'{{{W}}}val','').lstrip('-').isdigit()]
        next_id=max(existing_ids,default=0)+1
        for node in table.findall('.//w:sdtPr/w:id',NS):
            node.set(f'{{{W}}}val',str(next_id));next_id+=1
        for i,dt in enumerate(dates):
            d=next(d for d in payload['days'] if d['date']==dt);minutes=d.get('minutes',60);validate_slots(d['slots'],minutes)
            if minutes!=60 and not d.get('time_override'):
                raise ValueError('非60分钟须来自用户明确的时长变更，注明time_override')
            c1=rows[1+i*2].findall('w:tc',NS);c2=rows[2+i*2].findall('w:tc',NS)
            cell_text(c1[0],['周一','周二','周三','周四','周五','周六','周日'][i]+' '+dt[5:].replace('-','.'))
            for cell,value in zip([c1[1],c1[2],c2[1],c2[2]],d['slots']):cell_text(cell,value)
            for node in c1[-1].findall('.//w14:checked',NS):node.set(f'{{{NS["w14"]}}}val','0')
            for node in c1[-1].findall('.//w:sdtContent//w:t',NS):node.text='☐'
            state['plans'][dt]={'date':dt,'table':ti,'day_index':i,'original':d['slots'][:],'current':d['slots'][:],'minutes':minutes,'status':'planned','actual':None,'reason':payload['reason'],'history':[]}
        # Keep earlier tables and all their checkbox controls untouched.
        for element in [title,goal,criteria,table]:
            sect=body.find('w:sectPr',NS)
            body.insert(body.index(sect) if sect is not None else len(body),element)
        result=repack(data,ET.tostring(root,encoding='utf-8',xml_declaration=True))
        state['sources'][rel]=digest(result)
        state['adjustments'].append({'at':now(),'week_start':start,'reason':payload['reason'],'created_dates':dates})
        return self.commit(old,state,'extend_week_plan',{rel:result})

    def week_stats(self,state,start):
        start=iso(start);end=str(date.fromisoformat(start)+timedelta(days=6))
        subs=[sub for dt,sub in state['submissions'].items() if start<=dt<=end]
        items=[item for sub in subs for item in sub['items']]
        graded=[q for q in items if q['result']!='pending']
        total=sum(q['points'] for q in graded);score=sum(q['score'] for q in graded)
        determined=[q for q in graded if q['result']!='unanswered']
        plans=[plan for dt,plan in state['plans'].items() if start<=dt<=end]
        known=[p for p in plans if p['actual'] is not None]
        denom=sum(p['actual']['assigned_tasks'] for p in known)
        complete=sum(p['actual']['completed_tasks'] for p in known)
        distributed={}
        for e in self.observations(state,end):
            if start<=e['date'] and not e['success']:
                cat=state['knowledge'][e['knowledge_id']]['category'];distributed[cat]=distributed.get(cat,0)+1
        bank=self.bank(state,end)
        return {'week_start':start,'week_end':end,'updated_at':now(),'submission_days':len(subs),'total_questions':len(items),'graded_questions':len(graded),'pending_questions':len(items)-len(graded),'unanswered_questions':sum(q['result']=='unanswered' for q in items),'score':score,'possible_points':total,'score_rate':score/total if total else None,'accuracy':sum(q['result']=='correct' for q in determined)/len(determined) if determined else None,'errors_by_category':distributed,'planned_days':len(plans),'completed_days':sum(p['status']=='completed' for p in plans),'unknown_execution_days':sum(p['status'] in {'unconfirmed','planned'} for p in plans),'confirmed_task_completion_rate':complete/denom if denom else None,'confirmed_task_counts':{'completed':complete,'assigned':denom},'actual_minutes':sum(p['actual']['minutes'] for p in known),'knowledge':{k:{'label':v['label'],'state':v['mastery'],'last_detected':v['last_detected']} for k,v in bank.items()},'limits':'仅按实际可判定证据统计；初始摸底不混入本周正确率；未确认执行不当作未完成；听说未测不估分。'}

    def weekly(self,start):
        old=self.load();state=copy.deepcopy(old);start=iso(start)
        if date.fromisoformat(start).weekday()!=0:
            raise ValueError('周起点必须为周一')
        if not any(start<=dt<=str(date.fromisoformat(start)+timedelta(days=6)) for dt in state['plans']):
            raise ValueError('该周没有计划')
        stats=self.week_stats(state,start)
        comparable=lambda x:{k:v for k,v in x.items() if k!='updated_at'}
        if start in state['weekly'] and comparable(state['weekly'][start])==comparable(stats):
            return {'action':'unchanged','week_start':start}
        state['weekly'][start]=stats
        return self.commit(old,state,'weekly_assessment')

    def points(self):
        state=self.load()
        return state.get('key_points',{'document_path':'06_Learning Notes/English Learning_Key Points.docx','entries':{}})

    def vocabulary(self):
        return self.load().get('vocabulary_summary',{'document_path':'06_Learning Notes/English Learning_Vocabulary Summary.docx','total_unique_words':0,'days':{}})

    def summarize_vocabulary(self):
        from vocabulary_summary import prepare_summary
        old=self.load();state=copy.deepcopy(old)
        summary,files,changed=prepare_summary(self,state)
        if not changed:return {'action':'unchanged','file':summary['document_path'],'total_unique_words':summary['total_unique_words']}
        state['vocabulary_summary']=summary
        result=self.commit(old,state,'summarize_vocabulary',files)
        result.update(file=summary['document_path'],total_unique_words=summary['total_unique_words'])
        return result

    def record_points(self,payload):
        from key_points import prepare_record
        old=self.load()
        points,files,summary=prepare_record(self,old,payload)
        if points is None:return summary
        state=copy.deepcopy(old);state['key_points']=points
        result=self.commit(old,state,'record_key_points',files)
        result.update(summary)
        return result

    def progress(self,dt=None):
        state=self.load();dt=iso(dt or today());bank=self.bank(state,dt)
        return {'as_of':dt,'revision':state['revision'],'registered_worksheets':len(state['lessons']),'graded_days':len(state['submissions']),'completed_days':sum(p['status']=='completed' for p in state['plans'].values()),'mastery_counts':{s:sum(e['mastery']==s for e in bank.values()) for s in STATES},'due_reviews':[e for e in bank.values() if e['next_due'] and e['next_due']<=dt],'latest_assessment':state['weekly'][max(state['weekly'])] if state['weekly'] else None}

    def views(self,state):
        files={}
        bank=self.bank(state)
        files['05_Error Bank/Error Bank.json']=jd({'schema_version':1,'revision':state['revision'],'entries':bank})
        lines=['# 错题与间隔复习','', '从现有档案和最新批改记录生成。历史档案未注明发生日期的错误不虚构时间。', '', '| 知识点 | 分类 | 错误次数 | 最近检测 | 下次复习 | 状态 |','|---|---|---:|---|---|---|']
        for e in bank.values():
            lines.append(f"| {e['label']} | {e['category']} | {e['error_count']} | {e['last_detected'] or '历史日期未知/尚未检测'} | {e['next_due'] or '尚无错题'} | {e['mastery']} |")
        files['05_Error Bank/Error Bank.md']=('\n'.join(lines)+'\n').encode('utf-8')
        for dt,sub in state['submissions'].items():
            lesson=state['lessons'][dt]
            lines=[f"# {dt} Day {lesson['day']} 学习反馈",'',f"当前批改版本：{sub['revision']}。保留原始输入；本反馈由当前记录生成，历史版本在事务备份中。",'']
            for q in sub['items']:
                lines.extend([f"## {q['id']} {q['prompt']}",'',f"原答案：{q['original_answer'] if q['original_answer'] is not None else '无法辨认或未提供'}",f"判定：{q['result']}；得分：{q.get('score') if q.get('score') is not None else '待确认'}/{q['points']}",f"正确/参考答案：{q.get('expected_answer') or '待核对'}",f"讲解：{q.get('explanation') or q.get('pending_reason')}",''])
                for issue in q.get('issues',[]):
                    lines.extend([f"知识点：{issue['knowledge_id']}；错因：{issue['reason']}",f"说明：{issue['explanation']}",f"针对练习：{issue['practice']}",''])
            lines.extend(['下一次重点：根据当前到期错题，优先安排重复错误的主动回忆与独立造句，计入当天既定时长。',''])
            files[f"04_Learning Feedback/Daily/{dt.replace('-','')}_Learning Feedback_Day{lesson['day']}.md"]='\n'.join(lines).encode('utf-8')
        lines=['# 周计划执行与调整记录','','原计划永久保留；此文件记录实际内容、原因和完成情况，周计划Word仍为同一份文件。','']
        for dt,p in sorted(state['plans'].items()):
            if p['history']:
                lines.extend([f'## {dt}', '', '原计划：'+'；'.join(p['original']), '当前计划：'+'；'.join(p['current']),f"状态：{p['status']}；计划时间：{p['minutes']}分钟",'实际执行：'+json.dumps(p['actual'],ensure_ascii=False),''])
                for event in p['history']:
                    lines.append(event['at']+'｜'+event['reason']+'｜'+json.dumps(event['after'],ensure_ascii=False))
                lines.append('')
        files['02_Weekly Study Plan/Execution Log.md']='\n'.join(lines).encode('utf-8')
        for start,stats in state['weekly'].items():
            lines=[f"# {start} 至 {stats['week_end']} 每周能力评估",'',f"实际学习 {stats['actual_minutes']} 分钟；提交 {stats['submission_days']} 天；完成 {stats['completed_days']}/{stats['planned_days']} 天。", f"得分率：{stats['score_rate']:.1%}" if stats['score_rate'] is not None else '得分率：未评估',f"正确率：{stats['accuracy']:.1%}" if stats['accuracy'] is not None else '正确率：未评估',f"待确认 {stats['pending_questions']} 题；未确认执行 {stats['unknown_execution_days']} 天。",f"已确认任务完成率：{stats['confirmed_task_completion_rate']:.1%}" if stats['confirmed_task_completion_rate'] is not None else '任务完成率：尚无执行证据', '', '错误分布：'+json.dumps(stats['errors_by_category'],ensure_ascii=False),'', '| 知识点 | 掌握状态 | 最近检测 |','|---|---|---|']
            for k,v in stats['knowledge'].items():
                lines.append(f"| {v['label']} | {v['state']} | {v['last_detected'] or '未新增检测'} |")
            lines.extend(['',stats['limits'],'','下周建议：优先从未掌握/不稳定的重复错误中选择1—2个主线，在原60分钟内替换训练内容；听说另需真实材料诊断。实际计划变更使用adjust，保留原计划。',''])
            files[f"01_English Proficiency Profile/Weekly Assessments/{start.replace('-','')}_Assessment.md"]='\n'.join(lines).encode('utf-8')
            files[f"04_Learning Feedback/Weekly/{start.replace('-','')}_Weekly Summary.json"]=jd(stats)
        if state['weekly']:
            rel=state['config']['profile']
            original=self.safe(rel).read_text(encoding='utf-8')
            marker='\n<!-- english-learning-manager:start -->'
            # Only the managed footer changes; original diagnostic prose stays.
            latest=max(state['weekly'])
            footer=f"{marker}\n## English-learning-manager 最新评估\n\n最新评估周：{latest}。完整记录见 [每周评估](Weekly%20Assessments/{latest.replace('-','')}_Assessment.md)。\n\n初始评估正文保留；所有每周评估按周归档，修订前版本保存在 .learning-manager/transactions。\n<!-- english-learning-manager:end -->\n"
            pattern=r'\n<!-- english-learning-manager:start -->.*?<!-- english-learning-manager:end -->\n?'
            if marker in original:
                if '<!-- english-learning-manager:end -->' not in original:
                    raise ValueError('档案管理区标记不完整，保留现场')
                updated=re.sub(pattern,lambda _:footer,original,count=1,flags=re.DOTALL)
            else:
                updated=original+'\n'+footer
            files[rel]=updated.encode('utf-8')
        return files

    def audit(self):
        state=self.load();errors=[]
        if state.get('vocabulary_summary'):
            from vocabulary_summary import audit_summary
            errors.extend(audit_summary(self,state))
        if state.get('key_points'):
            from key_points import audit_points
            errors.extend(audit_points(self,state))
        rel=state['config']['plan']
        if not self.safe(rel).is_file() or digest(self.safe(rel).read_bytes())!=state['sources'][rel]:
            errors.append('周计划Word丢失/外部修改')
        days=[l['day'] for l in state['lessons'].values()]
        if len(days)!=len(set(days)):
            errors.append('重复Day')
        for dt,lesson in state['lessons'].items():
            p=self.safe(lesson['path'])
            if not p.is_file() or digest(p.read_bytes())!=lesson['sha256']:
                errors.append('试卷丢失/外部修改: '+dt)
        for dt,sub in state['submissions'].items():
            if dt not in state['lessons'] or len(sub['items'])!=len({q['id'] for q in sub['items']}):
                errors.append('作答身份冲突: '+dt)
            for e in sub.get('evidence',[]):
                p=self.safe(e['path'])
                if not p.exists() or digest(p.read_bytes())!=e['sha256']:
                    errors.append('原始作答证据损坏: '+e['path'])
        for rel,data in self.views(copy.deepcopy(state)).items():
            if not self.safe(rel).is_file() or self.safe(rel).read_bytes()!=data:
                errors.append('派生文件不同步: '+rel)
        for p in (self.home/'transactions').glob('*/manifest.json'):
            manifest=json.loads(p.read_text(encoding='utf-8'))
            for e in manifest['files']:
                for suffix,key in [('after','after_sha256'),('before','before_sha256')]:
                    if e[key] is not None:
                        snap=p.parent/f"{e['index']}.{suffix}"
                        if not snap.is_file() or digest(snap.read_bytes())!=e[key]:
                            errors.append('备份哈希失败: '+str(snap))
        baseline=self.home/'baseline/source-inventory.json'
        if baseline.is_file():
            originals=json.loads(baseline.read_text(encoding='utf-8'))
            for source_rel,snapshot_rel in [(state['config']['profile'],'profile.md'),(state['config']['plan'],'weekly-plan.docx')]:
                snap=self.home/'baseline'/snapshot_rel
                if not snap.is_file() or digest(snap.read_bytes())!=originals[source_rel]:
                    errors.append('初始档案/周计划备份损坏: '+source_rel)
        return {'passed':not errors,'errors':errors,'revision':state['revision'],'lessons':len(state['lessons']),'submission_days':len(state['submissions']),'baseline_error_observations':len(state['baseline_errors']),'historical_completed_tasks_frozen':True}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root',default=str(Path(__file__).resolve().parents[3]))
    sub=parser.add_subparsers(dest='command',required=True)
    init=sub.add_parser('init');init.add_argument('--worksheet-skill',required=True)
    for command in ['context','progress','due']:
        p=sub.add_parser(command);p.add_argument('--date',default=None)
        if command=='context':p.add_argument('--day',type=int)
    for command in ['register','grade','adjust','plan-week','record-points']:
        p=sub.add_parser(command);p.add_argument('--input',required=True)
    p=sub.add_parser('weekly');p.add_argument('--start',required=True)
    sub.add_parser('audit');sub.add_parser('recover');sub.add_parser('points');sub.add_parser('vocabulary');sub.add_parser('summarize-vocabulary')
    args=parser.parse_args();manager=Manager(args.root)
    if args.command=='init':result=manager.bootstrap(args.worksheet_skill)
    elif args.command=='context':result=manager.context(args.date,args.day)
    elif args.command=='progress':result=manager.progress(args.date)
    elif args.command=='due':result={'as_of':args.date or today(),'due_reviews':manager.progress(args.date)['due_reviews']}
    elif args.command=='weekly':result=manager.weekly(args.start)
    elif args.command=='audit':result=manager.audit()
    elif args.command=='recover':result=manager.recover()
    elif args.command=='points':result=manager.points()
    elif args.command=='vocabulary':result=manager.vocabulary()
    elif args.command=='summarize-vocabulary':result=manager.summarize_vocabulary()
    else:
        payload=json.loads(Path(args.input).read_text(encoding='utf-8'))
        result=getattr(manager,args.command.replace('-','_'))(payload)
    print(json.dumps(result,ensure_ascii=False,indent=2))
    if args.command=='audit' and not result['passed']:
        raise SystemExit(1)


if __name__=='__main__':
    try:
        main()
    except (ValueError,KeyError,FileExistsError,FileNotFoundError) as error:
        print(json.dumps({'error':str(error)},ensure_ascii=False),file=sys.stderr)
        raise SystemExit(2)
