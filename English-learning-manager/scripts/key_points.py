"""Personal key points: categorized notes, distinct from graded evidence."""
from copy import deepcopy
from datetime import date
from pathlib import Path
from zipfile import ZipFile
import hashlib,io,re
from lxml import etree as ET

DOCUMENT='06_Learning Notes/English Learning_Key Points.docx'
CATEGORIES=('词汇与搭配','语法与句型','易混辨析','常见错误')
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
N={'w':W}

def digest(data):return hashlib.sha256(data).hexdigest()
def required(value,label):
    if not isinstance(value,str) or not value.strip():raise ValueError(label+'不能为空')
    return value.strip()
def canonical(pattern):return re.sub(r'\s+',' ',pattern.strip()).casefold()
def unique(values):
    result=[]
    for value in values:
        if value not in result:result.append(deepcopy(value))
    return result

def validate_store(store):
    if store.get('schema_version')!=1 or store.get('document_path')!=DOCUMENT:
        raise ValueError('重点数据版本或唯一文件路径不正确')
    entries=store['entries'];patterns=set()
    for ident,item in entries.items():
        if item['id']!=ident or not re.fullmatch('[a-z0-9][a-z0-9_.-]{0,95}',ident):raise ValueError('重点标识无效')
        pattern=canonical(required(item['pattern'],'搭配'))
        if pattern in patterns:raise ValueError('相同重点搭配不得使用多个标识')
        patterns.add(pattern)
        category=item['category']
        if not isinstance(category,list) or not 2<=len(category)<=3 or category[0] not in CATEGORIES:
            raise ValueError('重点需要有效大类与子分类')
        for c in category:required(c,'分类')
        required(item['meaning'],'中文含义')
        if not isinstance(item['examples'],list) or not item['examples']:raise ValueError('重点缺少双语例句')
        for e in item['examples']:
            required(e['english'],'英文例句');required(e['chinese'],'例句翻译')
        if not isinstance(item['notes'],list):raise ValueError('用法提醒须为列表')
        for note in item['notes']:required(note,'用法提醒')
        for field in ['first_recorded_date','last_updated_date']:date.fromisoformat(item[field])
        if not item['originals']:raise ValueError('重点缺少用户原文')
        for original in item['originals']:
            required(original['text'],'原文');date.fromisoformat(original['date'])
            if original['source_path'] not in store['sources']:raise ValueError('重点原文来源未归档')

def build_document(store,existing=None):
    """Preserve styles and every non-document.xml ZIP part on later updates."""
    from docx import Document
    from docx.shared import Inches,Pt,RGBColor
    from docx.oxml.ns import qn
    doc=Document(io.BytesIO(existing)) if existing else Document()
    if not existing:
        sec=doc.sections[0]
        sec.page_width=Inches(8.5);sec.page_height=Inches(11)
        sec.top_margin=sec.bottom_margin=Inches(.65)
        sec.left_margin=sec.right_margin=Inches(.75)
        for name,size in [('Normal',10.5),('Title',22),('Heading 1',14),('Heading 2',12),('Heading 3',11.5)]:
            style=doc.styles[name]
            style.font.name='Times New Roman';style.font.size=Pt(size)
            style.font.color.rgb=RGBColor(0,0,0)
            fonts=style.element.get_or_add_rPr().get_or_add_rFonts()
            fonts.set(qn('w:eastAsia'),'宋体')
            fonts.attrib.pop(qn('w:asciiTheme'),None);fonts.attrib.pop(qn('w:hAnsiTheme'),None)
            fonts.attrib.pop(qn('w:eastAsiaTheme'),None)
            pf=style.paragraph_format
            pf.space_before=Pt(0 if name in ['Normal','Title'] else 9)
            pf.space_after=Pt(3 if name=='Normal' else 5)
            pf.line_spacing=1.12
            if name!='Normal':style.font.bold=True
            for border in list(style.element.findall('.//'+qn('w:pBdr'))):border.getparent().remove(border)
        doc.core_properties.title='英语学习重点'
        doc.core_properties.subject='分类整理个人英语学习重点'
    body=doc.element.body
    for child in list(body):
        if child.tag!=qn('w:sectPr'):body.remove(child)
    doc.add_paragraph('英语学习重点','Title')
    entries=store['entries']
    doc.add_paragraph('按知识点整理重要搭配、含义、双语例句和用法提醒，便于回顾与查找。')
    latest=max(i['last_updated_date'] for i in entries.values())
    p=doc.add_paragraph(f'共 {len(entries)} 条重点    最近更新 {latest}')
    p.runs[0].font.size=Pt(9)
    # Empty categories are omitted; category ordering remains consistent.
    groups={}
    for item in entries.values():groups.setdefault(tuple(item['category']),[]).append(item)
    last_category=None
    for category in sorted(groups,key=lambda c:(CATEGORIES.index(c[0]),c[1:])):
        if last_category!=category[0]:
            doc.add_paragraph(category[0],'Heading 1');last_category=category[0]
        doc.add_paragraph('  /  '.join(category[1:]),'Heading 2')
        for item in groups[category]:
            doc.add_paragraph(item['pattern'],'Heading 3')
            lines=[('含义',item['meaning'])]
            display=store.get('display_fields',{})
            if display.get('original',True):
                for original in item['originals']:lines.append(('原文',original['text']))
            for example in item['examples']:
                lines.append(('例句',example['english']))
                if display.get('translation',True):lines.append(('翻译',example['chinese']))
            for note in item['notes']:lines.append(('提醒',note))
            if display.get('first_recorded_date',True):lines.append(('首次收录',item['first_recorded_date']))
            for i,(label,value) in enumerate(lines):
                p=doc.add_paragraph()
                p.add_run(label+'：').bold=True;p.add_run(value)
                p.paragraph_format.keep_with_next=i<len(lines)-1
                p.paragraph_format.keep_together=True
                if label=='首次收录':
                    p.paragraph_format.space_after=Pt(7)
                    for run in p.runs:run.font.size=Pt(8.5)
    out=io.BytesIO();doc.save(out);generated=out.getvalue()
    if not existing:return generated
    document=ZipFile(io.BytesIO(generated)).read('word/document.xml')
    out=io.BytesIO()
    with ZipFile(io.BytesIO(existing)) as source,ZipFile(out,'w') as target:
        for entry in source.infolist():target.writestr(entry,document if entry.filename=='word/document.xml' else source.read(entry.filename))
    return out.getvalue()

def prepare_record(manager,old,payload):
    """Validate, merge and render in memory; Manager.commit persists atomically."""
    dt=date.fromisoformat(required(payload.get('date'),'收录日期')).isoformat()
    raw=payload.get('raw_input');required(raw,'用户原始输入')
    incoming=payload.get('entries')
    if not isinstance(incoming,list) or not incoming:raise ValueError('至少提供一条重点')
    current=old.get('key_points')
    store=deepcopy(current) if current else {'schema_version':1,'document_path':DOCUMENT,'entries':{},'sources':{}}
    existing=None;dest=manager.safe(DOCUMENT)
    if current:
        validate_store(store)
        if not dest.is_file() or digest(dest.read_bytes())!=store['document_sha256']:
            raise ValueError('重点Word丢失或被外部修改；先读取并合并，不能覆盖或直接重置校验值')
        existing=dest.read_bytes()
        # Never hide damaged original evidence by appending new notes.
        for rel,checksum in store['sources'].items():
            p=manager.safe(rel)
            if not p.is_file() or digest(p.read_bytes())!=checksum:raise ValueError('重点原始输入丢失或被修改')
    elif dest.exists():raise ValueError('已有未登记的重点Word，先读取合并，不得覆盖')
    source=f'.learning-manager/key-points/sources/{dt.replace("-","")}_{digest(raw.encode("utf-8"))}.txt'
    changed=[];added=[];batch_ids=set()
    for candidate in incoming:
        ident=required(candidate.get('id'),'稳定重点标识')
        if ident in batch_ids:raise ValueError('一次输入中同一重点应先合并')
        batch_ids.add(ident)
        original=candidate.get('original');required(original,'原文')
        if original not in raw:raise ValueError('逐条原文必须原样出现在用户输入中')
        previous=store['entries'].get(ident)
        entry=deepcopy(previous) if previous else {'id':ident,'first_recorded_date':dt,'originals':[],'examples':[],'notes':[]}
        for field in ['pattern','meaning','category']:entry[field]=deepcopy(candidate[field])
        examples=candidate.get('examples',[]);notes=candidate.get('notes',[])
        if not isinstance(examples,list) or not isinstance(notes,list):raise ValueError('例句和提醒须为列表')
        entry['examples']=unique(entry['examples']+examples)
        entry['notes']=unique(entry['notes']+notes)
        if not any(o['text']==original for o in entry['originals']):
            entry['originals'].append({'date':dt,'text':original,'source_path':source})
        comparable={k:v for k,v in entry.items() if k!='last_updated_date'}
        before={k:v for k,v in (previous or {}).items() if k!='last_updated_date'}
        if comparable!=before:
            entry['last_updated_date']=max(dt,entry.get('last_updated_date',dt))
            store['entries'][ident]=entry;changed.append(ident)
            if not previous:added.append(ident)
    if not changed:return None,{}, {'action':'unchanged','file':DOCUMENT,'total':len(store['entries']),'added':[],'updated':[]}
    files={}
    if any(o['source_path']==source for e in store['entries'].values() for o in e['originals']):
        store['sources'][source]=digest(raw.encode('utf-8'));files[source]=raw.encode('utf-8')
    validate_store(store)
    document=build_document(store,existing)
    store['document_sha256']=digest(document)
    files[DOCUMENT]=document
    summary={'file':DOCUMENT,'total':len(store['entries']),'added':added,'updated':[i for i in changed if i not in added]}
    return store,files,summary

def audit_points(manager,state):
    store=state.get('key_points')
    if not store:return []
    errors=[]
    try:validate_store(store)
    except (ValueError,KeyError,TypeError) as error:return ['重点数据无效: '+str(error)]
    doc=manager.safe(store['document_path'])
    if not doc.is_file() or digest(doc.read_bytes())!=store['document_sha256']:errors.append('重点Word丢失/外部修改')
    for rel,checksum in store['sources'].items():
        p=manager.safe(rel)
        if not p.is_file() or digest(p.read_bytes())!=checksum:
            errors.append('重点原文丢失/外部修改: '+rel)
    for entry in store['entries'].values():
        for original in entry['originals']:
            p=manager.safe(original['source_path'])
            if p.is_file() and original['text'] not in p.read_text(encoding='utf-8'):
                errors.append('重点原文与来源不一致: '+entry['id'])
    return errors
