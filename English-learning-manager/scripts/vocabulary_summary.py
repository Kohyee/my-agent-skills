"""Accumulate real new-word tables in one separate, traceable Word file."""
from copy import deepcopy
from datetime import date
from pathlib import Path
from zipfile import ZipFile
import hashlib,importlib.util,io,json
from lxml import etree as E

DOCUMENT='06_Learning Notes/English Learning_Vocabulary Summary.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
N={'w':W}
def w(name):return '{'+W+'}'+name
def sha(data):return hashlib.sha256(data).hexdigest()
def jd(data):return (json.dumps(data,ensure_ascii=False,sort_keys=True)+'\n').encode('utf-8')
def text(node):return ''.join(node.xpath('.//w:t/text()',namespaces=N))
def cell_text(cell):return '\n'.join(text(p) for p in cell.findall('w:p',N))

def extract(data,lesson):
    with ZipFile(io.BytesIO(data)) as z:root=E.fromstring(z.read('word/document.xml'))
    tables=[t for t in root.findall('w:body/w:tbl',N) if t.find('w:tr',N) is not None and 'Chinese definition' in text(t.find('w:tr',N))]
    if len(tables)!=1:raise ValueError('新增词汇来源缺少唯一五列主词表')
    table=tables[0];rows=table.findall('w:tr',N)
    if len(rows)!=21:raise ValueError('新增词汇来源应有20个主词')
    result=[]
    for row in rows[1:]:
        cells=row.findall('w:tc',N)
        if len(cells)!=5:raise ValueError('词汇来源应保留五列结构')
        paragraphs=cells[1].findall('w:p',N)
        headword=text(paragraphs[0]).strip()
        result.append({'word':headword.casefold(),'headword':headword,'pronunciation':'\n'.join(text(p) for p in paragraphs[1:]),'meaning':cell_text(cells[2]),'collocations':cell_text(cells[3]),'example':cell_text(cells[4])})
    if [r['word'] for r in result]!=[str(v).strip().casefold() for v in lesson['words']]:
        raise ValueError('词表主词与登记记录不一致；先核对，不猜测或修改历史登记')
    return table,result

def collect(manager,state,overrides=None):
    overrides=overrides or {};days={};tables={};sources={};seen=set();base=None
    for dt,lesson in sorted(state['lessons'].items(),key=lambda p:(p[0],p[1]['day'])):
        if lesson.get('word_mode')!='new' or date.fromisoformat(dt).weekday()>=5:continue
        source=overrides.get(lesson['path'])
        if source is None:source=manager.safe(lesson['path']).read_bytes()
        table,rows=extract(source,lesson)
        table=deepcopy(table);xml_rows=table.findall('w:tr',N)
        kept=[];duplicates=[]
        for row,xml_row in zip(rows,xml_rows[1:]):
            if row['word'] in seen:
                duplicates.append(row['word']);table.remove(xml_row)
            else:seen.add(row['word']);kept.append(row)
        if not kept:continue
        checksum=sha(source);snapshot=f'.learning-manager/vocabulary/sources/{checksum}.docx'
        days[dt]={'date':dt,'day':lesson['day'],'source_path':lesson['path'],'source_sha256':checksum,'snapshot_path':snapshot,'words':kept,'new_word_count':len(kept),'duplicate_words_skipped':duplicates,'table_sha256':sha(E.tostring(table,method='c14n'))}
        sources[snapshot]=source;tables[dt]=table
        if base is None:base=source
    signature=sha(jd([{k:v for k,v in d.items() if k not in {'source_sha256','snapshot_path'}} for d in days.values()]))
    return days,tables,sources,base,signature

def paragraph(value,size,style=None,page_break=False):
    p=E.Element(w('p'));pr=E.SubElement(p,w('pPr'))
    if style:E.SubElement(pr,w('pStyle')).set(w('val'),style)
    if page_break:E.SubElement(pr,w('pageBreakBefore'))
    E.SubElement(pr,w('keepNext'))
    spacing=E.SubElement(pr,w('spacing'));spacing.set(w('before'),'0');spacing.set(w('after'),'120')
    if style=='Title':
        borders=E.SubElement(pr,w('pBdr'))
        for side in ['top','left','bottom','right']:E.SubElement(borders,w(side)).set(w('val'),'nil')
    run=E.SubElement(p,w('r'));rp=E.SubElement(run,w('rPr'))
    fonts=E.SubElement(rp,w('rFonts'))
    for key,val in [('ascii','Times New Roman'),('hAnsi','Times New Roman'),('eastAsia','宋体')]:fonts.set(w(key),val)
    E.SubElement(rp,w('color')).set(w('val'),'000000')
    E.SubElement(rp,w('sz')).set(w('val'),str(int(size*2)))
    if size>=14:E.SubElement(rp,w('b'))
    t=E.SubElement(run,w('t'));t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=value
    return p

def render(base,tables,days,state):
    with ZipFile(io.BytesIO(base)) as source:root=E.fromstring(source.read('word/document.xml'))
    body=root.find('w:body',N);section=body.find('w:sectPr',N)
    if section is not None:body.remove(section)
    for child in list(body):body.remove(child)
    body.append(paragraph('英语新增单词汇总',22,'Title'))
    total=sum(d['new_word_count'] for d in days.values())
    body.append(paragraph(f'收录 {len(days)} 天，共 {total} 个新增主词。按日期和 Day 汇总，复习词不重复计入。',10.5))
    if not days:body.append(paragraph('目前没有已登记的新增词汇。',11))
    for i,(dt,day) in enumerate(days.items()):
        heading=paragraph(f'Day {day["day"]}    {dt}    新增 {day["new_word_count"]} 词',14)
        if i:heading.find('w:pPr/w:spacing',N).set(w('before'),'240')
        body.append(heading)
        body.append(deepcopy(tables[dt]))
    if section is not None:body.append(section)
    # Reuse the existing template border implementation; do not fork its rules.
    helper=Path(state['config']['worksheet_skill'])/'scripts/normalize_table_borders.py'
    spec=importlib.util.spec_from_file_location('vocabulary_black_borders',helper)
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    module.normalize_tables(root)
    assert not module.check_table_borders(root)
    out=io.BytesIO()
    with ZipFile(io.BytesIO(base)) as source,ZipFile(out,'w') as target:
        for item in source.infolist():target.writestr(item,E.tostring(root,encoding='utf-8',xml_declaration=True) if item.filename=='word/document.xml' else source.read(item.filename))
    return out.getvalue()

def prepare_summary(manager,state,overrides=None):
    previous=state.get('vocabulary_summary');existing=None
    dest=manager.safe(DOCUMENT)
    if previous:
        if previous['document_path']!=DOCUMENT:raise ValueError('单词汇总文件路径不一致')
        if not dest.is_file() or sha(dest.read_bytes())!=previous['document_sha256']:
            raise ValueError('单词汇总Word被外部修改或丢失；先读取合并，不直接覆盖')
        existing=dest.read_bytes()
    elif dest.exists():raise ValueError('已有未登记的单词汇总Word，先读取合并')
    days,tables,sources,base,signature=collect(manager,state,overrides)
    if previous and signature==previous['signature']:return previous,{},False
    if existing is not None:base=existing
    if base is None:base=(Path(state['config']['worksheet_skill'])/'assets/reference.docx').read_bytes()
    doc=render(base,tables,days,state)
    store={'schema_version':1,'document_path':DOCUMENT,'document_sha256':sha(doc),'signature':signature,'total_unique_words':sum(d['new_word_count'] for d in days.values()),'days':days}
    files={DOCUMENT:doc}
    for rel,data in sources.items():
        target=manager.safe(rel)
        if target.exists() and target.read_bytes()!=data:raise ValueError('单词来源档案校验失败')
        if not target.exists():files[rel]=data
    return store,files,True

def audit_summary(manager,state):
    store=state.get('vocabulary_summary')
    if not store:return []
    errors=[]
    if store.get('schema_version')!=1 or store.get('document_path')!=DOCUMENT:return ['单词汇总数据路径或版本无效']
    doc=manager.safe(DOCUMENT)
    if not doc.is_file() or sha(doc.read_bytes())!=store['document_sha256']:errors.append('单词汇总Word丢失/外部修改')
    seen=[]
    for day in store['days'].values():
        snapshot=manager.safe(day['snapshot_path'])
        if not snapshot.is_file() or sha(snapshot.read_bytes())!=day['source_sha256']:errors.append('单词汇总来源快照损坏: '+day['date'])
        seen.extend(r['word'] for r in day['words'])
    if len(seen)!=len(set(seen)) or len(seen)!=store['total_unique_words']:errors.append('单词汇总重复或总数不一致')
    try:
        if collect(manager,state)[-1]!=store['signature']:errors.append('单词汇总待同步：来源词表已改变')
    except (ValueError,FileNotFoundError,KeyError) as e:errors.append('单词汇总来源待核对: '+str(e))
    return errors
