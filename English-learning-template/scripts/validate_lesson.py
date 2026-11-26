"""Structural acceptance checks for EnglishLearning daily Word lessons.

Usage: python validate_lesson.py "/absolute/path/20261124_English worksheet_Day2.docx"
Use --reference only when inspecting the retained assets/reference.docx.
This checker does not replace a visual render of the generated document.
"""
from pathlib import Path
from datetime import datetime
from zipfile import ZipFile
import argparse,importlib.util,json,re
from lxml import etree as E

NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}

_border_spec=importlib.util.spec_from_file_location('english_template_border_rules',Path(__file__).with_name('normalize_table_borders.py'))
_border_rules=importlib.util.module_from_spec(_border_spec)
_border_spec.loader.exec_module(_border_rules)

def txt(n):return ''.join(n.xpath('.//w:t/text() | .//m:t/text()',namespaces=NS))
def check_lesson(path,reference=False):
 checks=[];errors=[]
 def check(condition,message):
  (checks if condition else errors).append(message)
 path=Path(path)
 with ZipFile(path) as z:
  check(z.testzip() is None,'DOCX ZIP package is intact')
  root=E.fromstring(z.read('word/document.xml'))
 body=root.find('w:body',NS);tables=body.findall('w:tbl',NS)
 border_errors=_border_rules.check_table_borders(root)
 check(not border_errors,'All table outer/inner edges are explicit black single lines; no gray/white/theme cell or row overrides')
 errors.extend(border_errors)
 check(re.fullmatch(r'DAY\s+[1-9]\d*\s*[·•]\s*ENGLISH WORKSHEET',txt(body.find('w:p',NS)).strip()) is not None,'Main title follows DAY N · ENGLISH WORKSHEET')
 name=re.fullmatch(r'(\d{8})_English worksheet_Day([1-9]\d*)\.docx',path.name)
 check(reference or name is not None,'Filename follows YYYYMMDD_English worksheet_DayN.docx')
 if name:
  date,day=name.groups();dt=datetime.strptime(date,'%Y%m%d')
  check(re.search(r'\bDAY\s+'+day+r'\b',txt(body.find('w:p',NS))) is not None,'Filename day matches the title')
  check('Date: '+dt.strftime('%d %B %Y').lstrip('0') in txt(body),'Filename date matches the Date paragraph')
 vocab=next((t for t in tables if 'Chinese definition' in txt(t) and 'Collocation' in txt(t)),None)
 check(vocab is not None,'Vocabulary table exists')
 words=[]
 if vocab is not None:
  rows=vocab.findall('w:tr',NS)
  check(len(rows)==21,'Vocabulary has 20 words and a header')
  check(len(vocab.find('w:tblGrid',NS))==5,'Vocabulary has five columns')
  check([txt(c) for c in rows[0].findall('w:tc',NS)]==['No.','Word','Chinese definition','Collocation','Example sentence'],'Vocabulary headers match the template')
  for i,row in enumerate(rows[1:],1):
   cells=row.findall('w:tc',NS)
   if len(cells)!=5:errors.append(f'Vocabulary row {i}: expected five cells');continue
   wordparas=cells[1].findall('w:p',NS)
   word=txt(wordparas[0]).strip() if wordparas else ''
   words.append(word.lower())
   check(len(wordparas)>=2 and '/' in txt(wordparas[1]),f'Word {i}: headword and phonetic transcription share one cell')
   # Only author-inserted line breaks separate collocations; visual line wrapping does not.
   parts=[]
   for p in cells[3].findall('w:p',NS):
    value=''
    for n in p.iter():
     if n.tag=='{'+NS['w']+'}t':value+=n.text or ''
     elif n.tag=='{'+NS['w']+'}br':value+='\n'
    parts.extend(s.strip() for s in value.splitlines() if s.strip())
   check(3<=len(parts)<=4,f'Word {i}: three or four collocations')
   check(all(re.search(r'[\u4e00-\u9fff]',s) for s in parts),f'Word {i}: every collocation has Chinese')
   check(len(cells[4].findall('w:p',NS))>=2 and re.search(r'[\u4e00-\u9fff]',txt(cells[4])) is not None,f'Word {i}: bilingual example')
 confused=next((t for t in tables if t.xpath('.//m:d',namespaces=NS)),None)
 check(confused is not None,'Confused-word table contains editable brace formulas')
 if confused is not None:
  check(len(confused.find('w:tblGrid',NS))==2,'Confused-word table retains two outer columns')
  for i,c in enumerate(confused.findall('w:tr/w:tc',NS),1):
   formula_ps=[p for p in c.findall('w:p',NS) if p.find('.//m:oMath',NS) is not None]
   check(len(formula_ps)>=2,f'Confused-word group {i}: multiple independent formulas')
   labels=[]
   for p in formula_ps:
    outer=p.find('m:oMathPara/m:oMath/m:m',NS)
    check(outer is not None and len(outer.findall('m:mr',NS))==1,f'Confused-word group {i}: category formula is independent')
    if outer is None:continue
    labels.append(txt(outer.find('m:mr/m:e',NS)))
    csp=outer.find('m:mPr/m:cSp',NS)
    check(csp is not None and csp.get('{'+NS['m']+'}val')=='640',f'Confused-word group {i}: fixed label width')
    check(re.search(r'注意[：:]|如[：:]|补充[：:]',txt(p)) is None,f'Confused-word group {i}: notes stay outside formulas')
   check(labels[:2]==['释义概括','辨析'],f'Confused-word group {i}: correct category labels')
 quizzes=[]
 for t in tables:
  first=t.find('w:tr',NS)
  if first is not None and [txt(c) for c in first.findall('w:tc',NS)]==['No.','中文提示','英文默写','英文提示','中文默写']:quizzes.append(t)
 check(len(quizzes)==1,'One five-column bilingual dictation table')
 if len(quizzes)==1:
  rows=quizzes[0].findall('w:tr',NS)
  check(len(rows)==21,'Dictation includes all 20 words in both directions')
  check(all(not txt(row.findall('w:tc',NS)[j]).strip() for row in rows[1:] for j in [2,4]),'Both answer columns are blank')
  english=[txt(row.findall('w:tc',NS)[3]).strip().lower() for row in rows[1:]]
  check(english==list(reversed(words)),'English dictation prompts use reversed vocabulary order')
 return {'file':str(path.resolve()),'passed':not errors,'passed_checks':len(checks),'errors':errors,'visual_review':'Not performed by this structural checker; render the actual DOCX separately.'}

if __name__=='__main__':
 ap=argparse.ArgumentParser();ap.add_argument('document');ap.add_argument('--reference',action='store_true');args=ap.parse_args()
 result=check_lesson(args.document,args.reference)
 print(json.dumps(result,ensure_ascii=False,indent=2))
 raise SystemExit(0 if result['passed'] else 1)
