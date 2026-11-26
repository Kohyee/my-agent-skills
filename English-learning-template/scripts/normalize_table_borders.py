"""Apply the worksheet's black single-line All Borders policy to a DOCX copy.

Only word/document.xml changes. Text, table grids, shading, geometry, math,
headers, footers and all other package parts are preserved.
"""
from pathlib import Path
from zipfile import ZipFile
import argparse,io,json
from lxml import etree as E

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
N={'w':W,'m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
TABLE_EDGES=('top','left','bottom','right','insideH','insideV')
CELL_EDGES=('top','left','bottom','right')
ORTHOGONAL=set(TABLE_EDGES)|{'start','end'}
THEMES=('themeColor','themeTint','themeShade')


def w(tag):return '{'+W+'}'+tag


def table_width(table):
    first=table.find('w:tr',N)
    text=''.join(first.xpath('.//w:t/text()',namespaces=N)) if first is not None else ''
    if '英文默写' in text and '中文默写' in text:return 2
    if 'Chinese definition' in text or table.find('.//m:oMath',N) is not None:return 4
    existing=[b.get(w('sz')) for b in table.findall('w:tblPr/w:tblBorders/*',N)]
    return next((int(size) for size in existing if size and size.isdigit() and int(size)>0),4)


def properties(parent,tag,first=True):
    found=parent.find('w:'+tag,N)
    if found is None:
        found=E.Element(w(tag))
        if first:parent.insert(0,found)
        else:
            following={'tcBorders':{'shd','noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark','headers','cellIns','cellDel','cellMerge','tcPrChange'},'tblBorders':{'shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription','tblPrChange'}}.get(tag,set())
            insert=next((i for i,child in enumerate(parent) if E.QName(child).localname in following),len(parent))
            parent.insert(insert,found)
    return found


def set_edges(container,required,width):
    names=set(required)|{E.QName(b).localname for b in container if E.QName(b).localname in ORTHOGONAL}
    for old in list(container):
        if E.QName(old).localname in ORTHOGONAL:container.remove(old)
    order=['top','left','start','bottom','right','end','insideH','insideV']
    for index,name in enumerate(n for n in order if n in names):
        edge=E.Element(w(name));container.insert(index,edge)
        edge.set(w('val'),'single');edge.set(w('color'),'000000');edge.set(w('sz'),str(width));edge.set(w('space'),'0')


def normalize_tables(root):
    tables=list(root.iter(w('tbl')))
    for table in tables:
        width=table_width(table)
        pr=properties(table,'tblPr')
        borders=properties(pr,'tblBorders',False)
        set_edges(borders,TABLE_EDGES,width)
        for row in table.findall('w:tr',N):
            # Row-level exceptions can otherwise override black table borders.
            exception=row.find('w:tblPrEx',N)
            if exception is not None:
                set_edges(properties(exception,'tblBorders',False),TABLE_EDGES,width)
            for cell in row.findall('w:tc',N):
                cp=properties(cell,'tcPr')
                set_edges(properties(cp,'tcBorders',False),CELL_EDGES,width)
    return len(tables)


def check_table_borders(root):
    issues=[]
    def inspect(edge,where):
        if edge.get(w('val'))!='single' or edge.get(w('color'),'').upper()!='000000':
            issues.append(where+': expected explicit 000000 single-line border')
        if any(edge.get(w(attr)) is not None for attr in THEMES):
            issues.append(where+': theme colour/tint/shade can override black')
        size=edge.get(w('sz'),'')
        if not size.isdigit() or int(size)<=0:issues.append(where+': border width must be positive')
    for ti,table in enumerate(root.iter(w('tbl')),1):
        borders=table.find('w:tblPr/w:tblBorders',N)
        for side in TABLE_EDGES:
            edges=borders.findall('w:'+side,N) if borders is not None else []
            if len(edges)!=1:issues.append(f'Table {ti} {side}: All Borders requires one explicit border')
            else:inspect(edges[0],f'Table {ti} {side}')
        for ri,row in enumerate(table.findall('w:tr',N),1):
            for edge in row.findall('w:tblPrEx/w:tblBorders/*',N):
                if E.QName(edge).localname in ORTHOGONAL:inspect(edge,f'Table {ti} row {ri} override {E.QName(edge).localname}')
            for ci,cell in enumerate(row.findall('w:tc',N),1):
                # An absent cell override inherits the explicit black table
                # grid. An existing gray/white/nil/theme override is a defect.
                for edge in cell.findall('w:tcPr/w:tcBorders/*',N):
                    if E.QName(edge).localname in ORTHOGONAL:
                        inspect(edge,f'Table {ti} row {ri} cell {ci} {E.QName(edge).localname}')
    return issues


def normalize_docx_bytes(data):
    output=io.BytesIO()
    with ZipFile(io.BytesIO(data)) as source,ZipFile(output,'w') as target:
        root=E.fromstring(source.read('word/document.xml'))
        normalize_tables(root)
        assert not check_table_borders(root)
        document=E.tostring(root,encoding='utf-8',xml_declaration=True)
        for item in source.infolist():target.writestr(item,document if item.filename=='word/document.xml' else source.read(item.filename))
    return output.getvalue()


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source');parser.add_argument('--output',required=True)
    args=parser.parse_args();source=Path(args.source);target=Path(args.output)
    if source.resolve()==target.resolve():raise SystemExit('Use a separate output copy; do not overwrite the retained reference.')
    if target.exists():raise SystemExit('Output exists; preserve it and choose a new path.')
    target.parent.mkdir(parents=True,exist_ok=True)
    target.write_bytes(normalize_docx_bytes(source.read_bytes()))
    print(json.dumps({'file':str(target.resolve()),'border_policy':'All Borders, single, 000000','passed':True},ensure_ascii=False,indent=2))
