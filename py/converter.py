"""Client-side legacy spreadsheet conversion. Compatible with Pyodide / PyScript.

Supported: SpreadsheetML 2003, HTML disguised as .xls, BIFF XLS, XLSX.
No file content is uploaded or sent to any server.
"""
from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from copy import copy
from html import unescape

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

NS = 'urn:schemas-microsoft-com:office:spreadsheet'
SS = '{' + NS + '}'


def detect(data: bytes):
    if data.startswith(bytes.fromhex('D0CF11E0A1B11AE1')):
        return 'biff'
    if data.startswith(b'PK\x03\x04') or data.startswith(b'PK\x05\x06'):
        return 'xlsx'
    if data[:4] in (b'\x09\x08\x10\x00', b'\x09\x04\x06\x00'):
        return 'biff'
    text = decode(data[:35000]).lstrip('\ufeff\u00a0 \n\r\t').lower()
    if 'schemas-microsoft-com:office:spreadsheet' in text and re.search(r'<(?:[a-z0-9_-]+:)?workbook\b', text):
        return 'xml'
    if '<html' in text or '<table' in text:
        return 'html'
    return None


def decode(data: bytes):
    if data.startswith(b'\xff\xfe'):
        return data.decode('utf-16le').lstrip('\ufeff')
    if data.startswith(b'\xfe\xff'):
        return data.decode('utf-16be').lstrip('\ufeff')
    return data.decode('utf-8-sig', errors='replace')


def hex_color(value):
    if not value:
        return None
    val = value.strip().lstrip('#')
    if re.fullmatch('[0-9a-fA-F]{6}', val):
        return 'FF' + val.upper()
    m = re.search(r'#[0-9a-fA-F]{6}', value)
    if m:
        return 'FF' + m.group(0)[1:].upper()
    return None


def px_to_points(raw, default=None):
    if not raw:
        return default
    match = re.match(r'\s*([\d.]+)(px|pt)?', str(raw), re.I)
    if not match:
        return default
    value = float(match.group(1))
    return value * (.75 if match.group(2) == 'px' else 1)


def read_excel_xml(data: bytes):
    root = ET.fromstring(data)
    if root.tag.split('}')[-1] != 'Workbook':
        raise ValueError('Not an Excel 2003 XML workbook')
    wb = Workbook()
    wb.remove(wb.active)
    definitions = {}
    for elt in root.findall('.//' + SS + 'Style'):
        definitions[elt.get(SS + 'ID') or 'Default'] = elt
    parent = definitions.get('Default')
    for ws_src in root.findall('.//' + SS + 'Worksheet'):
        name = (ws_src.get(SS + 'Name') or 'Sheet')[:31]
        ws = wb.create_sheet(name)
        table = ws_src.find(SS + 'Table')
        if table is None:
            continue
        ci = 0
        for col in table.findall(SS + 'Column'):
            ci = int(col.get(SS+'Index', ci+1))
            width = col.get(SS+'Width')
            if width:
                ws.column_dimensions[get_column_letter(ci)].width = max(4,min(80, float(width)/7.1))
            ci += 1
        next_ri = 1
        for row in table.findall(SS + 'Row'):
            ri = int(row.get(SS + 'Index', next_ri))
            height = row.get(SS+'Height')
            if height:
                ws.row_dimensions[ri].height = float(height)
            next_ci = 1
            for node in row.findall(SS+'Cell'):
                ci = int(node.get(SS+'Index',next_ci))
                dat = node.find(SS+'Data')
                value = dat.text if dat is not None else None
                value_type = dat.get(SS+'Type','String') if dat is not None else 'String'
                if value_type == 'Number' and value not in (None,''):
                    value = float(value)
                    if value.is_integer(): value = int(value)
                elif value_type == 'Boolean' and value not in (None,''):
                    value = (value == '1')
                cell = ws.cell(ri,ci,value)
                style_id = node.get(SS+'StyleID') or row.get(SS+'StyleID')
                style = definitions.get(style_id, parent)
                if style is not None:
                    apply_xml_style(cell,style,parent)
                formula = node.get(SS+'Formula')
                if formula and not ('R[' in formula or 'C[' in formula):
                    if formula.startswith('='):
                        cell.value = formula
                across = int(node.get(SS+'MergeAcross',0))
                down = int(node.get(SS+'MergeDown',0))
                if across or down:
                    ws.merge_cells(start_row=ri,start_column=ci,end_row=ri+down,end_column=ci+across)
                next_ci = ci + across + 1
            next_ri = ri + 1
        ws.sheet_view.showGridLines = False
    if not wb.worksheets:
        raise ValueError('Workbook contains no worksheets')
    return wb


def apply_xml_style(cell,style,default):
    def child(tag):
        item = style.find(SS + tag)
        if item is not None:
            return item
        return default.find(SS + tag) if default is not None and style is not default else None
    font=child('Font')
    if font is not None:
        cell.font=Font(name=font.get(SS+'FontName','Calibri'),size=float(font.get(SS+'Size','11')),
          bold=font.get(SS+'Bold')=='1',italic=font.get(SS+'Italic')=='1',
          underline='single' if font.get(SS+'Underline') else None,
          color=hex_color(font.get(SS+'Color')))
    fill=child('Interior')
    if fill is not None:
        col=hex_color(fill.get(SS+'Color'))
        if col:
            cell.fill=PatternFill('solid', fgColor=col)
    align=child('Alignment')
    if align is not None:
        h=align.get(SS+'Horizontal','').lower()
        v=align.get(SS+'Vertical','').lower()
        cell.alignment=Alignment(horizontal={'center':'center','right':'right','left':'left','justify':'justify'}.get(h),
                                 vertical={'center':'center','top':'top','bottom':'bottom'}.get(v),
                                 wrap_text=align.get(SS+'WrapText')=='1')
    border=child('Borders')
    if border is not None:
        edges={}
        for e in border.findall(SS+'Border'):
            pos=(e.get(SS+'Position') or '').lower()
            if pos in ['left','right','top','bottom']:
                edges[pos]=Side(style='thin' if e.get(SS+'LineStyle','Continuous') else 'thin',color=hex_color(e.get(SS+'Color')) or 'CCCCCC')
        cell.border=Border(**edges)
    fmt=child('NumberFormat')
    if fmt is not None and fmt.get(SS+'Format'):
        cell.number_format=fmt.get(SS+'Format')


def css_rules(soup):
    rules=defaultdict(dict)
    for sty in soup.find_all('style'):
        data=re.sub(r'/\*.*?\*/','',sty.get_text(),flags=re.S)
        for sel,body in re.findall(r'([^{}]+)\{([^{}]*)\}',data):
            for item in sel.split(','):
                rules[item.strip().lower()].update(declarations(body))
    return rules


def declarations(s):
    out={}
    for item in (s or '').split(';'):
        if ':' in item:
            k,v=item.split(':',1)
            out[k.strip().lower()]=v.strip().replace('!important','').strip()
    return out


def parse_number(raw):
    raw=raw.replace('\xa0',' ').strip()
    if not raw or not re.search(r'\d',raw): return None
    negative=raw.startswith('(') or raw.startswith('-')
    val=re.sub(r'[^\d.,]','',raw)
    if ',' in val and '.' in val:
        val=val.replace('.','').replace(',','.') if val.rfind(',')>val.rfind('.') else val.replace(',','')
    elif ',' in val:
        parts=val.split(',')
        val='.'.join(parts) if len(parts)==2 and len(parts[1])<=2 else ''.join(parts)
    elif val.count('.')>1:
        val=val.replace('.','')
    elif re.search(r'\.\d{3}$',val):
        val=val.replace('.','')
    try:
        n=float(val)*(-1 if negative else 1)
        return int(n) if n.is_integer() else n
    except ValueError:
        return None


def read_html_xls(data:bytes):
    from bs4 import BeautifulSoup
    soup=BeautifulSoup(decode(data),'html.parser')
    table=soup.find('table')
    if table is None:raise ValueError('HTML spreadsheet does not contain a table')
    rules=css_rules(soup)
    wb=Workbook();ws=wb.active;ws.title='Monitoring';ws.sheet_view.showGridLines=False
    widths=[6,14,29,12,25,12,19,19,13,20,20,22,17]
    for c,v in enumerate(widths,1): ws.column_dimensions[get_column_letter(c)].width=v
    taken=set();row_no=0
    for tr in table.find_all('tr'):
        elements=tr.find_all(['td','th'],recursive=False)
        if not elements:continue
        row_no+=1;col_no=1
        for elem in elements:
            while (row_no,col_no) in taken:col_no+=1
            colspan=max(1,min(60,int(elem.get('colspan',1) or 1)))
            rowspan=max(1,min(500,int(elem.get('rowspan',1) or 1)))
            classes=[x.lower() for x in elem.get('class',[])]
            css={**rules.get('body',{}),**rules.get(elem.name,{})}
            for cls in classes:
                css.update(rules.get('.'+cls,{}))
                css.update(rules.get(elem.name+'.'+cls,{}))
            css.update(declarations(elem.get('style','')))
            value=elem.get_text(' ',strip=True).replace('\xa0',' ')
            value=re.sub(r'\s+',' ',value).strip()
            numeric=('num' in classes or elem.has_attr('x:num')) and bool(value)
            parsed=parse_number(value) if numeric else None
            cell=ws.cell(row_no,col_no,parsed if parsed is not None else value)
            if parsed is not None:
                cell.number_format='"Rp. "#,##0.00' if ('Rp' in value or 'num' in classes) else '#,##0.00'
            if 'str' in classes or 'text-cell' in classes: cell.number_format='@'
            if elem.name=='th' or css.get('font-weight','').lower()=='bold' or (css.get('font-weight','').isdigit() and int(css.get('font-weight','0'))>=600):bold=True
            else:bold=False
            font_name=css.get('font-family','Calibri').split(',')[0].strip(' \'"')
            fsize=px_to_points(css.get('font-size'),10)
            col=hex_color(css.get('color'))
            cell.font=Font(name=font_name,size=fsize,bold=bold,color=col)
            bg=hex_color(css.get('background-color'))
            if bg:cell.fill=PatternFill('solid',fgColor=bg)
            align=css.get('text-align',None)
            if align not in ('center','left','right','justify'):align=None
            cell.alignment=Alignment(horizontal=align,vertical='center',wrap_text=css.get('white-space')!='nowrap')
            bs={}
            for side in ('left','right','top','bottom'):
                v=css.get('border-'+side) or css.get('border')
                if v and not re.search(r'\bnone\b',v):
                    bc=hex_color(v)
                    if bc:bs[side]=Side(style='thin',color=bc)
            if bs:cell.border=Border(**bs)
            if colspan>1 or rowspan>1:
                ws.merge_cells(start_row=row_no,start_column=col_no,end_row=row_no+rowspan-1,end_column=col_no+colspan-1)
                for rr in range(row_no,row_no+rowspan):
                    for cc in range(col_no,col_no+colspan):
                        if (rr,cc)!=(row_no,col_no):taken.add((rr,cc))
            col_no+=colspan
        ws.row_dimensions[row_no].height=17
    if row_no==0:raise ValueError('No spreadsheet rows found')
    return wb


def read_binary(data:bytes):
    import xlrd
    source=xlrd.open_workbook(file_contents=data,formatting_info=True)
    wb=Workbook();wb.remove(wb.active)
    for original in source.sheets():
        ws=wb.create_sheet((original.name or 'Sheet')[:31])
        for r in range(original.nrows):
            for c in range(original.ncols):
                src=original.cell(r,c)
                value=src.value
                if src.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):continue
                if src.ctype==xlrd.XL_CELL_BOOLEAN:value=bool(value)
                if src.ctype==xlrd.XL_CELL_NUMBER and float(value).is_integer():value=int(value)
                ws.cell(r+1,c+1,value)
        for rlo,rhi,clo,chi in original.merged_cells:
            ws.merge_cells(start_row=rlo+1,end_row=rhi,start_column=clo+1,end_column=chi)
        ws.sheet_view.showGridLines=False
    return wb


def convert_bytes(data:bytes,filename:str='input.xls'):
    kind=detect(data)
    if kind is None:raise ValueError('Unsupported Excel file format')
    if kind=='xlsx':
        wb=load_workbook(io.BytesIO(data),read_only=True)
        return data,'xlsx',len(wb.sheetnames)
    if kind=='xml':wb=read_excel_xml(data)
    elif kind=='html':wb=read_html_xls(data)
    elif kind=='biff':wb=read_binary(data)
    else:wb=load_workbook(io.BytesIO(data))
    buf=io.BytesIO();wb.save(buf)
    result=buf.getvalue()
    assert result[:2]==b'PK', 'XLSX output is invalid'
    check=load_workbook(io.BytesIO(result),read_only=False,data_only=False)
    assert len(check.sheetnames)==len(wb.sheetnames)
    return result,kind,len(wb.sheetnames)