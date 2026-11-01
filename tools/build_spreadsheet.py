"""Fill the ECON3700 unsealed roads workbook by editing sheet XML in place.

Every existing style, column width, freeze pane and web extension is kept; only
formulas (with cached values from a LibreOffice recalculation) are written into
the template's blank cells, plus a CPI workings block under the Inputs table.
Modelling conventions follow the course Q&A forum (06_course_QA/Case_Study_QA.md).

usage: python3 build_spreadsheet.py <template.xlsx> <output.xlsx> <work_dir>
"""
import os
import re
import subprocess
import sys
import zipfile
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import get_column_letter

SRC, OUT, WORK = sys.argv[1], sys.argv[2], sys.argv[3]
YEARS = [get_column_letter(c) for c in range(3, 33)]  # C..AF = FY2026..FY2055
assert YEARS[0] == 'C' and YEARS[-1] == 'AF'

inputs, dcf = {}, {}

# ---------- Inputs sheet: avoided abatement cost and CPI indexation ----------
inputs['C40'] = '1624'          # Doshi et al. (2026): Reef Trust Cape York gully abatement cost, FY2024 $
inputs['C41'] = 'C48/C45'       # FY2026 average CPI / FY2024 average CPI
inputs['C42'] = 'C40*C41'

# CPI workings block (new rows below the Inputs table)
CPI_ROWS = [
    (44, 'header', 'CPI indexation workings (ABS, All groups CPI, weighted average of eight capital cities)', '', None, ''),
    (45, 'row', 'FY2023-24 average CPI (2011-12 = 100)', 'index',
     'AVERAGE(135.3,136.1,137.4,138.8)', 'ABS 6401.0 quarterly index: Sep-23 135.3, Dec-23 136.1, Mar-24 137.4, Jun-24 138.8'),
    (46, 'row', 'FY2024-25 average CPI (2011-12 = 100)', 'index',
     'AVERAGE(139.1,139.4,140.7,141.7)', 'ABS 6401.0 quarterly index: Sep-24 139.1, Dec-24 139.4, Mar-25 140.7, Jun-25 141.7'),
    (47, 'row', 'FY2025-26 year-average CPI inflation', '%',
     'AVERAGE(3,3,3.5,3.8,3.4,3.8,3.8,3.7,4.6,4.2,4,3.8)',
     'Mean of ABS monthly annual CPI movements Jul-25 to Jun-26 (series re-referenced to Sep-25 = 100, so FY2025-26 is chained to FY2024-25)'),
    (48, 'row', 'FY2025-26 average CPI (2011-12 = 100 equivalent)', 'index',
     'C46*(1+C47/100)', 'Calculated'),
]

# ---------- DCF key parameters ----------
dcf['B5'] = 'Inputs!C2'
dcf['B6'] = 'Inputs!C3'
dcf['B7'] = 'Inputs!C8'
dcf['B8'] = 'Inputs!C9*Inputs!C8'          # crossings: 1 per km x 21.5 km (Q&A 8)
dcf['B9'] = 'Inputs!C12'                    # crest-to-crest length: sediment delivery (Q&A 8)
dcf['B10'] = 'Inputs!C42'
dcf['B11'] = 'Inputs!C22/100'

pv_rows = [17, 18, 19, 21, 22, 25, 26, 27, 28, 29, 31, 32,
           35, 36, 37, 38, 39, 40, 41, 42, 44, 45,
           48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 61, 62]
for r in pv_rows:
    dcf[f'B{r}'] = f'SUMPRODUCT(C{r}:AF{r},$C$14:$AF$14)'


def yr(c, offset=0):
    """Programme year number (FY2026 = 1) less an offset for delayed starts."""
    return f'({c}$13-Inputs!$C$4+1-{offset})' if offset else f'({c}$13-Inputs!$C$4+1)'


def grass(c, offset):
    n = yr(c, offset)
    return f'IF({n}<1,0,IF({n}<=2,Inputs!$C$25,Inputs!$C$26))*$B$7'


def woody(c, offset):
    n = yr(c, offset)
    return (f'IF({n}<1,0,CHOOSE(MIN({n},4),Inputs!$C$28,Inputs!$C$29,'
            f'Inputs!$C$30,Inputs!$C$31))*$B$7')


for c in YEARS:
    works = f'({c}$13=Inputs!$C$4)'           # TRUE in the initial works year, FY2026
    dcf[f'{c}14'] = f'1/(1+$B$5)^{yr(c)}'     # FY2026 = year 1 (Q&A 10, 11)
    # Scenario 1: business as usual
    dcf[f'{c}17'] = 'Inputs!$C$20*Inputs!$C$6'
    dcf[f'{c}18'] = 'Inputs!$C$21*Inputs!$C$7'
    dcf[f'{c}19'] = f'SUM({c}17:{c}18)'
    dcf[f'{c}20'] = 'Inputs!$C$35*$B$9'
    dcf[f'{c}21'] = f'{c}20*$B$10'
    dcf[f'{c}22'] = f'{c}19+{c}21'
    # Scenario 2: alternative maintenance (no works year; changes apply from FY2026)
    dcf[f'{c}25'] = f'{c}17*$B$11'
    dcf[f'{c}26'] = f'{c}18*$B$11'
    dcf[f'{c}27'] = grass(c, 0)
    dcf[f'{c}28'] = woody(c, 0)
    dcf[f'{c}29'] = f'SUM({c}25:{c}28)'
    dcf[f'{c}30'] = 'Inputs!$C$36*$B$9'
    dcf[f'{c}31'] = f'{c}30*$B$10'
    dcf[f'{c}32'] = f'{c}29+{c}31'
    # Scenario 3: capital works in FY2026; reduced maintenance, vegetation
    # management and erosion rate from FY2027 (Q&A 2, 3)
    dcf[f'{c}35'] = f'IF({works},Inputs!$C$14*$B$8,0)'
    dcf[f'{c}36'] = f'IF({works},Inputs!$C$15*$B$8,0)'
    dcf[f'{c}37'] = f'IF({works},Inputs!$C$16*$B$8,0)'
    dcf[f'{c}38'] = f'{c}17*IF({works},1,Inputs!$C$23/100)'
    dcf[f'{c}39'] = f'{c}18*IF({works},1,Inputs!$C$23/100)'
    dcf[f'{c}40'] = grass(c, 1)
    dcf[f'{c}41'] = woody(c, 1)
    dcf[f'{c}42'] = f'SUM({c}35:{c}41)'
    dcf[f'{c}43'] = f'IF({works},Inputs!$C$35,Inputs!$C$37)*$B$9'
    dcf[f'{c}44'] = f'{c}43*$B$10'
    dcf[f'{c}45'] = f'{c}42+{c}44'
    # Scenario 4: capital works and full BAU gravel maintenance in FY2026; bitumen
    # maintenance, drain cleaning, vegetation management and erosion rate from
    # FY2027 (Q&A 2, 4, 6, 7, 10)
    dcf[f'{c}48'] = f'IF({works},Inputs!$C$14*$B$8,0)'
    dcf[f'{c}49'] = f'IF({works},Inputs!$C$15*$B$8,0)'
    dcf[f'{c}50'] = f'IF({works},Inputs!$C$16*$B$8,0)'
    dcf[f'{c}51'] = f'IF({works},Inputs!$C$17*$B$7,0)'
    dcf[f'{c}52'] = f'IF({works},Inputs!$C$18*$B$7,0)'
    dcf[f'{c}53'] = f'IF({works},{c}17,0)'
    dcf[f'{c}54'] = f'IF({works},{c}18,0)'
    dcf[f'{c}55'] = f'IF({works},0,Inputs!$C$32*$B$7)'
    dcf[f'{c}56'] = f'IF({works},0,Inputs!$C$33*$B$7)'
    dcf[f'{c}57'] = grass(c, 1)
    dcf[f'{c}58'] = woody(c, 1)
    dcf[f'{c}59'] = f'SUM({c}48:{c}58)'
    dcf[f'{c}60'] = f'IF({works},Inputs!$C$35,Inputs!$C$38)*$B$9'
    dcf[f'{c}61'] = f'{c}60*$B$10'
    dcf[f'{c}62'] = f'{c}59+{c}61'

# ---------- Results summary ----------
fin = {'B': 19, 'C': 29, 'D': 42, 'E': 59}
tot = {'B': 22, 'C': 32, 'D': 45, 'E': 62}
env = {'B': 21, 'C': 31, 'D': 44, 'E': 61}
sed = {'B': 20, 'C': 30, 'D': 43, 'E': 60}
for col in 'BCDE':
    dcf[f'{col}67'] = f'$B${fin[col]}'
    dcf[f'{col}68'] = f'$B${tot[col]}'
for col in 'CDE':
    dcf[f'{col}71'] = f'{col}67-$B$67'
    dcf[f'{col}72'] = f'$B${env["B"]}-$B${env[col]}'
    dcf[f'{col}74'] = f'-{col}71'
    dcf[f'{col}75'] = f'$B$68-{col}68'
    dcf[f'{col}78'] = f'SUM($C${sed["B"]}:$AF${sed["B"]})-SUM($C${sed[col]}:$AF${sed[col]})'
    dcf[f'{col}79'] = f'{col}67/{col}78'   # total PV financial cost / tonnes avoided (Q&A 6)


def fmt_v(v):
    return f'<v>{repr(float(v))}</v>' if v is not None else ''


def inject(xml, cells, values):
    for ref, formula in cells.items():
        v = (values or {}).get(ref)
        f = escape(formula)
        pat = re.compile(r'<c r="%s"( s="\d+")?(?: t="\w+")?\s*(?:/>|>.*?</c>)' % ref, re.S)
        m = pat.search(xml)
        if not m:
            raise SystemExit(f'cell {ref} not present in template')
        xml = xml[:m.start()] + f'<c r="{ref}"{m.group(1) or ""}><f>{f}</f>{fmt_v(v)}</c>' + xml[m.end():]
    return xml


def cpi_rows(values):
    def txt(ref, s, text):
        return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'
    out = []
    for r, kind, label, unit, formula, source in CPI_ROWS:
        if kind == 'header':
            out.append(f'<row r="{r}" spans="1:4" s="2" customFormat="1">{txt(f"A{r}", 16, label)}'
                       f'<c r="B{r}" s="17"/><c r="C{r}" s="17"/><c r="D{r}" s="18"/></row>')
        else:
            v = (values or {}).get(f'C{r}')
            out.append(f'<row r="{r}" spans="1:4">{txt(f"A{r}", 9, label)}{txt(f"B{r}", 13, unit)}'
                       f'<c r="C{r}" s="15"><f>{escape(formula)}</f>{fmt_v(v)}</c>{txt(f"D{r}", 22, source)}</row>')
    return ''.join(out)


def build(dst, values=None):
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        name = item.filename
        if name == 'xl/calcChain.xml':
            continue  # Excel rebuilds it
        if name == '[Content_Types].xml':
            data = re.sub(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b'', data)
        if name == 'xl/_rels/workbook.xml.rels':
            data = re.sub(rb'<Relationship [^>]*Target="calcChain.xml"/>', b'', data)
        if name == 'xl/workbook.xml':
            s = data.decode()
            s = re.sub(r'<calcPr ([^/]*)/>', lambda m: '<calcPr ' + re.sub(r'\s*fullCalcOnLoad="\d"', '', m.group(1)) + ' fullCalcOnLoad="1"/>', s)
            data = s.encode()
        if name == 'xl/worksheets/sheet1.xml':
            s = inject(data.decode(), inputs, (values or {}).get('Inputs'))
            s = s.replace('</sheetData>', cpi_rows((values or {}).get('Inputs')) + '</sheetData>')
            s = s.replace('<dimension ref="A1:D42"/>', '<dimension ref="A1:D48"/>')
            data = s.encode()
        if name == 'xl/worksheets/sheet2.xml':
            data = inject(data.decode(), dcf, (values or {}).get('DCF')).encode()
        zout.writestr(item, data)
    zout.close()


os.makedirs(WORK, exist_ok=True)
stage1 = os.path.join(WORK, 'stage1.xlsx')
build(stage1)
# Recalculate with LibreOffice to obtain cached values, then write them into the untouched-format file.
subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir',
                os.path.join(WORK, 'lo'), stage1], check=True, capture_output=True)
wb = openpyxl.load_workbook(os.path.join(WORK, 'lo', 'stage1.xlsx'), data_only=True)
cpi_refs = {f'C{r}': None for r, kind, *_ in CPI_ROWS if kind == 'row'}
values = {'Inputs': {k: wb['Inputs'][k].value for k in list(inputs) + list(cpi_refs)},
          'DCF': {k: wb['DCF'][k].value for k in dcf}}
missing = [k for s in values for k, v in values[s].items() if not isinstance(v, (int, float))]
if missing:
    raise SystemExit(f'non-numeric results: {missing[:10]}')
build(OUT, values)
print('ok', OUT)
