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
inputs['C5'] = 'C4+C3-1'          # last appraisal year, FY2055 (Q&A 3)
inputs['C40'] = '1624'          # Doshi et al. (2026): Reef Trust Cape York gully abatement cost, FY2024 $
inputs['C41'] = 'C48/C45'       # FY2026 average CPI / FY2024 average CPI
inputs['C42'] = 'C40*C41'

# CPI workings block (new rows below the Inputs table)
Q_IDX = [('Sep-23', 135.3), ('Dec-23', 136.1), ('Mar-24', 137.4), ('Jun-24', 138.8),
         ('Sep-24', 139.1), ('Dec-24', 139.4), ('Mar-25', 140.7), ('Jun-25', 141.7)]
M_ANN = [('Jul-25', 3.0), ('Aug-25', 3.2), ('Sep-25', 3.6), ('Oct-25', 3.8), ('Nov-25', 3.4), ('Dec-25', 3.8),
         ('Jan-26', 3.8), ('Feb-26', 3.7), ('Mar-26', 4.6), ('Apr-26', 4.2), ('May-26', 4.0), ('Jun-26', 3.8)]
Q0 = 51                      # first quarterly index row
M0 = Q0 + len(Q_IDX) + 1     # first monthly movement row (after a sub-header)
CPI_ROWS = [
    (44, 'header', 'CPI indexation workings (ABS, All groups CPI, weighted average of eight capital cities)', '', None, ''),
    (45, 'row', 'FY2023-24 average CPI (2011-12 = 100)', 'index',
     f'AVERAGE(C{Q0}:C{Q0 + 3})', f'Average of the four FY2023-24 quarterly index numbers (rows {Q0}-{Q0 + 3})'),
    (46, 'row', 'FY2024-25 average CPI (2011-12 = 100)', 'index',
     f'AVERAGE(C{Q0 + 4}:C{Q0 + 7})', f'Average of the four FY2024-25 quarterly index numbers (rows {Q0 + 4}-{Q0 + 7})'),
    (47, 'row', 'FY2025-26 year-average CPI inflation', '%',
     f'AVERAGE(C{M0}:C{M0 + 11})',
     f'Average of the twelve monthly annual movements (rows {M0}-{M0 + 11}); the ABS re-referenced the CPI to Sep-25 = 100, so FY2025-26 is chained to FY2024-25'),
    (48, 'row', 'FY2025-26 average CPI (2011-12 = 100 equivalent)', 'index',
     'C46*(1+C47/100)', 'Calculated'),
    (50, 'header', 'ABS source data: quarterly All groups CPI index numbers (2011-12 = 100)', '', None, ''),
] + [(Q0 + k, 'data', f'CPI index, {q} quarter', 'index', v, 'ABS, Consumer Price Index, Australia (6401.0), quarterly index, 2011-12 = 100')
     for k, (q, v) in enumerate(Q_IDX)] + [
    (M0 - 1, 'header', 'ABS source data: monthly CPI, annual movement (%)', '', None, ''),
] + [(M0 + k, 'data', f'CPI annual movement, {m}', '%', v, 'ABS, Consumer Price Index, Australia (monthly releases Dec-25 to Aug-26)')
     for k, (m, v) in enumerate(M_ANN)]
CPI_LAST = M0 + len(M_ANN) - 1

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


# ---------- Sensitivity sheet (new sheet, all live formulas) ----------
# cell -> (content, style); content starting with '=' is a formula, str is text, number is an input
TITLE, HEAD, COLH, LABEL, GEN, PCT, CUR, CURT, PCTOUT = 3, 16, 55, 9, 13, 41, 46, 71, 11
sens = {}
SC = {'B': 'C', 'C': 'D', 'D': 'E'}                  # Sensitivity col -> DCF results col (S2, S3, S4)
TOT = {'B': 32, 'C': 45, 'D': 62}; FIN = {'B': 29, 'C': 42, 'D': 59}


def colhead(r, first):
    sens[f'A{r}'] = (first, COLH)
    for c, name in zip('BCD', ('S2 Alternative maintenance', 'S3 Major erosion control', 'S4 Full betterment')):
        sens[f'{c}{r}'] = (name, COLH)


sens['A1'] = ('Sensitivity analysis: NPV and cost-effectiveness relative to BAU (live-linked to Inputs and DCF)', TITLE)
sens['A3'] = ('1. Discount rate', HEAD)
sens['A4'] = ('Low discount rate', LABEL); sens['B4'] = (0.04, PCT)
sens['A5'] = ('High discount rate', LABEL); sens['B5'] = (0.10, PCT)
colhead(6, 'NPV relative to BAU ($)')
rows = [(7, 'NPV (financial + environmental) at base rate (DCF)', lambda c: f'=DCF!{SC[c]}75'),
        (8, 'NPV (financial + environmental) at low rate', lambda c: f'=NPV($B$4,DCF!$C$22:$AF$22)-NPV($B$4,DCF!$C${TOT[c]}:$AF${TOT[c]})'),
        (9, 'NPV (financial + environmental) at high rate', lambda c: f'=NPV($B$5,DCF!$C$22:$AF$22)-NPV($B$5,DCF!$C${TOT[c]}:$AF${TOT[c]})'),
        (10, 'Financial NPV at low rate', lambda c: f'=NPV($B$4,DCF!$C$19:$AF$19)-NPV($B$4,DCF!$C${FIN[c]}:$AF${FIN[c]})'),
        (11, 'Financial NPV at high rate', lambda c: f'=NPV($B$5,DCF!$C$19:$AF$19)-NPV($B$5,DCF!$C${FIN[c]}:$AF${FIN[c]})')]
sens['A13'] = ('2. Capital cost overrun (S3 and S4 capital works in FY2026)', HEAD)
sens['A14'] = ('Capital cost increase tested', LABEL); sens['B14'] = (0.6, PCT)
colhead(15, 'Measure')
CAP = {'B': '0', 'C': 'SUM(DCF!$B$35:$B$37)', 'D': 'SUM(DCF!$B$48:$B$52)'}
rows += [(16, 'PV of capital costs ($)', lambda c: f'={CAP[c]}'),
         (17, 'NPV (financial + environmental) with capital increase ($)', lambda c: f'={c}7-$B$14*{c}16')]
sens['A18'] = ('Capital increase at which S4 NPV falls to S2 NPV', LABEL); sens['D18'] = ('=(D7-B7)/D16', PCTOUT)
sens['A19'] = ('Capital increase at which NPV falls to zero', LABEL)
sens['C19'] = ('=C7/C16', PCTOUT); sens['D19'] = ('=D7/D16', PCTOUT)
sens['A21'] = ('3. Climate stress test: higher erosion rates in every scenario', HEAD)
sens['A22'] = ('Increase in erosion rates (all scenarios, all years)', LABEL); sens['B22'] = (0.2, PCT)
colhead(23, 'Measure')
rows += [(24, 'NPV (financial + environmental) with higher erosion ($)', lambda c: f'=DCF!{SC[c]}74+(1+$B$22)*DCF!{SC[c]}72'),
         (25, 'Cost-effectiveness with higher erosion ($/t)', lambda c: f'=DCF!{SC[c]}67/((1+$B$22)*DCF!{SC[c]}78)')]
sens['A26'] = ('Financial costs are unchanged; tonnes avoided and their value scale with the erosion rates.', LABEL)
sens['A28'] = ('4. Payback of upfront costs against BAU (undiscounted financial costs, see table 6)', HEAD)
colhead(29, 'Measure')
rows += [(30, 'First year cumulative savings repay the outlay', lambda c: f'=IF(COUNTIF({c}$58:{c}$87,">0")=30,"Beyond FY2055",Inputs!$C$4+COUNTIF({c}$58:{c}$87,">0"))')]
sens['A31'] = ('S4 discounted net cost still unrecovered at FY2055 ($)', LABEL); sens['D31'] = ('=E87', CUR)
sens['A33'] = ('5. NPV by avoided abatement cost value (Figure 3 data)', HEAD)
colhead(34, 'Value ($/t, FY2026)')
for i in range(15):
    r = 35 + i
    sens[f'A{r}'] = ((0 if i == 0 else f'=A{r - 1}+250'), GEN)
    for c in 'BCD':
        sens[f'{c}{r}'] = (f'=DCF!{SC[c]}$74+DCF!{SC[c]}$72*$A{r}/Inputs!$C$42', CUR)
sens['A51'] = ('Value at which S4 NPV overtakes S2 ($/t)', LABEL); sens['B51'] = ('=(DCF!C74-DCF!E74)/((DCF!E72-DCF!C72)/Inputs!C42)', CURT)
sens['A52'] = ('Value at which S4 NPV overtakes S3 ($/t)', LABEL); sens['B52'] = ('=(DCF!D74-DCF!E74)/((DCF!E72-DCF!D72)/Inputs!C42)', CURT)
sens['A53'] = ('Value at which S4 NPV is zero ($/t)', LABEL); sens['B53'] = ('=-DCF!E74/(DCF!E72/Inputs!C42)', CURT)
sens['A55'] = ('6. Cumulative net financial cost relative to BAU (positive = outlay not yet repaid)', HEAD)
colhead(56, 'Financial year')
sens['E56'] = ('S4 discounted', COLH)
for i in range(30):
    r = 58 + i
    n = f'($A{r}-Inputs!$C$4+1)'
    sens[f'A{r}'] = (f'=INDEX(DCF!$C$13:$AF$13,{i + 1})', GEN)
    for c in 'BCD':
        step = f'INDEX(DCF!$C${FIN[c]}:$AF${FIN[c]},{n})-INDEX(DCF!$C$19:$AF$19,{n})'
        sens[f'{c}{r}'] = ((f'={step}' if i == 0 else f'={c}{r - 1}+{step}'), CUR)
    dstep = f'(INDEX(DCF!$C$59:$AF$59,{n})-INDEX(DCF!$C$19:$AF$19,{n}))*INDEX(DCF!$C$14:$AF$14,{n})'
    sens[f'E{r}'] = ((f'={dstep}' if i == 0 else f'=E{r - 1}+{dstep}'), CUR)
for r, label, fn in rows:
    sens[f'A{r}'] = (label, LABEL)
    for c in 'BCD':
        sens[f'{c}{r}'] = (fn(c), GEN if r == 30 else (CURT if r == 25 else CUR))


def sheet3_xml(values):
    by_row = {}
    for ref, (content, style) in sens.items():
        col, row = re.match(r'([A-Z]+)(\d+)', ref).groups()
        by_row.setdefault(int(row), []).append((col, ref, content, style))
    out = []
    for r in sorted(by_row):
        cells = []
        for col, ref, content, style in sorted(by_row[r], key=lambda x: (len(x[0]), x[0])):
            if isinstance(content, str) and content.startswith('='):
                v = (values or {}).get(ref)
                if isinstance(v, str):
                    cells.append(f'<c r="{ref}" s="{style}" t="str"><f>{escape(content[1:])}</f><v>{escape(v)}</v></c>')
                else:
                    cells.append(f'<c r="{ref}" s="{style}"><f>{escape(content[1:])}</f>{fmt_v(v)}</c>')
            elif isinstance(content, str):
                cells.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(content)}</t></is></c>')
            else:
                cells.append(f'<c r="{ref}" s="{style}"><v>{content!r}</v></c>')
        out.append(f'<row r="{r}">{"".join(cells)}</row>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetPr><tabColor theme="5" tint="0.79998168889431442"/></sheetPr><dimension ref="A1:E87"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
            '<cols><col min="1" max="1" width="78" style="1" customWidth="1"/><col min="2" max="5" width="22" style="1" customWidth="1"/></cols>'
            f'<sheetData>{"".join(out)}</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


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
        elif kind == 'data':
            out.append(f'<row r="{r}" spans="1:4">{txt(f"A{r}", 9, label)}{txt(f"B{r}", 13, unit)}'
                       f'<c r="C{r}" s="38"><v>{formula!r}</v></c>{txt(f"D{r}", 22, source)}</row>')
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
            data = data.replace(b'</Types>', b'<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        if name == 'xl/_rels/workbook.xml.rels':
            data = re.sub(rb'<Relationship [^>]*Target="calcChain.xml"/>', b'', data)
            data = data.replace(b'</Relationships>', b'<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>')
        if name == 'xl/workbook.xml':
            s = data.decode()
            s = s.replace('<sheet name="DCF" sheetId="2" r:id="rId2"/>', '<sheet name="DCF" sheetId="2" r:id="rId2"/><sheet name="Sensitivity" sheetId="3" r:id="rId7"/>')
            s = re.sub(r'<calcPr ([^/]*)/>', lambda m: '<calcPr ' + re.sub(r'\s*fullCalcOnLoad="\d"', '', m.group(1)) + ' fullCalcOnLoad="1"/>', s)
            data = s.encode()
        if name == 'xl/worksheets/sheet1.xml':
            s = inject(data.decode(), inputs, (values or {}).get('Inputs'))
            s = s.replace('</sheetData>', cpi_rows((values or {}).get('Inputs')) + '</sheetData>')
            s = s.replace('<dimension ref="A1:D42"/>', f'<dimension ref="A1:D{CPI_LAST}"/>')
            data = s.encode()
        if name == 'docProps/app.xml':
            data = data.replace(b'<vt:i4>2</vt:i4>', b'<vt:i4>3</vt:i4>').replace(b'<vt:vector size="2" baseType="lpstr"><vt:lpstr>Inputs</vt:lpstr><vt:lpstr>DCF</vt:lpstr>', b'<vt:vector size="3" baseType="lpstr"><vt:lpstr>Inputs</vt:lpstr><vt:lpstr>DCF</vt:lpstr><vt:lpstr>Sensitivity</vt:lpstr>')
        if name == 'xl/worksheets/sheet2.xml':
            data = inject(data.decode(), dcf, (values or {}).get('DCF')).encode()
        zout.writestr(item, data)
    zout.writestr('xl/worksheets/sheet3.xml', sheet3_xml((values or {}).get('Sensitivity')).encode())
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
          'DCF': {k: wb['DCF'][k].value for k in dcf},
          'Sensitivity': {k: wb['Sensitivity'][k].value for k, (c, _) in sens.items() if isinstance(c, str) and c.startswith('=')}}
missing = [k for s in values for k, v in values[s].items() if not isinstance(v, (int, float)) and not (s == 'Sensitivity' and isinstance(v, str))]
if missing:
    raise SystemExit(f'non-numeric results: {missing[:10]}')
build(OUT, values)
print('ok', OUT)
