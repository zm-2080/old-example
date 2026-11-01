"""Replace chart placeholders in the report .docx with native Word charts.

Each chart is a DrawingML chart part (word/charts/chartN.xml) with cached data
and an embedded Excel workbook, so it can be edited in Word via "Edit Data".
Chart data are read from the completed spreadsheet so the figures always match it.

usage: python3 add_native_charts.py <report.docx> <spreadsheet.xlsx> <out.docx>
"""
import io
import re
import sys
import zipfile
from xml.sax.saxutils import escape


def attr(v):
    return escape(v, {'"': '&quot;'})

import openpyxl

DOCX, XLSX, OUT = sys.argv[1:4]

wb = openpyxl.load_workbook(XLSX, data_only=True)
d = wb['DCF']
val = wb['Inputs']['C42'].value
SC = ['S2 Alternative maintenance', 'S3 Major erosion control', 'S4 Full betterment']
COLS = 'CDE'
ce = [d[f'{c}79'].value for c in COLS]
tonnes_yr = [d['AF20'].value - d[f'AF{r}'].value for r in (30, 43, 60)]   # steady-state t/yr avoided
fin_npv = [d[f'{c}74'].value for c in COLS]
benefit = [d[f'{c}72'].value for c in COLS]
npv = [d[f'{c}75'].value for c in COLS]
k = [b / val for b in benefit]                       # NPV is linear in the per-tonne value
switch = (fin_npv[0] - fin_npv[2]) / (k[2] - k[0])  # value at which S4 overtakes S2
grid = list(range(0, 3501, 250))

C_NS = ('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
FONT = '<a:latin typeface="Times New Roman"/><a:cs typeface="Times New Roman"/>'


def txpr(sz=1000, b=0, rot=None):
    body = f'<a:bodyPr rot="{rot}" vert="horz"/>' if rot is not None else '<a:bodyPr/>'
    return (f'<c:txPr>{body}<a:lstStyle/><a:p><a:pPr><a:defRPr sz="{sz}" b="{b}">'
            f'<a:solidFill><a:srgbClr val="000000"/></a:solidFill>{FONT}</a:defRPr></a:pPr>'
            '<a:endParaRPr lang="en-US"/></a:p></c:txPr>')


def title(text, rot=None):
    body = f'<a:bodyPr rot="{rot}" vert="horz"/>' if rot is not None else '<a:bodyPr/>'
    return (f'<c:title><c:tx><c:rich>{body}<a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="0">{FONT}</a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="en-US" sz="1000" b="0">{FONT}</a:rPr><a:t>{escape(text)}</a:t></a:r></a:p></c:rich></c:tx>'
            '<c:overlay val="0"/></c:title>')


def fill(rgb):
    return f'<a:solidFill><a:srgbClr val="{rgb}"/></a:solidFill>'


def line(rgb, w=22225, dash=None):
    ds = f'<a:prstDash val="{dash}"/>' if dash else ''
    return f'<a:ln w="{w}" cap="rnd">{fill(rgb)}{ds}<a:round/></a:ln>'


NOLINE = '<a:ln><a:noFill/></a:ln>'
GRID = f'<c:majorGridlines><c:spPr><a:ln w="6350">{fill("D9D9D9")}</a:ln></c:spPr></c:majorGridlines>'
AXLINE = f'<c:spPr><a:ln w="9525">{fill("000000")}</a:ln></c:spPr>'


def strref(ref, items):
    pts = ''.join(f'<c:pt idx="{i}"><c:v>{escape(str(s))}</c:v></c:pt>' for i, s in enumerate(items))
    return f'<c:strRef><c:f>{ref}</c:f><c:strCache><c:ptCount val="{len(items)}"/>{pts}</c:strCache></c:strRef>'


def numref(ref, items, fmt='General'):
    pts = ''.join(f'<c:pt idx="{i}"><c:v>{v!r}</c:v></c:pt>' for i, v in enumerate(items))
    return (f'<c:numRef><c:f>{ref}</c:f><c:numCache><c:formatCode>{escape(fmt)}</c:formatCode>'
            f'<c:ptCount val="{len(items)}"/>{pts}</c:numCache></c:numRef>')


def dlbls(fmt, pos=None, sz=900, b=0):
    p = f'<c:dLblPos val="{pos}"/>' if pos else ''
    return (f'<c:dLbls><c:numFmt formatCode="{attr(fmt)}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            f'{txpr(sz, b)}{p}<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
            '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')


def cat_ax(ax, cross, delete=0, pos='b'):
    return (f'<c:catAx><c:axId val="{ax}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="{delete}"/>'
            f'<c:axPos val="{pos}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>'
            f'<c:minorTickMark val="none"/><c:tickLblPos val="low"/>{AXLINE}{txpr(1000)}<c:crossAx val="{cross}"/>'
            '<c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>')


def val_ax(ax, cross, ttl, fmt, pos='l', grid=True, crosses='autoZero', mn=None, mx=None, unit=None, cross_between='between'):
    sc = '<c:orientation val="minMax"/>' + (f'<c:max val="{mx}"/>' if mx is not None else '') + (f'<c:min val="{mn}"/>' if mn is not None else '')
    mu = f'<c:majorUnit val="{unit}"/>' if unit else ''
    return (f'<c:valAx><c:axId val="{ax}"/><c:scaling>{sc}</c:scaling><c:delete val="0"/><c:axPos val="{pos}"/>'
            f'{GRID if grid else ""}{title(ttl, rot=-5400000 if pos in "lr" else None)}'
            f'<c:numFmt formatCode="{attr(fmt)}" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
            f'<c:tickLblPos val="nextTo"/>{AXLINE}{txpr(1000)}<c:crossAx val="{cross}"/><c:crosses val="{crosses}"/>'
            f'<c:crossBetween val="{cross_between}"/>{mu}</c:valAx>')


def legend(pos='b'):
    return f'<c:legend><c:legendPos val="{pos}"/><c:overlay val="0"/>{txpr(1000)}</c:legend>'


def space(plot, leg=True):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace {C_NS}>'
            '<c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/>'
            f'<c:plotArea><c:layout/>{plot}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>'
            f'{legend() if leg else ""}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{txpr(1000)}'
            '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>')


def workbook(rows):
    w = openpyxl.Workbook()
    ws = w.active
    ws.title = 'Sheet1'
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    w.save(buf)
    return buf.getvalue()


SHORT = ['S2 Alternative\nmaintenance', 'S3 Major erosion\ncontrol', 'S4 Full\nbetterment']
CATS = [s.replace('\n', ' ') for s in SHORT]
BLUES = ['7FA7C9', '3D7AB5', '1B3F66']

# ---------- Figure 1: cost-effectiveness (columns) and tonnes avoided (line, secondary axis) ----------
dpts = ''.join(f'<c:dPt><c:idx val="{i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr>{fill(c)}</c:spPr></c:dPt>'
               for i, c in enumerate(BLUES))
fig1 = (
    '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
    f'<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>{strref("Sheet1!$B$1", ["Cost-effectiveness ($ per tonne avoided)"])}</c:tx>'
    f'<c:spPr>{fill(BLUES[1])}</c:spPr><c:invertIfNegative val="0"/>{dpts}{dlbls(chr(34) + "$" + chr(34) + "#,##0", "outEnd")}'
    f'<c:cat>{strref("Sheet1!$A$2:$A$4", CATS)}</c:cat><c:val>{numref("Sheet1!$B$2:$B$4", ce)}</c:val></c:ser>'
    '<c:gapWidth val="70"/><c:axId val="101"/><c:axId val="102"/></c:barChart>'
    '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
    f'<c:ser><c:idx val="1"/><c:order val="1"/><c:tx>{strref("Sheet1!$C$1", ["Fine sediment avoided per year at full effect (t/yr)"])}</c:tx>'
    f'<c:spPr>{line("C98B3D", 25400)}</c:spPr><c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr>{fill("C98B3D")}<a:ln><a:noFill/></a:ln></c:spPr></c:marker>'
    f'{dlbls("#,##0", "t")}<c:cat>{strref("Sheet1!$A$2:$A$4", CATS)}</c:cat><c:val>{numref("Sheet1!$C$2:$C$4", tonnes_yr)}</c:val>'
    '<c:smooth val="0"/></c:ser><c:marker val="1"/><c:axId val="103"/><c:axId val="104"/></c:lineChart>'
    + cat_ax(101, 102) + val_ax(102, 101, 'PV financial cost per tonne avoided ($/t)', '"$"#,##0', mn=0, mx=2000, unit=500)
    + cat_ax(103, 104, delete=1) + val_ax(104, 103, 'Fine sediment avoided (t/yr)', '#,##0', pos='r', grid=False, crosses='max', mn=0, mx=400, unit=100)
)
wb1 = workbook([['Scenario', 'Cost-effectiveness ($ per tonne avoided)', 'Fine sediment avoided per year at full effect (t/yr)']]
               + [[c, v, t] for c, v, t in zip(CATS, ce, tonnes_yr)])

# ---------- Figure 2: NPV composition (stacked columns) with total NPV markers ----------
m = [x / 1e6 for x in fin_npv]
bn = [x / 1e6 for x in benefit]
tt = [x / 1e6 for x in npv]
fig2 = (
    '<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
    f'<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>{strref("Sheet1!$B$1", ["Financial NPV (maintenance savings less capital and new costs)"])}</c:tx>'
    f'<c:spPr>{fill("C98B3D")}</c:spPr><c:invertIfNegative val="0"/>{dlbls("0.00", "ctr", 900)}'
    f'<c:cat>{strref("Sheet1!$A$2:$A$4", CATS)}</c:cat><c:val>{numref("Sheet1!$B$2:$B$4", m)}</c:val></c:ser>'
    f'<c:ser><c:idx val="1"/><c:order val="1"/><c:tx>{strref("Sheet1!$C$1", ["PV of avoided fine sediment"])}</c:tx>'
    f'<c:spPr>{fill("3D7AB5")}</c:spPr><c:invertIfNegative val="0"/>'
    f'<c:dLbls><c:numFmt formatCode="0.00" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    f'{FONT}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:dLblPos val="ctr"/><c:showLegendKey val="0"/>'
    '<c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
    f'<c:cat>{strref("Sheet1!$A$2:$A$4", CATS)}</c:cat><c:val>{numref("Sheet1!$C$2:$C$4", bn)}</c:val></c:ser>'
    '<c:gapWidth val="70"/><c:overlap val="100"/><c:axId val="201"/><c:axId val="202"/></c:barChart>'
    '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
    f'<c:ser><c:idx val="2"/><c:order val="2"/><c:tx>{strref("Sheet1!$D$1", ["NPV (financial + environmental)"])}</c:tx>'
    '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="diamond"/><c:size val="9"/>'
    f'<c:spPr>{fill("000000")}<a:ln><a:noFill/></a:ln></c:spPr></c:marker>{dlbls("0.00", "r", 1000, 1)}'
    f'<c:cat>{strref("Sheet1!$A$2:$A$4", CATS)}</c:cat><c:val>{numref("Sheet1!$D$2:$D$4", tt)}</c:val><c:smooth val="0"/></c:ser>'
    '<c:marker val="1"/><c:axId val="201"/><c:axId val="202"/></c:lineChart>'
    + cat_ax(201, 202) + val_ax(202, 201, 'AUD million (PV at 7%, FY2026 dollars)', '0', mn=-3, mx=9, unit=3)
)
wb2 = workbook([['Scenario', 'Financial NPV (maintenance savings less capital and new costs)', 'PV of avoided fine sediment',
                 'NPV (financial + environmental)']] + [[c, a, b, t] for c, a, b, t in zip(CATS, m, bn, tt)])

# ---------- Figure 3: NPV sensitivity to the per-tonne value (scatter lines) ----------
ys = [[(fin_npv[i] + k[i] * v) / 1e6 for v in grid] for i in range(3)]
n = len(grid)
sers = ''
for i in range(3):
    col = 'BCD'[i]
    sers += (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx>{strref(f"Sheet1!${col}$1", [CATS[i]])}</c:tx>'
             f'<c:spPr>{line(BLUES[i], 28575)}</c:spPr><c:marker><c:symbol val="none"/></c:marker>'
             f'<c:xVal>{numref(f"Sheet1!$A$2:$A${n + 1}", grid)}</c:xVal><c:yVal>{numref(f"Sheet1!${col}$2:${col}${n + 1}", ys[i])}</c:yVal>'
             '<c:smooth val="0"/></c:ser>')
refs = [(f'Base-case value (${val:,.0f}/t)', val, '595959', 'dash', 'F'), (f'S4 overtakes S2 (${switch:,.0f}/t)', switch, 'A6A6A6', 'sysDot', 'H')]
for j, (name, x, rgb, dash, col) in enumerate(refs):
    ycol = chr(ord(col) + 1)
    sers += (f'<c:ser><c:idx val="{3 + j}"/><c:order val="{3 + j}"/><c:tx>{strref(f"Sheet1!${col}$1", [name])}</c:tx>'
             f'<c:spPr>{line(rgb, 15875, dash)}</c:spPr><c:marker><c:symbol val="none"/></c:marker>'
             f'<c:xVal>{numref(f"Sheet1!${col}$2:${col}$3", [x, x])}</c:xVal><c:yVal>{numref(f"Sheet1!${ycol}$2:${ycol}$3", [-3, 15])}</c:yVal>'
             '<c:smooth val="0"/></c:ser>')
fig3 = (f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{sers}<c:axId val="301"/><c:axId val="302"/></c:scatterChart>'
        + val_ax(301, 302, 'Avoided abatement cost value ($ per tonne, FY2026)', '"$"#,##0', pos='b', grid=False, mn=0, mx=3500, unit=500, cross_between='midCat')
        + val_ax(302, 301, 'NPV relative to BAU (AUD million)', '0', mn=-3, mx=15, unit=3, cross_between='midCat'))
rows3 = [['Value ($/t)'] + CATS + ['', 'Base x', 'Base y', 'Switch x', 'Switch y']]
for idx, v in enumerate(grid):
    r = [v] + [ys[i][idx] for i in range(3)] + ['']
    r += ([val, -3, switch, -3] if idx == 0 else [val, 15, switch, 15] if idx == 1 else ['', '', '', ''])
    rows3.append(r)
# reference-series columns F/G (base) and H/I (switch); headers carry the series names
rows3[0][5:9] = [refs[0][0], 'y', refs[1][0], 'y']
wb3 = workbook(rows3)

CHARTS = {'%%CHART1%%': (space(fig1), wb1, 5760720, 3017520),
          '%%CHART2%%': (space(fig2), wb2, 5760720, 3200400),
          '%%CHART3%%': (space(fig3), wb3, 5760720, 3291840)}

zin = zipfile.ZipFile(DOCX)
files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
doc = files['word/document.xml'].decode()
rels = files['word/_rels/document.xml.rels'].decode()
ct = files['[Content_Types].xml'].decode()
if 'Extension="xlsx"' not in ct:
    # docx-js writes ContentType before Extension, so append rather than anchor on attribute order
    ct = ct.replace('</Types>', '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>', 1)

for n, (key, (xml, wbytes, cx, cy)) in enumerate(CHARTS.items(), start=1):
    rid = f'rIdChart{n}'
    files[f'word/charts/chart{n}.xml'] = xml.encode()
    files[f'word/charts/_rels/chart{n}.xml.rels'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet{n}.xlsx"/>'
        '</Relationships>').encode()
    files[f'word/embeddings/Microsoft_Excel_Worksheet{n}.xlsx'] = wbytes
    rels = rels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{n}.xml"/></Relationships>')
    ct = ct.replace('</Types>', f'<Override PartName="/word/charts/chart{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>')
    drawing = (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
               f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{900 + n}" name="Chart {n}" descr="Figure {n}"/><wp:cNvGraphicFramePr/>'
               '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
               f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="{rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    pat = re.compile(r'<w:r>(?:(?!<w:r>).)*?' + re.escape(key) + r'.*?</w:r>', re.S)
    doc, cnt = pat.subn(drawing, doc)
    assert cnt == 1, (key, cnt)

doc = re.sub(r'<w:highlightCs\b[^>]*/>', '', doc)  # docx-js emits this non-schema element
files['word/document.xml'] = doc.encode()
files['word/_rels/document.xml.rels'] = rels.encode()
files['[Content_Types].xml'] = ct.encode()
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    order = ['[Content_Types].xml'] + [f for f in files if f != '[Content_Types].xml']
    for f in order:
        z.writestr(f, files[f])
print('charts added', OUT, 'switch', round(switch, 2))
