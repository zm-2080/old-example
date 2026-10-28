const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, WidthType, BorderStyle, Header, PageNumber, PageBreak,
  LineRuleType, HighlightColor, VerticalAlign,
} = require('docx');

const FIG = process.argv[2];
const OUT = process.argv[3];
const FONT = 'Times New Roman';
const DOUBLE = { line: 480, lineRule: LineRuleType.AUTO, before: 0, after: 0 };
const SINGLE = { line: 240, lineRule: LineRuleType.AUTO, before: 0, after: 0 };
const bodyWords = [];

// *italic* and [[highlight]] inline markup
function runs(text, opts = {}) {
  const out = [];
  const re = /(\*[^*]+\*|\[\[[^\]]+\]\])/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size: opts.size || 24, bold: opts.bold }));
    const tok = m[0];
    if (tok.startsWith('*')) out.push(new TextRun({ text: tok.slice(1, -1), font: FONT, size: opts.size || 24, italics: true, bold: opts.bold }));
    else out.push(new TextRun({ text: tok.slice(2, -2), font: FONT, size: opts.size || 24, bold: opts.bold, highlight: HighlightColor.YELLOW }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, size: opts.size || 24, bold: opts.bold }));
  return out;
}
const count = (t) => bodyWords.push(t.replace(/[*\[\]]/g, ''));

const P = (t) => { count(t); return new Paragraph({ children: runs(t), spacing: DOUBLE, indent: { firstLine: 720 } }); };
const H1 = (t) => { count(t); return new Paragraph({ children: runs(t, { bold: true }), spacing: DOUBLE, alignment: AlignmentType.CENTER, keepNext: true }); };
const H2 = (t) => { count(t); return new Paragraph({ children: runs(t, { bold: true }), spacing: DOUBLE, keepNext: true }); };
const plain = (t, o = {}) => new Paragraph({ children: runs(t, o), spacing: DOUBLE, alignment: o.align || AlignmentType.LEFT });

// APA table/figure labels (not counted in body word count)
const label = (n) => new Paragraph({ children: runs(n, { bold: true }), spacing: DOUBLE, keepNext: true });
const caption = (t) => new Paragraph({ children: runs('*' + t + '*'), spacing: DOUBLE, keepNext: true });
const note = (t) => new Paragraph({ children: runs('*Note.* ' + t, { size: 20 }), spacing: { ...SINGLE, after: 240 } });

const thin = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
function table(widths, rows, leftCols = [0]) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0,
      cantSplit: true,
      children: r.map((c, j) => new TableCell({
        width: { size: widths[j], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        borders: {
          top: i === 0 ? thin : none,
          bottom: (i === 0 || i === rows.length - 1) ? thin : none,
          left: none, right: none,
        },
        children: [new Paragraph({
          children: runs(c, { size: 21, bold: i === 0 }),
          spacing: SINGLE,
          keepNext: true,
          alignment: leftCols.includes(j) ? AlignmentType.LEFT : AlignmentType.CENTER,
        })],
      })),
    })),
  });
}
function figure(file, w, h) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: SINGLE,
    children: [new ImageRun({ type: 'png', data: fs.readFileSync(path.join(FIG, file)), transformation: { width: w, height: h } })],
  });
}
const ref = (t) => new Paragraph({ children: runs(t), spacing: DOUBLE, indent: { left: 720, hanging: 720 } });

// ---------------- CONTENT ----------------
const body = [];

body.push(H1('Introduction'));
body.push(H2('Road Erosion and Climate Risk in Cook Shire'));
body.push(P('Cook Shire Council maintains an extensive network of unsealed roads in Cape York catchments that drain to the Great Barrier Reef (GBR). The region’s monsoonal climate delivers intense wet-season rainfall and near-annual cyclones, which strip fine material from road surfaces, table drains and batters (Shellberg et al., 2026). Business-as-usual (BAU) maintenance compounds the problem: graders reshape the road crown, re-cut drains and remove all verge vegetation, leaving bare soil to erode in the next storm. Highly erodible soils and high-intensity rainfall mean that drains and stream crossings scour in almost every wet season. Because disaster recovery funding generally reinstates roads only to their pre-disaster standard (Commonwealth of Australia, 2018), Council is locked into a cycle of washout, regrading and renewed erosion (Doshi, 2026).'));
body.push(H2('Environmental, Social and Economic Implications'));
body.push(P('Environmentally, the study roads lose about 38 t/ha/yr of fine sediment (<20 µm) toward the GBR, compared with natural rates below 2 t/ha/yr (Shellberg et al., 2026). Fine sediment raises turbidity, reducing the light available to seagrass and coral and degrading reef habitat (Waterhouse et al., 2024). Socially, washouts close roads that remote and Indigenous communities depend on, increase vehicle wear and create safety risks. Economically, Council spends about $20,000/km/yr re-grading these roads (Shellberg et al., 2026), while governments separately fund Reef programs to reduce the same sediment. Sediment delivery is therefore a negative externality: its damage is not borne by the road manager, so road maintenance decisions under-weight it (Field & Field, 2024).'));
body.push(H2('Project Goals and Scenarios'));
body.push(P('Natural Capital & Climate was engaged by South Cape York Catchments and Cook Shire Council to identify the most cost-effective and economically efficient way to reduce fine-sediment delivery from 21.5 km of Oaky Creek and Percy Roads (Doshi et al., 2026). Four scenarios were assessed over 30 years (Table 1).'));

const t1 = [
  label('Table 1'), caption('Road Management Scenarios and Modelled Fine-Sediment Delivery'),
  table([2000, 4826, 1300, 1300], [
    ['Scenario', 'Description', 't/km/yr', 't/yr (21.5 km)'],
    ['S1 Business-as-usual', 'Grade road crown, re-cut table and cut-off drains, remove all verge vegetation', '55.3', '1,189'],
    ['S2 Alternative maintenance', 'Cease repeated grading of drains, verges and batters; selective herbicide vegetation management; spot repairs (maintenance 70% of BAU)', '46.5', '1,000'],
    ['S3 Major erosion control', 'S2 plus rock/concrete bed-level floodways, rock mulch, grade control and gully chutes at stream crossings', '31.7', '682'],
    ['S4 Full betterment', 'S3 plus bitumen dust seal of the existing alignment', '17.6', '378'],
  ], [0, 1]),
  note('Erosion rates from Shellberg et al. (2026); scenario descriptions from Doshi (2026).'),
];
body.push(...t1);

body.push(H1('Cost-Effectiveness Analysis'));
body.push(P('Lifecycle costs (capital, maintenance and weed control) were discounted at 7% over FY2026–FY2055, consistent with Infrastructure Australia (2021), with capital works in FY2026 (Appendix A). Cost-effectiveness is the present value (PV) of each scenario’s net financial cost relative to BAU divided by the fine sediment avoided over the period (Table 2).'));

body.push(label('Table 2'), caption('Cost-Effectiveness of Fine-Sediment Reduction Relative to BAU'),
  table([2200, 1450, 1450, 1300, 1450, 1576], [
    ['Scenario', 'PV financial cost ($M)', 'PV net cost vs BAU ($M)', 'Avoided (t/yr)', 'Avoided, 30 yrs (t)', 'Cost-effectiveness ($/t)'],
    ['S1 BAU', '5.64', '–', '–', '–', '–'],
    ['S2 Alternative maintenance', '4.39', '−1.25', '189', '5,676', '−220.47'],
    ['S3 Major erosion control', '5.54', '−0.10', '507', '15,222', '−6.62'],
    ['S4 Full betterment', '5.99', '0.35', '811', '24,317', '14.53'],
  ]),
  note('PV at 7%, FY2026 dollars. Negative values are lifecycle savings relative to BAU. Calculations in the accompanying spreadsheet.'),
);

body.push(P('S2 and S3 dominate BAU: they cost less over their lifecycle and reduce sediment, giving negative cost-effectiveness ratios. S2 is the most cost-effective per tonne (−$220/t) because the 30% maintenance saving outweighs herbicide costs, but it avoids only 16% of BAU sediment. S3’s $1.15M of crossing works is more than repaid by maintenance savings. S4 is the only scenario with a positive net cost, yet at $14.53/t it is far below the $1,624/t Reef Trust paid for gully remediation on Cape York (Shellberg et al., 2026), while cutting sediment by 68% (Figure 1).'));
body.push(P('Average ratios cannot rank options of different scale, and negative ratios are hard to compare. The incremental cost of moving from S2 to S3 is $121 per additional tonne, and from S3 to S4 only $50/t, so each step up in ambition buys extra abatement cheaply. CEA also has limits: it measures a single outcome, ignoring access, safety and biodiversity co-benefits, and it divides discounted costs by undiscounted tonnes, implicitly valuing a tonne avoided in 2055 the same as one avoided today.'));

body.push(label('Figure 1'), caption('Sediment Avoided and Cost-Effectiveness by Scenario'),
  figure('fig1_cea.png', 600, 238),
  note('PV net financial cost per tonne avoided over 30 years (7% discount rate).'));

body.push(H1('Cost-Benefit Analysis'));
body.push(H2('Identifying and Valuing Benefits'));
body.push(P('Two benefits were quantified. First, reduced road maintenance is captured as lower financial costs than BAU. Second, avoided fine sediment was valued using an avoided abatement (replacement) cost: the amount society would otherwise pay to remove an equivalent tonne through Reef water-quality projects. The Cape York Reef Trust benchmark of $1,624/t (Shellberg et al., 2026) was indexed to FY2026 using ABS (2026) CPI growth (factor 1.0598), giving $1,721/t applied to each year’s avoided sediment. Weed and biosecurity control, biodiversity, road safety, fewer closures and lower vehicle wear were not quantified (Doshi, 2026), so the NPVs are conservative.'));
body.push(H2('Net Present Values'));
body.push(P('All three alternatives deliver positive NPVs relative to BAU, consistent in direction with Doshi et al. (2026). S4 has the highest NPV at $18.17M, followed by S3 ($11.70M) and S2 ($5.58M) (Table 3). Figure 2 shows that S4’s financial NPV is slightly negative (−$0.35M), so its case rests on the sediment benefit. Ranking is unchanged at 4% and 10% discount rates (S4 NPVs of $26.0M and $13.3M), although S4’s financial NPV becomes positive at 4% ($0.95M).'));

body.push(label('Table 3'), caption('Net Present Value Relative to BAU (AUD Million, 7%, FY2026 Dollars)'),
  table([2800, 1700, 1700, 1700, 1526], [
    ['Scenario', 'Financial NPV', 'PV sediment benefit', 'NPV (financial + environmental)', 'Rank'],
    ['S2 Alternative maintenance', '1.25', '4.32', '5.58', '3'],
    ['S3 Major erosion control', '0.10', '11.60', '11.70', '2'],
    ['S4 Full betterment', '−0.35', '18.52', '18.17', '1'],
  ]),
  note('A benefit–cost ratio is not reported because benefits are measured as avoided costs, making the ratio sensitive to how items are classified.'));

body.push(label('Figure 2'), caption('Composition of NPV by Scenario'),
  figure('fig2_npv.png', 520, 282), note('Financial NPV is avoided BAU maintenance less capital and new operating costs.'));

body.push(H2('Challenges in Valuing Non-Market Benefits'));
body.push(P('The avoided abatement cost is a cost-based proxy rather than a willingness-to-pay measure. It is valid only if society would otherwise fund equivalent abatement, and it may understate the non-use and tourism values of a healthier reef (Field & Field, 2024). The link between tonnes delivered and reef condition is also non-linear and uncertain, so a constant per-tonne value simplifies the damage function.'));
body.push(P('The value itself is uncertain because abatement costs vary widely between projects, and a 2015–2021 average treated as FY2024 dollars may understate current costs. Figure 3 shows S4 remains preferred for any value above $195/t; below that, S2 ranks first. The base value is nearly nine times this switching value, so the recommendation is robust.'));
body.push(P('Double counting must also be avoided. Maintenance savings are counted once, as lower costs, not again as benefits. The sediment value should not be added to separate estimates of reef tourism or habitat value, because both measure the same outcome. Erosion rates are long-term averages, so extreme events add further uncertainty.'));
body.push(P('Distribution also matters. Council and local ratepayers bear the capital cost, whereas Reef benefits accrue to Queensland, national and international beneficiaries, including tourism operators and people holding non-use values. This mismatch between who pays and who benefits helps explain under-investment and justifies cost sharing by Reef funders (Field & Field, 2024).'));

body.push(label('Figure 3'), caption('Sensitivity of NPV to the Avoided Abatement Cost Value'),
  figure('fig3_sensitivity.png', 520, 282), note('Lines show NPV (financial + environmental) as the per-tonne value changes; all other inputs held at base values.'));

body.push(H1('Comparison of CEA and CBA Results'));
body.push(label('Table 4'), caption('Summary of CEA and CBA Results'),
  table([3226, 2000, 2000, 2000], [
    ['Measure', 'S2', 'S3', 'S4'],
    ['CEA ratio, $/t (rank)', '−220 (1)', '−7 (2)', '15 (3)'],
    ['Incremental CEA, $/t', '–', '121 (vs S2)', '50 (vs S3)'],
    ['CBA NPV, $M (rank)', '5.58 (3)', '11.70 (2)', '18.17 (1)'],
    ['Financial NPV, $M', '1.25', '0.10', '−0.35'],
    ['FY2026 capital outlay, $M', '0', '1.15', '3.62'],
  ]),
  note('Derived from Tables 2 and 3.'));
body.push(P('The two methods rank the scenarios in opposite order (Table 4). CEA favours S2 because it measures cost per tonne without valuing how many tonnes are avoided. CBA monetises the additional abatement, and because the incremental costs of S3 ($121/t) and S4 ($50/t) are well below the $1,721/t value, the larger reductions generate greater net benefits.'));
body.push(P('Different decision-makers can use each analysis. Council, facing annual budget constraints, sees S2 as the cheapest and S4 as requiring $3.62M upfront. Reef investors such as Reef Trust, who allocate funds against fixed water-quality targets, would use CEA and find S4 at $15/t exceptional value compared with gully remediation. Disaster recovery agencies can use the CBA to justify funding betterment rather than like-for-like reinstatement, since S4 cuts ongoing maintenance from $425,000 to $148,000 per year. CEA therefore supports allocating a fixed budget, while CBA shows which option maximises net value to society from public funding.'));
body.push(P('The time frame also shapes conclusions. On an annual budget view, S4 costs $4.14M in FY2026 against $0.42M for BAU, and maintenance savings alone take until FY2040 to repay the outlay in undiscounted terms; at 7% they never fully do within 30 years. Agencies assessing options year to year will therefore reject betterment that a 30-year lifecycle analysis supports.'));

body.push(H1('Conclusion'));
body.push(P('S4 full betterment is recommended as the most economically efficient scenario, with an NPV of $18.17M and the lowest sediment delivery. It should be staged. Council can adopt S2 maintenance practices across the network immediately, because they reduce sediment while saving money; S3 crossing works follow, then the dust seal, funded by combining betterment and Reef water-quality funding to overcome S4’s negative financial NPV. The central trade-off is a large upfront cost and financial risk against larger, longer-term environmental gains; the opportunity cost of $3.62M of public capital is justified because S4 abates sediment at a fraction of the cost of alternative Reef investments.'));
body.push(P('Several factors affect long-term success. More intense rainfall under climate change (CSIRO & Bureau of Meteorology, 2024) will increase BAU erosion and repair costs, strengthening the case for betterment. Remoteness and short dry-season construction windows may raise capital costs, although the dust seal’s cost would need to rise by about 260% before S3 outranks S4. Disaster funding rules that require like-for-like rebuilding are the main barrier, and reform is needed. The transition will also affect local contractors who depend on recovery work. Post-works erosion monitoring should be required to verify the modelled reductions.'));

// ---------------- DOCUMENT ----------------
const words = bodyWords.join(' ').split(/\s+/).filter((w) => /[A-Za-z0-9$%]/.test(w)).length;
const titlePage = [
  plain(''), plain(''), plain(''),
  plain('Economic Appraisal of Unsealed Road Management Options in Cook Shire: A Cost-Effectiveness and Cost-Benefit Analysis', { bold: true, align: AlignmentType.CENTER }),
  plain(''),
  plain('[[Student Last Name, First Name]]', { align: AlignmentType.CENTER }),
  plain('[[Student ID]]', { align: AlignmentType.CENTER }),
  plain('School of Economics', { align: AlignmentType.CENTER }),
  plain('ECON3700: Environmental Economics', { align: AlignmentType.CENTER }),
  plain('Dr Ammarr Aripin', { align: AlignmentType.CENTER }),
  plain('9 November 2026', { align: AlignmentType.CENTER }),
  plain(''),
  plain('Word count: ' + words.toLocaleString('en-AU') + ' (excluding title page, tables, figures, notes, references and appendix)', { align: AlignmentType.CENTER }),
  new Paragraph({ children: [new PageBreak()] }),
  plain('Economic Appraisal of Unsealed Road Management Options in Cook Shire', { bold: true, align: AlignmentType.CENTER }),
];

const references = [
  new Paragraph({ children: [new PageBreak()] }),
  plain('References', { bold: true, align: AlignmentType.CENTER }),
  ref('Australian Bureau of Statistics. (2026). *Consumer price index, Australia, June 2026*. https://www.abs.gov.au/statistics/economy/price-indexes-and-inflation/consumer-price-index-australia/latest-release'),
  ref('Commonwealth of Australia. (2018). *Disaster recovery funding arrangements 2018*. Department of Home Affairs.'),
  ref('CSIRO & Bureau of Meteorology. (2024). *State of the climate 2024*. Commonwealth of Australia.'),
  ref('Doshi, A. (2026). *Unsealed road management in Cook Shire* [Guest lecture]. ECON3700 Environmental Economics, School of Economics.'),
  ref('Doshi, A., [[Erlandsen, X.]], [[Charlton-Henderson, X.]], [[Joyse, X.]], & [[Nyman, X.]] (2026). Incorporating fine sediment pollution impacts into unsealed road management: A case study from the Great Barrier Reef catchment area. *Environmental Management, 76*(4), Article 128. https://doi.org/10.1007/s00267-026-02423-2'),
  ref('Field, B. C., & Field, M. K. (2024). *Environmental economics: An introduction* (9th ed.). McGraw Hill.'),
  ref('Infrastructure Australia. (2021). *Guide to economic appraisal: Technical guide of the Assessment Framework*. Australian Government.'),
  ref('Shellberg, J. G., Klye, D., Price-Decle, J., Russell-Smith, P., Cook, K., Peter, T., Gambold, I., & Dobson, J. (2026). Quantifying fine sediment erosion from unsealed roads draining to the Great Barrier Reef before and after applying best management practices. *Environmental Management, 76*, Article 34. https://doi.org/10.1007/s00267-025-02327-7'),
  ref('Waterhouse, J., Pineda, M.-C., & Sambrook, K. (2024). 2022 Scientific Consensus Statement: Conclusions. In J. Waterhouse, M.-C. Pineda, & K. Sambrook (Eds.), *2022 Scientific Consensus Statement on land-based impacts on Great Barrier Reef water quality and ecosystem condition*. Commonwealth of Australia and Queensland Government.'),
];

const appendix = [
  new Paragraph({ children: [new PageBreak()] }),
  plain('Appendix A', { bold: true, align: AlignmentType.CENTER }),
  plain('Key Modelling Assumptions', { bold: true, align: AlignmentType.CENTER }),
  table([3000, 6226], [
    ['Item', 'Assumption'],
    ['Discount rate and period', '7% real; 30 years FY2026–FY2055; FY2026 is year 0 (undiscounted)'],
    ['Capital timing', 'All capital works occur in FY2026'],
    ['Stream crossings', '1 per km × 10.005 km crest-to-crest = 10.0 crossings; floodway and rock works priced per crossing'],
    ['Road length', '21.5 km (Oaky Creek 17 km, Percy 4.5 km); herbicide, dust seal, seal maintenance and drain cleaning per km of total length'],
    ['S2/S3 maintenance', '70% of BAU from FY2026'],
    ['S4 maintenance', 'BAU gravel maintenance in FY2026 only; seal maintenance ($5,500/km) and drain cleaning ($200/km) from FY2027'],
    ['Weed control', 'Grader grass $1,500/km (years 1–2) then $800/km; woody weeds $2,705, $1,122, $650, then $400/km'],
    ['Erosion rates', 'Scenario rates apply from FY2026 (Shellberg et al., 2026)'],
    ['Sediment value', '$1,624/t (FY2024) × CPI factor 1.0598 = $1,721/t (FY2026)'],
  ], [0, 1]),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })] })] }) },
    children: [...titlePage, ...body, ...references, ...appendix],
  }],
});

console.log('BODY WORDS', words);
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  fs.writeFileSync(OUT + '.wc', String(words));
});
