// Builds the Part 1 written report (.docx). Chart positions are left as
// %%CHARTn%% placeholders that add_native_charts.py replaces with native Word charts.
// usage: node build_report.js <out.docx>
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, Header, PageNumber, PageBreak,
  LineRuleType, HighlightColor, VerticalAlign,
} = require('docx');

const OUT = process.argv[2];
const FONT = 'Times New Roman';
const DOUBLE = { line: 480, lineRule: LineRuleType.AUTO, before: 0, after: 0 };
const SINGLE = { line: 240, lineRule: LineRuleType.AUTO, before: 0, after: 0 };
const bodyWords = [];

// *italic* and [[highlight]] inline markup
function runs(text, opts = {}) {
  const out = [];
  const re = /(\*[^*]+\*|\[\[[^\]]+\]\])/g;
  const base = { font: FONT, size: opts.size || 24, bold: opts.bold, color: opts.color };
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ ...base, text: text.slice(last, m.index) }));
    const tok = m[0];
    if (tok.startsWith('*')) out.push(new TextRun({ ...base, text: tok.slice(1, -1), italics: true }));
    else out.push(new TextRun({ ...base, text: tok.slice(2, -2), highlight: HighlightColor.YELLOW }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ ...base, text: text.slice(last) }));
  return out;
}
const count = (t) => bodyWords.push(t.replace(/[*\[\]]/g, ''));

const P = (t) => { count(t); return new Paragraph({ children: runs(t), spacing: DOUBLE, indent: { firstLine: 720 } }); };
const H1 = (t) => { count(t); return new Paragraph({ children: runs(t, { bold: true }), spacing: DOUBLE, alignment: AlignmentType.CENTER, keepNext: true }); };
const H3 = (t) => { count(t); return new Paragraph({ children: runs('*' + t + '*', { bold: true }), spacing: DOUBLE, keepNext: true }); };
const H2 = (t) => { count(t); return new Paragraph({ children: runs(t, { bold: true }), spacing: DOUBLE, keepNext: true }); };
const plain = (t, o = {}) => new Paragraph({ children: runs(t, o), spacing: DOUBLE, alignment: o.align || AlignmentType.LEFT });

// APA table/figure labels and notes (excluded from the body word count)
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
const chart = (n) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: SINGLE, keepNext: true, children: [new TextRun(`%%CHART${n}%%`)] });
const ref = (t) => new Paragraph({ children: runs(t), spacing: DOUBLE, indent: { left: 720, hanging: 720 } });

// ---------------- CONTENT ----------------
const body = [];
const LI = (n, t) => { count(t); return new Paragraph({ children: runs(`${n}.\t${t}`), spacing: DOUBLE, keepNext: true, keepLines: true, indent: { left: 720, hanging: 360 }, tabStops: [{ type: 'left', position: 720 }] }); };

body.push(H1('Executive Summary'));
body.push(P('Over 30 years at a 7% discount rate, full betterment (S4) has the lowest average cost per tonne of fine sediment avoided ($622/t) and the highest NPV ($5.93M relative to business-as-usual) of four options for 21.5 km of unsealed road in Cook Shire. However, S4 is financially infeasible for Council alone: its financial NPV is −$1.54M and it needs $4.95M of capital. S4 stays ahead when the discount rate, capital cost, erosion or realised reductions change one at a time, but not at a low sediment value ($74/t) or under a combined downside (60% higher capital with 80% of modelled reductions), which favour alternative maintenance (S2). We recommend adopting S2 now and staging S4 with Reef and disaster-recovery co-funding, monitoring results before sealing.'));

body.push(H1('Introduction'));
body.push(H2('Road Erosion and Climate Risk in Cook Shire'));
body.push(P('Cook Shire Council maintains an extensive network of unsealed roads in Cape York catchments draining to the Great Barrier Reef (GBR). Intense wet-season rainfall and near-annual cyclones strip fine material from road surfaces, table drains and batters (Shellberg et al., 2026). Business-as-usual (BAU) grading compounds this by stripping verge vegetation and leaving bare soil to scour. Because disaster recovery funding generally reinstates roads only to their pre-disaster standard (Commonwealth of Australia, 2018), Council is locked into a cycle of washout, regrading and renewed erosion (Doshi, 2026).'));
body.push(P('In December 2023, Tropical Cyclone Jasper, one of the wettest cyclones on Australian record, cut roads between Cairns and Cooktown and isolated communities (Australian Institute for Disaster Resilience [AIDR], 2024). Lecture photographs show the study roads, reconstructed under disaster funding in late 2021, scoured again (Doshi, 2026). Short-duration extreme rainfall has intensified by around 10% or more in some northern regions (CSIRO & Bureau of Meteorology, 2024).'));
body.push(H2('Environmental, Social and Economic Implications'));
body.push(H3('Environmental'));
body.push(P('The study roads lose about 38 t/ha/yr of fine sediment (<20 µm) toward the GBR, against natural rates below 2 t/ha/yr (Shellberg et al., 2026). Fine sediment raises turbidity, reducing light for seagrass and coral and degrading reef habitat (Waterhouse et al., 2024).'));
body.push(H3('Social'));
body.push(P('Washouts close roads that remote and Indigenous communities depend on for food, health care and evacuation; after Cyclone Jasper, residents of Wujal Wujal had to be evacuated to Cooktown by air (AIDR, 2024). Eroded surfaces also increase vehicle wear and crash risk (Doshi, 2026).'));
body.push(H3('Economic'));
body.push(P('Council spends about $20,000/km/yr re-grading these roads (Shellberg et al., 2026), while governments separately fund Reef programs to reduce the same sediment. Sediment delivery is therefore a negative externality that maintenance decisions under-weight (Field & Field, 2024).'));
body.push(H2('Project Goals and Scenarios'));
body.push(P('South Cape York Catchments and Cook Shire Council sought the most cost-effective and economically efficient way to reduce fine-sediment delivery from Oaky Creek and Percy Roads (Doshi et al., 2026). Four scenarios were assessed (Table 1).'));

body.push(label('Table 1'), caption('Road Management Scenarios and Modelled Fine-Sediment Delivery'),
  table([2000, 4926, 1150, 1150], [
    ['Scenario', 'Description', 't/km/yr', 'Delivered (t/yr)'],
    ['S1 Business-as-usual', 'Grade road crown, re-cut table and cut-off drains, remove all verge vegetation', '55.3', '553'],
    ['S2 Alternative maintenance', 'Cease repeated grading of drains, verges and batters; selective herbicide vegetation management; spot repairs (maintenance 70% of BAU)', '46.5', '465'],
    ['S3 Major erosion control', 'S2 plus concrete and rock bed-level floodways, rock mulch, grade control and gully chutes at 21.5 stream crossings', '31.7', '317'],
    ['S4 Full betterment', 'S3 plus bitumen dust seal of the existing 21.5 km alignment', '17.6', '176'],
  ], [0, 1]),
  note('Erosion rates from Shellberg et al. (2026); descriptions from Doshi (2026). Delivery is calculated over the 10.005 km of crest-to-crest stream-crossing sections that drain to waterways.'));

body.push(H1('Methodology'));
body.push(P('Each scenario is compared with BAU (Field & Field, 2024). Following the assessment template, cost-effectiveness analysis (CEA) divides each scenario’s total present value (PV) of financial costs by undiscounted tonnes of fine sediment avoided relative to BAU. Cost-benefit analysis (CBA) adds the value of avoided sediment to maintenance savings to estimate net present value (NPV). Values are real FY2026 dollars discounted at 7% over FY2026–FY2055, with FY2026 as year 1 (Infrastructure Australia, 2021).'));

body.push(H1('Key Assumptions and Data'));
body.push(P('Costs come from Cook Shire Council and erosion rates from Shellberg et al. (2026). Capital works occur in FY2026 at BAU maintenance and erosion; S3 and S4 savings, vegetation management and lower erosion begin in FY2027, while S2 changes apply from FY2026. Sediment avoided is valued at $1,624/t (FY2024; Doshi et al., 2026), CPI-indexed to $1,725/t (ABS, 2026). Resealing is not modelled. Appendix A lists all assumptions.'));

body.push(H1('Cost-Effectiveness Analysis'));
body.push(label('Table 2'), caption('Cost-Effectiveness of Fine-Sediment Reduction Relative to BAU'),
  table([1900, 1150, 1200, 1150, 1250, 1250, 1326], [
    ['Scenario', 'PV financial cost ($M)', 'Avoided per year at full effect (t/yr)', 'Avoided, 30 years (t)', 'Cost per tonne ($/t)', 'Cost per discounted tonne ($/t)', 'Net cost above BAU ($/t avoided)'],
    ['S1 BAU', '5.27', '–', '–', '–', '–', '–'],
    ['S2 Alternative maintenance', '4.10', '88', '2,641', '1,553', '3,754', '−443'],
    ['S3 Major erosion control', '6.50', '236', '6,847', '949', '2,399', '180'],
    ['S4 Full betterment', '6.81', '377', '10,938', '622', '1,573', '140'],
  ]),
  note('PV at 7%, FY2026 dollars. Annual avoidance applies from FY2026 for S2 and from FY2027 for S3 and S4. Cost-effectiveness uses undiscounted tonnes, following the assessment template; discounting tonnes at 7% raises every ratio but leaves the ranking unchanged, and external benchmarks may use either basis. Net cost above BAU is the PV of financial costs in excess of BAU per tonne avoided (negative values are savings); it differs from the incremental cost between scenarios reported in Table 5. Calculations are in the accompanying spreadsheet.'));
body.push(P('S4 is the most cost-effective scenario at $622/t: despite the highest lifecycle cost, it cuts delivery by 68%. S2 is the cheapest scenario but the least cost-effective ($1,553/t) because it avoids only 16% of BAU sediment, although it saves $443 per tonne relative to BAU (Figure 1).'));
body.push(P('Moving from S2 to S3 costs $571 per additional tonne, but S3 to S4 costs only $75, because the seal removes most remaining surface erosion while cutting ongoing costs from $323,000 to $148,000 per year. Because S3’s incremental cost exceeds S4’s, S3 is extendedly dominated: a combination of S2 and S4 could match its abatement at lower cost, so the real choice is between S2 and S4, at an incremental $326/t. Both steps up are well below the $1,725/t (FY2026) Reef Trust has paid for Cape York gully remediation (Doshi et al., 2026); although benchmark bases may differ, road betterment appears a competitive use of Reef water-quality funding. CEA, however, ignores co-benefits.'));

body.push(label('Figure 1'), caption('Cost-Effectiveness and Annual Fine Sediment Avoided by Scenario'),
  chart(1), note('Columns show PV financial cost per tonne avoided (left axis); the line shows annual sediment avoided relative to BAU once each scenario is fully in effect (right axis).'));

body.push(H1('Cost-Benefit Analysis'));
body.push(H2('Identifying and Valuing Benefits'));
body.push(P('Maintenance savings appear as lower financial costs. Avoided fine sediment is valued at the avoided abatement (replacement) cost: what society would otherwise pay to remove an equivalent tonne through Reef water-quality projects. The central value of $1,624/t (FY2024) is the average cost of Reef Trust gully remediation on Cape York (Doshi et al., 2026), chosen because it reflects the region’s actual alternative. Individual projects range from $70/t to $5,745/t (Doshi et al., 2026), or $74/t to $6,104/t in FY2026; at the low end only S2 has a positive NPV, while at the high end S4’s NPV reaches $24.88M. Weed control, biodiversity, road safety, fewer closures and vehicle wear were not quantified (Doshi, 2026). These and avoided disaster repairs mean the NPVs may understate benefits, but omitting resealing (about $0.57M PV if costed at the seal rate in FY2046, taking S4 to $5.36M) works the other way, so the net bias is uncertain.'));
body.push(H2('Net Present Values'));
body.push(P('All alternatives deliver positive NPVs (Table 3): S4 $5.93M, S3 $3.44M and S2 $3.05M. Only S2 has a positive financial NPV ($1.17M); S3 and S4 rely on the sediment benefit (Figure 2).'));

body.push(label('Table 3'), caption('Net Present Value Relative to BAU (AUD Million, 7%, FY2026 Dollars)'),
  table([2900, 1650, 1650, 1700, 1326], [
    ['Scenario', 'Financial NPV', 'PV sediment benefit', 'NPV (financial + environmental)', 'Rank'],
    ['S2 Alternative maintenance', '1.17', '1.89', '3.05', '3'],
    ['S3 Major erosion control', '−1.23', '4.67', '3.44', '2'],
    ['S4 Full betterment', '−1.54', '7.47', '5.93', '1'],
  ]),
  note('A benefit–cost ratio is not reported because benefits include avoided costs, making the ratio sensitive to how items are classified. Sensitivity results are in the spreadsheet’s Sensitivity sheet.'));

body.push(label('Figure 2'), caption('Composition of NPV by Scenario'),
  chart(2), note('Financial NPV is avoided BAU maintenance less capital and new operating costs. Diamonds show total NPV.'));

body.push(H2('Challenges in Valuing Non-Market Benefits'));
body.push(P('The avoided abatement cost is a cost-based proxy rather than a willingness-to-pay measure. It is valid only if society would otherwise fund equivalent abatement, and it may understate the non-use and tourism values of a healthier reef (Field & Field, 2024). The link between tonnes delivered and reef condition is also non-linear, so a constant per-tonne value simplifies the damage function.'));
body.push(P('Figure 3 shows that S4 has the highest NPV for any value above $836/t and a positive NPV above $355/t. The per-tonne value is therefore the most sensitive input, and refining it matters most (Recommendation 4). S4 stays first when other inputs change one at a time, but a combined downside favours S2 (Table 4).'));
body.push(label('Table 4'), caption('Robustness of the NPV Ranking (AUD Million, Relative to BAU)'),
  table([3826, 1250, 1250, 1250, 1650], [
    ['Test (NPV, $M)', 'S2', 'S3', 'S4', 'Highest NPV'],
    ['Base case', '3.05', '3.44', '5.93', 'S4'],
    ['Discount rate 4%', '4.29', '5.84', '10.30', 'S4'],
    ['Discount rate 10%', '2.30', '2.01', '3.32', 'S4'],
    ['Erosion rates +20%', '3.43', '4.38', '7.43', 'S4'],
    ['Capital costs +60%', '3.05', '2.06', '3.16', 'S4'],
    ['Reductions at 80% of modelled', '2.68', '2.51', '4.44', 'S4'],
    ['Sediment value $74/t', '1.25', '−1.03', '−1.21', 'S2'],
    ['Capital +60% and reductions at 80%', '2.68', '1.12', '1.67', 'S2']
  ]),
  note('Each test changes the stated input with all others at base values; the last row combines two downside tests. Full results, including the high sediment value, are in Appendix A and the spreadsheet’s Sensitivity sheet.'));
body.push(P('To avoid double counting, maintenance savings are counted once, as lower costs, and the sediment value is not added to reef tourism or habitat values measuring the same outcome. Council bears the cost while Reef benefits accrue nationally, justifying cost sharing.'));

body.push(label('Figure 3'), caption('Sensitivity of NPV to the Avoided Abatement Cost Value'),
  chart(3), note('Lines show NPV (financial + environmental) as the per-tonne value changes, with all other inputs at base values.'));

body.push(H1('Comparison of CEA and CBA Results'));
body.push(label('Table 5'), caption('Summary of CEA and CBA Results'),
  table([3426, 1900, 1900, 2000], [
    ['Measure', 'S2', 'S3', 'S4'],
    ['Cost-effectiveness, $/t (rank)', '1,553 (3)', '949 (2)', '622 (1)'],
    ['Incremental cost, $/additional t', '–', '571 (vs S2)', '75 (vs S3)'],
    ['NPV, $M (rank)', '3.05 (3)', '3.44 (2)', '5.93 (1)'],
    ['Financial NPV, $M', '1.17', '−1.23', '−1.54'],
    ['FY2026 capital outlay, $M', '0', '2.47', '4.95'],
  ]),
  note('Derived from Tables 2 and 3.'));
body.push(P('CEA and CBA rank the scenarios identically, with S4 first (Table 5). CEA finds the cheapest way to reach a water-quality outcome without valuing it; CBA shows whether benefits exceed costs, but its ranking depends on the per-tonne value. A purely financial view favours S2.'));
body.push(P('Decision-makers also weigh risks the base case omits. Timing: S4 costs $5.37M in FY2026 against $0.42M for BAU, and undiscounted savings repay this only by FY2045, which annual Council budgets cannot absorb. Overrun: S4 stays ahead of S2 only up to a 62% capital overrun. Performance: combined with higher capital costs, achieving only 80% of modelled reductions favours S2. Reef investors can use CEA to compare S4 with other abatement; disaster recovery agencies can use CBA to justify betterment.'));
body.push(P('Funders should fill S4’s funding gap. S4 needs $1.54M in PV of external funding relative to BAU ($2.71M, or $326/t, relative to S2), though its $4.95M outlay falls in FY2026; at $140/t the BAU-based gap is a fraction of the $1,725/t Reef Trust benchmark (Doshi et al., 2026).'));

body.push(H1('Conclusion'));
body.push(P('S4 is economically efficient but financially infeasible for Council without co-funding. It delivers the highest NPV and lowest cost per tonne, and this holds when the discount rate, capital costs, erosion rates or realised reductions change one at a time, including 20% higher erosion reflecting more intense rainfall under climate change (CSIRO & Bureau of Meteorology, 2024). However, a combined downside of 60% higher capital and only 80% of modelled reductions favours S2, and S4 leads only while abatement is valued above $836/t. These risks, with remote construction costs and the transition for local contractors, justify staging and monitoring before sealing. A staged pathway shifts costs and reductions in time, so it needs its own appraisal; these S4 NPVs assume immediate delivery.'));

body.push(H1('Recommendations'));
body.push(LI(1, 'Adopt S2 maintenance practices across the network now as an interim, no-regret step that saves $1.17M in PV while reducing sediment.'));
body.push(LI(2, 'Deliver S4 in stages, starting with crossing works, co-funded by Reef water-quality programs and disaster betterment funding.'));
body.push(LI(3, 'Seek reform of disaster funding rules so betterment can replace like-for-like reinstatement.'));
body.push(LI(4, 'Monitor post-works erosion and refine the abatement value before committing to the dust seal.'));

// ---------------- DOCUMENT ----------------
const words = bodyWords.join(' ').split(/\s+/).filter((w) => /[A-Za-z0-9$%]/.test(w)).length;
console.log('BODY WORDS', words);

const titlePage = [
  plain(''), plain(''), plain(''),
  plain('Economic Appraisal of Unsealed Road Management Options in Cook Shire: A Cost-Effectiveness and Cost-Benefit Analysis', { bold: true, align: AlignmentType.CENTER }),
  plain(''),
  plain('Zach Allen', { align: AlignmentType.CENTER }),
  plain('Student ID: 49610350', { align: AlignmentType.CENTER }),
  plain('School of Economics, The University of Queensland', { align: AlignmentType.CENTER }),
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
  ref('Australian Bureau of Statistics. (2026). *Consumer price index, Australia, June 2026*. https://www.abs.gov.au/statistics/economy/price-indexes-and-inflation/consumer-price-index-australia/jun-2026'),
  ref('Australian Institute for Disaster Resilience. (2024). *Cyclone – Tropical Cyclone Jasper, Queensland, December 2023 – January 2024*. Australian Disaster Resilience Knowledge Hub. https://knowledge.aidr.org.au/resources/cyclone-tropical-cyclone-jasper-queensland-2023/'),
  ref('Commonwealth of Australia. (2018). *Disaster recovery funding arrangements 2018*. National Emergency Management Agency. https://www.nema.gov.au/sites/default/files/2024-08/disaster-recovery-funding-arrangements-2018.pdf'),
  ref('CSIRO & Bureau of Meteorology. (2024). *State of the climate 2024*. Commonwealth of Australia. https://www.bom.gov.au/state-of-the-climate/'),
  ref('Doshi, A. (2026, October 21). *Unsealed road management in Cook Shire* [Guest lecture slides and recording]. ECON3700 Environmental Economics, The University of Queensland.'),
  ref('Doshi, A., Erlandsen, A., Charlton-Henderson, S., Joyse, K., & Nyman, P. (2026). Incorporating fine sediment pollution impacts into unsealed road management: A case study from the Great Barrier Reef catchment area. *Environmental Management, 76*(4), Article 128. https://doi.org/10.1007/s00267-026-02423-2'),
  ref('Field, B. C., & Field, M. K. (2024). *Environmental economics: An introduction* (9th ed.). McGraw Hill.'),
  ref('Infrastructure Australia. (2021). *Guide to economic appraisal: Technical guide of the Assessment Framework*. Australian Government.'),
  ref('Shellberg, J. G., Klye, D., Price-Decle, J., Russell-Smith, P., Cook, K., Peter, T., Gambold, I., & Dobson, J. (2026). Quantifying fine sediment erosion from unsealed roads draining to the Great Barrier Reef before and after applying best management practices. *Environmental Management, 76*, Article 34. https://doi.org/10.1007/s00267-025-02327-7'),
  ref('Waterhouse, J., Pineda, M.-C., & Sambrook, K. (2024). 2022 Scientific Consensus Statement: Conclusions. In J. Waterhouse, M.-C. Pineda, & K. Sambrook (Eds.), *2022 Scientific Consensus Statement on land-based impacts on Great Barrier Reef water quality and ecosystem condition*. Commonwealth of Australia and Queensland Government.'),
];

const appendix = [
  new Paragraph({ children: [new PageBreak()] }),
  plain('Appendix A', { bold: true, align: AlignmentType.CENTER }),
  label('Table A1'), caption('Key Modelling Assumptions'),
  table([2900, 6326], [
    ['Item', 'Assumption'],
    ['Discounting', '7% real; 30 years FY2026–FY2055; discount factor 1/(1.07)^n with FY2026 as year 1'],
    ['Road length and crossings', '21.5 km (Oaky Creek 17 km, Percy 4.5 km); 1 crossing per km = 21.5 crossings'],
    ['Sediment delivery', 'Erosion rate × 10.005 km of crest-to-crest stream-crossing sections'],
    ['Capital works (S3, S4)', 'FY2026: concrete floodway, rock floodway and rock mulch/grade control/chutes at every crossing; S4 adds dust seal and grading over 21.5 km'],
    ['S2', '70% of BAU maintenance, vegetation management and 46.5 t/km/yr from FY2026'],
    ['S3', 'FY2026 at BAU maintenance and erosion; from FY2027, 70% maintenance, vegetation management (year-1 rates in FY2027) and 31.7 t/km/yr'],
    ['S4', 'FY2026 at BAU gravel maintenance and erosion; from FY2027, seal maintenance ($5,500/km), drain cleaning ($200/km), vegetation management (year-1 rates in FY2027) and 17.6 t/km/yr; no reseal in the base case (one reseal in FY2046 at the dust seal rate: $0.57M PV, S4 NPV $5.36M)'],
    ['Vegetation management', 'Grader grass $1,500/km (years 1–2) then $800/km; woody weeds $2,705, $1,122, $650, then $400/km'],
    ['Sediment value', '$1,624/t in FY2024 dollars (Doshi et al., 2026) × CPI factor 1.0624 = $1,725/t (FY2026)'],
    ['CPI indexation', 'ABS All groups CPI, eight capital cities. FY2023-24 average 136.9 (2011-12 = 100). FY2025-26 average 101.01 (Sep-25 = 100) × link factor 143.6 / 99.73 (Sep-25 quarter on each base) = 145.45. Factor = 145.45 / 136.9 = 1.0624'],
    ['Sensitivity tests', 'See Table A2 and the spreadsheet’s Sensitivity sheet'],
    ['Cost-effectiveness', 'Total PV financial cost ÷ undiscounted tonnes avoided relative to BAU over 30 years (assessment template); cost per discounted tonne also reported'],
  ], [0, 1]),,
  plain(''),
  label('Table A2'), caption('Full Sensitivity Results: NPV Relative to BAU (AUD Million)'),
  table([3826, 1250, 1250, 1250, 1650], [
    ['Test (NPV, $M)', 'S2', 'S3', 'S4', 'Highest NPV'],
    ['Base case', '3.05', '3.44', '5.93', 'S4'],
    ['Discount rate 4%', '4.29', '5.84', '10.30', 'S4'],
    ['Discount rate 10%', '2.30', '2.01', '3.32', 'S4'],
    ['Erosion rates +20%', '3.43', '4.38', '7.43', 'S4'],
    ['Capital costs +60%', '3.05', '2.06', '3.16', 'S4'],
    ['Reductions at 80% of modelled', '2.68', '2.51', '4.44', 'S4'],
    ['Sediment value $74/t', '1.25', '−1.03', '−1.21', 'S2'],
    ['Sediment value $6,104/t', '7.84', '15.31', '24.88', 'S4'],
    ['Capital +60% and reductions at 80%', '2.68', '1.12', '1.67', 'S2']
  ]),
  note('Switching values: S4 overtakes S2 at $836/t and S3 at $189/t, and has a positive NPV above $355/t. S4 stays ahead of S2 up to a 62% capital overrun; undiscounted payback against BAU is FY2045 for S4 and FY2052 for S3. All values are live formulas in the spreadsheet’s Sensitivity sheet.')
];

const doc = new Document({
  creator: 'ECON3700 student',
  title: 'Economic Appraisal of Unsealed Road Management Options in Cook Shire',
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })] })] }) },
    children: [...titlePage, ...body, ...references, ...appendix],
  }],
});

Packer.toBuffer(doc).then((buf) => fs.writeFileSync(OUT, buf));
