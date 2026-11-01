// Builds the Part 2 reflective writing (.docx), 5Rs framework, 400-word limit.
// usage: node build_reflection.js <out.docx>
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LineRuleType, Header, PageNumber } = require('docx');

const OUT = process.argv[2];
const FONT = 'Times New Roman';
const DOUBLE = { line: 480, lineRule: LineRuleType.AUTO, before: 0, after: 0 };
const counted = [];

function runs(text, o = {}) {
  return text.split(/(\*[^*]+\*)/).filter(Boolean).map((t) => t.startsWith('*')
    ? new TextRun({ text: t.slice(1, -1), italics: true, bold: o.bold, font: FONT, size: 24 })
    : new TextRun({ text: t, bold: o.bold, font: FONT, size: 24 }));
}
const P = (t) => { counted.push(t); return new Paragraph({ children: runs(t), spacing: DOUBLE, indent: { firstLine: 720 } }); };
const H = (t) => { counted.push(t); return new Paragraph({ children: runs(t, { bold: true }), spacing: DOUBLE, keepNext: true }); };
const C = (t, bold) => new Paragraph({ children: runs(t, { bold }), spacing: DOUBLE, alignment: AlignmentType.CENTER });
const ref = (t) => new Paragraph({ children: runs(t), spacing: DOUBLE, indent: { left: 720, hanging: 720 } });

const body = [
  H('Reporting'),
  P('On 21 October 2026, Dr Ammar Doshi of Natural Capital and Climate presented a real appraisal of unsealed road management in Cook Shire. He described a cycle in which cyclones wash out roads, disaster funding rebuilds them only “like for like”, and local contractors come to depend on the repair work, while fine sediment keeps reaching the Great Barrier Reef (Doshi, 2026).'),
  H('Responding'),
  P('I was surprised that the main barrier was not a lack of solutions but a funding rule. Full betterment clearly cut erosion, yet the response Council kept hearing was “who’s gonna pay for this?” I had assumed that sound analysis leads naturally to good decisions; the lecture showed that institutions decide what gets built.'),
  H('Relating'),
  P('Before this assignment, CEA and CBA felt like separate textbook tools. Applying both to the case data changed that. In my report, full betterment had the lowest cost per tonne and the highest NPV, yet a negative financial NPV and a $4.95M upfront cost. Seeing one option look excellent for society but unaffordable for Council made the difference between economic efficiency and financial feasibility concrete (Field & Field, 2024).'),
  H('Reasoning'),
  P('Dr Doshi’s answer about the wide range of abatement costs, from $70 to $5,745 per tonne (Doshi et al., 2026), was the most valuable insight. Non-market benefits often rest on imperfect proxies, so he stressed transparency: present the range, test it, and tell decision-makers how likely a project is to deliver net benefits. My sensitivity analysis confirmed this, since at the low value only alternative maintenance had a positive NPV. CEA avoided that valuation problem but could not show whether reducing sediment was worth it at all, which is why the methods complement each other. His reminder that economists advise rather than decide also shaped my recommendations, which present trade-offs rather than a single answer.'),
  H('Reconstructing'),
  P('In future appraisals, I will ask who pays, when and under which rules, not only which option maximises NPV. I will report ranges and switching values rather than point estimates, and keep unquantified benefits, such as fewer road closures, in the narrative. For unsealed roads, I now see that protecting downstream water quality depends as much on reforming disaster funding to allow betterment, and on Reef co-funding, as on engineering. I would also consider the transition for local contractors, which a purely microeconomic CBA overlooks.'),
];

const words = counted.join(' ').split(/\s+/).filter((w) => /[A-Za-z0-9$%]/.test(w)).length;
console.log('REFLECTION WORDS', words);

const doc = new Document({
  creator: 'Zach Allen',
  title: 'ECON3700 Part 2 Reflective Writing',
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })] })] }) },
    children: [
      C('Reflecting on the Industry Expert: CEA and CBA in Unsealed Road Management', true),
      C('Zach Allen (Student ID: 49610350)'),
      C('ECON3700: Environmental Economics, The University of Queensland'),
      C(`Word count: ${words} (excluding title lines and references)`),
      ...body,
      C('References', true),
      ref('Doshi, A. (2026, October 21). *Unsealed road management in Cook Shire* [Guest lecture slides and recording]. ECON3700 Environmental Economics, The University of Queensland.'),
      ref('Doshi, A., Erlandsen, A., Charlton-Henderson, S., Joyse, K., & Nyman, P. (2026). Incorporating fine sediment pollution impacts into unsealed road management: A case study from the Great Barrier Reef catchment area. *Environmental Management, 76*(4), Article 128. https://doi.org/10.1007/s00267-026-02423-2'),
      ref('Field, B. C., & Field, M. K. (2024). *Environmental economics: An introduction* (9th ed.). McGraw Hill.'),
    ],
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync(OUT, b));
