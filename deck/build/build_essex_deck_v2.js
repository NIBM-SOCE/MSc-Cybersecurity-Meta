// NIBM x University of Essex - MSc Cyber Security business case deck
// Built to nibm_presentation_style.yaml v3.1 (warm-earth edition)
const pptxgen = require('pptxgenjs');
const fs = require('fs');

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Niranga Dharmaratna';
pres.title = 'MSc Cyber Security at NIBM - Business case for the University of Essex';

// ---------------------------------------------------------------- palette
const C = {
  brown: '7A5230', orange: 'EB5600', tan: 'C4A882', cream: 'F5EDE0', ivory: 'F0DFC0',
  sand: 'F2E6D0', caramel: 'A07848', navy: '1C3678', ink: '1A1A1A', grey: '595959',
  white: 'FFFFFF', warmgrey: 'BDBDAC', palecream: 'F9F0E4', orangetint: 'FFF5E8',
  border: 'CCCCCC', grid: 'E5DED2',
};
const F = 'Open Sans';
const LOGO = 'image/png;base64,' + fs.readFileSync('/mnt/user-data/uploads/NIBM-logo.png').toString('base64');
const FOOT = (process.env.DECK === 'appendix') ? 'NIBM SOCE - Appendix to the Essex proposal - September 2026' : 'NIBM SOCE - Proposal to the University of Essex - September 2026';

const MODE = process.env.DECK || 'main';
let slideNo = 0;

// ---------------------------------------------------------------- helpers
function rect(s, x, y, w, h, fill, extra) {
  s.addShape(pres.shapes.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { type: 'none' } }, extra || {}));
}
function txt(s, t, o) {
  s.addText(t, Object.assign({ fontFace: F, isTextBox: true, color: C.ink, valign: 'top', margin: 0 }, o));
}
function chrome(s, title, act, big) {
  slideNo++;
  const fs_ = big ? 28 : 24, TX = 0.45, TY = 0.62, TH = big ? 0.68 : 0.56;
  rect(s, 0, 0, 10, 0.53, C.cream);
  txt(s, title, { x: TX, y: TY, w: 9.1, h: TH, fontSize: fs_, bold: true, color: C.ink, valign: 'middle' });
  rect(s, TX, TY + TH + 0.05, 0.55, 0.07, C.brown);
  rect(s, TX + 0.60, TY + TH + 0.05, 0.28, 0.07, C.orange);
  // footer
  txt(s, act, { x: 0.45, y: 5.28, w: 4.5, h: 0.25, fontSize: 11, bold: true, color: C.caramel, charSpacing: 1 });
  txt(s, FOOT + '   ' + slideNo, { x: 4.0, y: 5.28, w: 5.55, h: 0.25, fontSize: 11, color: C.grey, align: 'right' });
  return TY + TH + 0.05 + 0.07 + 0.18; // content start y (about 1.55)
}
function card(s, x, y, w, h, stripColor, title, body, opts) {
  opts = opts || {};
  rect(s, x, y, w, h, opts.fill || C.white, { line: { color: C.border, width: 0.75 } });
  rect(s, x, y, 0.08, h, stripColor);
  txt(s, title, { x: x + 0.18, y: y + 0.10, w: w - 0.28, h: opts.titleH || 0.30, fontSize: opts.titleSize || 14, bold: true, color: opts.titleColor || stripColor });
  txt(s, body, { x: x + 0.18, y: y + 0.10 + (opts.titleH || 0.30), w: w - 0.28, h: h - 0.20 - (opts.titleH || 0.30), fontSize: opts.bodySize || 12, color: C.ink, valign: 'top' });
}
function notes(s, t) { s.addNotes(t); }

const ACT1 = 'ACT I - WHO NIBM IS', ACT2 = 'ACT II - WHAT THE MARKET SAYS', ACT3 = 'ACT III - HOW WE DELIVER';

if (MODE !== 'appendix') {
// ================================================================ 1. COVER
{
  const s = pres.addSlide(); slideNo++;
  s.background = { color: C.white };
  txt(s, 'National Institute of Business Management (NIBM)\nSchool of Computing and Engineering', { x: 0.42, y: 0.32, w: 7.0, h: 0.60, fontSize: 13, bold: true, color: C.ink });
  txt(s, 'PARTNERSHIP PROPOSAL TO THE UNIVERSITY OF ESSEX', { x: 0.42, y: 0.97, w: 6.0, h: 0.28, fontSize: 12, bold: true, color: C.brown, charSpacing: 1 });
  txt(s, 'MSc Cyber Security\nat NIBM', { x: 0.42, y: 1.30, w: 6.30, h: 2.10, fontSize: 38, bold: true, color: C.ink, valign: 'middle' });
  rect(s, 0.42, 3.44, 0.82, 0.09, C.brown);
  rect(s, 1.30, 3.44, 0.42, 0.09, C.orange);
  txt(s, 'with Postgraduate Diploma exit - a franchise business case', { x: 0.42, y: 3.62, w: 5.8, h: 0.48, fontSize: 20, italic: true, color: C.grey });
  txt(s, 'Niranga Dharmaratna - niranga@nibm.lk\nSeptember 2026 - 15 minutes', { x: 0.42, y: 4.95, w: 5.5, h: 0.5, fontSize: 13, color: C.grey });
  // right: logo and award-structure motif
  s.addImage({ data: LOGO, x: 6.55, y: 0.45, w: 3.05, h: 0.77 });
  // nested award structure graphic (flat vector)
  rect(s, 6.55, 1.75, 3.05, 3.05, C.brown);
  txt(s, 'MSc Cyber Security', { x: 6.70, y: 1.85, w: 2.8, h: 0.75, fontSize: 16, bold: true, color: C.white });
  rect(s, 6.80, 2.70, 2.55, 1.35, C.tan);
  txt(s, 'PgDip exit award', { x: 6.95, y: 2.80, w: 2.3, h: 0.6, fontSize: 16, bold: true, color: C.ink });
  rect(s, 6.95, 3.45, 2.25, 0.45, C.orange);
  txt(s, 'Taught at NIBM Colombo', { x: 6.95, y: 3.45, w: 2.25, h: 0.45, fontSize: 11, bold: true, color: C.white, align: 'center', valign: 'middle' });
  txt(s, 'Essex award - Essex standards - NIBM delivery', { x: 6.70, y: 4.25, w: 2.8, h: 0.45, fontSize: 12, italic: true, color: C.white, valign: 'middle' });
  notes(s, '0:00 to 0:45. One sentence: we are asking Essex to franchise its MSc Cyber Security, with the PgDip as an exit award, for evening and weekend delivery at NIBM Colombo. The next fifteen minutes answer the three things your partnerships framework asks: is the partner suitable, is the demand real, and can it be delivered and scaled.');
}

// ================================================================ 2. STATUTORY INSTITUTION (timeline)
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'A statutory institution since 1976', ACT1);
  txt(s, 'Created by statute, accountable to a ministry, running foreign university awards for thirty years. The claim is permanence, not age.', { x: 0.45, y: y0, w: 9.1, h: 0.42, fontSize: 13, italic: true, color: C.grey });
  const pts = [
    ['1968', 'Founded with UNDP and ILO support'],
    ['1976', 'Act of Parliament No. 23 of 1976'],
    ['1980', 'First Diploma in Computer System Design'],
    ['1996', 'First foreign award (UCD); Kandy campus'],
    ['2010', 'Galle campus opens'],
    ['2017', 'Coventry University degrees begin'],
    ['2026', '~120 computing graduates a year'],
  ];
  const lineY = 3.15, x0 = 0.85, x1 = 9.15, n = pts.length;
  rect(s, x0, lineY - 0.02, x1 - x0, 0.04, C.tan);
  pts.forEach((p, i) => {
    const cx = x0 + (x1 - x0) * i / (n - 1);
    const above = i % 2 === 0;
    s.addShape(pres.shapes.OVAL, { x: cx - 0.11, y: lineY - 0.11, w: 0.22, h: 0.22, fill: { color: i === 1 ? C.orange : C.brown }, line: { type: 'none' } });
    const ty = above ? lineY - 1.02 : lineY + 0.22;
    txt(s, p[0], { x: cx - 0.7, y: above ? ty : ty, w: 1.4, h: 0.28, fontSize: 14, bold: true, color: i === 1 ? C.orange : C.brown, align: 'center' });
    txt(s, p[1], { x: cx - 0.7, y: ty + 0.28, w: 1.4, h: 0.62, fontSize: 11, color: C.ink, align: 'center' });
  });
  // campus strip
  rect(s, 0.45, 4.58, 9.1, 0.55, C.sand);
  rect(s, 0.45, 4.58, 0.09, 0.55, C.caramel);
  txt(s, 'State-owned institute under the Ministry of Youth Affairs and Skills Development. Regional campuses: Colombo (1968), Kurunegala (1986), Kandy (1996), Galle (2010), Matara; two innovation centres with Limkokwing University.', { x: 0.68, y: 4.63, w: 8.75, h: 0.45, fontSize: 11, color: C.ink, valign: 'middle' });
  notes(s, '0:45 to 2:00. Essex\'s partnership principles put partner reputation and standing first, and the sustained reputation of Essex as paramount. A statutory body cannot quietly disappear, rebrand or be sold, which is the failure mode UK universities fear most in TNE. The two dates do different work: 1968 is origin, 1976 is legal personality. Computing since 1980, foreign awards since 1996.');
}

// ================================================================ 3. PARTNERSHIP PORTFOLIO (two column)
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'Thirty years of delivering external awards', ACT1);
  txt(s, 'Every degree NIBM delivers today already runs under an external awarding body\'s quality oversight. A franchise is a familiar operating model, not a new one.', { x: 0.45, y: y0, w: 9.1, h: 0.42, fontSize: 13, italic: true, color: C.grey });
  const colW = 4.52, gap = 0.11, x0 = 0.45, hy = y0 + 0.50;
  const heads = [['Academic awards and pathways', C.brown], ['Industry academies and centres', C.orange]];
  const cols = [
    [['Coventry University (UK)', 'Undergraduate degrees at NIBM since 2017, including BSc (Hons) Ethical Hacking and Network Security'],
     ['University College Dublin (Ireland)', 'First foreign award at NIBM: BSc Management Information Systems, 1996'],
     ['Griffith University (Australia)', 'Transfer pathway for NIBM students'],
     ['Ural Federal; Limkokwing University', 'Engineering links; two creative-technology innovation centres']],
    [['Cisco and Red Hat academies', 'Network and systems curricula delivered to vendor standard'],
     ['AWS Academy; EC-Council partnership', 'Cloud and security curricula, the base for cloud-hosted lab ranges'],
     ['Microsoft; Oracle', 'Platform partnerships supporting software and data teaching'],
     ['BCS accredited training provider', 'Accredited provider for BCS Agile and Business Analysis certifications in Sri Lanka']],
  ];
  heads.forEach((h, ci) => {
    const x = x0 + ci * (colW + gap);
    rect(s, x, hy, colW, 0.42, h[1]);
    txt(s, h[0], { x: x + 0.15, y: hy, w: colW - 0.3, h: 0.42, fontSize: 15, bold: true, color: C.white, valign: 'middle' });
    rect(s, x, hy + 0.42, colW, 2.70, C.white, { line: { color: C.border, width: 0.75 } });
    cols[ci].forEach((item, ri) => {
      const iy = hy + 0.42 + 0.10 + ri * 0.65;
      txt(s, item[0], { x: x + 0.15, y: iy, w: colW - 0.3, h: 0.24, fontSize: 12, bold: true, color: h[1] });
      txt(s, item[1], { x: x + 0.15, y: iy + 0.24, w: colW - 0.3, h: 0.40, fontSize: 11, color: C.ink });
    });
  });
  notes(s, '2:00 to 3:15. The point is operating maturity: external examiners, moderation, partner audits and vendor-standard curricula are already routine. NIBM matches each award to the partner strongest in it, which is why the portfolio has several names on it. If the Coventry question comes, this is the slide to return to: Coventry is the undergraduate partner and that continues; postgraduate cyber is a level the current arrangement does not cover.');
}

// ================================================================ 4. WHY ESSEX (ranking table)
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'Why Essex: top of the Colombo field', ACT1);
  const H = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.ink }, align: 'center' } });
  const E = (t, c) => ({ text: t, options: { bold: true, color: c || C.ink } });
  const N = (t, c) => ({ text: t, options: { color: c || C.ink, align: 'center' } });
  const rows = [
    [H('Awarding university'), H('Via'), H('CUG'), H('Guardian'), H('NCSC status, UK programme')],
    [E('University of Essex', C.brown), E('NIBM (proposed)', C.brown), E('29', C.brown), E('12', C.brown), E('Provisional, to Sep 2027', C.brown)],
    ['Lancashire (formerly UCLan)', 'UCL Sri Lanka', N('86'), N('98'), N('Not listed')],
    ['Kingston', 'ESOFT', N('89'), N('55'), N('Full, expires 30 Sep 2026')],
    ['Staffordshire', 'APIIT', N('94'), N('82'), N('Not listed')],
    ['Westminster', 'IIT', N('123'), N('122'), N('Full, to Sep 2028', C.orange)],
  ];
  rows[1][2].options.align = 'center'; rows[1][3].options.align = 'center'; rows[1][4].options.align = 'center';
  s.addTable(rows, { x: 0.45, y: y0, w: 5.8, colW: [1.55, 1.1, 0.6, 0.95, 1.6], fontFace: F, fontSize: 11, color: C.ink, border: { type: 'solid', color: C.border, pt: 0.5 }, rowH: [0.32, 0.4, 0.4, 0.4, 0.4, 0.4], valign: 'middle', margin: 0.04, fill: { color: C.white } });
  txt(s, 'UK awards delivered via Sri Lankan partners; 2026 league tables; NCSC status of the UK-delivered programme (May 2026 list). Franchise scope: for Essex to confirm.', { x: 0.5, y: y0 + 2.5, w: 5.7, h: 0.6, fontSize: 11, italic: true, color: C.grey });
  // right tiles
  const rx = 6.45, rw = 3.1;
  rect(s, rx, y0, rw, 1.2, C.brown);
  txt(s, '6th', { x: rx + 0.15, y: y0 + 0.08, w: rw - 0.3, h: 0.6, fontSize: 36, bold: true, color: C.white, align: 'center', valign: 'middle' });
  txt(s, 'UK research power in computer science (THE, REF 2021)', { x: rx + 0.15, y: y0 + 0.70, w: rw - 0.3, h: 0.45, fontSize: 11, bold: true, color: C.white, align: 'center' });
  rect(s, rx, y0 + 1.32, rw, 1.2, C.orange);
  txt(s, '120', { x: rx + 0.15, y: y0 + 1.40, w: rw - 0.3, h: 0.6, fontSize: 36, bold: true, color: C.white, align: 'center', valign: 'middle' });
  txt(s, 'credit PgDip exit in Essex\'s own award portfolio', { x: rx + 0.15, y: y0 + 2.02, w: rw - 0.3, h: 0.45, fontSize: 11, bold: true, color: C.white, align: 'center' });
  rect(s, 0.45, y0 + 3.14, 9.1, 0.48, C.ink);
  txt(s, 'Campus MSc profile: security by design, digital identity, digital forensics, network security. The natural continuation of NIBM\'s Ethical Hacking and Network Security stream.', { x: 0.62, y: y0 + 3.14, w: 8.8, h: 0.48, fontSize: 12, bold: true, color: C.white, valign: 'middle' });
  notes(s, '3:15 to 4:30. A fit argument, made concrete. Among the UK awards already delivered in Colombo, Essex leads both league tables, and by a wide margin in the Guardian. Research standing, a recognisable name, an award architecture that already includes the exit diploma, and a programme profile that continues what our graduates studied. Refresh the table with the 2027 editions before the meeting. Do not mention Coventry here; if asked, NIBM matches each award to the partner strongest in it, and postgraduate cyber is a level the current arrangement does not cover. Certification is stated honestly: Essex\'s provisional status beats three of the four local UK awards, which are not on the NCSC list at all, and trails only Westminster; Kingston\'s certification lapses on 30 September 2026 unless renewed. Whether any of this can be claimed for a franchised delivery is Essex\'s to confirm.');
}

// ================================================================ 5. THE BASE AND THE MISSING RUNG
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'One computing base, and no rung above Level 6', ACT2);
  const labels = ['Mar 2024', 'Jul 2024', 'Mar 2025', 'Jul 2025', 'Mar 2026', 'Sep 2026'];
  const data = [
    { name: 'Computing / Software Engineering', labels, values: [57, 9, 31, 49, 30, 34] },
    { name: 'IT for Business', labels, values: [15, 11, 15, 8, 18, 19] },
    { name: 'Ethical Hacking and Network Security', labels, values: [13, 2, 14, 7, 8, 6] },
  ];
  s.addChart(pres.charts.BAR, data, {
    x: 0.45, y: y0 - 0.05, w: 5.4, h: 3.0, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 55,
    chartColors: [C.tan, C.caramel, C.orange], showValue: true, dataLabelPosition: 'ctr', dataLabelFontSize: 11, dataLabelColor: C.white, dataLabelFontFace: F,
    catAxisLabelFontSize: 11, catAxisLabelColor: C.ink, catAxisLabelFontFace: F, valAxisLabelFontSize: 11, valAxisLabelColor: C.grey, valAxisLabelFontFace: F,
    valAxisMaxVal: 100, valAxisMajorUnit: 20, valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: true, legendPos: 'b', legendFontSize: 11, legendFontFace: F, legendColor: C.ink, showTitle: true, title: 'Graduates by convocation, SOCE', titleFontSize: 11, titleColor: C.grey, titleFontFace: F,
  });
  txt(s, 'Yearly totals 107 (2024), 124 (2025), 115 (2026). 346 graduates in the last six convocations, none with a postgraduate route at NIBM.', { x: 0.55, y: y0 + 2.98, w: 5.3, h: 0.6, fontSize: 11, italic: true, color: C.grey });
  const sx = 6.15, sw = 3.4;
  const stats = [['~120', 'computing graduates a year across three programmes', C.brown, C.white],
                 ['17 to 51', 'year-3 entries per intake to the cyber stream, Nov 2024 to Nov 2025; Nov 2026 expected above 25', C.orange, C.white],
                 ['0', 'postgraduate cyber places at NIBM today', C.tan, C.ink]];
  stats.forEach((st, i) => {
    const sy = y0 + 0.02 + i * 1.2;
    rect(s, sx, sy, sw, 1.08, st[2]);
    txt(s, st[0], { x: sx + 0.12, y: sy + 0.06, w: 1.35, h: 0.96, fontSize: 22, bold: true, color: st[3], valign: 'middle' });
    txt(s, st[1], { x: sx + 1.5, y: sy + 0.08, w: sw - 1.62, h: 0.92, fontSize: 11, bold: true, color: st[3], valign: 'middle' });
  });
  notes(s, '4:30 to 5:45. The base is established: about 120 computing graduates a year across Software Engineering, IT for Business and the cyber stream, and 346 in the last two and a half years with nowhere to go for a Masters at NIBM. The cyber stream itself is young and small at graduation today, about 20 a year, but degree intake at SOCE is entry to year 3, and those entries have jumped from 17 to 51 per intake, so the direct-entry pool roughly doubles from 2027. A curriculum review and rename to Applied Cybersecurity is pending. The programme is therefore pitched as the postgraduate exit for the whole computing base, with two entry lanes shown on the delivery slide.');
}

// ================================================================ 6. DEMAND (grouped bar)
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'Demand rises with a UK award', ACT2);
  const data = [
    { name: 'NIBM programme', labels: ['Very likely', 'Somewhat likely', 'Neither', 'Somewhat unlikely', 'Very unlikely'], values: [20, 22, 10, 3, 2] },
    { name: 'Same programme, UK award', labels: ['Very likely', 'Somewhat likely', 'Neither', 'Somewhat unlikely', 'Very unlikely'], values: [37, 11, 7, 1, 1] },
  ];
  s.addChart(pres.charts.BAR, data, {
    x: 0.45, y: y0 - 0.05, w: 5.9, h: 3.65, barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 60,
    chartColors: [C.tan, C.brown], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 11, dataLabelColor: C.ink, dataLabelFontFace: F,
    catAxisLabelFontSize: 11, catAxisLabelColor: C.ink, catAxisLabelFontFace: F, valAxisLabelFontSize: 11, valAxisLabelColor: C.grey, valAxisLabelFontFace: F,
    valAxisMaxVal: 40, valAxisMajorUnit: 8, valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: true, legendPos: 'b', legendFontSize: 11, legendFontFace: F, legendColor: C.ink,
  });
  // callout
  const cx = 6.6, cw = 2.95;
  rect(s, cx, y0, cw, 1.15, C.brown);
  txt(s, '74% to 84%', { x: cx + 0.15, y: y0 + 0.08, w: cw - 0.3, h: 0.5, fontSize: 20, bold: true, color: C.white });
  txt(s, 'likely or very likely to enrol, before and after the UK award', { x: cx + 0.15, y: y0 + 0.58, w: cw - 0.3, h: 0.55, fontSize: 11, bold: true, color: C.white });
  rect(s, cx, y0 + 1.27, cw, 1.15, C.orange);
  txt(s, '23 up, 0 down', { x: cx + 0.15, y: y0 + 1.35, w: cw - 0.3, h: 0.5, fontSize: 20, bold: true, color: C.white });
  txt(s, 'respondents who changed their answer on hearing the award is UK-issued', { x: cx + 0.15, y: y0 + 1.85, w: cw - 0.3, h: 0.55, fontSize: 11, bold: true, color: C.white });
  rect(s, cx, y0 + 2.54, cw, 1.1, C.sand);
  txt(s, 'n = 57, self-selected internal sample, Sep 2026; price untested. Raised on GCE O/L and A/L, Edexcel and Cambridge, this market knows UK awards.', { x: cx + 0.12, y: y0 + 2.58, w: cw - 0.24, h: 1.04, fontSize: 11, color: C.ink });
  notes(s, '5:45 to 7:00. The value of the partnership expressed as a number. The top box nearly doubles, 20 to 37. Twenty-three respondents moved up and none moved down; seventeen of them shifted into "very likely", which is exactly what a partnership case needs. The familiarity behind it: Sri Lanka\'s own school examinations are GCE O/L and A/L, and Edexcel and Cambridge run through the British Council and the international schools, so a UK award is the known currency. State the limits yourself: convenience sample, two days, no fee tested. The UK question followed the NIBM question, so treat the magnitude as slightly overstated and the direction as solid.');
}

// ================================================================ 7. AFFORDABILITY CORRIDOR
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'The affordability corridor', ACT2);
  const data = [{ name: 'LKR millions', labels: ['Local UK-partner band, upper', 'Essex Online MSc (international fee)', 'Essex campus MSc (international tuition only)'], values: [0.78, 5.9, 11.1] }];
  s.addChart(pres.charts.BAR, data, {
    x: 0.45, y: y0 - 0.05, w: 5.3, h: 2.5, barDir: 'bar', chartColors: [C.orange, C.caramel, C.brown], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 11, dataLabelColor: C.ink, dataLabelFontFace: F, dataLabelFormatCode: '0.0',
    catAxisLabelFontSize: 11, catAxisLabelColor: C.ink, catAxisLabelFontFace: F, valAxisLabelFontSize: 11, valAxisLabelColor: C.grey, valAxisLabelFontFace: F,
    valAxisMaxVal: 14, valAxisMajorUnit: 2, valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: 'none' }, showLegend: false, barGapWidthPct: 45,
    showTitle: true, title: 'Cost of a UK Cyber Security MSc to a Sri Lankan student, LKR millions', titleFontSize: 11, titleColor: C.grey, titleFontFace: F,
  });
  txt(s, 'GBP 24,675 and GBP 13,049 at about LKR 450 per pound (August 2026). Local band LKR 0.35 to 0.78 million from the NIBM competitor matrix; four of six local UK-partner routes publish no fee at all. Living costs excluded.', { x: 0.55, y: y0 + 2.5, w: 5.2, h: 0.5, fontSize: 11, italic: true, color: C.grey });
  rect(s, 0.45, y0 + 3.05, 5.3, 0.55, C.sand);
  rect(s, 0.45, y0 + 3.05, 0.09, 0.55, C.caramel);
  txt(s, 'Sterling bought about 264 rupees at the 2022 low and about 450 today: the same UK fee costs a Sri Lankan family roughly 70% more rupees than four years ago.', { x: 0.68, y: y0 + 3.08, w: 5.0, h: 0.5, fontSize: 11, color: C.ink, valign: 'middle' });
  const px = 6.05, pw = 3.5;
  rect(s, px, y0, pw, 3.6, C.orangetint, { line: { color: C.orange, width: 1 } });
  txt(s, 'Why the corridor matters now', { x: px + 0.18, y: y0 + 0.12, w: pw - 0.36, h: 0.3, fontSize: 13, bold: true, color: C.orange });
  txt(s, [
    { text: 'Entry-level IT pay starts near LKR 80,000 a month; the online Essex MSc equals about six years of it.', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'SLASSCOM: about 17,000 IT graduates a year; entry-level supply is no longer the bottleneck.', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'AI is compressing generalist roles: IFS cut 400 to 450 Sri Lankan staff from Nov 2025, citing AI.', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'The rational response is to specialise. A UK Masters priced locally is the only route within reach, a segment neither Essex product serves.', options: { bullet: true } },
  ], { x: px + 0.18, y: y0 + 0.46, w: pw - 0.36, h: 3.1, fontSize: 11, color: C.ink });
  notes(s, '7:00 to 8:15. This slide does two jobs. It shows that Essex\'s own two products sit an order of magnitude above what a Sri Lankan graduate can pay, which is the honest segmentation answer to Essex Online: the franchise serves people who cannot buy either. And it turns the labour-market news into a demand driver: entry-level generalist roles are oversupplied and being compressed by AI, so the graduate who wants to stay employable has to specialise, and the only affordable route to a UK Masters is a locally delivered one. Do not lead with layoffs; lead with specialise or stall.');
}

// ================================================================ 8. COMMITTED DEMAND (funnel + tiers)
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'Committed demand, discounted in the open', ACT2);
  const steps = [
    [57, 'responded', C.tan, C.ink, 5.4],
    [41, 'positive on a UK award and able to start within 12 months', C.caramel, C.white, 4.7],
    [29, 'central weighting (0.8 / 0.3)', C.brown, C.white, 4.0],
    [21, 'conservative weighting (0.6 / 0.2)', C.orange, C.white, 3.3],
  ];
  const cxm = 3.15;
  steps.forEach((st, i) => {
    const w = st[4], x = cxm - w / 2, y = y0 + 0.05 + i * 0.86, h = 0.74;
    rect(s, x, y, w, h, st[2]);
    txt(s, String(st[0]), { x: x + 0.12, y: y + 0.05, w: 1.0, h: h - 0.1, fontSize: 30, bold: true, color: st[3], valign: 'middle' });
    txt(s, st[1], { x: x + 1.15, y: y + 0.08, w: w - 1.3, h: h - 0.16, fontSize: 11, bold: true, color: st[3], valign: 'middle' });
  });
  const px = 6.15, pw = 3.4;
  rect(s, px, y0, pw, 3.55, C.orangetint, { line: { color: C.orange, width: 1 } });
  txt(s, 'Per intake, by entry lane', { x: px + 0.18, y: y0 + 0.12, w: pw - 0.36, h: 0.3, fontSize: 13, bold: true, color: C.orange });
  txt(s, [
    { text: 'Survey rates: 37.5% conservative, 50% central (21 and 29 of 57).', options: { breakLine: true, paraSpaceAfter: 6 } },
    { text: 'Direct lane, ~20 cyber graduates a year today:  7 to 10', options: { bold: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: 'Bridged lane, ~100 other computing graduates, rate halved for bridge friction:  19 to 25', options: { bold: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: 'Together:  26 to 35 per intake', options: { bold: true, breakLine: true, paraSpaceAfter: 6 } },
    { text: 'Not counted: 346 recent alumni, working professionals, and the direct lane doubling from 2027. Public benchmark: UCSC\'s state-sector Master of Cybersecurity took 28 in its 2023 intake.', options: { italic: true } },
  ], { x: px + 0.18, y: y0 + 0.46, w: pw - 0.36, h: 3.05, fontSize: 11, color: C.ink });
  notes(s, '8:15 to 9:30. Showing the discount is more persuasive than hiding it. Forty-one are positive on a UK award and could start within a year; apply purchase-intent discounting and the central figure is 29, the conservative 21. Applied by lane, the first intake clears a viable cohort on arithmetic Essex can check, before alumni, professionals, or the growth of the direct lane. Fifteen of the nineteen who said within 6 months are still undergraduates, so real intake pressure lands 6 to 12 months after graduation. The planning assumption agreed in negotiation for the first intake is fifteen; that number is not stated on any slide, and the case is that the floor sits comfortably above it.');
}

// ================================================================ 9. INDUSTRY IN ITS OWN WORDS
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'Industry describes a postgraduate gap', ACT2);
  const data = [{ name: 'Respondents (of 7)', labels: ['Would sponsor staff (yes)', 'Willing to contribute to delivery', 'UK Masters is a hiring advantage', 'Expect demand to rise'], values: [2, 6, 6, 7] }];
  s.addChart(pres.charts.BAR, data, {
    x: 0.45, y: y0 - 0.05, w: 4.55, h: 2.0, barDir: 'bar', chartColors: [C.brown], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 11, dataLabelColor: C.ink, dataLabelFontFace: F,
    catAxisLabelFontSize: 11, catAxisLabelColor: C.ink, catAxisLabelFontFace: F, valAxisLabelFontSize: 11, valAxisLabelColor: C.grey, valAxisLabelFontFace: F,
    valAxisMaxVal: 7, valAxisMajorUnit: 1, valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: 'none' }, showLegend: false, barGapWidthPct: 45,
  });
  txt(s, 'Sponsorship: 2 yes, 4 maybe, 0 no, 1 not applicable. Contribution: 6 yes, 1 maybe. n = 7 senior respondents, none NIBM staff: IT services, telecom, insurance and its regulator, three national-university academics.', { x: 0.55, y: y0 + 1.98, w: 4.45, h: 0.78, fontSize: 11, color: C.grey });
  rect(s, 0.45, y0 + 2.80, 4.55, 0.80, C.ivory);
  rect(s, 0.45, y0 + 2.80, 0.09, 0.80, C.caramel);
  txt(s, '"Graduates may identify vulnerabilities but struggle to validate findings, prioritise business risks, and recommend fixes that account for availability, cost, and operational constraints."  - IT services practitioner', { x: 0.68, y: y0 + 2.82, w: 4.25, h: 0.76, fontSize: 11, italic: true, color: C.ink, valign: 'middle' });
  const rows = [
    [{ text: 'What industry asked for', options: { bold: true, color: C.white, fill: { color: C.ink } } }, { text: 'Essex MSc profile', options: { bold: true, color: C.white, fill: { color: C.ink } } }],
    ['Digital forensics; incident response', { text: 'Forensics core; IR to map to modules', options: { color: C.brown, bold: true } }],
    ['Network security in production', { text: 'Core to the campus MSc', options: { color: C.brown, bold: true } }],
    ['Application and API security', { text: 'Security by design; extend locally', options: { color: C.brown, bold: true } }],
    ['Cloud security and architecture', { text: 'Local option or capstone', options: { color: C.orange, bold: true } }],
    ['GRC as enterprise risk', { text: 'Local option or capstone', options: { color: C.orange, bold: true } }],
    ['Employer-reviewed capstone', { text: 'Dissertation and capstone', options: { color: C.brown, bold: true } }],
  ];
  s.addTable(rows, { x: 5.25, y: y0, w: 4.3, colW: [2.25, 2.05], fontFace: F, fontSize: 11, color: C.ink, border: { type: 'solid', color: C.border, pt: 0.5 }, rowH: [0.32, 0.44, 0.44, 0.44, 0.44, 0.44, 0.44], valign: 'middle', margin: 0.05, fill: { color: C.white } });
  notes(s, '9:30 to 10:30. Seven named senior people, none NIBM staff, independently describe the same deficiency: graduates can operate tools but cannot exercise judgement in live environments. That is a description of a postgraduate learning outcome, which is the strongest thing an industry survey can say for a Masters. Forensics and network security overlap directly with the Essex programme profile as published; the mapping is to that profile, and a module-level mapping is the next step once Essex shares the specification. Cloud, application security and enterprise-risk governance are where local options and the capstone add flavour. All seven expect demand to rise; six of seven rate a UK Masters an advantage; six of seven volunteered to contribute. The seventh response, from a life insurer, adds AI automation and security operations to the curriculum ask, which sits naturally in the capstone and local options. If the single "no difference" is raised, one line only: it comes from within the state-university system, which is not the market this programme serves. Move on. Quote by sector only; respondents did not consent to attributed publication.');
}

// ================================================================ 10. DELIVERY MODEL
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'The panel already teaching the stream below', ACT3);
  const cw = 4.7, ch = 0.80, gap = 0.05;
  const panel = [
    [C.brown, 'Assoc. Prof. Manjula Sandirigama', 'Peradeniya; secure protocols; Attorney-at-Law. Leads Cybersecurity II and Research Project Delivery.'],
    [C.brown, 'Dr Krishanthmohan Ratnam', 'Sri Jayewardenepura; survivable network design. Leads Networked Systems and Advanced Network Systems.'],
    [C.orange, 'Industry practitioners and further academics', 'IT services, telecom, insurance and its regulator. Four industry and two further academic volunteers so far.'],
    [C.caramel, 'NIBM cyber faculty', 'Programme coordination, quality processes, assessment administration, student support.'],
  ];
  txt(s, 'Proposed panel, subject to availability and approval', { x: 0.45, y: y0 - 0.02, w: cw, h: 0.28, fontSize: 11, bold: true, color: C.brown });
  panel.forEach((p, i) => card(s, 0.45, y0 + 0.26 + i * (ch + gap), cw, ch, p[0], p[1], p[2], { titleSize: 11, bodySize: 11, titleH: 0.22 }));
  const rx = 5.35, rw = 4.2;
  rect(s, rx, y0 - 0.02, rw, 2.22, C.tan);
  txt(s, 'Structure and entry', { x: rx + 0.18, y: y0 + 0.06, w: rw - 0.36, h: 0.3, fontSize: 13, bold: true, color: C.ink });
  txt(s, [
    { text: 'MSc 180 credits, PgDip exit at 120; evening and weekend delivery', options: { bullet: true, breakLine: true, paraSpaceAfter: 3 } },
    { text: 'Employer-reviewed capstone; supervision from the university panel', options: { bullet: true, breakLine: true, paraSpaceAfter: 3 } },
    { text: 'Direct entry: network and cyber graduates', options: { bullet: true, breakLine: true, paraSpaceAfter: 3 } },
    { text: 'Bridged entry for other computing graduates, subject to Essex approval: professional qualification plus experience, or a priced NIBM bridge', options: { bullet: true } },
  ], { x: rx + 0.18, y: y0 + 0.38, w: rw - 0.36, h: 1.8, fontSize: 11, color: C.ink });
  rect(s, rx, y0 + 2.30, rw, 1.28, C.orangetint, { line: { color: C.orange, width: 1 } });
  txt(s, 'Laboratories', { x: rx + 0.18, y: y0 + 2.36, w: rw - 0.36, h: 0.3, fontSize: 13, bold: true, color: C.orange });
  txt(s, [
    { text: 'Offensive labs on BYOD; defensive, cloud and forensic labs on cloud-hosted ranges', options: { bullet: true, breakLine: true, paraSpaceAfter: 3 } },
    { text: 'AWS Academy, EC-Council partnership; dedicated facility requested', options: { bullet: true } },
  ], { x: rx + 0.18, y: y0 + 2.66, w: rw - 0.36, h: 0.9, fontSize: 11, color: C.ink });
  notes(s, '10:30 to 11:45. Continuity is the argument: the same industry-and-academia model that delivers the Coventry stream under external oversight extends one level up. Essex gets senior national-university academics with research standing, not lecturers hired for a launch. Essex approves every teacher by name, so a CV pack for every module leader and supervisor is ready on request. Titles as on the Peradeniya department page and Google Scholar; Dr Ratnam\'s head-of-department appointment is evidenced by the faculty announcement in the evidence pack while the USJ page is down. Both are current module leaders on the undergraduate stream; their place on the postgraduate panel is expected but depends on availability and on Essex\'s approval of each name, which is why the heading says proposed. Lane one entry has a precedent in Essex Online, which admits a relevant professional qualification; lane two is an NIBM certificate before the Essex programme, not credit within it.');
}

// ================================================================ 11. RISKS
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'Risks we would rather name first', ACT3);
  const H = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.ink } } });
  const R = (t) => ({ text: t, options: { bold: true, color: C.brown } });
  const rows = [
    [H('Risk'), H('Evidence today'), H('Mitigation'), H('Owner')],
    [R('Price untested'), 'No fee question; 40% would need a loan or scholarship', 'Fee-band question before any forecast; explore employer sponsorship (2 yes, 4 maybe)', 'NIBM'],
    [R('Staff approval'), 'Panel delivers the UG stream; Essex approves each teacher by name', 'CV pack for every module leader and supervisor, ready on request', 'NIBM, Essex'],
    [R('No dedicated facility yet'), 'BYOD and cloud labs; physical facility requested, pending', 'Cloud-first launch; physical lab as phase two', 'NIBM'],
    [R('Essex Online overlap'), 'Online MSc sold into Sri Lanka at about LKR 5.9 million', 'Different segment: locally priced cohort with labs; buyers who cannot reach the online fee', 'Joint'],
    [R('NCSC scope'), 'Campus MSc provisionally certified to Sep 2027; franchise scope unknown', 'Confirm what may be claimed locally; no claim until confirmed', 'Essex'],
    [R('Bridge cost and criteria'), 'ISC2 free CC route closed May 2026; criteria are Essex\'s to approve', 'Lane two priced by NIBM; lane one criteria proposed for Essex approval', 'NIBM, Essex'],
  ];
  s.addTable(rows, { x: 0.45, y: y0, w: 9.1, colW: [1.75, 2.75, 3.5, 1.1], fontFace: F, fontSize: 11, color: C.ink, border: { type: 'solid', color: C.border, pt: 0.5 }, rowH: [0.32, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48], valign: 'middle', margin: 0.05, fill: { color: C.white } });
  notes(s, '11:45 to 13:00. Essex\'s framework says the sustained reputation of Essex is paramount and asks for all risks identified. So we name them. Two are commercial and closable in days: price, staffing paperwork. One is capital and staged. Two need Essex\'s own answer: channel overlap and certification scope. The bridge row shows we have costed the entry design rather than assumed it. Have the Essex Online segmentation sentence ready aloud; it is the most likely question.');
}

// ================================================================ 12. SCALE AND THE ASK
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'Scale, and what we ask', ACT3);
  const data = [{ name: 'Indicative places per intake', labels: ['Both lanes, central', 'Both lanes, conservative', 'Direct lane only, today'], values: [35, 26, 10] }];
  s.addChart(pres.charts.BAR, data, {
    x: 0.45, y: y0 - 0.05, w: 4.7, h: 2.6, barDir: 'bar', chartColors: [C.brown, C.caramel, C.orange], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 12, dataLabelColor: C.ink, dataLabelFontFace: F,
    catAxisLabelFontSize: 11, catAxisLabelColor: C.ink, catAxisLabelFontFace: F, valAxisLabelFontSize: 11, valAxisLabelColor: C.grey, valAxisLabelFontFace: F,
    valAxisMaxVal: 60, valAxisMajorUnit: 10, valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: 'none' }, showLegend: false, barGapWidthPct: 70,
  });
  txt(s, 'Indicative places per intake at NIBM Colombo; alumni, professionals, the growing direct lane and regional campuses not counted; not price-tested. No local route today combines a UK award, certification standing, supervised labs and an evening-and-weekend timetable.', { x: 0.55, y: y0 + 2.65, w: 4.6, h: 0.9, fontSize: 11, color: C.grey });
  // ask panel
  const ax = 5.45, aw = 4.1;
  rect(s, ax, y0, aw, 3.57, C.orange);
  txt(s, 'What we ask of Essex', { x: ax + 0.22, y: y0 + 0.15, w: aw - 0.44, h: 0.35, fontSize: 16, bold: true, color: C.white });
  txt(s, [
    { text: '1.  Agreement in principle to scope a franchise of the MSc with PgDip exit', options: { breakLine: true, paraSpaceAfter: 10 } },
    { text: '2.  Exchange of due-diligence packs: our evidence repository for your partner requirements', options: { breakLine: true, paraSpaceAfter: 10 } },
    { text: '3.  A joint scoping visit to Colombo within the next quarter', options: {} },
  ], { x: ax + 0.22, y: y0 + 0.62, w: aw - 0.44, h: 2.8, fontSize: 14, bold: true, color: C.white });
  notes(s, '13:00 to 14:15. Essex\'s framework warns that small scale is a risk indicator and that too many small partnerships are neither strategic nor sustainable. So show the floor and the layers above it: the direct lane alone is thin today, both lanes together clear a viable cohort, and the uncounted layers sit above that. The first-intake planning assumption of fifteen is well below the conservative floor; do not state it as a cap. Then the ask, specific and dated. Close by pointing to the evidence pack.');
}

// ================================================================ 12. SOURCES AND METHODS
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'Sources and methods', ACT3);
  const items = [
    ['Survey A, students and graduates', 'Microsoft Forms, 6 items, anonymous, fielded 22 to 23 Sep 2026 within NIBM, n = 57. Convenience sample, no population denominator, no fee question. 86% current undergraduates. Purchase-intent weighting stated on each figure.'],
    ['Survey B, industry and external academics', 'Microsoft Forms, 9 items, named, fielded 22 to 23 Sep 2026, n = 7. Purposive sample of senior practitioners and academics already engaged with NIBM. Qualitative evidence; quoted by sector only, pending attribution consent.'],
    ['NIBM institutional facts', 'NIBM Law No. 23 of 1976; SOCE website (soce.nibm.ac.lk); SOCE convocation records Mar 2024 to Sep 2026 and year-3 intake records Nov 2024 to Nov 2026.'],
    ['University of Essex', 'Partnerships framework ("Our approach to new partnerships"); MSc Cyber Security course page and fees; Essex Online award structure, fees and entry criteria; REF 2021 research power (THE measure).'],
    ['Competitive landscape', 'NIBM competitor matrix, Sep 2026; NCSC certified degrees list (May 2026); Complete University Guide and Guardian league tables, 2026 editions.'],
    ['Market and economy', 'GBP/LKR rates (OFX, Aug 2026); ITPro.lk salary guide 2026; Ceylon Today, 12 Apr 2026 (SLASSCOM graduate figures); The Examiner, 15 May 2026 (IFS layoffs); ISC2, 22 Apr 2026 (end of free CC programme).'],
  ];
  items.forEach((it, i) => {
    const iy = y0 + i * 0.60;
    rect(s, 0.45, iy, 9.1, 0.54, i % 2 ? C.white : C.palecream, { line: { color: C.border, width: 0.5 } });
    txt(s, it[0], { x: 0.60, y: iy + 0.04, w: 2.3, h: 0.46, fontSize: 12, bold: true, color: C.brown, valign: 'middle' });
    txt(s, it[1], { x: 3.0, y: iy + 0.03, w: 6.4, h: 0.48, fontSize: 11, color: C.ink, valign: 'middle' });
  });
  notes(s, 'Not presented. Kept in the deck so that any figure challenged in questions can be traced on the spot.');
}

// ================================================================ 13. CLOSING
{
  const s = pres.addSlide(); slideNo++;
  s.background = { color: C.white };
  rect(s, 6.60, 0, 3.40, 5.625, C.brown);
  rect(s, 6.85, 2.05, 2.9, 0.95, C.white);
  s.addImage({ data: LOGO, x: 6.95, y: 2.18, w: 2.7, h: 0.68 });
  txt(s, 'Questions &', { x: 0.42, y: 1.20, w: 5.8, h: 0.95, fontSize: 48, bold: true, color: C.ink, valign: 'middle' });
  txt(s, 'Discussion', { x: 0.42, y: 2.12, w: 5.8, h: 0.95, fontSize: 48, bold: true, color: C.brown, valign: 'middle' });
  rect(s, 0.42, 3.12, 0.82, 0.08, C.brown);
  rect(s, 1.30, 3.12, 0.42, 0.08, C.orange);
  txt(s, 'The evidence pack behind every figure is available to the Essex partnerships team on request.', { x: 0.42, y: 3.32, w: 5.8, h: 0.72, fontSize: 16, italic: true, color: C.grey });
  txt(s, 'Niranga Dharmaratna - niranga@nibm.lk\nSchool of Computing and Engineering, NIBM, Colombo 07', { x: 0.42, y: 4.7, w: 5.8, h: 0.6, fontSize: 13, color: C.grey });
  notes(s, '14:15 onward. Questions. Likely first three: Essex Online overlap, Coventry, NCSC scope. Answers are in the appendix deck, A1 to A3, and in the question bank in the speaker script.');
}


}

if (MODE !== 'main') {
// ================================================================ APPENDIX
const ACTA = 'APPENDIX - HELD FOR QUESTIONS';
{
  const s = pres.addSlide(); slideNo++;
  s.background = { color: C.palecream };
  rect(s, 0.80, 1.20, 0.55, 0.05, C.brown);
  txt(s, 'Appendix', { x: 0.80, y: 1.45, w: 8.41, h: 0.9, fontSize: 36, bold: true, color: C.ink });
  txt(s, 'MSc Cyber Security at NIBM: proposal to the University of Essex. Held for questions, not presented.', { x: 0.80, y: 2.30, w: 8.41, h: 0.5, fontSize: 18, italic: true, color: C.grey });
  rect(s, 0.80, 3.05, 0.82, 0.09, C.brown);
  rect(s, 1.68, 3.05, 0.42, 0.09, C.orange);
  txt(s, 'A1  Essex Online versus the franchise\nA2  Partnerships, and the choice of Essex\nA3  NCSC certification: what may be claimed\nA4  The local field and the fee band\nA5  Survey method, weighting and limitations\nA6  Verification status', { x: 0.80, y: 3.35, w: 8.4, h: 1.9, fontSize: 13, color: C.ink });
  notes(s, 'Divider. Jump to the slide that answers the question asked; do not walk through the appendix.');
}

// A1 Essex Online segmentation
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'A1  Essex Online versus the franchise', ACTA);
  const colW = 4.52, gap = 0.11, x0 = 0.45;
  const heads = [['University of Essex Online', C.brown], ['Essex award franchised to NIBM', C.orange]];
  const cols = [
    ['Wholly online, self-directed, 15 to 20 hours a week per module', 'International fee GBP 13,049, about LKR 5.9 million at LKR 450 per pound', 'BCS accredited; multiple starts a year', 'Buyer: the disciplined remote learner who can fund a UK fee'],
    ['In-country cohort, supervised labs, employer-reviewed capstone', 'Locally priced against a LKR 0.35 to 0.78 million field', 'Evening and weekend timetable built for Sri Lankan working hours', 'Buyer: the graduate or professional who cannot reach the online fee and will not buy purely online study'],
  ];
  heads.forEach((h, ci) => {
    const x = x0 + ci * (colW + gap);
    rect(s, x, y0, colW, 0.44, h[1]);
    txt(s, h[0], { x: x + 0.15, y: y0, w: colW - 0.3, h: 0.44, fontSize: 14, bold: true, color: C.white, valign: 'middle' });
    rect(s, x, y0 + 0.44, colW, 2.2, C.white, { line: { color: C.border, width: 0.75 } });
    txt(s, cols[ci].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 3, paraSpaceAfter: 6 } })), { x: x + 0.15, y: y0 + 0.55, w: colW - 0.3, h: 2.0, fontSize: 12, color: C.ink });
  });
  rect(s, 0.45, y0 + 2.78, 9.1, 0.8, C.sand);
  rect(s, 0.45, y0 + 2.78, 0.09, 0.8, C.caramel);
  txt(s, 'Survey evidence: 84% likely or very likely with a UK award at an untested price; 40% would need a loan or scholarship; entry-level IT pay near LKR 80,000 a month makes the online fee about six years of salary. The franchise adds a segment; it does not move one.', { x: 0.68, y: y0 + 2.82, w: 8.75, h: 0.72, fontSize: 11, color: C.ink, valign: 'middle' });
  notes(s, 'The answer in one breath: the online product serves a buyer who can fund a UK fee and study alone; the franchise serves a buyer who can do neither. The price gap is an order of magnitude, so the segments do not overlap; if anything the franchise creates awareness of the Essex name that the online product can later convert among employers who sponsor staff.');
}

// A2 Partnership portfolio and choice of Essex
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'A2  Partnerships, and the choice of Essex', ACTA);
  rect(s, 0.45, y0, 4.4, 3.58, C.white, { line: { color: C.border, width: 0.75 } });
  rect(s, 0.45, y0, 0.09, 3.58, C.brown);
  txt(s, 'How NIBM chooses partners', { x: 0.68, y: y0 + 0.1, w: 4.0, h: 0.3, fontSize: 13, bold: true, color: C.brown });
  txt(s, [
    { text: 'One partner per specialism, chosen on strength in that award', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'Coventry University: undergraduate computing partner since 2017; that relationship continues unchanged', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'University College Dublin (1996), Griffith transfer pathway, Ural Federal, Limkokwing: each for a specific purpose', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'Postgraduate cyber is a level no current arrangement covers: this adds a partner at that level, it replaces none', options: { bullet: true } },
  ], { x: 0.68, y: y0 + 0.45, w: 4.0, h: 3.05, fontSize: 12, color: C.ink });
  rect(s, 5.1, y0, 4.45, 3.58, C.white, { line: { color: C.border, width: 0.75 } });
  rect(s, 5.1, y0, 0.09, 3.58, C.orange);
  txt(s, 'Why Essex on the evidence', { x: 5.33, y: y0 + 0.1, w: 4.0, h: 0.3, fontSize: 13, bold: true, color: C.orange });
  txt(s, [
    { text: 'Highest-ranked UK awarding university among those delivered in Colombo: CUG 29, Guardian 12 (2026)', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: '6th in the UK for research power in computer science (REF 2021, THE measure)', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'Programme profile continues the Ethical Hacking and Network Security stream: security by design, digital identity, forensics, network security', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'The PgDip exit and PgCert already exist in Essex\'s own award portfolio', options: { bullet: true, breakLine: true, paraSpaceAfter: 5 } },
    { text: 'Essex\'s TNE strategy names Sri Lanka as a priority market', options: { bullet: true } },
  ], { x: 5.33, y: y0 + 0.45, w: 4.0, h: 3.05, fontSize: 12, color: C.ink });
  notes(s, 'Use only if the Coventry question is asked. Say it plainly and once: Coventry is and remains the undergraduate partner; NIBM matches each award to the partner strongest in it; on ranking, research standing and programme profile, Essex is the fit for postgraduate cyber. If pressed on Coventry\'s own NCSC certification: certification is of the UK-delivered programme and does not automatically travel to a franchise for any university, so it was not the deciding variable; fit and standing were.');
}

// A3 NCSC
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'A3  NCSC certification: what may be claimed', ACTA);
  const H = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.ink } } });
  const rows = [
    [H('Route delivered in Sri Lanka'), H('UK awarding programme'), H('NCSC status (list of May 2026)'), H('Expiry')],
    [{ text: 'NIBM (proposed)', options: { bold: true, color: C.brown } }, 'Essex MSc Cyber Security', { text: 'Provisionally certified', options: { bold: true, color: C.brown } }, '30 Sep 2027'],
    ['IIT', 'Westminster MSc Cyber Security and Forensics', 'Fully certified', '30 Sep 2028'],
    ['ESOFT', 'Kingston MSc Network and Information Security (cyber pathway)', 'Fully certified', '30 Sep 2026'],
    ['UCL Sri Lanka', 'Lancashire MSc Cyber Security', 'Not listed (a different Lancashire programme is)', 'n/a'],
    ['APIIT', 'Staffordshire MSc Cyber Security', 'Not listed', 'n/a'],
    ['NEXT Campus (unverified)', 'London Metropolitan MSc Cyber Security', 'Not listed', 'n/a'],
  ];
  s.addTable(rows, { x: 0.45, y: y0, w: 9.1, colW: [1.9, 3.2, 2.6, 1.4], fontFace: F, fontSize: 11, color: C.ink, border: { type: 'solid', color: C.border, pt: 0.5 }, rowH: [0.32, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34], valign: 'middle', margin: 0.04, fill: { color: C.white } });
  rect(s, 0.45, y0 + 2.88, 9.1, 0.72, C.orangetint, { line: { color: C.orange, width: 1 } });
  txt(s, 'Certification attaches to the programme as delivered and assessed by the named UK academic team. Whether it extends to a franchised delivery is Essex\'s to answer; NIBM will claim nothing locally until Essex confirms what may be said. Essex Online\'s MSc is BCS accredited, not NCSC listed.', { x: 0.62, y: y0 + 2.9, w: 8.75, h: 0.68, fontSize: 11, color: C.ink, valign: 'middle' });
  notes(s, 'Two honest sentences. Essex\'s provisional status beats three of the four local UK awards and trails Westminster; Kingston\'s certification lapses on 30 September 2026 unless renewed. What can be claimed for a franchise is for Essex to decide, and we will follow that decision exactly.');
}

// A4 Local field and fee band
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'A4  The local field and the fee band', ACTA);
  const H = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.ink } } });
  const rows = [
    [H('Provider'), H('Award'), H('Model'), H('Published fee (LKR)'), H('Mode')],
    ['UCL Sri Lanka', 'Lancashire MSc Cyber Security', 'UK award, local delivery', '325,000 + GBP 1,000 (about 775,000)', '12 months, on campus; launched Feb 2026'],
    ['IIT', 'Westminster MSc Cyber Security and Forensics', 'UK award, local delivery', 'Not published', 'Part time, 2 years'],
    ['ESOFT', 'Kingston MSc Network and Information Security', 'UK award, local delivery', '350,000 (stale listing)', 'Listed full time'],
    ['APIIT', 'Staffordshire MSc Cyber Security', 'UK award, local delivery', 'Not published', 'Not published'],
    ['SLIIT', 'MSc IT (Cyber Security)', 'Own Sri Lankan award', '550,000 (stale listing)', 'Part time, 2 years'],
    ['UCSC', 'Master of Cybersecurity', 'State university award', '615,000', '2 yrs part time, evenings and Sat; 2023 intake 28'],
  ];
  s.addTable(rows, { x: 0.45, y: y0, w: 9.1, colW: [1.3, 2.6, 1.6, 1.8, 1.8], fontFace: F, fontSize: 11, color: C.ink, border: { type: 'solid', color: C.border, pt: 0.5 }, rowH: [0.32, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42], valign: 'middle', margin: 0.04, fill: { color: C.white } });
  txt(s, 'Source: NIBM competitor matrix, 12 Sep 2026, GBP component restated at LKR 450. Two further UK-partner routes (NEXT Campus, ANC) are unverified. Four of six UK-partner routes publish no fee. Almost all provision is part time, evenings and weekends.', { x: 0.5, y: y0 + 3.0, w: 9.0, h: 0.6, fontSize: 11, italic: true, color: C.grey });
  notes(s, 'The field is real but thin: seven verified routes across six providers, four UK-awarded, one of them only seven months old. Enough to prove demand, not enough to saturate. An Essex franchise does not need to undercut anyone; it needs to justify a figure in the upper part of the band through the award, certification standing and supervised delivery. No fee for the NIBM programme is proposed in this deck; that depends on Essex\'s commercial terms.');
}

// A5 Survey method
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'A5  Survey method, weighting and limitations', ACTA);
  const colW = 4.52, gap = 0.11, x0 = 0.45;
  const heads = [['Survey A: students, n = 57', C.brown], ['Survey B: industry and academics, n = 7', C.orange]];
  const cols = [
    ['Microsoft Forms, 6 items, anonymous; open link circulated within NIBM, 22 to 23 Sep 2026', '49 current undergraduates (86%), 8 graduates; no population denominator, so no response rate', 'Committed demand = positive on a UK award and start within 12 months, then weighted 0.8 / 0.3 (central) or 0.6 / 0.2 (conservative) on the top two bands', 'Limits: self-selected, no fee question, UK question asked after the NIBM question (possible ordering effect), timing aspirational for undergraduates'],
    ['Microsoft Forms, 9 items, named; purposive sample of senior people already engaged with NIBM, none NIBM staff', 'IT services, telecom, insurance and its regulator, three national-university academics', 'Qualitative evidence: quoted by sector only, because the form did not capture attribution consent', 'Limits: n = 7; willingness to contribute captured as yes or no, so mode of contribution and teaching capacity were followed up directly'],
  ];
  heads.forEach((h, ci) => {
    const x = x0 + ci * (colW + gap);
    rect(s, x, y0, colW, 0.44, h[1]);
    txt(s, h[0], { x: x + 0.15, y: y0, w: colW - 0.3, h: 0.44, fontSize: 13, bold: true, color: C.white, valign: 'middle' });
    rect(s, x, y0 + 0.44, colW, 3.14, C.white, { line: { color: C.border, width: 0.75 } });
    txt(s, cols[ci].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 3, paraSpaceAfter: 6 } })), { x: x + 0.15, y: y0 + 0.55, w: colW - 0.3, h: 2.95, fontSize: 11, color: C.ink });
  });
  notes(s, 'If method is challenged: agree with the limits before they are stated, because they are on the slide. The value of the survey is direction, not forecast: zero downward moves on the UK award across 57 people, and seven independent senior voices describing the same gap. The next step, already planned, is a single fee-band question to the same population.');
}

// A6 Verification status
{
  const s = pres.addSlide();
  const y0 = chrome(s, 'A6  Verification status', ACTA);
  const H = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.ink } } });
  const O = (t) => ({ text: t, options: { bold: true, color: C.orange } });
  const G = (t) => ({ text: t, options: { bold: true, color: C.brown } });
  const rows = [
    [H('Item'), H('Status'), H('Owner')],
    ['Whether NCSC certification extends to a franchised delivery', O('Open: ask Essex'), 'Partnerships team'],
    ['Essex\'s expected commercial terms and any fee floor', O('Open: ask Essex'), 'Partnerships team'],
    ['NIBM\'s regulatory standing to host a foreign master\'s award', O('Open: registrar, Coventry documentation'), 'Partnerships team'],
    ['Fee-band question to the student population', O('Planned: one day'), 'Academic delivery'],
    ['Delivery panel titles, CVs and consent to be named', G('Verified from public sources; consent in progress'), 'Academic delivery'],
    ['League tables refreshed to 2027 editions', O('Before the meeting'), 'Academic delivery'],
    ['Competitor fees for IIT, ESOFT, SLIIT by direct enquiry', O('Open'), 'Academic delivery'],
    ['Regulatory recognition of the award in Sri Lanka', O('To confirm'), 'Partnerships team'],
  ];
  s.addTable(rows, { x: 0.45, y: y0, w: 9.1, colW: [4.7, 2.6, 1.8], fontFace: F, fontSize: 11, color: C.ink, border: { type: 'solid', color: C.border, pt: 0.5 }, rowH: [0.32, 0.38, 0.38, 0.38, 0.38, 0.38, 0.38, 0.38, 0.38], valign: 'middle', margin: 0.05, fill: { color: C.white } });
  notes(s, 'Showing an open verification list is deliberate: it tells a due-diligence audience that the numbers they are seeing were checked, and that the ones still open are named rather than glossed. Every item here is also in the evidence repository with its source.');
}

}

const OUT = MODE === 'appendix' ? '/mnt/user-data/outputs/NIBM_Essex_MSc_Cyber_Business_Case_v2_Appendix.pptx' : '/mnt/user-data/outputs/NIBM_Essex_MSc_Cyber_Business_Case_v2.pptx';
pres.writeFile({ fileName: OUT }).then(f => console.log('written', f));
