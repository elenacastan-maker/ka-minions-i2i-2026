/**
 * Ka-Minions Business Case — Full Single Sheet
 * Covers all 6 slides: problem metrics, domain expansion, phase roadmap,
 * 10-quarter financial model, AI stack, enabler team, assumptions.
 *
 * Run: buildKaMinions()
 */

// ── Palette ──────────────────────────────────────────────────────────────────
var C = {
  darkGreen : '#1D4B00',
  midGreen  : '#609F28',
  lime      : '#B5E941',
  purple    : '#7c3aed',
  amber     : '#d97706',
  teal      : '#0891b2',
  red       : '#c0392b',
  white     : '#FFFFFF',
  lightGreen: '#f0fae6',
  offWhite  : '#F4F2EF',
  gray      : '#504E48',
  yellow    : '#fff9c4',
  purLight  : '#f3eeff',
  ambLight  : '#fff8e1',
  tealLight : '#e0f7fa',
  redLight  : '#fff5f5',
};

// ── Financial data ────────────────────────────────────────────────────────────
var PHASES = [
  { id:1, label:'Phase 1 — Hypothesis Validation', period:'Q4 2026',             teams:2,  savingYr:103718,  runCostQtr:3500,  setupCost:60000, aiStack:15500,  color:'#1D4B00', kpi:'Bug resolution 5d → <3h · 40× faster · n=206 issues' },
  { id:2, label:'Phase 2 — MVPs Creation',         period:'Q1–Q2 2027',          teams:7,  savingYr:463813,  runCostQtr:9000,  setupCost:15000, aiStack:21000,  color:'#7c3aed', kpi:'70% success rate · on-call −30% · ✓ Break-even Q2 2027' },
  { id:3, label:'Phase 3 — Growth & Iterations',   period:'Q3 2027–Q2 2028',     teams:22, savingYr:1457698, runCostQtr:25000, setupCost:25000, aiStack:37000,  color:'#d97706', kpi:'40% teams adoption · on-call savings scaling' },
  { id:4, label:'Phase 4 KPI — Scale Up',          period:'Q3 2028 → Always On', teams:30, savingYr:1989000, runCostQtr:20000, setupCost:0,     aiStack:76000,  color:'#0891b2', kpi:'70% adoption (30/43 teams) · 30% feature to prod · on-call −30%' },
];

var QLABELS = ["Q4 '26","Q1 '27","Q2 '27","Q3 '27","Q4 '27","Q1 '28","Q2 '28","Q3 '28","Q4 '28","Q1 '29"];
var QCAL    = ['Q4 2026','Q1 2027','Q2 2027','Q3 2027','Q4 2027','Q1 2028','Q2 2028','Q3 2028','Q4 2028','Q1 2029'];
var CUM_SAV = [25.9, 51.8, 167.8, 283.8, 648.2, 1012.6, 1377.0, 1874.2, 2371.4, 2868.6];
var CUM_INV = [87.0, 94.0, 119.0, 129.0, 178.0,  197.0,  216.0,  235.0,  323.0,  361.0];
var Q_SAV_I = [25.9, 25.9, 116.0, 116.0, 364.4,  364.4,  364.4,  497.2,  497.2,  497.2];
var Q_INV_I = [87.0,  7.0,  25.0,  10.0,  49.0,   19.0,   19.0,   19.0,   88.0,   38.0];

// ── Helpers ───────────────────────────────────────────────────────────────────
function hdr(sh, row, col, nCols, text, bgHex, fgHex, fontSize) {
  var r = sh.getRange(row, col, 1, nCols);
  if (nCols > 1) r.merge();
  r.setValue(text)
   .setBackground(bgHex || C.darkGreen)
   .setFontColor(fgHex || C.white)
   .setFontWeight('bold')
   .setFontSize(fontSize || 11)
   .setFontFamily('Arial')
   .setHorizontalAlignment('center');
  sh.setRowHeight(row, fontSize > 13 ? 44 : 26);
  return row + 1;
}

function colWidths(sh, widths) {
  widths.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
}

function tableHdr(sh, row, cols, values, bg) {
  var r = sh.getRange(row, cols[0], 1, values.length);
  var vals = [values];
  r.setValues(vals)
   .setBackground(bg || C.midGreen).setFontColor(C.white)
   .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(row, 24);
  return row + 1;
}

function dataRow(sh, row, values, bg, bold, color) {
  var r = sh.getRange(row, 1, 1, values.length);
  r.setValues([values])
   .setBackground(bg || C.white)
   .setFontFamily('Arial').setFontSize(9);
  if (bold) r.setFontWeight('bold');
  if (color) r.setFontColor(color);
  sh.setRowHeight(row, 22);
  return row + 1;
}

function sectionGap(sh, row) {
  sh.setRowHeight(row, 10);
  return row + 1;
}

// ─────────────────────────────────────────────────────────────────────────────
function buildKaMinions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shName = 'Ka-Minions Business Case';
  var sh = ss.getSheetByName(shName);
  if (sh) { sh.clear(); sh.clearFormats(); sh.clearNotes(); }
  else     { sh = ss.insertSheet(shName, 0); }

  colWidths(sh, [220, 130, 90, 115, 120, 120, 115, 190]);

  var R = 1;

  // ══════════════════════════════════════════════════════════════════════════
  // TITLE
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'Ka-Minions: Autonomous AI Dev Agents — Business Case · Ideas2Impact 2026', C.darkGreen, C.lime, 16);
  R = hdr(sh, R, 1, 8, 'Banana Squad · Adevinta / Kleinanzeigen · KPI: 70% of KA teams with active agents in production', C.midGreen, C.white, 9);
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 1 — KEY METRICS AT A GLANCE
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'KEY METRICS AT A GLANCE', C.darkGreen, C.lime, 12);

  var kpiHdr = sh.getRange(R, 1, 1, 3);
  kpiHdr.setValues([['KPI / Metric', 'Value', 'Context']]);
  kpiHdr.setBackground(C.midGreen).setFontColor(C.white).setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var kpis = [
    ['Teams at KPI target (30/43)',    '70% adoption',   'Phase 4 · Q3 2028 · 9 domains covered'],
    ['Annual savings at KPI',          '€1,989,000/yr',  'Incl. on-call savings (30% reduction · €8K/eng/yr)'],
    ['Capacity savings only (30t)',    '€1,557,000/yr',  'Engineering capacity freed: 19.5 FTE equiv'],
    ['On-call savings at KPI (30t)',   '€432,000/yr',    '30% reduction × 30t × 6 eng × €8K/yr · Phase 2 gate'],
    ['FTE equivalent freed',           '~25 engineers',  'Capacity + on-call premium · same headcount'],
    ['Break-even',                     'Q2 2027',        'Cum. savings (€167.8K) overtake cum. investment (€119K)'],
    ['ROI at Q10 (Q1 2029)',           '~7.9×',          '€2,868.6K saved ÷ €361K invested'],
    ['Net annual gain at KPI',         '>€1.91M/yr',     '€1.99M savings − €76K AI stack − enabler team'],
    ['Total investment (10 quarters)', '€361,000',       'Q4 2026 → Q1 2029 · setup + AI stack + maintenance'],
    ['AI stack cost vs savings',       '~4%',            'Paperclip + Claude API · holds at 10× token price'],
    ['Phase 1 single ask',             '€60–80K',        'Q4 2026 setup · gate-locked · no further spend without KPI'],
    ['Full potential (43 teams)',      '€2,849,137/yr',  'Capacity €2.23M + on-call €619K · 100% adoption'],
    ['Pilot baseline',                 '5 days → <3h',   '40× faster bug resolution · n=206 issues · analysis of 2 KA teams'],
  ];
  sh.getRange(R, 1, kpis.length, 3).setValues(kpis).setFontFamily('Arial').setFontSize(9);
  kpis.forEach(function(_, i) {
    sh.getRange(R + i, 1, 1, 3).setBackground(i % 2 === 0 ? C.lightGreen : C.white);
    sh.getRange(R + i, 2).setFontWeight('bold').setFontColor(C.darkGreen).setFontSize(10);
  });
  sh.getRange(R, 1, kpis.length, 3)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += kpis.length + 1;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 2 — PROBLEM METRICS (Slide 1)
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'THE PROBLEM TODAY — Slide 1 · Measured baseline (analysis of 2 KA teams · n=206 issues)', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 5).setValues([['Metric', 'Today (baseline)', 'Ka-Minion target', 'Improvement', 'Context']]);
  sh.getRange(R, 1, 1, 5).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var problems = [
    ['Avg bug resolution time',      '5 days',   '<3 hours',  '40× faster', 'n=206 issues · analysis of 2 KA teams · Phase 1 gate KPI'],
    ['Zombie tickets (>30 days)',     '22%',      '<8%',       '−14pp',      'Stale tickets inflating estimates every sprint · Phase 2 KPI'],
    ['Out-of-hours duty calls/wk',   '~3/team',  '−30%',      '30% fewer',  'Sentinels watch 24/7 · triage & silence false alerts · Phase 2 KPI'],
    ['Sprint capacity on ops/triage','30–40%',   'Freed',     '+20% output','Reclaimed for feature & product delivery'],
    ['On-call premium cost',         '€8K/eng/yr','−30%',     '€2,400/eng/yr','30% reduction in hands-on out-of-hours incidents'],
  ];
  sh.getRange(R, 1, problems.length, 5).setValues(problems).setFontFamily('Arial').setFontSize(9);
  problems.forEach(function(_, i) {
    sh.getRange(R + i, 1, 1, 5).setBackground(i % 2 === 0 ? C.redLight : C.white);
  });
  sh.getRange(R, 3, problems.length, 1).setFontColor(C.darkGreen).setFontWeight('bold');
  sh.getRange(R, 1, problems.length, 5)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += problems.length + 1;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 3 — DOMAIN EXPANSION TABLE (Slide 2)
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'DOMAIN EXPANSION — Full-org rollout by domain · Slide 2', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 6).setValues([['Domain', 'Phase', 'Teams', 'Hours / yr', 'Capacity savings / yr', 'Notes']]);
  sh.getRange(R, 1, 1, 6).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var domains = [
    ['TRX',                '1 — Pilot',    2,  '2,282',  '€103,718',   'Phase 1 · Q4 2026 · pilot'],
    ['Consumer Flywheel',  '2 — MVPs',     4,  '4,564',  '€207,436',   'Phase 2 · Q1 2027'],
    ['D&A',                '2 — MVPs',     1,  '1,141',  '€51,859',    'Phase 2 · Q1 2027'],
    ['Consumer Growth',    '3 — Growth',   5,  '5,704',  '€259,295',   'Phase 3 · Q3 2027'],
    ['Advertising',        '3 — Growth',   6,  '6,845',  '€311,154',   'Phase 3 · Q3 2027'],
    ['PRO',                '3 — Growth',   4,  '4,564',  '€207,436',   'Phase 3 · Q3 2027'],
    ['RE & Motors',        '4 — Scale',    6,  '6,845',  '€311,154',   'Phase 4 · Q3 2028'],
    ['Platform',           '4 — Scale',    9,  '10,268', '€466,731',   'Phase 4 · largest domain'],
    ['Data',               '4 — Scale',    6,  '6,845',  '€311,154',   'Phase 4 · Q3 2028'],
    ['TOTAL — all 43 teams', 'All phases', 43, '49,059', '€2,229,937', 'Full potential (capacity only)'],
    ['KPI target — 30 teams (70%)', 'Ph4 gate', 30, '34,230', '€1,557,000', 'Capacity only · +€432K on-call = €1,989,000 total'],
  ];
  var domPhBg = {'1 — Pilot':C.lightGreen,'2 — MVPs':C.purLight,'3 — Growth':C.ambLight,'4 — Scale':C.tealLight,'All phases':'#f6faf0','Ph4 gate':C.lightGreen};
  domains.forEach(function(row, i) {
    var ry = R + i;
    var isTotal = i === 9; var isKPI = i === 10;
    var bg = isKPI ? C.lightGreen : isTotal ? '#f6faf0' : (domPhBg[row[1]] || C.white);
    sh.getRange(ry, 1, 1, 6).setValues([row]).setBackground(bg).setFontFamily('Arial').setFontSize(9);
    if (isKPI || isTotal) sh.getRange(ry, 1, 1, 6).setFontWeight('bold');
    if (isKPI) sh.getRange(ry, 5).setFontColor(C.darkGreen);
    sh.setRowHeight(ry, 22);
  });
  sh.getRange(R, 1, domains.length, 6)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += domains.length + 1;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 4 — PHASE ROADMAP SUMMARY (Slide 2 / 6)
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'PHASE ROADMAP — 4-phase rollout · Slides 2 & 6', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 7).setValues([['Phase', 'Period', 'Teams', 'Annual savings (incl. on-call)', 'AI stack/yr', 'Net/yr', 'Gate KPI']]);
  sh.getRange(R, 1, 1, 7).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  PHASES.forEach(function(ph) {
    var net = ph.savingYr - ph.aiStack;
    sh.getRange(R, 1, 1, 7).setValues([[
      ph.label, ph.period, ph.teams,
      '€' + ph.savingYr.toLocaleString(),
      '~€' + (ph.aiStack/1000).toFixed(1) + 'K',
      '~€' + Math.round(net/1000) + 'K',
      ph.kpi
    ]]);
    sh.getRange(R, 1, 1, 7)
      .setBackground(ph.color).setFontColor(C.white)
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
    sh.setRowHeight(R, 24); R++;
  });

  sh.getRange(R, 1, 1, 7).merge()
    .setValue('✓ Break-even Q2 2027 — cumulative savings overtake total investment mid-Phase 2 · On-call savings activate at Phase 2 Sentinel gate')
    .setBackground(C.yellow).setFontColor('#7a5c00').setFontWeight('bold')
    .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
  sh.setRowHeight(R, 24); R += 2;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 5 — 10-QUARTER FINANCIAL MODEL (Slide 5)
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'CUMULATIVE INVESTMENT vs SAVINGS — 10 Quarters · Q4 2026 → Q1 2029 · values in €K · Slide 5', C.darkGreen, C.lime, 12);

  var q10hdr = [['Q#','Calendar','Active Teams','Qtr Savings €K','Qtr Investment €K','Cum Savings €K','Cum Investment €K','Milestone']];
  sh.getRange(R, 1, 1, 8).setValues(q10hdr);
  sh.getRange(R, 1, 1, 8).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var phColors = { 1:'#e8f5e1', 2:'#f3eeff', 3:'#fff8e1', 4:'#e0f7fa' };
  var Q10_START = R;

  QLABELS.forEach(function(ql, i) {
    var q = i + 1;
    var teams = 0, phId = 0, milestone = '';
    PHASES.forEach(function(ph) {
      var qs = ph.id===1?1 : ph.id===2?2 : ph.id===3?4 : 8;
      var qe = ph.id===1?1 : ph.id===2?3 : ph.id===3?7 : 10;
      if (q >= qs && q <= qe) { teams = ph.teams; phId = ph.id; milestone = ph.label.split('—')[0].trim(); }
    });
    if (q === 3) milestone += ' · ✓ Break-even Q2 2027';
    if (q === 8) milestone += ' · KPI target reached';

    sh.getRange(R, 1, 1, 8).setValues([['Q'+q, QCAL[i], teams, Q_SAV_I[i], Q_INV_I[i], CUM_SAV[i], CUM_INV[i], milestone]]);
    sh.getRange(R, 4, 1, 4).setNumberFormat('€#,##0.0"K"');
    var bg = q === 3 ? C.yellow : (phColors[phId] || C.white);
    sh.getRange(R, 1, 1, 8).setBackground(bg).setFontFamily('Arial').setFontSize(9);
    if (q === 3) sh.getRange(R, 1, 1, 8).setFontWeight('bold');
    sh.setRowHeight(R, 22); R++;
  });

  var roi = Math.round(CUM_SAV[9] / CUM_INV[9] * 10) / 10;
  sh.getRange(R, 1, 1, 8).setValues([['TOTAL', 'Q4 2026 → Q1 2029', '', '', '', CUM_SAV[9], CUM_INV[9], 'ROI: ' + roi + '×  ·  Break-even Q2 2027']]);
  sh.getRange(R, 6, 1, 2).setNumberFormat('€#,##0.0"K"');
  sh.getRange(R, 1, 1, 8).setBackground(C.darkGreen).setFontColor(C.lime)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 26);
  sh.getRange(Q10_START - 1, 1, R - Q10_START + 2, 8)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += 2;

  // ── Chart 1: Cumulative savings vs investment line chart ────────────────
  var chartDataRow = Q10_START - 1;
  var lineChart = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange(chartDataRow, 2, 12, 1))  // Calendar labels (incl header)
    .addRange(sh.getRange(chartDataRow, 6, 12, 1))  // Cum Savings
    .addRange(sh.getRange(chartDataRow, 7, 12, 1))  // Cum Investment
    .setOption('title', 'Cumulative Savings vs Investment (€K) — Break-even Q2 2027 · ROI ~7.9× at Q10')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: C.darkGreen })
    .setOption('hAxis', { title: 'Quarter (Q1 = Q4 2026)', textStyle: { color: C.gray } })
    .setOption('vAxis', { title: '€K', textStyle: { color: C.gray } })
    .setOption('series', {
      0: { color: C.darkGreen, lineWidth: 3, pointSize: 6, labelInLegend: 'Cumulative savings' },
      1: { color: C.red,       lineWidth: 2, pointSize: 4, lineDashStyle: [4,2], labelInLegend: 'Cumulative investment' },
    })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', { fill: '#f9fdf5' })
    .setOption('width', 700)
    .setOption('height', 380)
    .setPosition(R, 1, 0, 0)
    .build();
  sh.insertChart(lineChart);
  R += 23;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 6 — ANNUAL SAVINGS RAMP BY PHASE (Slide 5)
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'ANNUAL SAVINGS RAMP BY PHASE — €K/yr at each rollout snapshot · Slide 5', C.darkGreen, C.lime, 12);

  var RAMP_START = R;
  sh.getRange(R, 1, 1, 5).setValues([['Milestone snapshot', 'Phase 1 (2t)', 'Phase 2 (7t)', 'Phase 3 (22t)', 'Phase 4 KPI (30t)']]);
  sh.getRange(R, 1, 1, 5).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var rampData = [
    ["Q4 '26 — Ph1 live",  103.7,     0,       0,       0     ],
    ["Q1 '27 — Ph2 live",  103.7,  463.8,      0,       0     ],
    ["Q3 '27 — Ph3 live",  103.7,  463.8,  1457.7,      0     ],
    ["Q3 '28 — Ph4 KPI",   103.7,  463.8,  1457.7,  1989.0   ],
    ["2028+ full scale",   103.7,  463.8,  1457.7,  2849.1   ],
  ];
  var rampNotes = [
    'Phase 1 only · 2 teams · TRX domain',
    'On-call savings activate (Sentinel gate) · 7 teams',
    '22 teams · majority of capacity unlocked',
    'KPI target · 30/43 teams · incl. full on-call savings',
    'Full 43-team potential · capacity + on-call',
  ];
  var rampBgs = [C.lightGreen, C.purLight, C.ambLight, C.tealLight, '#f0f9ff'];
  sh.getRange(R, 1, rampData.length, 5).setValues(rampData);
  sh.getRange(R, 2, rampData.length, 4).setNumberFormat('€#,##0.0"K"');
  rampData.forEach(function(_, i) {
    sh.getRange(R + i, 1, 1, 5)
      .setBackground(rampBgs[i]).setFontFamily('Arial').setFontSize(9);
    sh.getRange(R + i, 1).setNote(rampNotes[i]);
    sh.setRowHeight(R + i, 22);
  });
  sh.getRange(RAMP_START, 1, rampData.length + 1, 5)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += rampData.length + 1;

  // ── Chart 2: Stacked bar — savings ramp by phase ────────────────────────
  var barChart = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange(RAMP_START, 1, rampData.length + 1, 5))
    .setOption('title', 'Annual Savings Ramp — Stacked by Phase (€K/yr)')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: C.darkGreen })
    .setOption('isStacked', true)
    .setOption('hAxis', { title: '€K / year', textStyle: { color: C.gray } })
    .setOption('vAxis', { title: 'Rollout snapshot', textStyle: { color: C.gray } })
    .setOption('series', {
      0: { color: C.darkGreen, labelInLegend: 'Phase 1 (2t)' },
      1: { color: C.purple,    labelInLegend: 'Phase 2 (7t)' },
      2: { color: C.amber,     labelInLegend: 'Phase 3 (22t)' },
      3: { color: C.teal,      labelInLegend: 'Phase 4 KPI (30t)' },
    })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', { fill: '#f9fdf5' })
    .setOption('width', 700)
    .setOption('height', 360)
    .setPosition(R, 1, 0, 0)
    .build();
  sh.insertChart(barChart);
  R += 23;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 7 — AI STACK COST MODEL (Slide 4 / 5)
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'AI STACK COST MODEL — Paperclip + Claude API · Slides 4 & 5', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 7).setValues([['Phase', 'Teams', 'Paperclip/yr', 'Claude API/yr', 'Total AI stack/yr', 'Savings/yr', 'Stack as % of savings']]);
  sh.getRange(R, 1, 1, 7).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var aiStack = [
    ['Phase 1 · Q4 2026 · 2 teams',  2,  '~€12,000', '~€3,500',  '~€15,500', '€103,718',  '15%*'],
    ['Phase 2 · Q1–Q2 2027 · 7t',    7,  '~€15,000', '~€9,000',  '~€21,000', '€463,813',  '4.5%'],
    ['Phase 3 · Q3 2027 · 22t',      22, '~€20,000', '~€25,000', '~€37,000', '€1,457,698','2.5%'],
    ['Phase 4 KPI · Q3 2028 · 30t',  30, '~€24,000', '~€52,000', '~€76,000', '€1,989,000','3.8%'],
  ];
  var aiBgs = [C.lightGreen, C.purLight, C.ambLight, C.tealLight];
  aiStack.forEach(function(row, i) {
    sh.getRange(R + i, 1, 1, 7).setValues([row])
      .setBackground(aiBgs[i]).setFontFamily('Arial').setFontSize(9);
    sh.setRowHeight(R + i, 22);
  });
  sh.getRange(R, 1, aiStack.length, 7)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += aiStack.length;

  sh.getRange(R, 1, 1, 7).merge()
    .setValue('* Ph1 % includes €60–80K one-time setup amortised over year 1 · ongoing run rate drops to 3.4% · ratio holds at 10× token price inflation')
    .setBackground('#fff8e8').setFontColor('#7a5c00').setFontSize(8).setFontFamily('Arial')
    .setHorizontalAlignment('left');
  sh.setRowHeight(R, 20); R += 2;
  R = sectionGap(sh, R);

  // ── Chart 3: AI stack cost vs savings column chart ──────────────────────
  // Build a small data range on the side for the chart
  var aiChartDataRow = R;
  sh.getRange(R, 1, 1, 3).setValues([['Phase', 'AI Stack Cost €K', 'Savings €K']]);
  sh.getRange(R, 1, 1, 3).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;
  var aiChartData = [
    ['Phase 1', 15.5, 103.7],
    ['Phase 2', 21.0, 463.8],
    ['Phase 3', 37.0, 1457.7],
    ['Phase 4 KPI', 76.0, 1989.0],
  ];
  sh.getRange(R, 1, aiChartData.length, 3).setValues(aiChartData)
    .setBackground(C.lightGreen).setFontFamily('Arial').setFontSize(9);
  sh.getRange(R, 1, aiChartData.length, 3)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += aiChartData.length + 1;

  var aiColChart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(aiChartDataRow, 1, aiChartData.length + 1, 3))
    .setOption('title', 'AI Stack Cost vs Savings by Phase (€K)')
    .setOption('titleTextStyle', { fontSize: 12, bold: true, color: C.darkGreen })
    .setOption('isStacked', false)
    .setOption('hAxis', { textStyle: { color: C.gray } })
    .setOption('vAxis', { title: '€K', textStyle: { color: C.gray } })
    .setOption('series', {
      0: { color: C.red,       labelInLegend: 'AI Stack Cost' },
      1: { color: C.darkGreen, labelInLegend: 'Annual Savings' },
    })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', { fill: '#f9fdf5' })
    .setOption('width', 520)
    .setOption('height', 320)
    .setPosition(R, 1, 0, 0)
    .build();
  sh.insertChart(aiColChart);
  R += 21;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 8 — ENABLER TEAM COSTS (Slide 4)
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'ENABLER TEAM — Headcount & Cost · Slide 4 · Existing reallocation — no net new hires in Phase 1', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 6).setValues([['Phase', 'Team composition', 'Eng cost/yr', 'EM cost/yr', 'Domain champions (10%)', 'Total people cost/yr']]);
  sh.getRange(R, 1, 1, 6).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var enabler = [
    ['Phase 1 · Q4 2026',   '0.5 EM + 1 eng (reallocation)', '~€40,000', '~€20,000', '2t × €8K = €16K',  '~€76,000'],
    ['Phase 2 · Q1–Q2 2027','1 EM + 2 eng',                  '~€80,000', '~€40,000', '7t × €8K = €56K',  '~€176,000'],
    ['Phase 3 · Q3 2027',   '1 EM + 2 eng',                  '~€80,000', '~€40,000', '22t × €8K = €176K','~€296,000'],
    ['Phase 4 KPI (30t)',   '1 EM + 2 eng',                  '~€80,000', '~€40,000', '30t × €8K = €240K','~€360,000'],
  ];
  var enaColors = ['#1D4B00','#7c3aed','#d97706','#0891b2'];
  enabler.forEach(function(row, i) {
    sh.getRange(R + i, 1, 1, 6).setValues([row])
      .setBackground(aiBgs[i]).setFontFamily('Arial').setFontSize(9);
    sh.getRange(R + i, 1).setFontColor(enaColors[i]).setFontWeight('bold');
    sh.setRowHeight(R + i, 22);
  });
  sh.getRange(R, 1, enabler.length, 6)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += enabler.length;

  sh.getRange(R, 1, 1, 6).merge()
    .setValue('People cost is the largest line item — but sourced entirely from existing headcount reallocation. Net incremental cash cost in Phase 1 = €0 on headcount.')
    .setBackground('#fff8e8').setFontColor('#7a5c00').setFontSize(8).setFontFamily('Arial');
  sh.setRowHeight(R, 20); R += 2;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 9 — GOVERNANCE & RISK (Slide 3)
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'GOVERNANCE & RISK MITIGATION · Slide 3', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 3).setValues([['Control / Risk', 'Status', 'Detail']]);
  sh.getRange(R, 1, 1, 3).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var gov = [
    ['Scope guardrail',          'ACTIVE',    'Ka-Minions only act on ≤1 Story Point issues. No merge without CI green + Paperclip audit gate.'],
    ['Full audit trail',         'ACTIVE',    'Every agent action logged in Paperclip: timestamp, token cost, outcome. Exportable for compliance.'],
    ['Kill switch',              'ACTIVE',    'Any Ka-Minion paused/terminated instantly from dashboard — per-agent, per-team, or org-wide.'],
    ['Phase-gated investment',   'ACTIVE',    'No spend unlocks without hitting the prior phase KPI. Zero risk of over-investing.'],
    ['LLM hallucination risk',   'MITIGATED', 'CI green + reviewer approval required before any merge. Ka-Minion proposes; human approves.'],
    ['API cost overrun risk',    'MITIGATED', 'Per-agent monthly budget caps in Paperclip. Teams need EM override to exceed allocated spend.'],
    ['Single point of failure',  'LOW',       'Paperclip provides fallback + circuit breaker. Outage = tickets remain unautomated, no damage.'],
  ];
  gov.forEach(function(row, i) {
    var bg = row[1] === 'ACTIVE' ? C.lightGreen : row[1] === 'MITIGATED' ? C.ambLight : C.white;
    var fc = row[1] === 'ACTIVE' ? C.darkGreen : row[1] === 'MITIGATED' ? '#92400e' : C.gray;
    sh.getRange(R + i, 1, 1, 3).setValues([row]).setBackground(bg).setFontFamily('Arial').setFontSize(9);
    sh.getRange(R + i, 2).setFontWeight('bold').setFontColor(fc);
    sh.setRowHeight(R + i, 22);
  });
  sh.getRange(R, 1, gov.length, 3)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += gov.length + 1;
  R = sectionGap(sh, R);

  // ══════════════════════════════════════════════════════════════════════════
  // BLOCK 10 — CALCULATION ASSUMPTIONS
  // ══════════════════════════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'CALCULATION ASSUMPTIONS', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 3).setValues([['Assumption', 'Value', 'Detail']]);
  sh.getRange(R, 1, 1, 3).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');
  sh.setRowHeight(R, 24); R++;

  var assumptions = [
    ['Avg engineer cost (fully loaded)',   '€80,000/yr',    '€6,667/month · includes salary + benefits'],
    ['Capacity freed per engineer',        '25%',           '0.25 FTE equivalent = €20,000/yr saved per eng'],
    ['Phase 1 · 2 teams × 6 eng',         '€103,718/yr',   '= 2 × 6 × €20K × 0.93 ramp adjustment · capacity only'],
    ['Phase 2 · 7 teams × 6 eng',         '€463,813/yr',   '= €363K capacity + €100.8K on-call (30% × 7t × 6eng × €8K)'],
    ['Phase 3 · 22 teams × 6 eng',        '€1,457,698/yr', '= €1.14M capacity + €316.8K on-call savings'],
    ['Phase 4 KPI · 30 teams × 6 eng',    '€1,989,000/yr', '= €1.557M capacity + €432K on-call savings'],
    ['On-call premium saving',             '€8K/eng/yr',    '30% reduction · activates at Phase 2 Sentinel mode gate'],
    ['On-call reduction trigger',          'Phase 2 KPI',   'Sentinel mode covers nights & weekends · −30% hands-on incidents'],
    ['LLM cost (Claude API)',              '~€1/task',      '1,750 tasks/team/yr · 50K input tokens + 2K output avg · scales linearly'],
    ['Paperclip platform',                 '~€12–24K/yr',   'Team plan · audit trail · governance · phase 1 → phase 4 KPI'],
    ['AI stack vs savings (Phase 4)',      '~3.8%',         '€76K AI cost vs €1.99M savings at Phase 4 KPI'],
    ['Break-even quarter',                 'Q2 2027',       'Cum. savings €167.8K > cum. investment €119K (Q3 of model)'],
    ['ROI at Q10 (Q1 2029)',              '~7.9×',          '€2,868.6K saved ÷ €361K invested = 7.95×'],
    ['Full potential · 43 teams',         '€2,849,137/yr',  'Capacity €2.23M + on-call €619K · if all KA teams adopt'],
    ['Pilot measurement basis',           'n=206 issues',   'Analysis of 2 KA teams · 6-month backlog · bug + task data'],
    ['Cost of waiting (per quarter)',      '~€497K lost',    'Recoverable capacity not captured without Ka-Minions at Phase 4 run-rate'],
  ];

  sh.getRange(R, 1, assumptions.length, 3).setValues(assumptions).setFontFamily('Arial').setFontSize(9);
  assumptions.forEach(function(_, i) {
    sh.getRange(R + i, 1, 1, 3).setBackground(i % 2 === 0 ? C.lightGreen : C.white);
    sh.getRange(R + i, 2).setFontWeight('bold').setFontColor(C.darkGreen);
    sh.setRowHeight(R + i, 22);
  });
  sh.getRange(R - 1, 1, assumptions.length + 1, 3)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += assumptions.length + 1;

  // ══════════════════════════════════════════════════════════════════════════
  // FINALISE
  // ══════════════════════════════════════════════════════════════════════════
  ss.setActiveSheet(sh);
  sh.setFrozenRows(0);
  SpreadsheetApp.flush();
  Logger.log('Ka-Minions full sheet built — ' + (R - 1) + ' rows · 3 charts');
}
