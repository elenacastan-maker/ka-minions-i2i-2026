/**
 * Ka-Minions Business Case — Single Sheet
 * Run: buildKaMinions()
 */

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
};

var PHASES = [
  { id:1, label:'Phase 1 — Hypothesis Validation', period:'Q4 2026',          teams:2,  savingYr:103718,  runCostQtr:3500,  setupCost:60000, color:'#1D4B00', kpi:'Bug resolution 3d → <3h' },
  { id:2, label:'Phase 2 — MVPs Creation',         period:'Q1–Q2 2027',       teams:7,  savingYr:363013,  runCostQtr:9000,  setupCost:15000, color:'#7c3aed', kpi:'70% success rate · ✓ Break-even Q2 2027' },
  { id:3, label:'Phase 3 — Growth & Iterations',   period:'Q3 2027–Q2 2028',  teams:22, savingYr:1140898, runCostQtr:25000, setupCost:25000, color:'#d97706', kpi:'40% teams adoption' },
  { id:4, label:'Phase 4 KPI — Scale Up',          period:'Q3 2028 → Always On', teams:30, savingYr:1557000, runCostQtr:20000, setupCost:0,     color:'#0891b2', kpi:'70% adoption (30/43 teams) · 30% feature to prod' },
];

var QLABELS = ["Q4 '26","Q1 '27","Q2 '27","Q3 '27","Q4 '27","Q1 '28","Q2 '28","Q3 '28","Q4 '28","Q1 '29"];
var QCAL    = ['Q4 2026','Q1 2027','Q2 2027','Q3 2027','Q4 2027','Q1 2028','Q2 2028','Q3 2028','Q4 2028','Q1 2029'];
var CUM_SAV = [25.9, 51.9, 142.6, 233.4, 518.6, 803.8, 1089.0, 1374.3, 1764.3, 2154.3];
var CUM_INV = [87.0, 94.0, 119.0, 129.0, 178.0, 197.0, 216.0, 235.0, 323.0, 361.0];
var Q_SAV_I = [25.9, 26.0, 90.7, 90.8, 285.2, 285.2, 285.2, 285.3, 390.0, 390.0];
var Q_INV_I = [87.0,  7.0, 25.0, 10.0,  49.0,  19.0,  19.0,  19.0,  88.0,  38.0];

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
  sh.setRowHeight(row, (fontSize || 11) > 13 ? 40 : 26);
  return row + 1;
}

function colWidths(sh, widths) {
  widths.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
}

// ─────────────────────────────────────────────────────────────────────────────

function buildKaMinions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shName = 'Ka-Minions Business Case';
  var sh = ss.getSheetByName(shName);
  if (sh) { sh.clear(); sh.clearFormats(); }
  else     { sh = ss.insertSheet(shName, 0); }

  colWidths(sh, [220, 140, 90, 120, 130, 130, 130, 200]);

  var R = 1; // current row cursor

  // ══════════════════════════════════════════════════════
  // BLOCK 1 — TITLE
  // ══════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'Ka-Minions: Autonomous AI Dev Agents — Business Case', C.darkGreen, C.lime, 16);
  R = hdr(sh, R, 1, 8, 'IDEAS2IMPACT 2026 · Banana Squad · Adevinta / Kleinanzeigen', C.midGreen, C.white, 9);
  sh.getRange(R, 1, 1, 8).setBackground(C.offWhite); R++;

  // ══════════════════════════════════════════════════════
  // BLOCK 2 — KPI TILES (one row per KPI, 3 columns)
  // ══════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'Key Metrics at a Glance', C.darkGreen, C.lime, 12);

  var kpiHdrRange = sh.getRange(R, 1, 1, 3);
  kpiHdrRange.setValues([['KPI / Metric', 'Value', 'Context']]);
  kpiHdrRange.setBackground(C.midGreen).setFontColor(C.white).setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
  R++;

  var kpis = [
    ['Teams at KPI target',           '30 / 43 (70%)', 'Full potential at 43 teams: €2.23M/yr'],
    ['Annual savings at KPI (30t)',    '€1,557,000',    'From Q3 2028 · Phase 4 · Always On'],
    ['FTE capacity freed',            '19.5 engineers','Same headcount · more product delivery'],
    ['Break-even',                    'Q2 2027',       'Cumulative savings overtake total investment'],
    ['ROI at Q10',                    '~6×',           '€2.15M saved vs €361K invested'],
    ['Net value 2028+',               '>€2.1M/yr',     '30t KPI · 19.5 FTE freed'],
    ['Total investment Q1–Q10',       '€361,000',      'Setup + AI stack + enabler team (10 quarters)'],
    ['AI stack cost vs savings',      '~3%',           'Paperclip + Claude API · scales linearly'],
    ['Phase 1 single ask',            '€60–80K',       'Q4 2026 setup only · gate-locked before next phase'],
    ['Avg engineer saving/yr',        '€80,000',       '25% capacity × €80K avg fully-loaded cost'],
  ];
  sh.getRange(R, 1, kpis.length, 3).setValues(kpis).setFontFamily('Arial').setFontSize(10);
  kpis.forEach(function(_, i) {
    sh.getRange(R + i, 1, 1, 3).setBackground(i % 2 === 0 ? C.lightGreen : C.white);
  });
  sh.getRange(R, 2).setFontWeight('bold').setFontColor(C.darkGreen).setFontSize(13);
  sh.getRange(R + 1, 2).setFontWeight('bold').setFontColor(C.darkGreen).setFontSize(13);
  sh.getRange(R, 1, kpis.length, 3)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += kpis.length + 1;

  // ══════════════════════════════════════════════════════
  // BLOCK 3 — PHASE SUMMARY TABLE
  // ══════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'Phase Summary — Investment & Savings', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 6).setValues([['Phase', 'Period', 'Teams', 'Annual Savings', 'Running Cost/Q', 'KPI Gate']]);
  sh.getRange(R, 1, 1, 6).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
  R++;

  PHASES.forEach(function(ph) {
    sh.getRange(R, 1, 1, 6).setValues([[
      ph.label, ph.period, ph.teams,
      '€' + ph.savingYr.toLocaleString(),
      '€' + ph.runCostQtr.toLocaleString() + '/Q',
      ph.kpi
    ]]);
    sh.getRange(R, 1, 1, 6)
      .setBackground(ph.color).setFontColor(C.white)
      .setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
    R++;
  });

  // Break-even callout
  sh.getRange(R, 1, 1, 6).merge()
    .setValue('✓ Break-even Q2 2027 — cumulative savings overtake total investment mid-Phase 2')
    .setBackground(C.yellow).setFontColor('#7a5c00').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  sh.getRange(R, 1, 1, 6)
    .setBorder(true,true,true,true,false,false,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += 2;

  // ══════════════════════════════════════════════════════
  // BLOCK 4 — 10-QUARTER MODEL TABLE
  // ══════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'Cumulative Investment vs Savings — 10 Quarters (Q1 = Q4 2026, values in €K)', C.darkGreen, C.lime, 12);

  var q10hdr = [['Q#','Calendar','Active Teams','Qtr Savings €K','Qtr Investment €K','Cum Savings €K','Cum Investment €K','Milestone']];
  sh.getRange(R, 1, 1, 8).setValues(q10hdr);
  sh.getRange(R, 1, 1, 8).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
  R++;

  var phColors = { 1:'#e8f5e1', 2:'#f3eeff', 3:'#fff8e1', 4:'#e0f7fa' };
  var Q10_START = R;

  QLABELS.forEach(function(ql, i) {
    var q = i + 1;
    var teams = 0, phId = 0, milestone = '';
    PHASES.forEach(function(ph) {
      var qs = ph.id === 1 ? 1 : ph.id === 2 ? 2 : ph.id === 3 ? 4 : 8;
      var qe = ph.id === 1 ? 1 : ph.id === 2 ? 3 : ph.id === 3 ? 7 : 10;
      if (q >= qs && q <= qe) { teams = ph.teams; phId = ph.id; milestone = ph.label.split('—')[0].trim(); }
    });
    if (q === 3) milestone += ' · ✓ Break-even';
    if (q === 8) milestone += ' · KPI target';

    sh.getRange(R, 1, 1, 8).setValues([['Q'+q, QCAL[i], teams, Q_SAV_I[i], Q_INV_I[i], CUM_SAV[i], CUM_INV[i], milestone]]);
    sh.getRange(R, 4, 1, 4).setNumberFormat('€#,##0.0"K"');
    var bg = q === 3 ? C.yellow : (phColors[phId] || C.white);
    sh.getRange(R, 1, 1, 8).setBackground(bg).setFontFamily('Arial').setFontSize(10);
    if (q === 3) sh.getRange(R, 1, 1, 8).setFontWeight('bold');
    R++;
  });

  // Totals
  sh.getRange(R, 1, 1, 8).setValues([['TOTAL','Q4 2026 → Q1 2029','','','', CUM_SAV[9], CUM_INV[9],
    'ROI: ' + Math.round(CUM_SAV[9]/CUM_INV[9]*10)/10 + '×']]);
  sh.getRange(R, 6, 1, 2).setNumberFormat('€#,##0.0"K"');
  sh.getRange(R, 1, 1, 8).setBackground(C.darkGreen).setFontColor(C.lime)
    .setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
  sh.getRange(Q10_START - 1, 1, R - Q10_START + 2, 8)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += 2;

  // ── LINE CHART: Investment vs Savings ──────────────────────────────────────
  // Chart data range: calendar col (B) + cum savings (F) + cum investment (G)
  var chartDataRow = Q10_START - 1; // header row
  var lineChart = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange(chartDataRow, 2, 11, 1))  // Calendar labels
    .addRange(sh.getRange(chartDataRow, 6, 11, 1))  // Cum Savings
    .addRange(sh.getRange(chartDataRow, 7, 11, 1))  // Cum Investment
    .setOption('title', 'Cumulative Savings vs Investment (€K) — Break-even Q2 2027')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: C.darkGreen })
    .setOption('hAxis', { title: 'Quarter', textStyle: { color: C.gray } })
    .setOption('vAxis', { title: '€K', textStyle: { color: C.gray } })
    .setOption('series', {
      0: { color: C.darkGreen, lineWidth: 3, pointSize: 6 },
      1: { color: C.red,       lineWidth: 2, pointSize: 4 },
    })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', { fill: '#f9fdf5' })
    .setOption('width', 680)
    .setOption('height', 360)
    .setPosition(R, 1, 0, 0)
    .build();
  sh.insertChart(lineChart);
  R += 22; // leave space for chart

  // ══════════════════════════════════════════════════════
  // BLOCK 5 — SAVINGS RAMP TABLE + CHART
  // ══════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'Annual Savings Ramp by Phase (€K/yr at each snapshot)', C.darkGreen, C.lime, 12);

  var RAMP_START = R;
  sh.getRange(R, 1, 1, 5).setValues([['Milestone snapshot','Phase 1 (2t)','Phase 2 (7t)','Phase 3 (22t)','Phase 4 KPI (30t)']]);
  sh.getRange(R, 1, 1, 5).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
  R++;

  var rampData = [
    ["Q4 '26 — Ph1 live",  103.7,    0,      0,       0     ],
    ["Q1 '27 — Ph2 live",  103.7,  363.0,    0,       0     ],
    ["Q3 '27 — Ph3 live",  103.7,  363.0, 1140.9,     0     ],
    ["Q3 '28 — Ph4 KPI",   103.7,  363.0, 1140.9,  1557.0   ],
    ["2028+ full scale",   103.7,  363.0, 1140.9,  2100.0   ],
  ];
  var rampBgs = [C.lightGreen,'#f3eeff','#fff8e1','#e0f7fa','#f0f9ff'];
  sh.getRange(R, 1, rampData.length, 5).setValues(rampData);
  sh.getRange(R, 2, rampData.length, 4).setNumberFormat('€#,##0.0"K"');
  rampData.forEach(function(_, i) {
    sh.getRange(R + i, 1, 1, 5).setBackground(rampBgs[i]).setFontFamily('Arial').setFontSize(10);
  });
  var RAMP_END_ROW = R + rampData.length - 1;
  sh.getRange(RAMP_START, 1, rampData.length + 1, 5)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
  R += rampData.length + 1;

  // ── STACKED BAR CHART ──────────────────────────────────────────────────────
  var barChart = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange(RAMP_START, 1, rampData.length + 1, 5))
    .setOption('title', 'Annual Savings Ramp — by Phase (€K/yr, stacked)')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: C.darkGreen })
    .setOption('isStacked', true)
    .setOption('hAxis', { title: '€K / year', textStyle: { color: C.gray } })
    .setOption('vAxis', { title: 'Phase snapshot', textStyle: { color: C.gray } })
    .setOption('series', {
      0: { color: C.darkGreen },
      1: { color: C.purple },
      2: { color: C.amber },
      3: { color: C.teal },
    })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', { fill: '#f9fdf5' })
    .setOption('width', 680)
    .setOption('height', 340)
    .setPosition(R, 1, 0, 0)
    .build();
  sh.insertChart(barChart);
  R += 22;

  // ══════════════════════════════════════════════════════
  // BLOCK 6 — CALCULATION ASSUMPTIONS
  // ══════════════════════════════════════════════════════
  R = hdr(sh, R, 1, 8, 'Calculation Assumptions', C.darkGreen, C.lime, 12);

  sh.getRange(R, 1, 1, 3).setValues([['Assumption', 'Value', 'Detail']]);
  sh.getRange(R, 1, 1, 3).setBackground(C.midGreen).setFontColor(C.white)
    .setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
  R++;

  var assumptions = [
    ['Avg engineer cost (fully loaded)', '€80,000/yr',    '€6,667/month · includes salary + benefits'],
    ['Capacity freed per engineer',      '25%',           '0.25 FTE equivalent = €20,000/yr saved per eng'],
    ['Phase 1 (2 teams × 6 eng)',        '€103,718/yr',   '= 2 × 6 × €20K × 0.93 ramp adjustment'],
    ['Phase 2 (7 teams × 6 eng)',        '€363,013/yr',   '= 7 × 6 × €20K × 0.865 adj'],
    ['Phase 3 (22 teams × 6 eng)',       '€1,140,898/yr', '= 22 × 6 × €20K × 0.864 adj'],
    ['Phase 4 KPI (30 teams × 6 eng)',   '€1,557,000/yr', '= 30 × 6 × €20K × 0.865 adj'],
    ['LLM cost (Claude API)',            '~€1/task',      '1,750 tasks/team/yr · scales linearly with teams'],
    ['Paperclip platform',               '~€12,000/yr',   'Team plan · includes audit trail + governance'],
    ['AI stack vs savings at scale',     '~3%',           '€52K AI cost vs €1.56M savings at Phase 4 KPI'],
    ['Break-even quarter',               'Q2 2027',       'Cum. savings €142.6K > cum. investment €119K'],
    ['ROI at Q10 (Q1 2029)',             '~6×',           '€2,154K saved ÷ €361K invested = 5.97×'],
    ['Full potential (43 teams)',        '€2,229,937/yr', 'If all KA teams adopt · not in KPI scope'],
  ];

  sh.getRange(R, 1, assumptions.length, 3).setValues(assumptions).setFontFamily('Arial').setFontSize(10);
  assumptions.forEach(function(_, i) {
    sh.getRange(R + i, 1, 1, 3).setBackground(i % 2 === 0 ? C.lightGreen : C.white);
    sh.getRange(R + i, 2).setFontWeight('bold').setFontColor(C.darkGreen);
  });
  sh.getRange(R - 1, 1, assumptions.length + 1, 3)
    .setBorder(true,true,true,true,true,true,'#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);

  ss.setActiveSheet(sh);
  SpreadsheetApp.flush();
  Logger.log('✅ Done — Ka-Minions Business Case built in a single sheet.');
}
