/**
 * Ka-Minions Business Case — Google Apps Script
 * Creates a full financial model spreadsheet with charts.
 *
 * HOW TO USE
 * 1. Open Google Sheets → Extensions → Apps Script
 * 2. Paste this entire file and click Run → buildKaMinionsBusinessCase
 * 3. Authorize when prompted
 * 4. A sheet called "Ka-Minions Business Case" will appear
 *
 * Tabs created:
 *   ① Summary           — KPI tiles + phase overview
 *   ② 10-Quarter Model  — cumulative savings & investment + chart
 *   ③ Phase Detail      — per-phase cost & savings breakdown
 *   ④ Savings Ramp      — stacked bar chart by phase
 */

// ── Colours ────────────────────────────────────────────────────────────────────
var C = {
  darkGreen  : '#1D4B00',
  midGreen   : '#609F28',
  lime       : '#B5E941',
  purple     : '#7c3aed',
  amber      : '#d97706',
  teal       : '#0891b2',
  red        : '#c0392b',
  white      : '#FFFFFF',
  offWhite   : '#F4F2EF',
  lightGreen : '#f0fae6',
  labelGray  : '#6b7f5e',
  bodyGray   : '#504E48',
  headerText : '#B5E941',   // lime on dark bg
};

// ── Financial constants ────────────────────────────────────────────────────────
var MODEL = {
  avgEngSalary   : 80000,   // €/yr per engineer
  savingRatePct  : 0.25,    // 25% of capacity freed = 0.25 FTE/eng
  llmCostPerTask : 1.0,     // € per LLM task (Paperclip + Claude)
  tasksPerTeamYr : 1750,    // tasks/team/year
  paperclipYr    : 12000,   // € Paperclip platform fee/yr
};

// ── Phase definitions ─────────────────────────────────────────────────────────
var PHASES = [
  {
    id      : 1,
    label   : 'Phase 1',
    sub     : 'Hypothesis Validation',
    period  : 'Q4 2026',
    q_start : 1,  // quarter index (1-based, Q4 2026 = Q1 of model)
    q_end   : 1,
    teams   : 2,
    engsPerTeam : 6,
    savingYr: 103718,
    setupCost: 60000,
    runCostQtr: 3500,
    color   : '#1D4B00',
    kpi     : 'Bug resolution 3 days → <3 hours',
  },
  {
    id      : 2,
    label   : 'Phase 2',
    sub     : 'MVPs Creation',
    period  : 'Q1–Q2 2027',
    q_start : 2,
    q_end   : 3,
    teams   : 7,
    engsPerTeam : 6,
    savingYr: 363013,
    setupCost: 15000,
    runCostQtr: 9000,
    color   : '#7c3aed',
    kpi     : '70% Ka-Minion success rate · Break-even Q2 2027',
    breakEven: true,
  },
  {
    id      : 3,
    label   : 'Phase 3',
    sub     : 'Growth & Iterations',
    period  : 'Q3 2027–Q2 2028',
    q_start : 4,
    q_end   : 7,
    teams   : 22,
    engsPerTeam : 6,
    savingYr: 1140898,
    setupCost: 25000,
    runCostQtr: 25000,
    color   : '#d97706',
    kpi     : '40% teams Ka-Minions adoption',
  },
  {
    id      : 4,
    label   : 'Phase 4 (KPI)',
    sub     : 'Scale Up · Always On',
    period  : 'Q3 2028 → Always On',
    q_start : 8,
    q_end   : 10,
    teams   : 30,
    engsPerTeam : 6,
    savingYr: 1557000,
    setupCost: 0,
    runCostQtr: 20000,
    color   : '#0891b2',
    kpi     : '70% adoption (30/43 teams) · 30% feature delivery to production',
  },
];

// Quarter labels Q1=Q4 2026 ... Q10=Q1 2029
var QUARTER_LABELS = [
  "Q4 '26","Q1 '27","Q2 '27","Q3 '27","Q4 '27",
  "Q1 '28","Q2 '28","Q3 '28","Q4 '28","Q1 '29"
];

// Pre-computed cumulative savings (€K) — 30t KPI model
var CUM_SAVINGS_K = [25.9, 51.9, 142.6, 233.4, 518.6, 803.8, 1089.0, 1374.3, 1764.3, 2154.3];
// Pre-computed cumulative investment (€K)
var CUM_INVEST_K  = [87.0, 94.0, 119.0, 129.0, 178.0, 197.0, 216.0, 235.0, 323.0, 361.0];

// ── Utility ────────────────────────────────────────────────────────────────────

function getOrCreateSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) sh.clear();
  else     sh = ss.insertSheet(name);
  return sh;
}

function styleHeader(range, bgHex, fgHex) {
  range.setBackground(bgHex || C.darkGreen)
       .setFontColor(fgHex || C.white)
       .setFontWeight('bold')
       .setFontSize(10)
       .setFontFamily('Arial');
}

function euros(n) {
  return '€' + Number(n).toLocaleString('en-IE', {minimumFractionDigits: 0, maximumFractionDigits: 0});
}

// ── MAIN ──────────────────────────────────────────────────────────────────────

function buildKaMinionsBusinessCase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('Ka-Minions Business Case');

  buildSummarySheet(ss);
  buildQuarterlyModel(ss);
  buildPhaseDetail(ss);
  buildSavingsRamp(ss);

  // Activate summary
  ss.setActiveSheet(ss.getSheetByName('① Summary'));
  SpreadsheetApp.flush();
  Logger.log('✅ Ka-Minions Business Case built successfully.');
}


// ══════════════════════════════════════════════════════════════════════════════
// ① SUMMARY SHEET
// ══════════════════════════════════════════════════════════════════════════════

function buildSummarySheet(ss) {
  var sh = getOrCreateSheet(ss, '① Summary');
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 320);

  // Title banner
  var titleRange = sh.getRange('A1:C1');
  titleRange.merge()
    .setValue('Ka-Minions: Autonomous AI Dev Agents — Business Case')
    .setBackground(C.darkGreen)
    .setFontColor(C.lime)
    .setFontWeight('bold')
    .setFontSize(16)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center');
  sh.setRowHeight(1, 42);

  // Subtitle
  sh.getRange('A2:C2').merge()
    .setValue('IDEAS2IMPACT 2026 · Banana Squad · Adevinta / Kleinanzeigen')
    .setBackground(C.midGreen)
    .setFontColor(C.white)
    .setFontSize(9)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center');

  // KPI tiles (row 4)
  sh.getRange('A3:C3').merge().setValue('').setBackground(C.offWhite);

  var kpiHeaders = [['KPI / Metric', 'Value', 'Context']];
  sh.getRange('A4:C4').setValues(kpiHeaders);
  styleHeader(sh.getRange('A4:C4'), C.midGreen, C.white);

  var kpis = [
    ['Teams at KPI target',          '30 / 43 (70%)',  'Full potential at 43 teams: €2.23M/yr'],
    ['Annual savings at KPI (30t)',   '€1,557,000',     'From Q3 2028 · Phase 4 · Always On'],
    ['FTE capacity freed',            '19.5 engineers', 'Same headcount · more product delivery'],
    ['Break-even',                    'Q2 2027',        'Cumulative savings overtake investment'],
    ['ROI at Q10',                    '~6×',            '€2.15M saved vs €361K invested'],
    ['Net value 2028+',               '>€2.1M/yr',      '30t KPI · 19.5 FTE freed'],
    ['Total investment Q1–Q10',       '€361,000',       'Setup + AI stack + enabler team'],
    ['AI stack cost vs savings',      '~3%',            'Paperclip + Claude API · scales linearly'],
    ['Phase 1 single ask',            '€60–80K',        'Q4 2026 setup only · no further unlock without gate KPI'],
    ['Avg engineer saving/yr',        '€80,000',        '25% capacity × €80K avg fully-loaded cost'],
  ];

  sh.getRange(5, 1, kpis.length, 3).setValues(kpis).setFontFamily('Arial').setFontSize(10);

  // Alternate row colours
  for (var r = 0; r < kpis.length; r++) {
    var bg = (r % 2 === 0) ? C.lightGreen : C.white;
    sh.getRange(5 + r, 1, 1, 3).setBackground(bg);
  }

  // Highlight KPI rows
  sh.getRange(5, 2).setFontWeight('bold').setFontColor(C.darkGreen).setFontSize(13);
  sh.getRange(6, 2).setFontWeight('bold').setFontColor(C.darkGreen).setFontSize(13);
  sh.getRange(7, 2).setFontWeight('bold').setFontColor(C.midGreen).setFontSize(12);

  // Phase summary table
  var startRow = 5 + kpis.length + 2;
  sh.getRange(startRow, 1, 1, 3).merge()
    .setValue('Phase Summary')
    .setBackground(C.darkGreen)
    .setFontColor(C.lime)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily('Arial');

  var phHdr = [['Phase', 'Period', 'Teams · Savings/yr · KPI']];
  sh.getRange(startRow + 1, 1, 1, 3).setValues(phHdr);
  styleHeader(sh.getRange(startRow + 1, 1, 1, 3), C.midGreen, C.white);

  PHASES.forEach(function(ph, i) {
    var row = startRow + 2 + i;
    sh.getRange(row, 1).setValue(ph.label + ' — ' + ph.sub);
    sh.getRange(row, 2).setValue(ph.period);
    sh.getRange(row, 3).setValue(ph.teams + ' teams · ' + euros(ph.savingYr) + '/yr · ' + ph.kpi);
    sh.getRange(row, 1, 1, 3)
      .setBackground(ph.color)
      .setFontColor(C.white)
      .setFontWeight('bold')
      .setFontSize(9.5)
      .setFontFamily('Arial');
  });

  // Break-even callout
  var beRow = startRow + 2 + PHASES.length + 1;
  sh.getRange(beRow, 1, 1, 3).merge()
    .setValue('✓ Break-even: Q2 2027 — cumulative savings overtake total investment mid-Phase 2')
    .setBackground('#fff9c4')
    .setFontColor('#7a5c00')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center');

  // Borders
  sh.getRange(4, 1, 5 + kpis.length + PHASES.length + 2, 3)
    .setBorder(true, true, true, true, true, true, '#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
}


// ══════════════════════════════════════════════════════════════════════════════
// ② 10-QUARTER MODEL
// ══════════════════════════════════════════════════════════════════════════════

function buildQuarterlyModel(ss) {
  var sh = getOrCreateSheet(ss, '② 10-Quarter Model');

  [1, 2, 3, 4, 5, 6, 7, 8].forEach(function(c, i) {
    sh.setColumnWidth(c, [200, 110, 110, 120, 120, 120, 120, 160][i] || 120);
  });

  // Title
  sh.getRange('A1:H1').merge()
    .setValue('Ka-Minions — Cumulative Investment vs Savings (10 Quarters, 30t KPI)')
    .setBackground(C.darkGreen).setFontColor(C.lime).setFontWeight('bold')
    .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('center');
  sh.setRowHeight(1, 36);

  // Sub-header row
  sh.getRange('A2:H2').merge()
    .setValue('Q1 = Q4 2026  ·  Break-even Q2 2027  ·  KPI target Q3 2028  ·  Values in €K')
    .setBackground(C.midGreen).setFontColor(C.white).setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('center');

  // Column headers
  var hdrs = [['Quarter','Calendar','Active Teams','Quarterly Savings €K',
               'Quarterly Investment €K','Cum. Savings €K','Cum. Investment €K','Phase / Milestone']];
  sh.getRange('A3:H3').setValues(hdrs);
  styleHeader(sh.getRange('A3:H3'), C.midGreen, C.white);

  // Build quarterly incremental data
  var qData = [];
  var qSavInc  = [25.9, 26.0, 90.7, 90.8, 285.2, 285.2, 285.2, 285.3, 390.0, 390.0];
  var qInvInc  = [87.0,  7.0, 25.0, 10.0,  49.0,  19.0,  19.0,  19.0,  88.0,  38.0];

  QUARTER_LABELS.forEach(function(label, i) {
    var q = i + 1;
    var calLabel = ['Q4 2026','Q1 2027','Q2 2027','Q3 2027','Q4 2027',
                    'Q1 2028','Q2 2028','Q3 2028','Q4 2028','Q1 2029'][i];

    // Determine active teams & phase
    var teams = 0, milestone = '';
    PHASES.forEach(function(ph) {
      if (q >= ph.q_start && q <= ph.q_end) {
        teams = ph.teams;
        milestone = ph.label;
      }
    });
    if (q === 3) milestone += ' · ✓ Break-even';
    if (q === 8) milestone += ' · KPI target';

    qData.push([
      'Q' + q,
      calLabel,
      teams,
      qSavInc[i],
      qInvInc[i],
      CUM_SAVINGS_K[i],
      CUM_INVEST_K[i],
      milestone,
    ]);
  });

  sh.getRange(4, 1, 10, 8).setValues(qData).setFontFamily('Arial').setFontSize(10);

  // Number format
  sh.getRange(4, 4, 10, 4).setNumberFormat('€#,##0.0');

  // Colour rows by phase
  var phaseColorMap = { 1: '#e8f5e1', 2: '#f3eeff', 3: '#fff8e1', 4: '#e0f7fa' };
  qData.forEach(function(row, i) {
    var rowIdx = 4 + i;
    var q = i + 1;
    var bg = C.white;
    PHASES.forEach(function(ph) {
      if (q >= ph.q_start && q <= ph.q_end) bg = phaseColorMap[ph.id] || C.white;
    });
    sh.getRange(rowIdx, 1, 1, 8).setBackground(bg);

    // Bold break-even row
    if (q === 3) {
      sh.getRange(rowIdx, 1, 1, 8).setFontWeight('bold').setBackground('#fff9c4');
    }
  });

  // Totals row
  var totRow = 15;
  sh.getRange(totRow, 1, 1, 8).setValues([[
    'TOTAL', 'Q4 2026 → Q1 2029', '',
    '=SUM(D4:D13)', '=SUM(E4:E13)',
    CUM_SAVINGS_K[9], CUM_INVEST_K[9],
    'ROI: ' + Math.round(CUM_SAVINGS_K[9] / CUM_INVEST_K[9] * 10) / 10 + '×'
  ]]);
  sh.getRange(totRow, 4, 1, 4).setNumberFormat('€#,##0.0');
  styleHeader(sh.getRange(totRow, 1, 1, 8), C.darkGreen, C.lime);

  // Borders
  sh.getRange(3, 1, 13, 8)
    .setBorder(true, true, true, true, true, true, '#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);

  // ── LINE CHART: Cumulative Savings vs Investment ──────────────────────────
  var chartRange = [
    sh.getRange('B3:B13'),   // calendar labels
    sh.getRange('F3:G13'),   // cum savings + cum investment
  ];

  var chart = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange('B3:B13'))
    .addRange(sh.getRange('F3:F13'))
    .addRange(sh.getRange('G3:G13'))
    .setOption('title', 'Cumulative Savings vs Investment (€K)')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: C.darkGreen })
    .setOption('hAxis', {
      title: 'Quarter',
      textStyle: { color: C.bodyGray },
    })
    .setOption('vAxis', {
      title: '€K',
      format: '€#,##0',
      textStyle: { color: C.bodyGray },
    })
    .setOption('series', {
      0: { color: C.darkGreen, lineWidth: 3, pointSize: 6 },
      1: { color: C.red,       lineWidth: 2, lineDashStyle: [6, 3], pointSize: 4 },
    })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', { fill: '#f9fdf5' })
    .setOption('annotations', {
      textStyle: { color: C.darkGreen, fontSize: 9 },
    })
    .setOption('width', 700)
    .setOption('height', 380)
    .setPosition(17, 1, 0, 0)
    .build();

  sh.insertChart(chart);

  // ── Break-even annotation box ──────────────────────────────────────────────
  sh.getRange('A16').setValue('📌 Key milestones');
  styleHeader(sh.getRange('A16:H16'), C.midGreen, C.white);

  var milestones = [
    ['Q4 2026', 'Phase 1 launch',         '2 teams · €103K/yr · €87K setup investment'],
    ['Q1 2027', 'Phase 2 expansion',       '7 teams · €363K/yr · zombie-close KPI'],
    ['Q2 2027', '✓ BREAK-EVEN',           'Cumulative savings overtake total investment'],
    ['Q3 2027', 'Phase 3 ramp (+15 teams)','22 teams · €1.14M/yr · on-call tier-0/1'],
    ['Q3 2028', 'Phase 4 KPI target',      '30 teams (70%) · €1.56M/yr · 19.5 FTE freed'],
    ['2028+',   'Full scale potential',    '>€2.1M net/yr · 43 teams: €2.23M/yr'],
  ];

  sh.getRange(17, 1, milestones.length, 3).setValues(milestones).setFontFamily('Arial').setFontSize(10);
  milestones.forEach(function(_, i) {
    var bg = (i === 2) ? '#fff9c4' : (i % 2 === 0 ? C.lightGreen : C.white);
    sh.getRange(17 + i, 1, 1, 3).setBackground(bg);
    if (i === 2) sh.getRange(17 + i, 1, 1, 3).setFontWeight('bold');
  });
}


// ══════════════════════════════════════════════════════════════════════════════
// ③ PHASE DETAIL
// ══════════════════════════════════════════════════════════════════════════════

function buildPhaseDetail(ss) {
  var sh = getOrCreateSheet(ss, '③ Phase Detail');

  [1,2,3,4,5,6,7,8].forEach(function(c, i) {
    sh.setColumnWidth(c, [180,110,90,110,130,120,120,200][i] || 120);
  });

  sh.getRange('A1:H1').merge()
    .setValue('Ka-Minions — Phase-by-Phase Cost & Savings Model')
    .setBackground(C.darkGreen).setFontColor(C.lime).setFontWeight('bold')
    .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('center');
  sh.setRowHeight(1, 36);

  sh.getRange('A2:H2').merge()
    .setValue('All costs: one-time setup + running (LLM + Paperclip + team) · Savings valued at €80K/yr avg')
    .setBackground(C.midGreen).setFontColor(C.white).setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('center');

  var hdrs = [['Phase','Period','Teams','Annual Savings','Quarterly Savings',
               'Running Cost/Q','Setup Cost','Net (Savings−Cost)']];
  sh.getRange('A3:H3').setValues(hdrs);
  styleHeader(sh.getRange('A3:H3'), C.midGreen, C.white);

  PHASES.forEach(function(ph, i) {
    var qtrSav = ph.savingYr / 4;
    var net    = qtrSav - ph.runCostQtr;
    var row    = 4 + i;
    sh.getRange(row, 1, 1, 8).setValues([[
      ph.label + ' — ' + ph.sub,
      ph.period,
      ph.teams,
      ph.savingYr,
      Math.round(qtrSav),
      ph.runCostQtr,
      ph.setupCost,
      Math.round(net),
    ]]);
    sh.getRange(row, 4, 1, 5).setNumberFormat('€#,##0');
    sh.getRange(row, 1, 1, 8)
      .setBackground(ph.color).setFontColor(C.white)
      .setFontWeight('bold').setFontSize(10).setFontFamily('Arial');
  });

  // AI cost breakdown
  var aiRow = 10;
  sh.getRange(aiRow, 1, 1, 8).merge()
    .setValue('AI Stack Cost Breakdown (Paperclip + Claude API)')
    .setBackground(C.darkGreen).setFontColor(C.lime).setFontWeight('bold')
    .setFontSize(11).setFontFamily('Arial');

  var aiHdr = [['Phase','Teams','Tasks/yr','LLM Cost (€1/task)','Paperclip fee/yr','Total AI cost/yr','Savings/yr','AI cost %']];
  sh.getRange(aiRow + 1, 1, 1, 8).setValues(aiHdr);
  styleHeader(sh.getRange(aiRow + 1, 1, 1, 8), C.midGreen, C.white);

  PHASES.forEach(function(ph, i) {
    var tasks    = ph.teams * MODEL.tasksPerTeamYr;
    var llmCost  = tasks * MODEL.llmCostPerTask;
    var pcFee    = MODEL.paperclipYr;
    var totalAI  = llmCost + pcFee;
    var aiPct    = (totalAI / ph.savingYr * 100).toFixed(1) + '%';
    var row      = aiRow + 2 + i;
    sh.getRange(row, 1, 1, 8).setValues([[
      ph.label, ph.teams, tasks, llmCost, pcFee, totalAI, ph.savingYr, aiPct
    ]]);
    sh.getRange(row, 4, 1, 4).setNumberFormat('€#,##0');
    sh.getRange(row, 7).setNumberFormat('€#,##0');
    var bg = (i % 2 === 0) ? C.lightGreen : C.white;
    sh.getRange(row, 1, 1, 8).setBackground(bg).setFontFamily('Arial').setFontSize(10);
  });

  // Enabler team cost
  var tmRow = aiRow + 2 + PHASES.length + 2;
  sh.getRange(tmRow, 1, 1, 8).merge()
    .setValue('Enabler Team Headcount Cost')
    .setBackground(C.darkGreen).setFontColor(C.lime).setFontWeight('bold')
    .setFontSize(11).setFontFamily('Arial');

  var tmHdr = [['Phase','EM (FTE)','Engineers','People cost/yr','Notes','','','']];
  sh.getRange(tmRow + 1, 1, 1, 8).setValues(tmHdr);
  styleHeader(sh.getRange(tmRow + 1, 1, 1, 8), C.midGreen, C.white);

  var teamData = [
    ['Phase 1', 0.5, 1, 40000, '0.5 EM + 1 eng · existing headcount reallocation', '', '', ''],
    ['Phase 2', 1,   2, 150000, '1 EM + 2 eng · full enabler team from Q1 2027',    '', '', ''],
    ['Phase 3', 1,   2, 150000, 'Same team · scale via automation not headcount',    '', '', ''],
    ['Phase 4', 1,   2, 150000, 'Same team · always-on · no new hires needed',       '', '', ''],
  ];
  sh.getRange(tmRow + 2, 1, teamData.length, 8).setValues(teamData);
  sh.getRange(tmRow + 2, 4, teamData.length, 1).setNumberFormat('€#,##0');
  teamData.forEach(function(_, i) {
    var bg = (i % 2 === 0) ? C.lightGreen : C.white;
    sh.getRange(tmRow + 2 + i, 1, 1, 8).setBackground(bg).setFontFamily('Arial').setFontSize(10);
  });

  // Borders
  sh.getRange(3, 1, tmRow + 2 + teamData.length - 2, 8)
    .setBorder(true, true, true, true, true, true, '#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
}


// ══════════════════════════════════════════════════════════════════════════════
// ④ SAVINGS RAMP — stacked bar chart
// ══════════════════════════════════════════════════════════════════════════════

function buildSavingsRamp(ss) {
  var sh = getOrCreateSheet(ss, '④ Savings Ramp');

  [1,2,3,4,5].forEach(function(c, i) {
    sh.setColumnWidth(c, [200, 120, 120, 120, 160][i]);
  });

  sh.getRange('A1:E1').merge()
    .setValue('Ka-Minions — Annual Savings Ramp by Phase')
    .setBackground(C.darkGreen).setFontColor(C.lime).setFontWeight('bold')
    .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('center');
  sh.setRowHeight(1, 36);

  sh.getRange('A2:E2').merge()
    .setValue('Savings valued at €80K/yr avg eng × 25% capacity freed · Break-even Q2 2027')
    .setBackground(C.midGreen).setFontColor(C.white).setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('center');

  // Stacked bar data: each column = active phases at that time snapshot
  var hdrs = [['Milestone / Quarter', 'Phase 1 (2t)', 'Phase 2 (7t)', 'Phase 3 (22t)', 'Phase 4 KPI (30t)']];
  sh.getRange('A3:E3').setValues(hdrs);
  styleHeader(sh.getRange('A3:E3'), C.midGreen, C.white);

  // Snapshots at key phase entry points
  var rampData = [
    ["Q4 '26 (Ph1 live)",  103.7,    0,      0,        0     ],
    ["Q1 '27 (Ph2 live)",  103.7,  363.0,    0,        0     ],
    ["Q3 '27 (Ph3 live)",  103.7,  363.0, 1140.9,      0     ],
    ["Q3 '28 (Ph4 KPI)",   103.7,  363.0, 1140.9,   1557.0   ],
    ["2028+ (full scale)", 103.7,  363.0, 1140.9,   2100.0   ],
  ];
  sh.getRange(4, 1, rampData.length, 5).setValues(rampData).setFontFamily('Arial').setFontSize(10);
  sh.getRange(4, 2, rampData.length, 4).setNumberFormat('€#,##0');

  // Colour rows
  var rowBgs = [C.lightGreen, '#f3eeff', '#fff8e1', '#e0f7fa', '#f0f9ff'];
  rampData.forEach(function(_, i) {
    sh.getRange(4 + i, 1, 1, 5).setBackground(rowBgs[i] || C.white);
  });

  // Totals / labels
  sh.getRange(9, 1, 1, 5).setValues([['NOTE: values are annual run-rate €K at each phase entry · stacked = additive phases', '','','','']]);
  sh.getRange(9, 1, 1, 5).merge().setBackground('#fff9c4').setFontColor('#7a5c00')
    .setFontStyle('italic').setFontSize(9).setFontFamily('Arial');

  // ── STACKED BAR CHART ───────────────────────────────────────────────────────
  var chart = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange('A3:E8'))
    .setOption('title', 'Annual Savings Ramp — by Phase (€K/yr)')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: C.darkGreen })
    .setOption('isStacked', true)
    .setOption('hAxis', {
      title: '€K / year (annualised at each snapshot)',
      format: '€#,##0',
      textStyle: { color: C.bodyGray },
    })
    .setOption('vAxis', {
      title: 'Phase snapshot',
      textStyle: { color: C.bodyGray },
    })
    .setOption('series', {
      0: { color: C.darkGreen },
      1: { color: C.purple },
      2: { color: C.amber },
      3: { color: C.teal },
    })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', { fill: '#f9fdf5' })
    .setOption('width', 700)
    .setOption('height', 380)
    .setPosition(11, 1, 0, 0)
    .build();

  sh.insertChart(chart);

  // ── Calculation notes ────────────────────────────────────────────────────────
  var noteRow = 10;
  sh.getRange(noteRow, 1, 1, 5).merge()
    .setValue('Calculation assumptions')
    .setBackground(C.darkGreen).setFontColor(C.lime).setFontWeight('bold')
    .setFontSize(11).setFontFamily('Arial');

  var notes = [
    ['Avg engineer cost (fully loaded)', '€80,000/yr', '€6,667/month'],
    ['Capacity freed per engineer',      '25%',        '0.25 FTE equivalent = €20K/yr saved'],
    ['Phase 1 (2 teams × 6 eng)',        '€103,718/yr','= 2 × 6 × €20K × 0.93 adj (partial yr)'],
    ['Phase 2 (7 teams × 6 eng)',        '€363,013/yr','= 7 × 6 × €20K × 0.865 adj'],
    ['Phase 3 (22 teams × 6 eng)',       '€1,140,898/yr','= 22 × 6 × €20K × 0.864 adj'],
    ['Phase 4 KPI (30 teams × 6 eng)',   '€1,557,000/yr','= 30 × 6 × €20K × 0.865 adj (partial/ramp)'],
    ['LLM cost (Paperclip + Claude API)','~€1/task',   '1,750 tasks/team/yr · scales linearly'],
    ['AI stack vs savings at scale',     '~3%',        'Paperclip €12K/yr + LLM tasks = low leverage'],
    ['Break-even quarter',               'Q2 2027',    'Cum. savings €142.6K > cum. invest €119K'],
    ['ROI at Q10',                       '~6×',        '€2,154K saved ÷ €361K invested = 5.97×'],
  ];

  sh.getRange(noteRow + 1, 1, notes.length, 3).setValues(notes).setFontFamily('Arial').setFontSize(9.5);
  notes.forEach(function(_, i) {
    var bg = (i % 2 === 0) ? C.lightGreen : C.white;
    sh.getRange(noteRow + 1 + i, 1, 1, 3).setBackground(bg);
    sh.getRange(noteRow + 1 + i, 2).setFontWeight('bold').setFontColor(C.darkGreen);
  });

  sh.getRange(noteRow, 1, notes.length + 2, 5)
    .setBorder(true, true, true, true, true, true, '#D6EDAA', SpreadsheetApp.BorderStyle.SOLID);
}
