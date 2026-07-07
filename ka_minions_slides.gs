/**
 * Ka-Minions — Google Slides Builder
 * Replica fiel del HTML ka-minions-business-case.html
 *
 * USO:
 * 1. Extensions → Apps Script en cualquier Google Sheets/Docs (o script.google.com)
 * 2. Pega este fichero completo
 * 3. Run → buildKaMinionsSlides
 * 4. Se crea una presentación nueva en tu Drive
 */

// Paleta 
var G = '#1D4B00';
var GM = '#609F28';
var LIME = '#B5E941';
var PUR = '#7c3aed';
var AMB = '#d97706';
var TEAL = '#0891b2';
var RED = '#c0392b';
var WH = '#FFFFFF';
var BK = '#0e1a08';
var GY = '#6b7f5e';
var GY2 = '#504E48';
var BDR = '#D6EDAA';
var S2 = '#f7faf2';
var GL = '#f0fae6';
var YEL = '#fffbeb';
var OFFWH= '#F4F2EF';

// Dimensiones diseño (960×540) → slide real (720×540 pt = 10"×7.5") 
var SW = 960; // diseño: ancho
var SH = 540; // diseño: alto
var SX = 720/960; // scale x = 0.75
var SY = 540/540; // scale y = 1.0

// Helpers 

function newPres() {
 return SlidesApp.create('Ka-Minions Business Case · Ideas2Impact 2026');
}

function addSlide(pres, idx) {
 if (idx === 0) return pres.getSlides()[0];
 return pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
}

function clearSlide(slide) {
 slide.getPageElements().forEach(function(e) { e.remove(); });
 slide.getBackground().setSolidFill(WH);
}

// Convierte coordenadas de diseño → puntos reales del slide
function px(v) { return v * SX; }
function py(v) { return v * SY; }

function box(slide, x, y, w, h, bg, border, borderW) {
 var s = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, px(x), py(y), px(w), py(h));
 s.getFill().setSolidFill(bg || WH);
 if (border) {
 s.getBorder().getLineFill().setSolidFill(border);
 s.getBorder().setWeight(borderW || 1);
 } else {
 s.getBorder().setTransparent();
 }
 return s;
}

function txt(slide, text, x, y, w, h, opts) {
 opts = opts || {};
 var s = slide.insertTextBox(text, px(x), py(y), px(w), py(h));
 s.getFill().setTransparent();
 s.getBorder().setTransparent();
 var tf = s.getText();
 var st = tf.getTextStyle();
 st.setFontFamily('Arial');
 st.setFontSize(opts.size || 10);
 st.setBold(opts.bold || false);
 st.setItalic(opts.italic || false);
 st.setForegroundColor(opts.color || BK);
 var ps = tf.getParagraphStyle();
 ps.setParagraphAlignment(
 opts.align === 'center' ? SlidesApp.ParagraphAlignment.CENTER :
 opts.align === 'right' ? SlidesApp.ParagraphAlignment.END :
 SlidesApp.ParagraphAlignment.START
 );
 return s;
}

function hdrBand(slide, title, eyebrow, sub) {
 // dark green band
 box(slide, 0, 0, SW, 76, G);
 // lime accent line
 box(slide, 0, 76, SW, 3, LIME);
 if (eyebrow) txt(slide, eyebrow, 36, 8, SW-72, 14, {size:8, color:LIME, bold:true});
 txt(slide, title, 36, 22, SW-72, 36, {size:17, bold:true, color:WH});
 if (sub) txt(slide, sub, 36, 58, SW-72, 18, {size:8.5, color:'#ccffcc'});
}

function footerPhases(slide) {
 box(slide, 0, SH-38, SW, 38, S2, BDR, 0.5);
 var phases = [
 {t:'Ph1 — Q4 2026', s:'TRX · 2t · €103K/yr', c:G},
 {t:'Ph2 — Q1–Q2 2027', s:'7t · €464K/yr · on-call −30% · Break-even Q2 2027', c:PUR},
 {t:'Ph3 — Q3 2027–Q2 2028', s:'22t · €1.46M/yr · incl. on-call savings', c:AMB},
 {t:'Ph4 — Q3 2028 → KPI',s:'30t · €1.99M/yr · >€2.85M net 2028+', c:TEAL},
 ];
 phases.forEach(function(ph, i) {
 var x = 8 + i * 236;
 box(slide, x, SH-38, 3, 38, ph.c);
 txt(slide, ph.t, x+6, SH-35, 228, 16, {size:7.5, bold:true, color:BK});
 txt(slide, ph.s, x+6, SH-20, 228, 16, {size:7, color:GY});
 });
}

function statTile(slide, x, y, w, h, label, value, sub, bg, valColor, valSize) {
 box(slide, x, y, w, h, bg || S2, BDR, 1);
 txt(slide, label, x+8, y+7, w-16, 13, {size:7.5, bold:true, color:GY});
 txt(slide, value, x+8, y+20, w-16, 24, {size:valSize||22, bold:true, color:valColor||G});
 txt(slide, sub, x+8, y+45, w-16, 14, {size:8, color:GY2});
}

function phaseCard(slide, x, y, w, h, phase, title, desc, saving, kpi, accentColor, chipBg) {
 box(slide, x, y, w, h, WH, BDR, 1);
 box(slide, x, y, w, 3.5, accentColor);
 box(slide, x, y, w, 38, chipBg || GL);
 txt(slide, phase, x+9, y+5, w-18, 14, {size:7.5, bold:true, color:accentColor});
 txt(slide, title, x+9, y+42, w-18, 28, {size:11, bold:true, color:accentColor});
 txt(slide, desc, x+9, y+72, w-18, 30, {size:8.5, italic:true, color:GY2});
 box(slide, x+8, y+106, w-16, 22, chipBg || GL, accentColor, 1);
 txt(slide, saving, x+12, y+109, w-24, 16, {size:8.5, bold:true, color:accentColor});
 box(slide, x+8, y+132, w-16, 20, GL);
 txt(slide, kpi, x+12, y+134, w-24, 16, {size:7.5, bold:true, color:G});
}

function roiFooter(slide, tiles) {
 box(slide, 0, SH-52, SW, 52, G);
 box(slide, 0, SH-52, SW, 3, LIME);
 tiles.forEach(function(t, i) {
 var x = 10 + i * (SW/tiles.length);
 var w = SW/tiles.length - 10;
 txt(slide, t.label, x, SH-48, w, 12, {size:6.5, bold:true, color:'#aaaaaa', align:'center'});
 txt(slide, t.value, x, SH-36, w, 18, {size:14, bold:true, color:LIME, align:'center'});
 txt(slide, t.sub, x, SH-18, w, 14, {size:6.5, color:'#99bb99', align:'center'});
 });
}


// 
// SLIDE 1 — The Problem & 3 Value Pillars
// 
function buildSlide1(pres) {
 var sl = addSlide(pres, 0);
 clearSlide(sl);

 hdrBand(sl,
 'Engineers are spending ~35% of their sprint on work AI agents can own — across all 43 KA teams',
 'Ka-Minions · Business Case · Ideas2Impact 2026',
 'Pilot: Teams AZ & MO · 6-month backlog · KPI target: 70% adoption = 30 teams = €1.56M/yr · 19.5 FTE freed'
 );

 // STAT ROW 
 var statY = 84;
 statTile(sl, 8, statY, 140, 64, 'Avg bug resolution today', '5 days', 'Ka-Minion target: <3 hours (40×) · n=206 issues · 2 KA teams', S2, AMB, 20);
 statTile(sl, 153, statY, 140, 64, 'Zombie tickets >30 days', '22%', 'Backlog noise every sprint review', S2, RED, 20);
 statTile(sl, 298, statY, 140, 64, 'Sprint capacity on ops & triage', '30–40%', 'vs feature & product delivery', S2, RED, 20);
 statTile(sl, 443, statY, 175, 64, 'Out-of-hours duty calls & alerts/wk','~3 calls','Nights · weekends · ~€8K/eng/yr premium', S2, AMB, 20);
 // Lime recoverable tile
 box(sl, 622, statY, 330, 64, '#f5ffe0', LIME, 1.5);
 txt(sl, 'Recoverable capacity · KPI target (30/43 teams)', 624, statY+6, 326, 13, {size:7.5, bold:true, color:GY});
 txt(sl, '€1.99M / yr', 624, statY+19, 200, 26, {size:24, bold:true, color:G});
 txt(sl, 'From Phase 4 (Q3 2028) · incl. on-call savings (30% · €8K/eng) · full potential 43t: €2.85M', 624, statY+46, 326, 14, {size:7.5, color:GY2});
 // Timeline dots
 var tl = [[G,'Q4 2026','€103K'],[GM,'Q1 2027','€363K'],[AMB,'Q3 2027','€1.14M'],[TEAL,'Q3 2028','€1.56M'],['#B5E941','2028+','>€2.1M']];
 tl.forEach(function(d, i) {
 var dx = 622 + i * 66;
 box(sl, dx+4, statY+43, 8, 8, d[0], d[0], 0).setDescription(d[1]);
 txt(sl, d[1], dx-4, statY+53, 56, 10, {size:6, bold:true, color:d[0], align:'center'});
 txt(sl, d[2], dx-4, statY+62, 56, 10, {size:6, bold:true, color:BK, align:'center'});
 });

 // 3 PILLARS 
 var py = 154, ph = 220;
 var pillars = [
 { n:'Pillar 01', h:'Faster delivery of bugs & routine tasks',
 bad:'Bugs sit 5 days before a human opens them — measured across 206 issues in 2 KA teams. 80% of ≤1SP bugs are automatable.',
 good:'Ka-Minion picks up any ≤1SP ticket the moment it lands in Jira. Reads context, writes fix, runs CI, opens PR — 24/7, no queue, no handoff.',
 kv:'5d → <3h', kl:'Phase 1 gate KPI · bug resolution time · 40× faster · n=206 issues'},
 { n:'Pillar 02', h:'Zombie task closure — healthy backlog, zero noise',
 bad:'1 in 5 Jira tickets is a zombie (>30 days). Across 43 teams this is hundreds of stale issues inflating estimates every sprint.',
 good:'Sentinel mode runs nightly: scans stale tickets, auto-enriches context, resolves if automatable, flags for human if not. Backlog clean by default.',
 kv:'22% → <8%', kl:'Phase 2 KPI · stale ticket rate across all active teams'},
 { n:'Pillar 03', h:'Out-of-hours duty calls & alerts — engineers get their weekends back',
 bad:'Engineers rotate on-duty shifts covering nights & weekends. ~3 hands-on calls & alerts/team/week — restarts, rollbacks, known error patterns.',
 good:'Ka-Minion Sentinels watch 24/7. When a known pattern fires they act immediately. Only genuine unknowns wake a human. Engineers stay off-duty unless it truly matters.',
 kv:'−30%', kl:'Phase 2 KPI · hands-on out-of-hours calls & alerts per team/week'},
 ];

 pillars.forEach(function(p, i) {
 var px = 8 + i * 317, pw = 310;
 box(sl, px, py, pw, ph, S2, BDR, 1);
 txt(sl, p.n, px+10, py+8, pw-20, 12, {size:8, bold:true, color:GY});
 txt(sl, p.h, px+10, py+20, pw-20, 28, {size:11, bold:true, color:G});
 // red box
 box(sl, px+8, py+52, pw-16, 54, '#fff5f5');
 txt(sl, p.bad, px+12, py+54, pw-24, 50, {size:8.5, color:'#7f1d1d'});
 // green box
 box(sl, px+8, py+110, pw-16, 70, GL);
 txt(sl, p.good, px+12, py+112, pw-24, 66, {size:8.5, color:BK});
 // KPI
 box(sl, px+8, py+186, pw-16, 1, BDR);
 txt(sl, p.kv, px+10, py+192, pw-20, 18, {size:16, bold:true, color:G});
 txt(sl, p.kl, px+10, py+210, pw-20, 12, {size:8, color:GY});
 });

 // BOTTOM CALLOUTS 
 var cy = 380;
 box(sl, 8, cy, 626, 34, GL, BDR, 0.5);
 txt(sl, '', 14, cy+8, 18, 18, {size:12});
 txt(sl, 'Scope guardrail: Ka-Minions only touch ≤1SP issues. No merges to main without CI + Paperclip audit gate. Human approval for anything above scope. PoC running inside Banana Squad.', 34, cy+5, 594, 26, {size:8.5, color:'#14532d'});

 box(sl, 638, cy, 314, 34, '#fff5f5', BDR, 0.5);
 txt(sl, '', 644, cy+8, 18, 18, {size:12});
 txt(sl, 'Cost of waiting: Each quarter without Ka-Minions = ~€557K in recoverable capacity lost across all 43 teams. The zombie rate grows ~2pp/quarter at current trend.', 664, cy+5, 280, 26, {size:8.5, color:'#7f1d1d'});

 footerPhases(sl);
 txt(sl, '1 / 7', SW-40, SH-52, 36, 12, {size:9, color:GY});
}


// 
// SLIDE 2 — ROI & Expansion
// 
function buildSlide2(pres) {
 var sl = addSlide(pres, 1);
 clearSlide(sl);

 hdrBand(sl,
 '€1.99M/yr at 70% adoption (30 teams) — phased by domain, proven at each gate before scaling',
 'Ka-Minions · Business Case · Ideas2Impact 2026',
 'Phase 1: TRX pilot (2 teams, €103K/yr) · KPI target: 30/43 teams (70%), 9 domains · Full potential (incl. on-call): €2.85M'
 );

 // LEFT: domain table 
 var lx = 8, tw = 438, ty = 84;
 txt(sl, 'FULL-ORG EXPANSION — BY DOMAIN', lx, ty, tw, 12, {size:7.5, bold:true, color:GY});

 // Table header
 box(sl, lx, ty+14, tw, 16, G);
 ['Domain','Teams','Hours/yr','Savings/yr','Coverage'].forEach(function(h, i) {
 var xs = [lx+4, lx+160, lx+215, lx+280, lx+360];
 txt(sl, h, xs[i], ty+17, 70, 12, {size:7, bold:true, color:WH});
 });

 var domains = [
 {d:'TRX · Phase 1 pilot', t:2, h:'2,282', s:'€103,718', pct:22, c:G},
 {d:'Consumer Flywheel', t:4, h:'4,564', s:'€207,436', pct:44, c:GM},
 {d:'D&A', t:1, h:'1,141', s:'€51,859', pct:11, c:GM},
 {d:'Consumer Growth', t:5, h:'5,704', s:'€259,295', pct:56, c:PUR},
 {d:'Advertising', t:6, h:'6,845', s:'€311,154', pct:67, c:PUR},
 {d:'PRO', t:4, h:'4,564', s:'€207,436', pct:44, c:PUR},
 {d:'RE & Motors', t:6, h:'6,845', s:'€311,154', pct:67, c:TEAL},
 {d:'Platform', t:9, h:'10,268', s:'€466,731', pct:100,c:TEAL},
 {d:'Data', t:6, h:'6,845', s:'€311,154', pct:67, c:TEAL},
 {d:'TOTAL — all 43 teams', t:43, h:'49,059', s:'€2,229,937', pct:0, c:GY, tot:true},
 {d:'KPI target — 30 teams (70%)', t:30, h:'34,230', s:'€1,557,000', pct:0, c:G, kpi:true},
 ];
 domains.forEach(function(row, i) {
 var ry = ty + 30 + i * 16;
 var bg = row.kpi ? GL : row.tot ? '#f6faf0' : (i%2===0 ? WH : '#f6faf0');
 box(sl, lx, ry, tw, 16, bg);
 txt(sl, row.d, lx+4, ry+2, 154, 12, {size:8, bold:row.kpi||row.tot, color:row.kpi?G:BK});
 txt(sl, ''+row.t,lx+160, ry+2, 50, 12, {size:8, bold:row.kpi, color:row.kpi?G:BK, align:'right'});
 txt(sl, row.h, lx+215, ry+2, 60, 12, {size:8, color:GY2, align:'right'});
 txt(sl, row.s, lx+280, ry+2, 76, 12, {size:8, bold:row.kpi||row.tot, color:G, align:'right'});
 if (row.pct > 0) {
 box(sl, lx+360, ry+4, 72, 8, '#e5f0d2');
 box(sl, lx+360, ry+4, Math.round(72*row.pct/100), 8, row.c);
 }
 });

 // FTE + LLM tiles
 statTile(sl, lx, ty+214, 214, 54, 'FTE equivalent at KPI target (30 teams)', '19.5 FTE', 'Same headcount · more output · 27.9 FTE at full 43-team potential', '#f5ffe0', G, 20);
 statTile(sl, lx+218, ty+214, 218, 54, 'LLM running cost vs savings', '~3%', 'Claude API ≈ €1/task · math holds at 10× token price', '#f5ffe0', G, 20);

 // RIGHT: phase roadmap 
 var rx = 452, rw = SW-rx-8;
 txt(sl, '4-PHASE ROLLOUT — DOMAIN BY DOMAIN', rx, ty, rw, 12, {size:7.5, bold:true, color:GY});

 var rphases = [
 {n:'1', t:'Hypothesis Validation · Q4 2026 — 1 quarter',
 s:'Domains: TRX (2 teams) · €103,718/yr · Automate ≤1SP bugs & BAU tasks · Revenue starts Q4 2026',
 gate:'Gate: 3d→<3h', active:true},
 {n:'2', t:'MVPs & Expansion · Q1–Q2 2027 — 2 quarters',
 s:'Consumer Flywheel (4) + D&A (1) → 7 teams · €363,013/yr · Sentinel zombie-close · Break-even Q2 2027',
 gate:'Gate: 70% success'},
 {n:'3', t:'Growth & Iterations · Q3 2027–Q2 2028 — 4 quarters',
 s:'Consumer Growth (5) + Advertising (6) + PRO (4) → 22 teams · €1,140,898/yr · on-call duty calls & alerts coverage',
 gate:'Gate: 40% adoption', dim:true},
 {n:'4', t:'Scale — Always On · Q3 2028 → full scale 2028+',
 s:'RE & Motors (6) + Platform (9) + Data (6) → 30 KPI teams · €1.56M/yr from Q3 2028 · >€2.1M net in 2028+',
 gate:' KPI: 70% teams active agents', dim:true},
 ];

 rphases.forEach(function(rp, i) {
 var ry = ty+14+i*58, rh2 = 54;
 var bg = rp.active ? GL : WH;
 var op = rp.dim ? 0.7 : 1;
 box(sl, rx, ry, rw, rh2, bg, BDR, 1);
 box(sl, rx, ry, 22, 22, rp.dim ? '#c8d9be' : G);
 txt(sl, rp.n, rx+5, ry+4, 12, 14, {size:10, bold:true, color:WH});
 txt(sl, rp.t, rx+26, ry+4, rw-100, 14, {size:9, bold:true, color:rp.dim?GY:G});
 txt(sl, rp.s, rx+26, ry+18, rw-100, 28, {size:7.5, color:rp.dim?GY:GY2});
 txt(sl, rp.gate, rx+rw-90, ry+16, 88, 14, {size:7.5, bold:true, color:GY, align:'right'});
 });

 // ROI mini-table
 var mty = ty+248;
 box(sl, rx, mty, rw, 14, G);
 ['Phase','Teams','Annual Savings','AI Cost','Net Gain'].forEach(function(h, i) {
 var xs = [rx+4, rx+80, rx+130, rx+210, rx+280];
 txt(sl, h, xs[i], mty+2, 72, 10, {size:7, bold:true, color:WH});
 });
 var rtData = [
 ['Phase 1 · Q4 2026', '2', '€103,718', '~€3,500', '~€100K'],
 ['Phase 2 · Q1–Q2 2027 · +on-call', '7', '€463,813', '~€9,000', '~€455K'],
 ['Phase 3 · Q3 2027–Q2 2028 · +on-call', '22', '€1,457,698', '~€25K', '~€1.43M'],
 ['Phase 4 KPI · 30t Q3 2028+ · +on-call','30', '€1,989,000', '~€80K', '>€1.91M/yr'],
 ];
 rtData.forEach(function(r, i) {
 var ry = mty+14+i*16;
 box(sl, rx, ry, rw, 16, i%2===0?WH:'#f6faf0');
 var xs = [rx+4, rx+80, rx+130, rx+210, rx+280];
 r.forEach(function(v, ci) {
 var isVal = ci >= 2;
 txt(sl, v, xs[ci], ry+3, 72, 10, {size:8, bold:ci===4||ci===0, color:ci===4?G:BK, align:isVal?'right':'left'});
 });
 });
 // KPI row
 var kpiry = mty+14+64;
 box(sl, rx, kpiry, rw, 18, GL, BDR, 0.5);
 txt(sl, 'Phase 4 KPI target · 70% adoption = 30 teams · from Q3 2028 · incl. on-call', rx+4, kpiry+3, 260, 12, {size:7.5, bold:true, color:G});
 txt(sl, '>€1.91M / yr', rx+rw-70, kpiry+3, 66, 12, {size:9, bold:true, color:G, align:'right'});

 // Single ask callout
 box(sl, rx, mty+96, rw, 28, '#fff5f5', RED, 0.5);
 txt(sl, 'Single ask today: Approve Phase 1 (€60–80K setup · Q4 2026). No further investment unlocks without hitting the gate KPI. Break-even Q2 2027. €1.56M/yr from Q3 2028. ROI >6× by 2028+.', rx+8, mty+100, rw-16, 22, {size:8, color:'#7f1d1d'});

 footerPhases(sl);
 txt(sl, '2 / 7', SW-40, SH-52, 36, 12, {size:9, color:GY});
}


// 
// SLIDE 3 — Costs, ROI & Governance
// 
function buildSlide3(pres) {
 var sl = addSlide(pres, 2);
 clearSlide(sl);

 hdrBand(sl,
 '€361K total investment over 10 quarters — ~6× ROI at 70% adoption target, break-even Q2 2027',
 'Ka-Minions · Business Case · Ideas2Impact 2026',
 'KPI target: 30/43 teams (70%) · €1.56M/yr savings · LLM cost ≈ 3% of savings · governance via Paperclip from day one'
 );

 var bodyY = 84;

 // LEFT: Operational costs table 
 txt(sl, 'OPERATIONAL COSTS — PER PHASE', 8, bodyY, 400, 12, {size:7.5, bold:true, color:GY});
 box(sl, 8, bodyY+14, 464, 14, G);
 ['Phase','Scope','Setup','LLM API/yr','Maintenance/yr','Total/yr'].forEach(function(h, i) {
 var xs = [10,70,160,228,310,390];
 txt(sl, h, xs[i], bodyY+17, 80, 10, {size:7, bold:true, color:WH});
 });
 var costData = [
 ['Phase 1','2 teams (TRX)', '€60–80K','~€3,500', '~€10K','~€74K'],
 ['Phase 2','7 teams', '—', '~€9,000', '~€15K','~€24K'],
 ['Phase 3','22 teams', '—', '~€25,000','~€30K','~€55K'],
 ['Phase 4','43 teams (full)', '—', '~€65,000','~€50K','~€115K/yr'],
 ];
 costData.forEach(function(r, i) {
 var ry = bodyY+28+i*18;
 var bg = i===3 ? GL : i%2===0 ? WH : '#f6faf0';
 box(sl, 8, ry, 464, 18, bg);
 var xs = [10,70,160,228,310,390];
 r.forEach(function(v, ci) {
 txt(sl, v, xs[ci], ry+4, 76, 12, {size:8.5, bold:ci===0||ci===5, color:ci===5?G:BK});
 });
 });

 // ROI Summary 
 txt(sl, 'ROI SUMMARY', 8, bodyY+104, 400, 12, {size:7.5, bold:true, color:GY});
 statTile(sl, 8, bodyY+118, 112, 52, 'Break-even', 'Q2 2027', 'Cum. savings overtake investment mid-Ph2', '#f5ffe0', G, 18);
 statTile(sl, 124, bodyY+118, 112, 52, 'ROI at Q10 (30t)', '~6×', '~€2.15M saved vs €361K invested', '#f5ffe0', G, 18);
 statTile(sl, 240, bodyY+118, 112, 52, 'Net gain Year 1 (Ph1)', '~€30K', '€103K savings − €74K yr-1 cost', S2, G, 18);
 statTile(sl, 356, bodyY+118, 116, 52, 'Net gain at KPI target (30t)', '>€1.48M', '€1.56M savings − €80K running cost', S2, G, 18);

 // RIGHT: Governance grid 
 txt(sl, 'GOVERNANCE & RISK MITIGATION', 484, bodyY, 460, 12, {size:7.5, bold:true, color:GY});
 var govItems = [
 {icon:'', t:'Scope guardrail:', b:'Ka-Minions only act on ≤1 Story Point issues. No merge to main without CI + Paperclip audit gate. Hard-coded at orchestration layer.', good:true},
 {icon:'', t:'Full audit trail:', b:'Every agent action (code read, diff, PR) is logged in Paperclip with timestamp, token cost, and outcome. Exportable for compliance.', good:true},
 {icon:'', t:'Kill switch:', b:'Any Ka-Minion can be paused or terminated instantly from Paperclip dashboard — per-agent, per-team, or org-wide. No runaway agents.', good:true},
 {icon:'', t:'LLM hallucination risk:', b:'Mitigated by requiring CI green + reviewer approval before any merge. Ka-Minion proposes; human (or automated gate) approves.', good:false},
 {icon:'', t:'API cost overrun risk:', b:'Per-agent monthly budget caps in Paperclip. Teams cannot exceed allocated token spend without explicit override from their EM.', good:false},
 {icon:'', t:'Phase-gated investment:', b:'No further spend unlocks without hitting the prior phase KPI. Zero risk of over-investing before value is proven.', good:true},
 ];
 govItems.forEach(function(g, i) {
 var gx = 484 + (i%2)*238, gy = bodyY+14+Math.floor(i/2)*76;
 var bg = g.good ? GL : YEL;
 var tc = g.good ? '#14532d' : '#78350f';
 box(sl, gx, gy, 232, 70, bg, BDR, 0.5);
 txt(sl, g.icon, gx+6, gy+8, 20, 20, {size:13});
 txt(sl, g.t, gx+28, gy+8, 198, 14, {size:9, bold:true, color:tc});
 txt(sl, g.b, gx+28, gy+22, 198, 44, {size:8, color:tc});
 });

 // AI cost model summary
 txt(sl, 'AI STACK COST MODEL (Paperclip + Claude API)', 8, bodyY+178, 464, 12, {size:7.5, bold:true, color:GY});
 var aiData = [
 {ph:'Phase 1 · 2t · ~3,500 tasks/yr', cost:'€3,500/yr', pct:'3.4% of €103K savings', c:G},
 {ph:'Phase 2 · 7t · ~12,250 tasks/yr',cost:'€12,250/yr', pct:'3.4% of €363K savings', c:PUR},
 {ph:'Phase 3 · 22t · ~38,500 tasks/yr',cost:'€38,500/yr',pct:'3.4% of €1.14M savings',c:AMB},
 {ph:'Phase 4 KPI · 30t · ~52,500 tasks',cost:'€52,500/yr',pct:'3.4% of €1.56M savings',c:TEAL},
 ];
 aiData.forEach(function(a, i) {
 var ax = 8 + i*116, ay = bodyY+192, aw = 112;
 box(sl, ax, ay, aw, 52, S2, BDR, 1);
 box(sl, ax, ay, aw, 3, a.c);
 txt(sl, a.ph, ax+6, ay+8, aw-12, 18, {size:7, color:GY2});
 txt(sl, a.cost, ax+6, ay+24, aw-12, 16, {size:12, bold:true, color:a.c});
 txt(sl, a.pct, ax+6, ay+40, aw-12, 12, {size:7, color:GY});
 });

 footerPhases(sl);
 txt(sl, '3 / 7', SW-40, SH-52, 36, 12, {size:9, color:GY});
}


// 
// SLIDE 4 — Team & AI Cost Model
// 
function buildSlide4(pres) {
 var sl = addSlide(pres, 3);
 clearSlide(sl);

 hdrBand(sl,
 'Enabler team + two SaaS tools — total ~€115K/yr vs €1.56M/yr freed · 1 EM + 2 engineers',
 'Ka-Minions · Business Case · Ideas2Impact 2026',
 '1 EM + 2 engineers (existing headcount reallocation) · Paperclip + Claude API · cost ≈ 7% of KPI savings'
 );

 var bodyY = 84;
 var cols = [
 {t:'Enabler Team', c:G, items:['EM (50% Phase 1 → full from Ph2)','2 Engineers (1 from Ph1, +1 from Ph2)','Prioritises backlog, manages phase gates','Trains domain-specific agent context','→ No new hire needed until Phase 2']},
 {t:'Paperclip Platform', c:PUR, items:['Orchestration & governance layer','~€12–24K/yr (team plan)','Full audit trail per agent action','Human-in-the-loop approval gates','Instant kill-switch per agent or org-wide']},
 {t:'Claude API', c:AMB, items:['LLM engine for agent reasoning','~€1/task average across all types','Phase 1 (3.5K tasks): ~€3.5K/yr','Phase 4 (52K tasks): ~€52K/yr','Cost ≈ 3.3% of savings at scale']},
 {t:'Governance & Risk', c:TEAL, items:['Scope guardrail: ≤1 SP tasks only','Budget cap per agent (EM override)','CI must be green before any merge','Per-agent cost dashboard in Paperclip','Rollback in 1 click — zero blast radius']},
 ];

 cols.forEach(function(col, i) {
 var cx = 8 + i * 236, cw = 232, ch = 340;
 box(sl, cx, bodyY, cw, ch, S2, BDR, 1);
 box(sl, cx, bodyY, cw, 3, col.c);
 box(sl, cx, bodyY, cw, 24, '#f5f5f5');
 txt(sl, col.t, cx+8, bodyY+6, cw-16, 16, {size:11, bold:true, color:col.c});
 col.items.forEach(function(item, j) {
 var iy = bodyY + 34 + j*48;
 box(sl, cx+8, iy, 4, 34, col.c);
 txt(sl, '› ' + item, cx+16, iy+2, cw-24, 30, {size:9, color:BK});
 });
 });

 // Phase investment insight
 box(sl, 8, bodyY+350, SW-16, 36, '#f5ffe0', LIME, 1);
 txt(sl, ' Phase-gated: no spend unlocks without hitting the prior gate KPI. Single ask today: €60–80K Phase 1 setup (Q4 2026). Break-even Q2 2027. €1.56M/yr from Q3 2028. ROI >6× by 2028+.', 20, bodyY+358, SW-36, 24, {size:9, color:G});

 footerPhases(sl);
 txt(sl, '4 / 7', SW-40, SH-52, 36, 12, {size:9, color:GY});
}


// 
// SLIDE 5 — Impact at a Glance (matches ka-minions-business-case.html #s5)
// 
function buildSlide5(pres) {
 var sl = addSlide(pres, 4);
 clearSlide(sl);

 // HEADER (imp-hdr) 
 box(sl, 0, 0, SW, 68, G);
 box(sl, 0, 68, SW, 4, LIME);
 txt(sl, 'Ka-Minions · Ideas2Impact 2026 · Banana Squad', 28, 6, 580, 12, {size:8.5, bold:true, color:LIME});
 txt(sl, 'Autonomous AI agents that free engineering capacity worth €1.56M/yr —', 28, 18, 680, 18, {size:14, bold:true, color:WH});
 txt(sl, 'same headcount, radically more product delivery', 28, 36, 680, 16, {size:14, bold:true, color:WH});

 // ibadge tiles — top right
 var bdata = [['30/43','Teams · KPI target'],['Q3','Break-even'],['~6×','ROI at Q10']];
 bdata.forEach(function(b, i) {
 var bx = 720 + i * 80;
 box(sl, bx, 5, 76, 56, '#2a5a10', LIME, 1.5);
 txt(sl, b[0], bx+2, 11, 72, 26, {size:16, bold:true, color:LIME, align:'center'});
 txt(sl, b[1], bx+2, 40, 72, 14, {size:7, color:'#cccccc', align:'center'});
 });
 // kpi-pill
 box(sl, 720, 3, 236, 20, '#2a5a10', '#6b9e30', 1);
 box(sl, 724, 8, 7, 7, LIME);
 txt(sl, 'Top KPI: 70% of KA teams with active agents', 736, 5, 214, 16, {size:8, bold:true, color:WH});

 // BODY: 3 columns 
 // Col1 x=0..320 | div | Col2 x=322..641 | div | Col3 x=643..960
 var C1X = 0, C1W = 320;
 var C2X = 322, C2W = 318;
 var C3X = 642, C3W = 318;
 var BY = 72; // body top (below header+border)
 var BH = SH - BY - 38; // body height (above footer)

 // Column backgrounds
 box(sl, C1X, BY, C1W, BH, WH);
 box(sl, C2X, BY, C2W, BH, WH);
 box(sl, C3X, BY, C3W, BH, WH);
 // Dividers
 box(sl, C2X-1, BY, 1.5, BH, '#e5f0d2');
 box(sl, C3X-1, BY, 1.5, BH, '#e5f0d2');

 // COLUMN 1 — The Problem 
 var P = 14; // col padding
 txt(sl, 'The Problem today', C1X+P, BY+P, C1W-P*2, 10, {size:7.5, bold:true, color:'#609F28'});
 txt(sl, 'Engineers are doing work', C1X+P, BY+P+12, C1W-P*2, 14, {size:11, bold:true, color:G});
 txt(sl, 'AI agents can own', C1X+P, BY+P+26, C1W-P*2, 14, {size:11, bold:true, color:G});

 // spair rows — each 68pt tall
 var spairs = [
 {bef:'5 d', bl:'Today · n=206', aft:'<3 h', al:'Ka-Minion', kpi:'Phase 1 gate KPI', sub:'Avg bug resolution · 40× faster · 206 issues · 2 KA teams', red:true},
 {bef:'22%', bl:'Zombie tickets', aft:'<8%', al:'Target', kpi:'Phase 2 KPI', sub:'Stale tickets (>30 days) across all active teams', red:true},
 {bef:'3/wk', bl:'Calls & Alerts', aft:'−30%', al:'Target', kpi:'Phase 2 KPI', sub:'Out-of-hours duty calls & alerts · Sentinels watch 24/7', red:true},
 {bef:'30–40%', bl:'Sprint on ops', aft:'+20%', al:'Capacity', kpi:'Freed', sub:'Sprint capacity reclaimed for feature & product delivery', red:true},
 ];
 spairs.forEach(function(sp, i) {
 var sy = BY + P + 44 + i * 72;
 box(sl, C1X+P, sy, C1W-P*2, 60, '#fff5f5');
 // Before
 txt(sl, sp.bef, C1X+P+2, sy+6, 56, 28, {size:22, bold:true, color:RED, align:'center'});
 txt(sl, sp.bl, C1X+P+2, sy+36, 56, 12, {size:7.5, bold:true, color:GY, align:'center'});
 // Arrow + KPI
 txt(sl, '→', C1X+P+62, sy+14, 18, 18, {size:14, color:GM, align:'center'});
 txt(sl, sp.kpi, C1X+P+62, sy+34, 68, 12, {size:6.5, bold:true, color:GY2, align:'center'});
 // After
 txt(sl, sp.aft, C1X+P+134,sy+6, 56, 28, {size:22, bold:true, color:G, align:'center'});
 txt(sl, sp.al, C1X+P+134,sy+36, 56, 12, {size:7.5, bold:true, color:GY, align:'center'});
 // Sub
 txt(sl, sp.sub, C1X+P, sy+50, C1W-P*2, 16, {size:7.5, color:GY2});
 });

 // Alert callout
 var alertY = BY + P + 44 + 4*72 + 2;
 box(sl, C1X+P, alertY, C1W-P*2, 30, '#fffbeb');
 txt(sl, ' Each quarter without Ka-Minions = ~€390K in recoverable capacity lost across the 30 KPI-target teams', C1X+P+6, alertY+6, C1W-P*2-12, 22, {size:8, color:'#78350f'});

 // COLUMN 2 — Financial Impact + charts 
 txt(sl, 'Financial Impact · Q4 2026 → 2028+', C2X+P, BY+P, C2W-P*2, 10, {size:7.5, bold:true, color:'#609F28'});
 txt(sl, 'Same headcount. Radically more output.', C2X+P, BY+P+12, C2W-P*2, 14, {size:11, bold:true, color:G});

 // 3 headline tiles
 var ftW2 = (C2W - P*2 - 8) / 3;
 var ftdata = [
 {l:'Savings at KPI target', v:'€1.99M/yr', s:'30 teams · incl. on-call · Q3 2028', lime:true},
 {l:'Engineering capacity freed', v:'~25 eng', s:'capacity + on-call reduction', lime:true},
 {l:'ROI at Q10 · 2028', v:'~7.9×', s:'€2.87M vs €361K'},
 ];
 ftdata.forEach(function(ft, i) {
 var fx = C2X+P + i*(ftW2+4);
 var fy = BY+P+28;
 box(sl, fx, fy, ftW2, 50, ft.lime?'#f0fae6':S2, ft.lime?LIME:BDR, ft.lime?1.5:1);
 txt(sl, ft.l, fx+5, fy+4, ftW2-10, 10, {size:6.5, bold:true, color:GY});
 txt(sl, ft.v, fx+5, fy+14, ftW2-10, 22, {size:16, bold:true, color:G});
 txt(sl, ft.s, fx+5, fy+38, ftW2-10, 10, {size:7, color:GY2});
 });

 // Chart 1: Savings ramp stacked bars 
 var ch1Y = BY+P+86;
 txt(sl, 'Cumulative annual savings ramp — by phase', C2X+P, ch1Y, C2W-P*2, 10, {size:7, bold:true, color:GY});
 // 5 snapshots: Q4'26, Q1'27, Q3'27, Q3'28, 2028+
 // Ph1=103K, Ph2=260K, Ph3=777K, Ph4=417K (stacked layers in each bar)
 var sMax = 2100; // €K max
 var barAreaH1 = 80, barAreaY1 = ch1Y+12;
 var bars1 = [
 {label:"Q4'26", ph:[103.7,   0,      0,      0    ]},
 {label:"Q1'27", ph:[103.7, 360.1,    0,      0    ]},
 {label:"Q3'27", ph:[103.7, 360.1,  993.9,    0    ]},
 {label:"Q3'28", ph:[103.7, 360.1,  993.9,  531.3  ]},
 {label:"2028+", ph:[103.7, 360.1,  993.9,  860.1  ]},
 ];
 var phCols1 = [G, PUR, AMB, TEAL];
 var bW1 = 44, bGap1 = 10;
 var chartLeft1 = C2X + P + 28;
 bars1.forEach(function(b, bi) {
 var bx = chartLeft1 + bi*(bW1+bGap1);
 var stackBottom = barAreaY1 + barAreaH1;
 b.ph.forEach(function(v, pi) {
 if (v === 0) return;
 var bH = Math.round(v / sMax * barAreaH1);
 stackBottom -= bH;
 box(sl, bx, stackBottom, bW1, bH, phCols1[pi]);
 });
 txt(sl, b.label, bx, barAreaY1+barAreaH1+2, bW1, 10, {size:7, color:G, bold:true, align:'center'});
 var total = b.ph.reduce(function(a,c){return a+c;},0);
 txt(sl, '€'+Math.round(total/100)/10+'M', bx, barAreaY1+barAreaH1+10, bW1, 10, {size:6.5, color:GY2, align:'center'});
 });
 // Y axis labels
 [0,500,1000,1500,2000].forEach(function(v) {
 var yy = barAreaY1 + barAreaH1 - Math.round(v/sMax*barAreaH1);
 box(sl, chartLeft1-2, yy, chartLeft1-C2X-P-2, 0.5, '#e5f0d2');
 txt(sl, v===0?'€0':'€'+v/1000+'M', C2X+P, yy-4, 26, 10, {size:6, color:GY, align:'right'});
 });
 // Break-even marker on Q1'27 bar
 var beX = chartLeft1 + 1*(bW1+bGap1) + bW1/2;
 box(sl, beX-0.5, barAreaY1, 1.5, barAreaH1, AMB);
 box(sl, beX-58, barAreaY1-16, 56, 14, '#fff8e8', AMB, 0.5);
 txt(sl, 'Break-even Q2\'27', beX-57, barAreaY1-15, 54, 12, {size:7.5, bold:true, color:'#92400e'});

 // Phase legend
 var legY1 = barAreaY1 + barAreaH1 + 22;
 ['Ph1','Ph2','Ph3','Ph4'].forEach(function(lbl, i) {
 var lx = chartLeft1 + i*70;
 box(sl, lx, legY1, 8, 6, phCols1[i]);
 txt(sl, lbl, lx+10, legY1-2, 58, 10, {size:6.5, color:GY2});
 });

 // Chart 2: Investment vs Savings sparkline 
 var ch2Y = ch1Y + barAreaH1 + 44;
 txt(sl, 'Investment vs savings — 10 quarters', C2X+P, ch2Y, C2W-P*2, 10, {size:7, bold:true, color:GY});

 var savPts = [0, 25.9, 51.8, 167.8, 283.8, 648.2, 1012.6, 1377.0, 1874.2, 2371.4, 2868.6];
 var invPts = [0, 87.0, 94.0, 119.0, 129.0, 178.0,  197.0,  216.0,  235.0,  323.0,  361.0];
 var ch2Max = 2200;
 var ch2H = 68, ch2Y2 = ch2Y+12;
 var ch2Left = chartLeft1;
 var ch2W2 = C2W - P*2 - 28;
 var xStep = ch2W2 / 10;

 // Use insertLine() for proper diagonal lines
 function line(slide, x1, y1, x2, y2, colorHex, weight) {
 var ln = slide.insertLine(SlidesApp.LineCategory.STRAIGHT,
 px(x1), py(y1), px(x2), py(y2));
 ln.getLineFill().setSolidFill(colorHex);
 ln.setWeight(weight || 1.5);
 return ln;
 }

 // Savings line segments (green, solid, weight 2)
 for (var qi = 0; qi < 10; qi++) {
 line(sl,
 ch2Left + qi*xStep, ch2Y2 + ch2H - savPts[qi] /ch2Max*ch2H,
 ch2Left + (qi+1)*xStep, ch2Y2 + ch2H - savPts[qi+1]/ch2Max*ch2H,
 G, 2
 );
 }
 // Investment line segments (red, dashed simulation via alternating draw/skip)
 for (var qi2 = 0; qi2 < 10; qi2++) {
 line(sl,
 ch2Left + qi2*xStep, ch2Y2 + ch2H - invPts[qi2] /ch2Max*ch2H,
 ch2Left + (qi2+1)*xStep, ch2Y2 + ch2H - invPts[qi2+1]/ch2Max*ch2H,
 RED, 1.5
 );
 }
 // Baseline
 box(sl, ch2Left, ch2Y2+ch2H, ch2W2, 0.8, BDR);
 // End labels
 txt(sl, '>€2.1M', ch2Left+ch2W2+2, ch2Y2, 36, 10, {size:6.5, bold:true, color:G});
 txt(sl, '€361K', ch2Left+ch2W2+2, ch2Y2+ch2H-6, 36, 10, {size:6.5, bold:true, color:RED});
 // Break-even dot + label
 var beX2 = ch2Left + 2.5*xStep;
 var beY2 = ch2Y2 + ch2H - 90.0/ch2Max*ch2H;
 box(sl, beX2-5, beY2-5, 10, 10, AMB);
 box(sl, beX2+8, beY2-12, 82, 14, '#fff8e8', AMB, 0.5);
 txt(sl, 'Break-even Q2\'27', beX2+9, beY2-11, 80, 12, {size:8, bold:true, color:'#92400e'});
 // Legend
 var legY2 = ch2Y2 + ch2H + 4;
 box(sl, ch2Left, legY2, 12, 3, G);
 txt(sl, 'Cumulative savings', ch2Left+14, legY2-2, 90, 10, {size:6.5, color:GY2});
 box(sl, ch2Left+110, legY2, 12, 3, RED);
 txt(sl, 'Cumulative investment', ch2Left+124, legY2-2, 90, 10, {size:6.5, color:RED});

 // Phase-gated callout
 var gateY = ch2Y2 + ch2H + 18;
 box(sl, C2X+P, gateY, C2W-P*2, 20, GL);
 txt(sl, ' Phase-gated: no spend unlocks without hitting prior KPI. Single ask today: €60–80K Phase 1 · Q4 2026.', C2X+P+6, gateY+4, C2W-P*2-12, 14, {size:7.5, bold:true, color:'#14532d'});

 // COLUMN 3 — AI Stack costs & amortisation 
 txt(sl, 'AI STACK — COSTS & AMORTISATION', C3X+P, BY+P, C3W-P*2, 10, {size:7.5, bold:true, color:'#609F28'});
 txt(sl, '€76K runs €1.56M/yr —', C3X+P, BY+P+12, C3W-P*2, 14, {size:11, bold:true, color:G});
 txt(sl, 'cost ~5% of savings at full scale', C3X+P, BY+P+26, C3W-P*2, 14, {size:11, bold:true, color:G});

 // 2 stack tiles
 var stW = (C3W - P*2 - 6) / 2;
 var stiles = [
 {l:'Paperclip platform', v:'€12–24K/yr', s:'€12K Ph1 → €24K Ph4 · governance + audit'},
 {l:'Claude API · ~€1/task', v:'€3.5–52K/yr', s:'3,500 tasks Ph1 → 52,000 tasks Ph4'},
 ];
 stiles.forEach(function(st, i) {
 var sx = C3X+P + i*(stW+6), sy = BY+P+44;
 box(sl, sx, sy, stW, 48, S2, BDR, 1);
 txt(sl, st.l, sx+6, sy+5, stW-12, 10, {size:6.5, bold:true, color:GY});
 txt(sl, st.v, sx+6, sy+15, stW-12, 18, {size:14, bold:true, color:G});
 txt(sl, st.s, sx+6, sy+34, stW-12, 14, {size:7, color:GY2});
 });

 // Amortisation table
 txt(sl, 'COST vs SAVINGS — PHASE BY PHASE', C3X+P, BY+P+100, C3W-P*2, 10, {size:7.5, bold:true, color:GY});
 // Header row
 box(sl, C3X+P, BY+P+112, C3W-P*2, 14, G);
 var thX = [C3X+P+3, C3X+P+70, C3X+P+130, C3X+P+185, C3X+P+228];
 ['Phase','AI stack/yr','Savings/yr','Stack %','Net/yr'].forEach(function(h, i) {
 txt(sl, h, thX[i], BY+P+114, 60, 10, {size:6.5, bold:true, color:WH, align: i>0 ? 'right' : 'left'});
 });

 var amorRows = [
 {ph:'Ph1 · 2t · Q4\'26', pc:G,    ai:'~€15.5K', sav:'€103K',  pct:'15%*', net:'~€88K',   bg:WH},
 {ph:'Ph2 · 7t +on-call', pc:PUR,  ai:'~€21K',   sav:'€464K',  pct:'4.5%', net:'~€443K',  bg:'#f6faf0'},
 {ph:'Ph3 · 22t +on-call',pc:AMB,  ai:'~€37K',   sav:'€1.46M', pct:'2.5%', net:'~€1.42M', bg:WH},
 {ph:'Ph4 · 30t +on-call',pc:TEAL, ai:'~€76K',   sav:'€1.99M', pct:'3.8%', net:'>€1.91M', bg:GL},
 ];
 amorRows.forEach(function(r, i) {
 var ry = BY+P+126 + i*18;
 box(sl, C3X+P, ry, C3W-P*2, 18, r.bg);
 txt(sl, r.ph, thX[0], ry+3, 66, 12, {size:7.5, bold:true, color:r.pc});
 txt(sl, r.ai, thX[1], ry+3, 54, 12, {size:7.5, color:BK, align:'right'});
 txt(sl, r.sav, thX[2], ry+3, 54, 12, {size:7.5, bold:true, color:G, align:'right'});
 txt(sl, r.pct, thX[3], ry+3, 42, 12, {size:7.5, color:i===0?AMB:GM, align:'right'});
 txt(sl, r.net, thX[4], ry+3, 52, 12, {size:7.5, bold:true, color:G, align:'right'});
 });
 txt(sl, '* Ph1 % includes €60–80K one-time setup; ongoing run rate drops to 3.4%', C3X+P, BY+P+200, C3W-P*2, 10, {size:6.5, color:GY});

 // Amortisation callout
 box(sl, C3X+P, BY+P+212, C3W-P*2, 46, GL, BDR, 0.5);
 txt(sl, ' Setup cost fully amortised by Q2 2027.', C3X+P+6, BY+P+215, C3W-P*2-12, 12, {size:8.5, bold:true, color:'#14532d'});
 txt(sl, '€60–80K one-time Q4 2026 recovered in <2 quarters. From Q3 2027 AI stack cost stays flat at 3–5% of savings regardless of team count.', C3X+P+6, BY+P+228, C3W-P*2-12, 28, {size:8, color:'#14532d'});

 // 10x price resilience callout
 box(sl, C3X+P, BY+P+262, C3W-P*2, 34, '#fffbeb');
 txt(sl, ' Ratio holds even at 10× token price inflation —', C3X+P+6, BY+P+265, C3W-P*2-12, 12, {size:8.5, bold:true, color:'#78350f'});
 txt(sl, '€1/task → €10/task still leaves net savings >€1M/yr at Phase 4 KPI. Cost scales linearly; savings scale with headcount.', C3X+P+6, BY+P+278, C3W-P*2-12, 18, {size:8, color:'#78350f'});

 // FOOTER (imp-footer) 
 var FY = SH - 38;
 box(sl, 0, FY, SW, 38, S2, BDR, 0.5);
 txt(sl, 'Rollout →', 8, FY+10, 48, 16, {size:7.5, bold:true, color:GY});
 var iphases = [
 {t:'Phase 1 — Hypothesis · Q4 2026', s:'TRX · 2 teams · €103K/yr · savings start immediately', k:'Gate: 3d → <3h · 1 quarter', bc:G, tc:'#1a2e0d', kc:'#609F28'},
 {t:'Phase 2 — MVPs · Q1–Q2 2027', s:'7 teams · €363K/yr · break-even crossed Q2 2027', k:'Gate: 70% success rate · 2 quarters', bc:PUR, tc:'#5b21b6', kc:'#609F28'},
 {t:'Phase 3 — Growth · Q3 2027–Q2 2028', s:'22 teams · €1.14M/yr · Sentinels nights & weekends', k:'Gate: 40% adoption · 4 quarters', bc:AMB, tc:'#92400e', kc:'#609F28'},
 {t:'Phase 4 — Scale · Q3 2028 → full scale', s:'30 teams (70% KPI) · €1.56M/yr · >€2.1M net in 2028', k:' Top KPI: 70% teams with active agents', bc:TEAL, tc:'#0e7490', kc:'#609F28'},
 ];
 var iphW = (SW - 60) / 4;
 iphases.forEach(function(ip, i) {
 var ipx = 58 + i * iphW;
 box(sl, ipx, FY, 3, 38, ip.bc);
 txt(sl, ip.t, ipx+5, FY+2, iphW-10, 12, {size:7.5, bold:true, color:ip.tc});
 txt(sl, ip.s, ipx+5, FY+14, iphW-10, 12, {size:7, color:GY});
 txt(sl, ip.k, ipx+5, FY+26, iphW-10, 10, {size:6.5, bold:true, color:ip.kc});
 // Arrow between phases
 if (i < 3) txt(sl, '›', ipx+iphW-6, FY+12, 12, 14, {size:11, color:'#D6EDAA', align:'center'});
 });
 txt(sl, '5 / 7', SW-40, FY-10, 36, 10, {size:8, color:GY});
}


// 
// SLIDE 6 — Path to Implementation
// 
function buildSlide6(pres) {
 var sl = addSlide(pres, 5);
 clearSlide(sl);

 box(sl, 0, 0, SW, SH, OFFWH);
 txt(sl, 'IDEAS2IMPACT 2026 · ka-i2i-autonomous-minions-agents', 28, 6, SW-56, 14, {size:8, bold:true, color:GM});
 txt(sl, 'Ka-Minions path to implementation', 28, 20, SW-56, 40, {size:26, bold:true, color:G});

 // NOW/NEXT banner
 box(sl, 28, 68, SW/2-36, 32, PUR);
 box(sl, SW/2+4, 68, SW/2-36, 32, GM);
 txt(sl, 'NOW: Training & learning', 34, 73, SW/2-44, 22, {size:12, bold:true, color:WH});
 txt(sl, 'NEXT: Drive business value delivery', SW/2+10, 73, SW/2-44, 22, {size:12, bold:true, color:WH});

 // 4 phase cards
 var cards = [
 {ph:'PHASE 1\nHypothesis validation', period:'Q4 2026 · 1Q',
 title:'Automate repetitive\noperations',
 desc:'Pilot teams (1–2) creating and\ntraining their own Ka-Minions',
 saving:'€103K/yr · 2 teams',
 kpi:'KPI: Bug resolution 5 days → <3 hours (40×) · n=206', c:G, chipBg:'#e8f5e1'},
 {ph:'PHASE 2\nMVPs creation', period:'Q1–Q2 2027 · 2Q',
 title:'Boosting KA Impact:\nfuel teams with AI Agents',
 desc:'Five selected teams implementing\nAI agents for product delivery',
 saving:'€464K/yr · 7 teams · on-call −30% · Break-even Q2 2027',
 kpi:'KPI: 70% success rate · on-call reduction gate', c:PUR, chipBg:'#f3eeff'},
 {ph:'PHASE 3\nGrowth & Iterations', period:'Q3 2027–Q2 2028 · 4Q',
 title:'AI-Agents\nAcceleration',
 desc:'Driving efficiency by integrating\nAI agents within existing headcounts',
 saving:'€1.46M/yr · 22 teams · incl. on-call savings',
 kpi:'KPI: 40% Teams Ka-Minions adoption', c:AMB, chipBg:'#fff8e1'},
 {ph:'PHASE 4\nScale Up', period:'Q3 2028 → Always on',
 title:'AI agentic teams',
 desc:'Building Features and Services\nwith AI Agents as Team Members',
 saving:'€1.99M/yr · 30 teams · KPI target · incl. on-call',
 kpi:'KPI: 70% adoption + 30% feature delivery to production', c:TEAL, chipBg:'#e0f7fa'},
 ];

 cards.forEach(function(card, i) {
 var cx = 8 + i*238, cy = 106, cw = 234, ch = 340;
 box(sl, cx, cy, cw, ch, WH, BDR, 1);
 box(sl, cx, cy, cw, 3.5, card.c);
 // phase header bg
 box(sl, cx, cy, cw, 36, card.chipBg || GL);
 txt(sl, card.ph, cx+8, cy+4, cw-16, 22, {size:8, bold:true, color:card.c});
 txt(sl, card.period,cx+8, cy+26, cw-16, 14, {size:7, color:GY2});
 txt(sl, card.title, cx+8, cy+44, cw-16, 30, {size:11, bold:true, color:card.c});
 txt(sl, card.desc, cx+8, cy+76, cw-16, 28, {size:8.5,italic:true, color:GY2});
 // saving chip
 box(sl, cx+6, cy+108, cw-12, 26, card.chipBg || GL, card.c, 1);
 txt(sl, card.saving, cx+10, cy+113, cw-20, 16, {size:8.5, bold:true, color:card.c});
 // KPI box
 box(sl, cx+6, cy+138, cw-12, 22, GL);
 txt(sl, card.kpi, cx+10, cy+141, cw-20, 18, {size:7.5, bold:true, color:G});
 // Phase 4 extra KPI pills
 if (i === 3) {
 box(sl, cx+6, cy+164, cw-12, 16, BK);
 txt(sl, '70% adoption (30/43 teams)', cx+10, cy+167, cw-20, 12, {size:7, bold:true, color:WH});
 box(sl, cx+6, cy+182, cw-12, 16, BK);
 txt(sl, '30% feature delivery to production', cx+10, cy+185, cw-20, 12, {size:7, bold:true, color:WH});
 }
 // Duration dot
 txt(sl, i===0?'1Q':i===1?'2 Qs':i===2?'4 Qs':'Always on', cx+cw/2-20, cy+ch-36, 40, 30, {size:i===3?11:16, bold:true, color:card.c, align:'center'});
 });

 // ROI footer bar
 roiFooter(sl, [
 {label:'Total investment', value:'€361K', sub:'Q4 2026 → Q10'},
 {label:'Break-even', value:'Q2 2027', sub:'savings overtake cost'},
 {label:'Savings at KPI', value:'€1.99M/yr', sub:'30 teams · incl. on-call'},
 {label:'ROI at Q10', value:'~7.9×', sub:'€2.87M vs €361K invested'},
 {label:'Net value 2028+', value:'>€2.85M', sub:'full potential 43 teams'},
 ]);
 txt(sl, '6 / 7', SW-40, SH-52, 36, 12, {size:9, color:GY});
}


// 
// SLIDE 7 — Financial Charts (Investment vs Savings + Savings Ramp)
// 
function buildSlide7(pres) {
 var sl = addSlide(pres, 6);
 clearSlide(sl);

 hdrBand(sl,
 'Financial Model — 10-Quarter Cumulative & Savings Ramp by Phase',
 'Ka-Minions · Business Case · Ideas2Impact 2026',
 'Break-even Q2 2027 · ROI ~6× at Q10 · Total investment €361K vs €2.15M saved'
 );

 var bodyY = 84;

 // CHART 1: Cumulative Investment vs Savings (bar chart using boxes) 
 txt(sl, 'CUMULATIVE SAVINGS vs INVESTMENT — 10 QUARTERS (€K)', 8, bodyY, 700, 12, {size:7.5, bold:true, color:GY});

 var CUM_SAV = [25.9, 51.8, 167.8, 283.8, 648.2, 1012.6, 1377.0, 1874.2, 2371.4, 2868.6];
 var CUM_INV = [87.0, 94.0, 119.0, 129.0, 178.0,  197.0,  216.0,  235.0,  323.0,  361.0];
 var QLABELS = ["Q4'26","Q1'27","Q2'27","Q3'27","Q4'27","Q1'28","Q2'28","Q3'28","Q4'28","Q1'29"];
 var maxVal = 2200; // max for scaling
 var chartX = 8, chartY = bodyY + 14, chartW = 700, chartH = 160;
 var colW = 64, gap = 4;
 var barAreaH = chartH - 20; // leave 20pt for labels below

 // Axes
 box(sl, chartX, chartY, chartW, chartH, '#f9fdf5', BDR, 0.5);

 // Y-axis gridlines + labels
 [0, 500, 1000, 1500, 2000].forEach(function(v) {
 var gy = chartY + barAreaH - Math.round(v / maxVal * barAreaH);
 box(sl, chartX + 30, gy, chartW - 38, 0.5, '#d6edaa');
 txt(sl, '€'+v+'K', chartX+2, gy-4, 28, 10, {size:6, color:GY, align:'right'});
 });

 // Bars
 QLABELS.forEach(function(ql, i) {
 var bx = chartX + 32 + i * (colW + gap);
 var savH = Math.round(CUM_SAV[i] / maxVal * barAreaH);
 var invH = Math.round(CUM_INV[i] / maxVal * barAreaH);
 var savY = chartY + barAreaH - savH;
 var invY = chartY + barAreaH - invH;

 // Savings bar (dark green, thicker)
 box(sl, bx, savY, 28, savH, G);
 // Investment bar (red, thinner, offset)
 box(sl, bx + 30, invY, 20, invH, RED);
 // Q label
 txt(sl, ql, bx, chartY + barAreaH + 2, 52, 10, {size:6.5, color:GY2, align:'center'});
 // Break-even marker
 if (i === 2) {
 box(sl, bx - 2, chartY, 2, barAreaH, AMB);
 txt(sl, 'BE', bx - 14, chartY + 2, 30, 10, {size:6, bold:true, color:AMB});
 }
 });

 // Legend
 box(sl, chartX + 32, chartY + barAreaH + 14, 12, 8, G);
 txt(sl, 'Cum. Savings', chartX + 46, chartY + barAreaH + 12, 80, 10, {size:7, color:G});
 box(sl, chartX + 140, chartY + barAreaH + 14, 12, 8, RED);
 txt(sl, 'Cum. Investment', chartX + 154, chartY + barAreaH + 12, 90, 10, {size:7, color:RED});
 txt(sl, 'BE = Break-even Q2 2027', chartX + 260, chartY + barAreaH + 12, 120, 10, {size:7, color:AMB});

 // CHART 2: Annual Savings Ramp — Stacked 
 var c2Y = bodyY + 198;
 txt(sl, 'ANNUAL SAVINGS RAMP — STACKED BY PHASE (€K/yr)', 8, c2Y, 700, 12, {size:7.5, bold:true, color:GY});

 var snapshots = ["Q4'26\nPh1 live", "Q1'27\nPh2 live", "Q3'27\nPh3 live", "Q3'28\nPh4 KPI", "2028+\nfull scale"];
 var phColors2 = [G, PUR, AMB, TEAL];
 // Each row: [Ph1, Ph2, Ph3, Ph4] stacked
 var rampStacks = [
 [103.7,   0,       0,       0     ],
 [103.7, 463.8,     0,       0     ],
 [103.7, 463.8,  1457.7,     0     ],
 [103.7, 463.8,  1457.7,  1989.0  ],
 [103.7, 463.8,  1457.7,  2849.1  ],
 ];
 var maxRamp = 3800;
 var c2BarH = 120, c2X = 8;
 var barH2 = c2BarH - 20;
 var snW = 132, snGap = 6;

 box(sl, c2X, c2Y + 14, SW - 16, c2BarH, '#f9fdf5', BDR, 0.5);

 // Gridlines
 [0, 1000, 2000, 3000].forEach(function(v) {
 var gy = c2Y + 14 + barH2 - Math.round(v / maxRamp * barH2);
 box(sl, c2X + 52, gy, SW - 68, 0.5, '#d6edaa');
 txt(sl, '€'+v+'K', c2X+2, gy-4, 48, 10, {size:6, color:GY, align:'right'});
 });

 snapshots.forEach(function(sn, i) {
 var bx = c2X + 54 + i * (snW + snGap);
 var stack = rampStacks[i];
 var prevH = 0;
 stack.forEach(function(val, pi) {
 if (val === 0) return;
 var bH = Math.round(val / maxRamp * barH2);
 var bY = c2Y + 14 + barH2 - prevH - bH;
 box(sl, bx, bY, snW, bH, phColors2[pi]);
 if (bH > 10) txt(sl, '€'+Math.round(val)+'K', bx+2, bY+2, snW-4, bH-2, {size:6.5, bold:true, color:WH, align:'center'});
 prevH += bH;
 });
 // Label
 var snLabel = sn.replace('\n', ' ');
 txt(sl, snLabel, bx, c2Y + 14 + barH2 + 2, snW, 16, {size:7, color:GY2, align:'center'});
 });

 // Legend row
 var legY = c2Y + 14 + c2BarH + 2;
 ['Phase 1 (TRX · 2t)', 'Phase 2 (7t)', 'Phase 3 (22t)', 'Phase 4 KPI (30t)'].forEach(function(lbl, i) {
 var lx = 8 + i * 180;
 box(sl, lx, legY, 12, 8, phColors2[i]);
 txt(sl, lbl, lx+14, legY-2, 164, 12, {size:7, color:GY2});
 });

 roiFooter(sl, [
 {label:'Total investment', value:'€361K', sub:'Q4 2026 → Q10'},
 {label:'Break-even', value:'Q2 2027', sub:'savings overtake cost'},
 {label:'Savings at KPI', value:'€1.99M/yr', sub:'30 teams · incl. on-call'},
 {label:'ROI at Q10', value:'~7.9×', sub:'€2.87M vs €361K invested'},
 {label:'Net value 2028+', value:'>€2.85M', sub:'full potential 43 teams'},
 ]);
 txt(sl, '7 / 7', SW-40, SH-52, 36, 12, {size:9, color:GY});
}


// 
// MAIN
// 
function buildKaMinionsSlides() {
 var pres = newPres();

 buildSlide1(pres);
 buildSlide2(pres);
 buildSlide3(pres);
 buildSlide4(pres);
 buildSlide5(pres);
 buildSlide6(pres);
 buildSlide7(pres);

 // Remove default blank first slide if extra
 var slides = pres.getSlides();
 if (slides.length > 7) slides[slides.length-1].remove();

 var url = pres.getUrl();
 Logger.log(' Presentación creada: ' + url);
}
