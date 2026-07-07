# Ka-Minions — Autonomous AI Dev Agents
### Ideas2Impact 2026 · Banana Squad · Kleinanzeigen / Adevinta

Business case for deploying autonomous AI agents across all 43 KA engineering teams — freeing **~25 FTE equivalents** and recovering **€1.99M/yr** in engineering capacity at 70% adoption (capacity + on-call savings).

---

## Live business case

**[View the full business case deck](https://elenacastan-maker.github.io/ka-minions-i2i-2026/)**

6-slide C-level deck covering:
- The problem & 3 value pillars (5d → <3h bug resolution · n=206 issues · 2 KA teams)
- ROI & domain expansion (30/43 teams · 70% adoption KPI · incl. on-call savings)
- Costs, ROI & governance (€361K total investment · ~7.9× ROI at Q10)
- Enabler team & AI stack (Paperclip + Claude API · cost ~4% of savings)
- Impact at a glance (financial model · savings ramp · investment vs savings)
- Path to implementation (4-phase rollout · Q4 2026 → Q3 2028)

---

## Key numbers

| Metric | Value |
|---|---|
| KPI target | 30 / 43 teams (70%) |
| Annual savings at KPI | €1,989,000 / yr (incl. on-call savings) |
| On-call premium saving | 30% reduction · €8K/eng/yr · from Phase 2 gate |
| FTE capacity freed | ~25 engineers (capacity + on-call) |
| Break-even | Q2 2027 |
| ROI at Q10 | ~7.9× |
| Total investment (10Q) | €361,000 |
| AI stack cost vs savings | ~4% |
| Phase 1 single ask | €60–80K setup · Q4 2026 |
| Avg bug resolution (measured) | 5 days → <3 hours · 40× faster · n=206 issues |

---

## Repository structure

```
index.html                      Business case deck (6 slides, GitHub Pages)
build_business_case_pptx.py     Generate PPTX from scratch (python-pptx)
build_business_case_deck.py     Build full analysis deck (Python)
build_clevel_deck.py            C-level one-pager deck (Python)
push_business_case_sheets.py    Push financial model to Google Sheets (Python)
ka_minions_slides.gs            Google Slides builder (Apps Script) — 7 slides
ka_minions_single_sheet.gs      Google Sheets single-tab financial model (Apps Script)
ka_minions_business_case.gs     Google Sheets multi-tab business case (Apps Script)
```

---

## Google Apps Script usage

### Slides (ka_minions_slides.gs)
1. Go to [script.google.com](https://script.google.com) or open any Google Sheet → Extensions → Apps Script
2. Paste the full contents of `ka_minions_slides.gs`
3. Run → `buildKaMinionsSlides`
4. A new presentation is created in your Drive with 7 slides

### Sheets (ka_minions_single_sheet.gs)
1. Open any Google Sheet → Extensions → Apps Script
2. Paste the full contents of `ka_minions_single_sheet.gs`
3. Run → `buildKaMinions`
4. A tab named "Ka-Minions Business Case" is created with the full financial model and charts

---

## Python scripts

### Requirements
```bash
pip install python-pptx
```

### Generate PPTX
```bash
python3 build_business_case_pptx.py
# → ka-minions-business-case.pptx
```

---

## Financial model — 4-phase rollout

| Phase | Period | Teams | Annual savings | AI stack/yr | Net/yr |
|---|---|---|---|---|---|
| Phase 1 — Hypothesis | Q4 2026 | 2 | €103,718 | ~€15.5K | ~€88K |
| Phase 2 — MVPs | Q1–Q2 2027 | 7 | €463,813 (+on-call) | ~€21K | ~€443K |
| Phase 3 — Growth | Q3 2027–Q2 2028 | 22 | €1,457,698 (+on-call) | ~€37K | ~€1.42M |
| Phase 4 — Scale (KPI) | Q3 2028+ | 30 | €1,989,000 (+on-call) | ~€76K | >€1.91M |

Break-even Q2 2027 — cumulative savings overtake total investment after 3 quarters. On-call savings (30% reduction · €8K/eng/yr premium) activate at Phase 2 gate (Sentinel mode).

---

## Assumptions

- Avg engineer cost (fully loaded): €80,000/yr
- Capacity freed per engineer: 25% = €20,000/yr
- Claude API: ~€1/task · 1,750 tasks/team/yr
- Paperclip platform: ~€12,000/yr (team plan)
- Bug resolution baseline: **5 days** (measured across 206 issues in 2 KA teams, July 2026)
- Ka-Minion target: <3 hours = **40× faster**

---

## Contact

Elena Castán — Product Manager, KA PRO · Kleinanzeigen / Adevinta
Ideas2Impact 2026 · Banana Squad
