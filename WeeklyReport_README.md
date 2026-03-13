# ⚙️ Weekly Sourcing Report Automation

> Built with **Microsoft Excel + VBA** | Covers: SLA Tracking, ATS Audit, Training Score Monitoring, Week-over-Week Reporting

---

## 🗂️ Overview

This automation was built to eliminate manual, repetitive weekly reporting tasks for a recruitment operations team. What previously took **1–2 hours of manual work** per report cycle was reduced to a **single macro run** — freeing up the team to focus on actual sourcing work instead of data wrangling.

> 📝 **Context:** This was built during a two-week coverage period for a Senior team member. Rather than just maintaining the status quo, the workflow was fully automated and handed off as a reusable template for future reporting cycles.

---

## ✨ What It Does

### `Run_All_Updates()` — One-Click Master Macro
Runs all sub-macros in sequence with a single click:

1. **CleanPrevTSSheet** — Clears previous week's training score columns to prep for fresh data
2. **UpdateTrainingScoreSheet** — Adds SLA flags, pulls previous training scores via XLOOKUP, and flags score changes
3. **UpdateQIACandidatesSheet** — Audits candidate records against ATS data via VLOOKUP (by name and email)
4. **UpdateKATSCandidatesSheet** — Applies SLA tracking logic for K-ATS pipeline candidates
5. **UpdateAndCopyAllTasks** — Cleans task timestamps, applies SLA formulas, and auto-copies completed tasks to a separate sheet

---

## 📊 Key Features

### SLA Traffic Light System
Automatically color-codes performance metrics based on thresholds:

| Color | Threshold | Meaning |
|---|---|---|
| 🟢 Green | ≥ 85% | On Track |
| 🟡 Yellow | 76% – 84% | Needs Attention |
| 🔴 Red | < 75% | At Risk |

### Week-over-Week Comparison (`UpdateWeekXX`)
- Automatically calculates current week's metrics across QIA Audit, Task SLA, and Training Score
- Populates a structured weekly report sheet
- Adds contextual notes (e.g., *"X/Y tasks completed within SLA"*)
- Previous week data sits in Column D, current week in Column E — instant comparison view

### Dual ATS Audit
Cross-references candidate records against two ATS systems using both **name** and **email** — catches duplicates and missing entries that manual checking would miss.

---

## 🛠️ Tools Used

- **Microsoft Excel** — report structure and formatting
- **VBA (Visual Basic for Applications)** — full automation
- **XLOOKUP + VLOOKUP** — cross-sheet data validation
- **FormulaR1C1** — dynamic SLA formula injection
- **TextToColumns** — automated date/time field parsing

---

## 📁 Repository Contents

```
/
├── README.md                        ← You are here
├── weekly_sourcing_automation.bas   ← Exported VBA module
└── dashboard_preview.png            ← Optional: screenshot of report output
```

> 🔒 **Data Privacy:** All sheet references have been anonymized. Internal system names have been replaced with generic identifiers (ATS, K-ATS). No candidate, client, or company data is included in this repository.

---

## ⚠️ How to Use

1. Open your weekly sourcing Excel report
2. Import the `.bas` module via **VBA Editor → File → Import**
3. Ensure your sheet names match the expected names in the code (see comments)
4. Run `Run_All_Updates()` from the macro menu
5. Double-check outputs — a confirmation prompt will appear when complete

---

## 🙋 About This Project

Built independently during a two-week senior coverage period as a People Ops initiative. The goal was simple: stop doing manually what a macro can do in seconds.

Built under pressure. Handed off anyway. 💪

---

*Feel free to connect on [LinkedIn](#) for questions about the methodology!*
