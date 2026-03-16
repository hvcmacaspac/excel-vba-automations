# ⚙️ Weekly Sourcing Report Automation

> Built with **Microsoft Excel + VBA** | Covers: SLA Tracking, ATS Audit, Training Score Monitoring, Week-over-Week Reporting, and Template Reset

---

## 🗂️ Overview

This automation was built to eliminate manual, repetitive weekly reporting tasks for a recruitment operations team. What previously took **1–2 hours of manual work** per report cycle was reduced to a **single macro run** — freeing up the team to focus on actual sourcing work instead of data wrangling.

The tool was designed to be reusable week-over-week, with a built-in reset macro that clears all data and restores the template to its original state for the next reporting cycle.

> 📝 **Context:** Originally built during a two-week senior coverage period as a People Ops initiative. The template was designed for department-wide adoption, with a dedicated instruction sheet for non-technical users.

---

## 🚀 How to Use

1. Paste all source reports into their corresponding sheets (see instruction sheet)
2. Run **`Run_All_Updates`** to process all data in one click
3. Perform manual QIA audit on the QIA Candidates sheet
4. Run **`UpdateWeekXX`** to populate the weekly report sheet
5. Rename the Week XX sheet to the current week (e.g., `SEP 24-OCT 1`)
6. Save a copy as `.xlsx` for archiving
7. Run **`Clear_Template`** to reset for next week's report

---

## ⚙️ Macro Reference

### `Run_All_Updates` — Master Macro
Runs all update sub-macros in sequence with a single click:

| Order | Macro | What It Does |
|---|---|---|
| 1 | `CleanPrevTSSheet` | Clears previous week's training score columns to prep for fresh data |
| 2 | `UpdateTrainingScoreSheet` | Adds SLA flags, pulls previous training scores via XLOOKUP, flags score changes |
| 3 | `UpdateQIACandidatesSheet` | Audits candidate records against ATS data via VLOOKUP (by name and email) |
| 4 | `UpdateKATSCandidatesSheet` | Applies SLA tracking logic for K-ATS pipeline candidates |
| 5 | `UpdateAndCopyAllTasks` | Cleans task timestamps, applies SLA formulas, auto-copies completed tasks to separate sheet |

---

### `UpdateWeekXX` — Weekly Report Populator
Automatically calculates and populates the weekly performance report sheet:

- Computes QIA Audit pass rate, Task SLA compliance, and Training Score metrics
- Populates current week's data in **Column E** (Column D = previous week for comparison)
- Adds contextual notes (e.g., *"X/Y tasks completed within SLA"*)
- Applies traffic light color coding per metric

---

### `Clear_Template` — Weekly Reset
Clears all pasted data across all sheets and restores the template to its original state:

- Deletes data from all source sheets
- Clears and resets the weekly report sheet
- Renames the active week sheet back to **"Week XX"**
- Displays a confirmation prompt when complete

---

## 📊 SLA Traffic Light System

Performance metrics are automatically color-coded based on thresholds:

| Color | Threshold | Status |
|---|---|---|
| 🟢 Green | ≥ 85% | On Track |
| 🟡 Yellow | 76% – 84% | Needs Attention |
| 🔴 Red | < 75% | At Risk |

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
├── WeeklyReport_README.md           ← You are here
├── weekly_sourcing_automation.bas   ← Full exported VBA module
├── clear_automation_template.bas    ← Clear template module
└── instructions_preview.png         ← Screenshot of the instruction sheet
```

> 🔒 **Data Privacy:** All sheet references have been anonymized. Internal system names have been replaced with generic identifiers (ATS, K-ATS). No candidate, client, or company data is included in this repository.

---

## 🙋 About This Project

Built independently during a two-week senior coverage period. Rather than just maintaining existing workflows, the entire reporting process was automated and handed off as a reusable, self-documented template — designed so that any team member, technical or not, could run it independently.

Built under pressure. Handed off anyway. 💪

---

*Feel free to connect on linkedin.com/in/hvcmacaspac/ for questions about the methodology!*
