# 📋 Interview Dispatch Automation

> Built with **Microsoft Excel + VBA** | Covers: Interview Request Processing, Timezone Adjustment, TA Code Mapping, and Automated Dispatch Preparation

---

## 🗂️ Overview

Interview scheduling requests came in through a team channel with **50+ requests per day**, processed by 2–3 team members at fixed dispatch windows: **8 AM, 12 PM, and 3 PM** — with a final sweep before end of day at 4 PM to ensure no requests were missed before coordinators signed off.

Each dispatch window required manually copying and pasting every candidate detail from MS Teams into a tracker, then dividing and sending out assignments to coordinators — a process that was slow, repetitive, and highly prone to human error at high volumes.

A raw export file existed, but the data was unstructured and still required significant manual cleanup before it could be used.

This automation was built to solve that — transforming a messy raw export into a clean, dispatch-ready tracker in a single macro run.

> 📝 **Origin Story:** Learned VBA from scratch specifically to solve this problem. Built out of equal parts necessity and spite for manual work. 💪

---

## ❌ Before Automation

- **50+ requests per day** manually copied from MS Teams into a tracker
- Each field (name, phone, job title, interview type, etc.) copy-pasted one by one by 2–3 team members
- Process repeated at **8 AM, 12 PM, 3 PM, and 4 PM** daily
- High error risk at volume — missed fields, misread details, incorrect assignments
- Raw export file existed but required heavy manual cleanup before use

## ✅ After Automation

- Raw export pasted once into source sheet
- Single macro run populates all 27 columns instantly
- Formulas converted to static values automatically — no recalculation issues
- Dispatch-ready output in seconds, not minutes

---

## ⚙️ What the Macro Does

### `Interview_Requests` — Single-Run Dispatch Populator

Processes raw interview request data and populates the Dispatch sheet with:

| Column | Field | Method |
|---|---|---|
| A | Request ID | Direct copy |
| B | Created Date | INDEX/MATCH lookup |
| C | TA Code | IFS mapping by vertical |
| D | Date | DATEVALUE extraction |
| E | Adjusted Time | UTC → Local with DST handling |
| F | Full Name | CONCAT from first + last name |
| G | Recruiter | INDEX/MATCH lookup |
| H | Dispatch Group | VLOOKUP from reference table |
| K | Phone Number | INDEX/MATCH lookup |
| L | Candidate Type | INDEX/MATCH lookup |
| M | Req Number | INDEX/MATCH lookup |
| N | Job Title | INDEX/MATCH lookup |
| O | Interview Type | INDEX/MATCH lookup |
| P | Interview Channel | IFS logic (In-Person/Video/Phone) |
| Q | Hiring Manager | INDEX/MATCH lookup |
| R | Interview Round | INDEX/MATCH lookup |
| U | First Name | INDEX/MATCH lookup |
| V | Last Name | INDEX/MATCH lookup |
| W | Scheduling Type | INDEX/MATCH lookup |
| Y | Location | INDEX/MATCH lookup |
| Z | Member Org | INDEX/MATCH lookup |
| AA | Vertical | INDEX/MATCH lookup |

After all formulas are applied, the entire range is **converted to static values** — ensuring the output is stable and portable.

---

## 🌐 Timezone Handling

Automatically adjusts interview times from UTC to local time, accounting for Daylight Saving Time:

```vb
' UTC-4 during DST (Mar–Nov), UTC-5 outside DST
=RC[15]-IF(AND(MONTH(TODAY())>=3,MONTH(TODAY())<=11),TIME(4,0,0),TIME(5,0,0))
```

---

## 🛠️ Tools Used

- **Microsoft Excel** — tracker structure and output formatting
- **VBA (Visual Basic for Applications)** — full automation
- **INDEX/MATCH** — dynamic cross-sheet data lookup
- **IFS + VLOOKUP** — vertical and channel mapping logic
- **DATEVALUE + TIMEVALUE** — date and time parsing from raw strings
- **FormulaR1C1** — dynamic formula injection across variable row ranges

---

## 📁 Repository Contents

```
/
├── InterviewDispatchREADME.md     ← You are here
└── interview_dispatch.bas         ← Exported VBA module
```

> 🔒 **Data Privacy:** Variable names and business logic identifiers have been generalized for portfolio use. No candidate, client, or company data is included in this repository.

---

## 🙋 About This Project

Built independently after identifying a high-frequency, error-prone manual process. Rather than accepting the inefficiency, the workflow was automated from scratch — learning VBA specifically for this purpose and using it to eliminate hours of repetitive work per week.

Sometimes the best projects are born out of spite.

---

*Feel free to connect on linkedin.com/in/hvcmacaspac/ for questions about the methodology!*
