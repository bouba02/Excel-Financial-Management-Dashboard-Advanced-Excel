# Excel Financial Management Dashboard | Advanced Excel

> **6-tab integrated system · SUMIFS · INDEX/MATCH · Conditional alerts · Accounting export**  
> Sales · Expenses · Budget · Profitability · N vs N-1 comparison · Windows & Mac compatible

![Dashboard Preview](Dashboard_Ngroup_HD.png)

🇫🇷 [Version française disponible ici](README_FR.md)

---

## Business Problem

SMEs and micro-businesses manage their finances from scattered Excel files —
no consolidated view, time-consuming manual closing, and no automatic alerts
on financial overruns.

**4 questions this dashboard answers:**

| Question | Axis |
|---|---|
| Where does my cash stand this month? | Real-time |
| Are my expenses drifting from budget? | Control |
| Which clients generate the most revenue? | Commercial |
| What is my profitability vs last year? | Comparative |

---

## Architecture — 6 Tabs

```
┌─────────────────────────────────────────────────┐
│  TAB 1 : PARAMETERS                             │
│  Expense categories · Accounting rules ·        │
│  Reporting periods · Data validation             │
├─────────────────────────────────────────────────┤
│  TAB 2 : SALES_REVENUE                          │
│  Date | Client | Amount | Category | Status     │
│  Duplicate detection · Client pivot              │
├─────────────────────────────────────────────────┤
│  TAB 3 : EXPENSES                               │
│  Date | Supplier | Amount | Category            │
│  Amount validation · Abnormal spend alerts       │
├─────────────────────────────────────────────────┤
│  TAB 4 : DASHBOARD ⭐ Core of the system         │
│  Zone 1 : KPIs (Revenue · Expenses · Profit · %)│
│  Zone 2 : Charts (N vs N-1 · Revenue vs Exp.)  │
│  Zone 3 : Quarterly · Expense breakdown         │
│  Zone 4 : Top 10 clients · Monthly detail       │
├─────────────────────────────────────────────────┤
│  TAB 5 : ACCOUNTING_EXPORT                      │
│  Standard accountant format                     │
├─────────────────────────────────────────────────┤
│  TAB 6 : USER_GUIDE                             │
│  Step-by-step manual · FAQ · Troubleshooting    │
└─────────────────────────────────────────────────┘
```

---

## Key Formulas

```excel
// Revenue by client over period
=SUMIFS(Sales[Amount], Sales[Client], [@Client],
        Sales[Date], ">="&StartDate,
        Sales[Date], "<="&EndDate)

// N vs N-1 comparison
=IFERROR(
    SUMIFS(Sales[Amount], Sales[Year], YearRef)
    / SUMIFS(Sales[Amount], Sales[Year], YearRef-1) - 1,
    0)

// Dynamic category lookup
=INDEX(Parameters[Label],
       MATCH([@Category], Parameters[Code], 0))
```

---

## Conditional Alerts

```
🔴 Budget variance > 10%   → Automatic red highlight
🟠 Abnormal expense         → Alert vs historical average
🟢 All under control        → Green indicator
```

---

## Dashboard — 4 Zones

![Dashboard](Dashboard_Ngroup_HD.png)

| Zone | Content |
|---|---|
| Zone 1 — KPIs | Total Revenue · Expenses · Net Profit · Margin % · N vs N-1 |
| Zone 2 — Charts | Monthly revenue N vs N-1 (line) · Revenue vs Expenses (stacked bars) |
| Zone 3 — Analysis | Quarterly performance Q1→Q4 · Expense breakdown by category |
| Zone 4 — Details | Top 10 clients (ranked + % revenue) · Monthly summary Jan–Dec |

---

## Tech Stack

| Tool | Usage |
|---|---|
| **Advanced Excel** | SUMIFS · INDEX/MATCH · Pivot tables |
| **Conditional formatting** | Visual alerts on business thresholds |
| **Dynamic charts** | Auto-update on every data entry |
| **Data validation** | Dropdown lists · Input controls |

**Capacity:** 10,000+ rows · Windows & Mac compatible · Excel 2016+

---

## Repository Structure

```
Excel-Dashboard-Pilotage-Financier/
├── README.md
├── README_FR.md
├── Dashboard_Ngroup_HD.png
└── Pilotage_Financier.xlsx
    ├── PARAMETERS
    ├── SALES_REVENUE
    ├── EXPENSES
    ├── DASHBOARD
    ├── ACCOUNTING_EXPORT
    └── USER_GUIDE
```

---

## Quick Start

```bash
git clone https://github.com/bouba02/Excel-Dashboard-Pilotage-Financier.git
```

Open `Pilotage_Financier.xlsx` in Excel 2016+ or Office 365.

---

## Author

**Boubacar Nikiema** — Data Analyst & BI Consultant

Specialized in financial dashboards, advanced Excel and SME performance management
using Power BI, SQL, Python and Excel. Based in Morocco, working with clients
across Africa and French-speaking Europe.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-boubacar--nikiema-blue?logo=linkedin)](https://linkedin.com/in/boubacar-nikiema)
[![YouTube](https://img.shields.io/badge/YouTube-BoubacarDataAnalyst-red?logo=youtube)](https://youtube.com/@BoubacarDataAnalyst)
[![Email](https://img.shields.io/badge/Email-nikiemaboubacar%40gmail.com-gray?logo=gmail)](mailto:nikiemaboubacar@gmail.com)
[![Portfolio](https://img.shields.io/badge/Portfolio-data.ngroupmediadigital.com-green)](https://data.ngroupmediadigital.com)

---

*Code: MIT License · Compatible Excel 2016+ and Office 365*
