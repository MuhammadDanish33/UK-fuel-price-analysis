# 🇬🇧 UK Weekly Road Fuel Price Analysis

> **Junior Data Analyst Portfolio Project — Microsoft Excel**

---

## 📌 Overview / Business Question

UK motorists, transport operators, and policy stakeholders need clear, timely answers to practical questions about pump prices:

- What are the current average petrol and diesel prices?
- How much have prices changed week-on-week and year-on-year?
- When did prices peak or hit their lowest point?
- What has the typical diesel premium over petrol been at any given time?
- Can I quickly look up the exact price for any historical week?

This project turns **official UK government fuel price data** into a clean, structured, and queryable Excel workbook that answers all of the above.

---

## 🎯 Objective

Build a well-documented Excel workbook that:
1. Preserves the original government dataset without modification
2. Delivers a clean, analysis-ready time series with standardised column names and derived helper fields
3. Calculates key weekly and monthly price KPIs directly from the data
4. Provides a simple interactive dashboard for historical week lookups

---

## 📂 Dataset / Data Source

| Field | Detail |
|---|---|
| **Source** | UK Department for Energy Security and Net Zero (Official Statistics) |
| **File Used** | `Weekly_Fuel_Prices_24-08-26.xlsx` — Sheet: *Data* |
| **Coverage** | 9 June 2003 – 24 August 2026 |
| **Observations** | 1,212 weekly records |
| **Market Sample** | ~60% of UK road-fuel volume (4 oil majors + 2 supermarket chains) |
| **Fuels Covered** | ULSP (Unleaded Petrol) and ULSD (Diesel) |

**Original columns used:** Pump price, week-on-week change, year-on-year change, duty rate, and VAT rate — for both fuels.  
No regional breakdowns or supermarket-vs-forecourt splits are available in this source file and none are claimed.

---

## 🛠️ Tools & Skills

**Tool:** Microsoft Excel only — no VBA, no Power Query, no external libraries.

| Skill Area | Detail |
|---|---|
| **Data organisation** | 5-sheet workbook structure with clear naming conventions |
| **Column standardisation** | Renamed 11 original headers to analyst-readable names with a data dictionary |
| **Helper column engineering** | `YEAR()`, `MONTH()`, `TEXT()`, `WEEKNUM()` to derive Year, Month, MonthName, WeekID |
| **KPI calculation** | `AVERAGE()`, `MAX()`, `MIN()` for all-time statistics |
| **Lookup functions** | `INDEX()`, `MATCH()`, `IFERROR()` for the interactive dashboard |
| **Aggregation** | Pivot Table for monthly average and peak prices by year |
| **Data validation** | Dropdown list to constrain WeekID input in the dashboard |
| **Charting** | Line charts for petrol and diesel price trends over time |
| **Documentation** | In-workbook ReadMe sheet with glossary and usage guide |

---

## 🔬 Methodology / Analysis

### Workbook Structure

| Sheet | Content |
|---|---|
| `01_ReadMe` | Project overview, data source, column glossary, how-to-use guide |
| `02_RawData` | Original government data — preserved and unmodified |
| `03_CleanData` | Full cleaned time series: renamed columns + derived helper columns |
| `04_Analysis & Findings` | Latest-week KPIs, all-time statistics, monthly Pivot Table, trend charts, written findings |
| `05_Dashboard` | Interactive lookup: enter a WeekID → instantly retrieve all price and change data |

### Data Cleaning Steps
- Renamed 11 long source headers to concise, underscore-separated names (e.g. `ULSP: Pump price (p/litre)` → `Petrol_Price (p/litre)`)
- Added four derived helper columns: `Year`, `Month`, `MonthName`, `WeekID` (format: `YYYY-WW`, e.g. `2026-34`)
- No source values were altered — all original figures are preserved in `02_RawData`

---

## 📊 Key Findings

All figures are drawn directly from `04_Analysis & Findings`. All prices are in **pence per litre (p/litre)**.

### Latest Snapshot (week of 24 August 2026)

| Metric | Petrol (ULSP) | Diesel (ULSD) |
|---|---|---|
| **Latest Price** | **161.63p** | **182.82p** |
| Week-on-Week Change | −0.17p | +0.37p |
| Year-on-Year Change | +27.71p | +40.90p |

### All-Time Statistics (June 2003 – August 2026)

| Metric | Petrol | Diesel |
|---|---|---|
| **All-Time Average** | 119.25p | 125.03p |
| **All-Time High** | **191.55p** *(4 Jul 2022)* | **199.22p** *(4 Jul 2022)* |
| **All-Time Low** | **74.21p** *(14 Jul 2003)* | **76.41p** *(14 Jul 2003)* |

### Notable Periods

**📈 2022 Energy Crisis**  
Petrol rose from ~145p/litre (January 2022) to an all-time peak of **191.55p** (4 July 2022) — a **+32.1% surge in six months** — driven by post-pandemic demand recovery and the Russia–Ukraine conflict. Diesel peaked simultaneously at **199.22p**. By December 2022, petrol had retreated to 151.94p and diesel to 175.52p.

**💷 Fuel Duty Cut (March 2022)**  
Fuel duty was cut from **57.95p to 52.95p/litre**, first reflected in data from 28 March 2022. Despite the 5p reduction, pump prices continued rising for a further 14 weeks — confirming that the global commodity price surge completely overwhelmed the duty relief during this period.

**📉 2025 Stability Window**  
After three years of volatility, prices stabilised across 2025. Petrol ranged narrowly between 131.35p and 139.62p (annual average: **135.05p**) and diesel between 137.54p and 146.88p (annual average: **142.70p**) — the calmest sustained pricing environment since 2019.

**📈 2026 Price Surge (Two Phases)**  
From 9 March 2026, prices surged sharply: petrol climbed **+19.7%** to 158.17p and diesel **+35.2%** to 192.14p by 13 April 2026 — the steepest six-week rise since the 2022 crisis. Diesel then retreated, while petrol continued rising in a second phase, reaching **162.15p** on 10 August 2026.

---

## 🖥️ Dashboard / Outputs

**`05_Dashboard` — Interactive Lookup**

Enter any valid `WeekID` (format: `YYYY-WW`, e.g. `2022-27`) into the dropdown to instantly retrieve:

- Petrol price and week-on-week change
- Diesel price and week-on-week change
- Year-on-year change for both fuels
- Fuel duty rate
- Month and year of the selected week

Built using `INDEX`/`MATCH` with `IFERROR` for clean error handling — no macros or VBA required.

**`04_Analysis & Findings` — Summary KPIs + Monthly Pivot Table**

The Pivot Table summarises data by year and month, showing average and peak petrol/diesel prices along with the count of weekly observations per period — enabling fast month-by-month comparisons across the full 2003–2026 dataset.

---

## 💡 Business Recommendations

*The following recommendations are supported directly by the data in the workbook:*

1. **Fleet and transport budget planning should account for year-on-year volatility.** The data shows petrol prices varied by more than 100p/litre across the full dataset (74p–191p), with multi-month surges possible — as seen in both 2022 and 2026. Fixed-price fuel contracts or hedging strategies may reduce exposure.

2. **Duty changes alone are insufficient to counteract commodity-driven price surges.** The March 2022 5p duty cut had no visible effect on pump prices for 14+ weeks, as global oil price movements dominated. Stakeholders should not rely on duty policy as a near-term price relief mechanism when commodity markets are elevated.

3. **The 2025 stability window (avg petrol 135p, avg diesel 143p) may represent a reference baseline** for modelling normal-market fuel cost scenarios — before the 2026 surge re-emerged.

---

## ✅ Conclusion

This project demonstrates the ability to take a real, publicly available government dataset — messy headers, multi-year weekly time series — and turn it into a structured, documented, and usable analytical resource using core Excel skills. The workbook is self-contained: a new analyst can open it, read the `01_ReadMe` sheet, and immediately understand the data, the methodology, and how to use the dashboard without any external guidance.

The scope is deliberately realistic for a junior analyst role: solid data foundations, clean outputs, and one practical interactive feature — built without over-engineering.

---

## 📁 Files

| File | Description |
|---|---|
| `UK_Fuel_Price_Analysis_Project.xlsx` | Main project workbook (5 sheets) |

---

*Data Source: UK Department for Energy Security and Net Zero — Official Statistics*  
*Coverage: 9 June 2003 – 24 August 2026 · 1,212 weekly observations*
