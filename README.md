# CHART OF ACCOUTNS VARIANCE ANALYSIS AUTOMATION

## Overview

This repository contains a Python-based automation solution developed for Treasury Reporting to streamline **Variance Analysis**, **Account Movement Analysis**, and **interactive dashboard generation**.

The project is designed to reduce manual work in month-end reporting by automating the update of **Quick Report master data**, processing **GL movement data** across selected reporting periods, comparing movements against **Balance Sheet Exposure** balances, and generating both **Excel-based analytical outputs** and a **self-contained HTML dashboard**.

The solution supports a structured reporting workflow from raw source files to management-ready outputs.

---

## What This Project Does

This automation package performs two main functions:

### 1. Variance / Account Movement Analysis Engine
The variance engine:

- updates the **Quick Report master data** with new monthly Quick Report files,
- processes entity-level and period-level **GL movement data**,
- compares GL activity against **Balance Sheet Exposure** balances,
- calculates variance logic across reporting periods,
- creates structured Excel outputs for analysis and reporting,
- exports BI-ready datasets for dashboard consumption.

### 2. HTML Dashboard Generator
The dashboard engine:

- reads the BI-ready Excel output,
- processes both **Balance Sheet Exposure** and **Variance Movement** data,
- generates a fully interactive **HTML dashboard**,
- provides executive-level views, drill-down analysis, and trend visualisations,
- saves the dashboard as a standalone HTML file that can be opened locally in a browser.

---

## Key Features

### Variance Analysis Engine
- Optional update of **Quick Report master data**
- Multi-period variance analysis
- Entity-based filtering and reporting
- Comparison of **Opening**, **Movement (GL)**, and **Closing** balances
- Variance % calculation
- Main driver explanation logic
- Entity-specific variance sheets
- Cross-month comparison views
- Pivot-style account movement details
- Full consolidation output across entities
- BI-ready Excel export

### HTML Dashboard
- Executive summary KPIs
- Exposure and Unrealized G/L drill-down views
- FX sensitivity scenario analysis
- Currency-level exposure charts
- Entity-level exposure and Unrealized G/L views
- Pivot-style summary sections
- Variance account movement trend analysis
- Detailed variance movement table
- Standalone HTML output with no server requirement

---

## End-to-End Reporting Flow

The reporting flow is designed as follows:

1. **Quick Report data** is used as the source for GL movement analysis.
2. **Balance Sheet Exposure master data** is used as the reference exposure dataset.
3. The **Variance Analysis engine** updates master data and produces:
   - Excel-based analytical reports
   - BI-ready output datasets
4. The **HTML Dashboard engine** reads the BI-ready workbook and generates a browser-based dashboard for executive review.

In short:

**Source files → Variance Analysis Engine → BI Excel Output → HTML Dashboard Generator → Interactive Dashboard**

---

## Main Inputs

The solution typically requires the following inputs:

### For the Variance Analysis Engine
- Quick Report master data file
- Updated Balance Sheet Exposure report
- Variance / Account Movement Analysis template
- Optional new Quick Report files for master data update

### For the HTML Dashboard Engine
A master Excel workbook containing the following BI sheets:

- `BALANCE_SHEET_EXPOSURE_BI_DATA`
- `VARIANCE_MOVEMENTS_BI_DATA`

---

## Main Outputs

### Variance Analysis Outputs
- Updated Quick Report master data file
- Variance / Account Movement Analysis Excel report
- BI-ready Excel output containing:
  - `VARIANCE_MOVEMENTS_BI_DATA`
  - `BALANCE_SHEET_EXPOSURE_BI_DATA`

### Dashboard Output
- Standalone interactive HTML dashboard saved locally and opened in the default browser

---

## Reporting Logic in Business Terms

From a Treasury Reporting perspective, this solution helps answer questions such as:

- Which accounts moved materially compared to the previous period?
- Are account movements supported by Balance Sheet Exposure balances?
- Which currencies are driving exposure and Unrealized G/L?
- Which entities contribute most to total exposure?
- How would exposure metrics change under alternative FX scenarios?
- Which accounts or currencies require management attention?

This makes the process more suitable for both **analytical review** and **management-level reporting**.

---

## Repository Purpose

This repository is intended to document and store the automation logic behind the **Variance Analysis** and **HTML Dashboard** reporting workflow.

It serves as a central place for:

- Python automation scripts
- reporting documentation
- dashboard generation logic
- reporting process standardization
- future maintenance and enhancement

---
