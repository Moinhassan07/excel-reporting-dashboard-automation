# Excel Reporting Dashboard Automation

An end-to-end reporting automation project that generates Excel-based EV fleet reports and extends the workflow into Grafana for interactive analytics and monitoring.

## Overview

This project automates EV fleet reporting using Python and Excel. It reads operational data, merges lookup data, generates weekly and monthly summaries, and exports a formatted multi-sheet Excel workbook for reporting and dashboard analysis.

The project was later extended into Grafana by loading EV operations data into PostgreSQL and building an interactive dashboard for KPI tracking, operational monitoring, and filter-based exploration.

## Features

- Automated Excel report generation using Python
- Multi-sheet workbook creation
- Weekly KM summary
- Monthly KM summary
- City-level summary
- Vehicle-type summary
- Excel formatting with styled headers, freeze panes, and adjusted column widths
- Dashboard sheet with KPI cards and charts
- Grafana dashboard extension for interactive analytics
- Filters for city and bike type in Grafana

## Tech Stack

- Python
- Pandas
- openpyxl
- XlsxWriter
- Microsoft Excel
- PostgreSQL
- Grafana

## Project Structure

```text
excel-reporting-dashboard-automation/
├── assets/
│   ├── dashboard.png
│   └── grafana-dashboard.png
├── data/
├── docs/
├── output/
├── create_excel_report.py
├── format_excel_report.py
├── requirements.txt
└── README.md
```

## Workflow

1. Load EV operations data and lookup data
2. Clean and merge the datasets
3. Generate weekly and monthly summaries
4. Export results to a multi-sheet Excel workbook
5. Apply workbook formatting using openpyxl
6. Build a dashboard sheet in Excel with KPI cards and charts
7. Load EV operations data into PostgreSQL
8. Connect Grafana to PostgreSQL
9. Build an interactive Grafana dashboard with filters and charts

## Output Sheets

The generated Excel workbook contains the following sheets:

- `Raw_Data`
- `Lookup_Data`
- `Cleaned_Data`
- `Weekly_KM`
- `Monthly_KM`
- `City_Summary`
- `Vehicle_Summary`
- `Dashboard`

## Project Screenshots

### Excel Dashboard
![Excel Dashboard](./assets/dashboard.png)

### Grafana Dashboard
![Grafana Dashboard](./assets/grafana-dashboard.png)

## Business Use Case

This project simulates a reporting workflow for EV operations teams that need accurate weekly and monthly reporting for distance traveled, battery usage, charge cycles, and city-level fleet monitoring.

It demonstrates how a traditional Excel reporting process can be automated and then extended into a dashboard-driven analytics workflow for better monitoring and decision-making.

## How to Run

### 1. Create and activate a virtual environment

```bash
python3 -m venv venv
source venv/bin/activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Generate the Excel report

```bash
python3 create_excel_report.py
python3 format_excel_report.py
```

## Grafana Extension

The reporting workflow was extended into Grafana by loading EV operations data into PostgreSQL and creating an interactive dashboard with:

- KPI cards
- Daily trend charts
- City-level analysis
- SLA and operational status tracking
- Filters for city and bike type
- Detailed recent operations table

This shows how the same reporting pipeline can support both Excel-based business reporting and dashboard-based operational monitoring.

## Resume Bullet

Built an Excel Reporting Dashboard Automation project using Python, Pandas, openpyxl, and XlsxWriter to generate multi-sheet EV fleet operational reports and KPI-based Excel dashboards, later extending the workflow into Grafana with PostgreSQL for interactive analytics and monitoring.

## Project Highlights

- Automated repetitive reporting work using Python
- Created clean, presentation-ready Excel outputs
- Combined data processing, Excel reporting, SQL-based analytics, and dashboard visualization in one workflow
- Designed a business-focused reporting solution for EV fleet operations
- Extended static Excel reporting into interactive Grafana monitoring

## Future Improvements

- Add slicers and pivot charts in Excel
- Add conditional formatting for KPI alerts
- Export charts automatically from Python
- Extend the reporting workflow to Power BI
- Add Docker-based deployment for PostgreSQL and Grafana
- Add automated screenshot generation and sample output files for portfolio demos
