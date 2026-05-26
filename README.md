# Excel Reporting Dashboard Automation

## Overview
This project automates EV fleet reporting using Python and Excel. It reads operational data, merges lookup data, generates weekly and monthly summaries, and exports a formatted multi-sheet Excel workbook for reporting and dashboard analysis.

![Dashboard Screenshot](assets/dashboard.png)

## Features
- Automated Excel report generation using Python
- Multi-sheet workbook creation
- Weekly KM summary
- Monthly KM summary
- City-level summary
- Vehicle-type summary
- Excel formatting with styled headers, freeze panes, and adjusted column widths
- Dashboard sheet with KPI cards and charts

## Tech Stack
- Python
- Pandas
- openpyxl
- XlsxWriter
- Excel

## Project Structure
```bash
excel-reporting-dashboard-automation/
├── assets/
│   └── dashboard.png
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
2. Clean and merge datasets
3. Generate weekly and monthly summaries
4. Export results to a multi-sheet Excel workbook
5. Apply formatting using openpyxl
6. Build a dashboard in Excel with KPI cards and charts

## Output Sheets
- Raw_Data
- Lookup_Data
- Cleaned_Data
- Weekly_KM
- Monthly_KM
- City_Summary
- Vehicle_Summary
- Dashboard

## Business Use Case
This project simulates a reporting workflow for EV operations teams that need accurate weekly and monthly reports for distance traveled, battery usage, charge cycles, and city-level fleet monitoring.

## How to Run
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 create_excel_report.py
python3 format_excel_report.py
```

## Resume Bullet
Built an Excel Reporting Dashboard Automation project using Python, Pandas, openpyxl, and XlsxWriter to generate multi-sheet operational reports and KPI-based Excel dashboards for EV fleet monitoring.

## Future Improvements
- Add slicers and pivot charts
- Add conditional formatting for KPI alerts
- Export charts automatically from Python
- Extend to Power BI or Grafana