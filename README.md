# Automated Sales & Inventory MIS Report Generator

An automated Python data pipeline built with **Pandas** and **OpenPyXL** to clean raw multi-sheet Excel sales/inventory datasets and generate structured Management Information System (MIS) reports.

## Key Features
- **Data Cleaning & Validation:** Automatically removes duplicate records and filters out incomplete rows with missing product names or sales quantities.
- **Automated MIS Aggregations:** Computes region-wise revenue and product-level sales performance using group-by operations.
- **Inventory Reorder Alerts:** Flags low-stock products by comparing current stock levels against predefined reorder thresholds.
- **Multi-Tab Excel Export:** Outputs a clean, management-ready `.xlsx` workbook (`Automated_MIS_Report.xlsx`) with separate tabs for summaries, alerts, and cleaned data.

## Tech Stack
- Python
- Pandas
- OpenPyXL
- Microsoft Excel

## How to Run
1. Install dependencies:
   ```bash
   pip install pandas openpyxl
