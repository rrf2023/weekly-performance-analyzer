Weekly Performance Analyzer

A desktop ETL and reporting tool for generating weekly performance reports and anomaly detection from PostgreSQL transactional data.

The application produces structured Excel reports per interviewer and a consolidated “suspicious households” report using rule-based fraud and consistency detection.

Project Structure

This project is intentionally implemented as a single-file ETL application (app/main.py) to demonstrate full pipeline clarity without framework fragmentation.

Features
Data Processing
Extracts transactional data from PostgreSQL
Supports weekly filtering by selectable date range
Normalizes:
household IDs
product categories
EAN / SKU codes
Reporting
Generates per-interviewer Excel reports
Produces consolidated suspicious household reports

Applies business rules:

households with 1–2 product categories are flagged
duplicate receipt detection via SKU signature matching
Excel Output Features
Formatted sheets with column sizing
Hierarchical grouping (expand/collapse rows)
Cross-file hyperlinks between summary and detail reports
Aggregated purchase summaries
Requirements
Python 3.9+
PostgreSQL database access
Excel reference files (must be placed in working directory):
Supervisors.xlsx
Справочник.xlsx
Installation
git clone https://github.com/rrf2023/weekly-performance-analyzer.git
cd weekly-performance-analyzer

python -m venv venv
source venv/bin/activate   # Linux / Mac
venv\Scripts\activate      # Windows

pip install -r requirements.txt
Run
python app/main.py
Configuration

Database credentials are stored in setup.txt (auto-generated on first run):

host=localhost
port=5432
database=your_db
user=your_user
password=your_password
Input Data Schema (PostgreSQL)

Required columns:

f0103 → household ID
f0105 → reserved / unused in current pipeline
f0122 → SKU / EAN
prod_group → product category
date → transaction date
Outputs
1. Interviewer Reports

Filename format:

{interviewer_name}_{week_start}.xlsx

Contains:

household-level purchase summaries
category breakdowns
enriched descriptions from reference dictionary
2. Suspicious Households Report

Filename:

Suspicious_households_{week_start}.xlsx

Contains:

flagged households
anomaly indicators
SKU-level breakdown
grouped hierarchical structure for review
Business Logic

A household is flagged if:

It has 1–2 product categories, OR
It shares an identical SKU signature with another household
Tech Stack
Python
Tkinter (desktop UI)
Pandas (data processing)
PostgreSQL (data source)
OpenPyXL (Excel generation)
Limitations
Requires Excel reference files in working directory
Large datasets may slow UI (no async execution)
Credentials stored locally in plaintext (setup.txt)
Future Improvements
Modular architecture refactor (service separation)
Async database operations
Web-based dashboard (Streamlit / FastAPI)
Secure credential storage
Config-driven rule engine
License

MIT
