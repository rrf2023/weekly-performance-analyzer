Weekly Performance Analyzer

Automated ETL and anomaly detection pipeline for retail household transaction analysis.

This project ingests PostgreSQL transactional data, applies rule-based behavioral analysis, and generates structured weekly Excel reports for supervisors. It detects potentially anomalous household purchasing patterns using transparent, explainable business rules rather than black-box models.

Built as a portfolio-grade data engineering and analytics project, it demonstrates practical ETL workflows, business-rule systems, relational data processing, and automated reporting in Python.

Key Features
ETL Pipeline
Extracts transactional data from PostgreSQL
Normalizes household IDs, SKU/EAN codes, and product categories
Applies weekly time-window filtering
Reporting Engine
Generates per-supervisor Excel reports
Produces consolidated anomaly report
Cross-references household activity across datasets
Business Rule Engine (Anomaly Detection)

A household is flagged if:

It has only 1–2 product categories (low diversity signal), OR
It shares identical SKU signature with other households
Excel Output Features
Structured worksheets with formatted columns
Hierarchical grouping (expand/collapse rows)
Cross-report navigation via internal links
Aggregated summaries per household and category
Tech Stack
Python 3.9+
PostgreSQL
psycopg2
pandas
openpyxl
Tkinter (optional desktop UI layer)
Project Structure
app/
  main.py                 # ETL pipeline orchestrator
  Supervisors.xlsx       # Supervisor-to-household mapping
  Справочник.xlsx        # Category reference dictionary
  EAN.xlsx               # SKU/EAN reference data
  setup.example.txt      # Configuration template

docs/
  schema.sql             # PostgreSQL schema definition

output/
  sample_reports/        # Example generated Excel reports (not required to run)

requirements.txt
README.md
Data Model
PostgreSQL Table: sample_transactions
Column	Description
f0103	Household ID
f0122	SKU / EAN code
prod_group	Product category code
date	Transaction timestamp
Outputs
1. Supervisor Reports

Generated per supervisor:

sup_{id}_{week_start}.xlsx

Contains:

Household-level purchase summaries
Category breakdown
Referenced product descriptions
2. Suspicious Households Report
Suspicious_households_{week_start}.xlsx

Contains:

Flagged households
Anomaly indicators
SKU signature clustering results
Setup & Installation
1. Clone repository
git clone https://github.com/rrf2023/weekly-performance-analyzer.git
cd weekly-performance-analyzer
2. Create virtual environment
python3 -m venv venv
source venv/bin/activate
3. Install dependencies
pip install -r requirements.txt
4. Configure database

Copy configuration template:

cp app/setup.example.txt app/setup.txt

Edit with your PostgreSQL credentials:

host=localhost
port=5432
database=your_db
user=your_user
password=your_password
5. Initialize database

Run schema setup:

psql -U your_user -d your_db -f docs/schema.sql
6. Run pipeline
python app/main.py
Example Workflow
Load transactional data from PostgreSQL
Join with reference dictionaries (Supervisors + categories + SKU metadata)
Apply weekly filtering
Generate supervisor-level reports
Detect anomalies using rule engine
Export structured Excel outputs
Design Philosophy

This project intentionally avoids heavy frameworks to demonstrate:

Transparent ETL logic
Explicit business rule implementation
Reproducible reporting pipeline
Minimal dependency surface

The goal is clarity over abstraction.

Limitations
Single-process execution (no async pipeline)
Local configuration-based setup
Designed for batch weekly runs, not streaming
Excel output is file-based (not BI tool integration)
Future Improvements
Modular architecture refactor (ETL / rules / reporting separation)
Async PostgreSQL queries for scalability
Config-driven rule engine
Web dashboard (Streamlit / FastAPI)
Secure credential storage (.env integration)
Incremental processing (delta loads instead of full refresh)
License

MIT

Notes

This project is intended as a portfolio-grade demonstration of:

ETL pipeline design
Data transformation logic
Business rule-based anomaly detection
Automated reporting systems
