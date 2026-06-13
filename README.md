Weekly Performance Analyzer

Overview

Weekly Performance Analyzer is a production-style ETL and behavioral analytics system designed to process retail transaction data and identify anomalous household purchasing patterns using transparent, rule-based logic.

Unlike typical machine learning–driven anomaly detection systems, this project prioritizes interpretability and auditability, making it suitable for operational retail analytics environments where explainability is critical.

The system ingests PostgreSQL transaction data, enriches it using reference datasets, applies structured transformation logic, and generates automated weekly Excel reports for supervisors and analytics teams.

Why This Project Matters

In real-world retail and operational analytics, stakeholders often require:

Fully explainable anomaly detection logic (no black-box models)
Reproducible weekly reporting pipelines
Clear traceability from raw data → business insight
Lightweight, maintainable ETL systems

This project demonstrates how these requirements can be implemented using a modular Python-based data pipeline.

System Architecture
PostgreSQL → ETL Layer → Data Enrichment → Rule Engine → Reporting Layer (Excel)
Core Components
1. ETL Pipeline
Extracts transactional data from PostgreSQL
Normalizes household IDs, SKU/EAN codes, and product categories
Applies weekly time-window filtering for reporting cycles
2. Data Enrichment Layer
Joins transaction data with reference datasets:
Supervisor → household mapping
Product category dictionary
SKU/EAN reference metadata
Ensures consistent categorization across all outputs
3. Rule-Based Anomaly Detection Engine

Households are flagged based on interpretable behavioral heuristics:

Low category diversity
Households with only 1–2 product categories within a weekly window
SKU signature similarity
Households sharing identical or highly similar SKU purchase patterns

These rules are designed to surface:

Unusual purchasing behavior
Potentially coordinated or duplicated activity patterns
Outliers in consumption structure
4. Reporting Engine
Generates per-supervisor weekly Excel reports
Produces consolidated anomaly reports
Includes structured summaries and hierarchical grouping
Supports cross-referencing between households and categories
Outputs
Supervisor Reports

Generated per supervisor:

sup_{id}_{week_start}.xlsx

Includes:

Household-level purchase summaries
Category breakdowns
Product-level aggregation
Anomaly Report
Suspicious_households_{week_start}.xlsx

Includes:

Flagged households
Detected anomaly indicators
SKU signature clustering results
Data Model
Table: sample_transactions
Column	Description
f0103	Household ID
f0122	SKU / EAN code
prod_group	Product category code
date	Transaction timestamp
Tech Stack
Python 3.9+
PostgreSQL
psycopg2
pandas
openpyxl
Tkinter (optional UI layer)
Project Structure
app/
├── main.py              # Pipeline orchestrator
├── db.py                # PostgreSQL connection layer
├── processors.py       # Transformation + rule engine
├── excel_export.py     # Reporting engine
├── reference_loader.py # Reference dataset loader
├── config.py           # Configuration management
├── gui.py              # Optional desktop UI

docs/
└── schema.sql          # Database schema

examples/
├── example_1.png
└── example_2.png

requirements.txt
README.md
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
cp app/setup.example.txt app/setup.txt

Edit configuration:

host=localhost
port=5432
database=your_db
user=your_user
password=your_password
5. Initialize database
psql -U your_user -d your_db -f docs/schema.sql
6. Run pipeline
python app/main.py
Example Workflow
Load transactional data from PostgreSQL
Enrich data using reference dictionaries
Apply weekly filtering logic
Generate supervisor-level reports
Execute anomaly detection rules
Export structured Excel outputs
Design Philosophy

This project is intentionally built with minimal dependencies to emphasize:

Transparent ETL logic
Explicit business rule implementation
Reproducible reporting pipelines
Maintainable modular structure

The focus is on clarity, auditability, and practical data engineering design, rather than abstraction-heavy frameworks.

Limitations
Designed for batch weekly processing (not real-time streaming)
Single-process execution model
Local configuration-based setup
Excel-based output instead of BI dashboard integration

These constraints reflect a focus on simplicity and interpretability.

Future Improvements
Modular ETL pipeline refactor with stricter separation of concerns
Async PostgreSQL query execution for performance scaling
Config-driven rule engine (external rule definitions)
Web-based dashboard (Streamlit or FastAPI)
Secure credential management via environment variables
Incremental data processing (delta loads instead of full refresh)
License

MIT
