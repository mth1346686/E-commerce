# E-Commerce Data Integration & Analytics Pipeline

[![DEPI](https://img.shields.io/badge/Initiative-DEPI_Egypt-0052CC.svg)](https://depi.gov.eg/)
[![Track](https://img.shields.io/badge/Track-Data_Engineering-orange.svg)](#)
[![Documentation](https://img.shields.io/badge/Documentation-PDF_Available-red.svg)](./Documentation.pdf)
[![Airflow](https://img.shields.io/badge/Orchestration-Apache_Airflow-017CEE.svg?logo=apache-airflow&logoColor=white)](https://airflow.apache.org/)
[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Docker](https://img.shields.io/badge/Packaging-Docker-2496ED.svg?logo=docker&logoColor=white)](https://www.docker.com/)

Official Graduation Project for the **Digital Egypt Pioneers Initiative (DEPI)** — Data Engineering Track.

An end-to-end automated ETL and data warehousing pipeline that extracts multi-source sales and catalog feeds, cleanses and validates records, and populates a performance-optimized **Kimball Star Schema** data warehouse orchestrated via **Apache Airflow**.

---

## 📄 Project Documentation

The complete, official 5-page graduation project documentation is available in this repository:

👉 **[Download Project Documentation (PDF)](./Documentation.pdf)**

---

## 👥 Team Members & Assigned Roles

| Member | Role | Key Responsibilities |
| :--- | :--- | :--- |
| **Mahmoud Tarek Mahmoud Hafez** | **Team Leader, Data Architect & Pipeline Lead** | Overall project supervision, system architecture design, Kimball star schema modeling, DDL generation, Airflow DAG workflow implementation, and GitHub governance. |
| **Mohamed Ahmed Fathy El-Deeb** | **Core Extraction & Ingestion** | Development of ingestion scripts, REST API connectors, web scrapers, schema validation logic, and raw landing stage operations. |
| **Mahmoud Abdelnaby Refaat Ibrahim** | **Core Transformation & Data Warehouse** | End-to-end data cleansing scripts, complex Pandas business transformations, SQL aggregation views, indexing strategies, and incremental staging loads. |
| **Associate Team Member 4** | **Support & QA Documentation** | Predefined data-quality check scripts, technical summaries, data dictionaries, pipeline run logs monitoring, and presentation deliverables. |

---

## 🏗️ System Architecture & Workflow

```
[ Multi-Source Ingestion ]
   ├── Product Catalogs (REST APIs & Web Scraping)
   ├── Transactions & Customers (CSV / JSON streams)
               │
               ▼
[ Pre-Ingestion & Cleansing Layer (Python & Pandas) ]
   ├── Schema Validation & Drift Detection
   ├── Deduplication & Normalization
   ├── Imputation & UTC Timestamp Standardization
   ├── Dead-Letter Queue (Quarantine for Malformed Records)
               │
               ▼
[ Data Warehouse & Analytical Modeling ]
   ├── Fact_Sales
   ├── Dim_Customer
   ├── Dim_Product
   └── Dim_Date
               │
               ▼
[ Business Intelligence & Analytics Views ]
   ├── Optimized Indexes & Materialized SQL Views
   └── High-Frequency KPI Aggregations (Sub-second query latency)
```

---

## 🗓️ Project Plan (Timeline & Milestones)

| Phase | Timeline | Milestones & Key Deliverables |
| :--- | :--- | :--- |
| **Phase 1: Planning & Requirements** | Oct 2 – Oct 16, 2026 | Finalized project charter, team roles, requirements document, and GitHub repository setup. |
| **Phase 2: System Analysis & Modeling** | Oct 16 – Nov 6, 2026 | Conceptual architecture diagram, ERD design, DFD mapping, and Star Schema modeling. |
| **Phase 3: ETL & Pipeline Development** | Nov 6 – Nov 30, 2026 | Python ingestion modules, transformation scripts, SQL DDL/DML, and functional Airflow DAGs. |
| **Phase 4: Testing, QA & Final Delivery** | Nov 30 – Dec 4, 2026 | Data quality assertion tests, Docker containerization, final presentation, infographic, and reports. |

---

## 🎯 Key Performance Indicators (KPIs)

* **Pipeline Execution SLA:** Scheduled daily DAG run completion time under **15 minutes**.
* **Data Quality & Integrity Rate:** Above **99.5% validity** across primary keys, foreign keys, and non-nullable transaction values.
* **Query Latency:** Sub-second execution time (**< 1.0s**) for standard analytical reporting on the relational star schema.
* **System Resilience:** Automated recovery and idempotency ensuring **zero duplicate transactions** upon retry.

---

## 🛡️ Risk Assessment & Mitigation

1. **External API Rate Limiting / Outages:** Mitigated via exponential backoff, circuit-breaker logic, Airflow retries, and cached landing storage.
2. **Schema Drift & Corrupt Feeds:** Validated at ingestion; malformed records are routed to a dead-letter quarantine table without halting the pipeline.
3. **Analytical Query Bottlenecks:** Addressed with composite indexes on foreign keys and date dimensions, alongside pre-aggregated SQL views.

---

## 📚 References & Standards
* Kimball, R., & Ross, M. (2013). *The Data Warehouse Toolkit: The Definitive Guide to Dimensional Modeling*.
* Apache Airflow Documentation: *Architecture, Scheduling & Orchestration Best Practices*.
* McKinney, W. (2022). *Python for Data Analysis: Data Wrangling with pandas, NumPy, and Jupyter*.

---

## 📄 License
This project is released under the [MIT License](LICENSE).
