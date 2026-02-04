# Transport Work Reconciliation Automator

## Context
I built this tool to automate the parsing of legal documents (`.docx` addendums) and the calculation of quarterly transport metrics. By automating the reconciliation of trips and mileage, **I eliminated the manual entry of hundreds of route intervals and reduced report generation time from hours to minutes**, ensuring 100% consistency with the legal supplemental agreements.

---

## How it works
The project is structured as a pipeline that moves from document parsing to financial planning and final reconciliation.

### 1. Document Parsing (`read_docx.py`)
This module extracts route data directly from the appendices of legal acts. It handles three specific date logic types:
* **Start dates:** Records a single "start" entry.
* **Intervals:** Creates two entries (start and end), automatically adding `+1 day` to the end date to correctly close the interval for calculation.
* **Single day markers:** Treated as a 1-day interval.
* **Deduplication:** The script automatically cleans overlapping route data, prioritizing interval logic to prevent double-counting.

### 2. Financial Planning (`plan_by_last_add_aggs.py`)
This script generates trip and mileage forecasts for the finance department based on the latest supplemental agreements. It accepts a list of agreement numbers and a specific month range to output the expected transport volume.

### 3. Core Reconciliation (`main.py`)
The engine of the project. It integrates everything:
* **Data Loading:** Pulls coefficients, vehicle capacities, route plans, and holiday schedules.
* **Holiday Logic:** Incorporates `holidays.xlsx` to account for route changes during non-working days that aren't specified in the main addendums.
* **Processing Loop:** Iterates through quarterly periods using `pandas.date_range`, calculating trips, mileage (from start and end points), and total costs per route.
* **Aggregation:** Outputs two distinct Excel reports—one detailed by route and one summarized by quarter.

---

## Tech Stack
* **Python 3.x**
* **Pandas:** For the heavy lifting on data manipulation and date arithmetic.
* **TQDM:** For progress tracking during processing loops.
* **python-docx:** To extract structured data from Word-based legal appendices.
* **OpenPyXL:** For generating final Excel reports.

---

## Why it matters
Accuracy in transport reconciliation is the difference between getting paid on time and dealing with legal disputes. This tool ensures that every kilometer calculated matches the latest signed supplemental agreement. It provides the finance department with a "single source of truth," bridging the gap between legal document text and raw operational data.
