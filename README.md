
## Project: Personal Vacation Leave Management Scripts

This repository contains a set of personal scripts designed to help manage and analyze annual leave data. The scripts automate steps from extracting information to updating a leave report.

### Directory Structure
```
mohamed-ah/vacation/
├── 01 vac.js
├── 02 filter.py
├── 03 checkmissing.py
├── 04 approved.py
├── 05 checkmissingApproved.py
├── 06 analyze.py
└── 07 leaveadjust.py
```
### Script Overview

Here's a brief description of what each script does:

* **`01 vac.js`**: This JavaScript script is for use in a web browser. It helps extract annual leave data by inputting employee IDs into a web form and then collecting the relevant table information. The extracted data is printed to the browser's console.

* **`02 filter.py`**: This Python script takes raw leave data (likely from the `vac.js` output) and filters it based on a list of valid employee codes. It cleans up the data and prepares it for further processing.
    * **Input**: `employee.csv`, `output.csv`
    * **Output**: `filtered_output.csv`

* **`03 checkmissing.py`**: This script checks if any employee codes from your main employee list are missing from the `filtered_output.csv` file. It helps ensure all employees are accounted for.
    * **Input**: `employee.csv`, `filtered_output.csv`

* **`04 approved.py`**: This script specifically extracts "Approved" annual leave records for the year 2025 from the filtered data.
    * **Input**: `filtered_output.csv`
    * **Output**: `approved_2025_output.csv`

* **`05 checkmissingApproved.py`**: Similar to `03 checkmissing.py`, this script checks for any employee codes missing from the *approved* leave data file (`approved_2025_output.csv`).
    * **Input**: `employee.csv`, `approved_2025_output.csv`

* **`06 analyze.py`**: This script analyzes the approved leave records for a specific period (June 14-20, 2025). It determines if employees have no leave, are fully covered, or partially covered during this time.
    * **Input**: `employee.csv`, `approved_2025_output.csv`
    * **Output**: `leave_analysis_report.csv`

* **`07 leaveadjust.py`**: This script updates an Excel file (`AL HARAM PROJECTS.xlsx`) using the analysis from `leave_analysis_report.csv`. It populates "Start Date" and "End Date" columns in the Excel sheet based on the leave status of each employee.
    * **Input**: `AL HARAM PROJECTS.xlsx`, `leave_analysis_report.csv`
    * **Output**: An updated Excel file (e.g., `updated_AL_HARAM_PROJECTS.xlsx`).

### How to Use

The scripts are generally designed to be run in a sequence:

1.  Use `01 vac.js` in your web browser to get raw leave data. Save this as `output.csv`.
2.  Run `02 filter.py` to clean and filter the `output.csv`.
3.  Optionally, run `03 checkmissing.py` to see if any employees were missed in the initial filtering.
4.  Run `04 approved.py` to get only the approved leave records for 2025.
5.  Optionally, run `05 checkmissingApproved.py` to check for missing employees in the approved list.
6.  Run `06 analyze.py` to generate a report on leave coverage for the specific June 2025 period.
7.  Finally, run `07 leaveadjust.py` to update your main Excel project file.

### Prerequisites

* **Python**: Version 3.x
* **Python Libraries**:
    * `pandas`
    * `openpyxl` (for Excel file handling)
    * You can install them using pip: `pip install pandas openpyxl`
* **Input Files**: You'll need `employee.csv`, `output.csv` (from the JS script), and `AL HARAM PROJECTS.xlsx` in your working directory.

### Important Notes on Dates

The Python scripts are currently set to process data for the year **2025** and analyze a reference period from **June 14, 2025, to June 20, 2025**. If you use these scripts for different dates, you will need to adjust the date variables within `04 approved.py`, `06 analyze.py`, and `07 leaveadjust.py`.
```
