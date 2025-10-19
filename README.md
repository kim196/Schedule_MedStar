NP/PA Schedule Processor

A Python utility to parse an Excel staff schedule and automatically produce:

A CSV of precomputed P/W (Point / W) assignments across a date range

Populated NP/PA assignment .docx files for individual dates using a Word template

The project helps automate scheduling for NP/PA staff by balancing point and W assignments fairly across weekdays.

Overview:
This repository includes two main components:

Compute_PW.py – scans the Excel schedule across a date range and produces precomputed_pw_assignments.csv

Filled_Assignment.py – reads the schedule for a single date and fills a Word template (NPPA Assignments.docx), producing daily .docx files in the Created_Docs folder

Project Structure:

Schedule.xlsx (input Excel schedule)

NPPA Assignments.docx (Word template)

Compute_PW.py (precompute P/W assignments)

Filled_Assignment.py (generate DOCX for a single date)

precomputed_pw_assignments.csv (output file)

Created_Docs/ (folder for generated DOCX files)

Requirements:

Python 3.8+

Libraries: pandas, python-docx

Install with:
pip install pandas python-docx

Configuration:
You can change paths in the Python scripts:
EXCEL_PATH = "Schedule.xlsx"
TEMPLATE_PATH = "NPPA Assignments.docx"
ASSIGNMENT_OUTPUT_PATH = "precomputed_pw_assignments.csv"
OUTPUT_DIR = "Created_Docs"

Usage:

To precompute P/W assignments:
python Compute_PW.py
Then enter a start and end date when prompted.
Output → precomputed_pw_assignments.csv

To generate NP/PA DOCX assignments:
python Filled_Assignment.py
Enter a date when prompted.
Output → Created_Docs/NPPA_Assignments_<DATE>.docx

Input Details:

The Excel file should be structured in 17-row blocks per day.

The second row in each block contains the date.

Names containing “NP” or “PA” are processed.

Shifts like FSH, PTO, SH, OT are ignored.

Assignment Rules:

Weekdays only (Mon–Fri)

Valid shifts are like 6, 6-6, or 0630-1630

P/W must be evenly distributed

No one can have both P and W on the same day

P/W only assigned for shifts starting before 12 PM

Late shifts (like those with a 9) can only be W, not P

Known Limitations:

Single-digit shifts like “6” may not parse properly

Simple string matching for “9” may misread times like 19

No handling for missing or invalid Excel data

Assignment order is deterministic, not random

No unit tests yet

Suggested .gitignore:
pycache/
*.pyc
*.xlsx
*.docx
Created_Docs/
precomputed_pw_assignments.csv
.DS_Store
.env
