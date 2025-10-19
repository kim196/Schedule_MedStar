# NP/PA Schedule Processor

A Python tool that automates scheduling for NP/PA staff.  
It reads from an Excel schedule and:

- Balances **Point (P)** and **W** assignments fairly across weekdays  
- Exports a CSV of all P/W assignments  
- Generates daily NP/PA assignment `.docx` files from a Word template  

Includes two main scripts:
- **Compute_PW.py** → Precomputes P/W assignments for a date range  
- **Filled_Assignment.py** → Fills out the NP/PA assignment template for a given date  

Built with **pandas** and **python** to make staff scheduling faster, more accurate, and easier to manage.

---

## Overview

This repository includes two main components:

1. **`Compute_PW.py`** – scans the Excel schedule across a date range and produces `precomputed_pw_assignments.csv`  
2. **`Filled_Assignment.py`** – reads the schedule for a single date and fills a Word template (`NPPA Assignments.docx`), producing daily `.docx` files in the `Created_Docs` folder

---

## File Structure

Schedule.xlsx # Input Excel schedule
NPPA Assignments.docx # Word template
Compute_PW.py # Precompute P/W assignments
Filled_Assignment.py # Generate DOCX for a single date
precomputed_pw_assignments.csv # Output CSV
Created_Docs/ # Folder for generated DOCX files
README.md # This file

