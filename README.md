# Intro / Issue

This project was created to solve a real workflow issue at **MedStar**.  
Originally, the scheduling team only had access to a large Excel spreadsheet showing who was working, when, and on what type of shift.  

Each day, someone had to **manually review that spreadsheet**, **write out a daily assignment sheet**, and **assign “P” (Point) and “W”** roles, ensuring that over time, every provider received an equal number of each assignment.  

However, there were many constraints:
- Certain shifts (like **OT**, **FSH**, or **PTO**) could not be assigned P or W  
- Assignments only applied to **weekday day shifts** (before 12 PM)  
- Late shifts (like those with a 9 PM end) could only be assigned **W**, not **P**

This tool automates that entire process.

It reads the **Excel schedule.xlsx**, applies all assignment rules, and automatically:
- Balances **P/W (Point / W)** roles fairly across providers  
- Exports a **CSV** with all precomputed P/W assignments  
- Generates **daily `.docx` assignment sheets** from a Word template

Together, these scripts replace hours of manual scheduling work with a few simple commands.

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

1. **`Compute_PW.py`** – scans the Excel schedule across a date range and produces `precomputed_pw_assignments.csv`, utilizing an algorithm that will assign P-Point person and W - a person assigned with a doctor that day 
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

