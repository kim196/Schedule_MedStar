# NP/PA Schedule Processor

A Python utility to parse an Excel staff schedule and automatically produce:

- A CSV of precomputed **P/W (Point / W)** assignments across a date range  
- Populated **NP/PA assignment `.docx` files** for individual dates using a Word template  

This project automates scheduling for NP/PA staff by balancing point and W assignments fairly across weekdays.

---

## 🧠 Overview

This repository includes two main components:

1. **`Compute_PW.py`** – scans the Excel schedule across a date range and produces `precomputed_pw_assignments.csv`  
2. **`Filled_Assignment.py`** – reads the schedule for a single date and fills a Word template (`NPPA Assignments.docx`), producing daily `.docx` files in the `Created_Docs` folder

---

## 📁 Project Structure

