import sys
import os
from datetime import datetime, timedelta
import pandas as pd

# Add main/ to path so we can import your logic
sys.path.append(os.path.abspath("../main"))

from Filled_Assignment import extract_daily_assignments, fill_half_sheet, p_assignments, w_assignments

EXCEL_PATH = "../main/Schedule.xlsx"

def test_distribution(start_date_str, num_days=42):
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
    xl = pd.ExcelFile(EXCEL_PATH)

    for i in range(num_days):
        current_date = start_date + timedelta(days=i)
        if current_date.weekday() > 4:  # Skip weekends
            continue
        date_str = current_date.strftime("%Y-%m-%d")

        for sheet_name in xl.sheet_names:
            sheet = xl.parse(sheet_name, header=None)
            assignments = extract_daily_assignments(sheet, date_str)
            if assignments:
                fill_half_sheet(date_str, assignments)
                break

    print("\n=== Point (P) Assignments ===")
    for name, count in sorted(p_assignments.items(), key=lambda x: -x[1]):
        print(f"{name}: {count}")

    print("\n=== Wang (W) Assignments ===")
    for name, count in sorted(w_assignments.items(), key=lambda x: -x[1]):
        print(f"{name}: {count}")

if __name__ == "__main__":
    test_distribution("2025-06-09", num_days=42)
