import o
import re
from datetime import datetime, timedelta
import pandas as pd
from collections import defaultdict
#---------Still working on alg----------------
# Alg to assign equal numbers of p and w each day
EXCEL_PATH = "Schedule.xlsx"
ASSIGNMENT_OUTPUT_PATH = "precomputed_pw_assignments.csv"

p_count = defaultdict(int)
w_count = defaultdict(int)
pw_schedule = defaultdict(lambda: {"P": None, "W": None})

# ----Shift Rules-------------------------
def is_valid_shift(shift):
    if not isinstance(shift, str):
        return False
    shift = shift.upper().strip()
    if any(bad in shift for bad in ["FSH", "PTO", "SH", "OT"]):
        return False
    return bool(re.search(r"\d+\s*-\s*\d+", shift))

def is_late_shift(shift):
    if not isinstance(shift, str):
        return False
    return "9" in shift

def is_eligible(name, shift, kind):
    if not is_valid_shift(shift):
        return False
    shift = shift.upper().strip()
    start_match = re.match(r"\d+", shift)
    if not start_match:
        return False
    start_time = int(start_match.group())
    hour = start_time // 100 if start_time >= 100 else start_time
    if hour >= 12:
        return False
    if kind == "P" and is_late_shift(shift):
        return False
    return True

#--------extract data-----------
def extract_assignments(sheet, date):
    assignments = []
    for block_start in range(0, len(sheet), 17):
        if block_start + 2 >= len(sheet):
            break
        date_row = sheet.iloc[block_start + 1]
        name_rows = sheet.iloc[block_start + 2:block_start + 15]
        for col_idx, cell in enumerate(date_row):
            if isinstance(cell, datetime) and cell.date() == date:
                for _, row in name_rows.iterrows():
                    name = row[0]
                    time = row[col_idx]
                    if pd.notna(time) and isinstance(name, str) and ("NP" in name or "PA" in name):
                        assignments.append((name.strip(), str(time).strip()))
                return assignments
    return []

# ---Main ---------------
def precompute_pw_assignments(start_date_str, end_date_str):
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
    xl = pd.ExcelFile(EXCEL_PATH)

    current = start_date
    while current <= end_date:
        if current.weekday() < 5:  # Weekdays only
            for sheet_name in xl.sheet_names:
                sheet = xl.parse(sheet_name, header=None)
                assigns = extract_assignments(sheet, current)
                if not assigns:
                    continue

                eligible_p = [n for n, t in assigns if is_eligible(n, t, "P")]
                eligible_w = [n for n, t in assigns if is_eligible(n, t, "W")]

                eligible_p = sorted(set(eligible_p), key=lambda n: (p_count[n], w_count[n]))
                eligible_w = sorted(set(eligible_w), key=lambda n: (w_count[n], p_count[n]))

                p_choice = next((n for n in eligible_p if p_count[n] <= w_count[n]), None)
                w_choice = next((n for n in eligible_w if w_count[n] <= p_count[n] and n != p_choice), None)

                if p_choice:
                    pw_schedule[str(current)]["P"] = p_choice
                    p_count[p_choice] += 1
                if w_choice:
                    pw_schedule[str(current)]["W"] = w_choice
                    w_count[w_choice] += 1
                break
        current += timedelta(days=1)

    # Save to CSV
    with open(ASSIGNMENT_OUTPUT_PATH, "w") as f:
        f.write("Date,P_Assignment,W_Assignment\n")
        for date_str in sorted(pw_schedule.keys()):
            p = pw_schedule[date_str]["P"] or ""
            w = pw_schedule[date_str]["W"] or ""
            f.write(f"{date_str},{p},{w}\n")
    print(f"Assignments saved to {ASSIGNMENT_OUTPUT_PATH}")

# === ENTRY POINT ===
if __name__ == "__main__":
    start = input("Enter start date (YYYY-MM-DD): ").strip()
    end = input("Enter end date (YYYY-MM-DD): ").strip()
    precompute_pw_assignments(start, end)

# every person has one week off 
# then every other week they must be P or W 
