"""
add_admins_sheet.py
-------------------
One-time migration: adds an 'Admins' sheet to the existing grades.xlsx
without touching Students, Grades, or Teachers sheets.

Run once:
    python add_admins_sheet.py
"""

import os
import pandas as pd

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "grades.xlsx")

if not os.path.exists(EXCEL_PATH):
    print("[ERROR] grades.xlsx not found. Run init_project.py first.")
    raise SystemExit(1)

# Load all existing sheets so nothing is lost
students_df  = pd.read_excel(EXCEL_PATH, sheet_name="Students",  engine="openpyxl", dtype=str)
grades_df    = pd.read_excel(EXCEL_PATH, sheet_name="Grades",    engine="openpyxl", dtype=str)
teachers_df  = pd.read_excel(EXCEL_PATH, sheet_name="Teachers",  engine="openpyxl", dtype=str)

# Check whether Admins sheet already exists
try:
    existing_admins = pd.read_excel(EXCEL_PATH, sheet_name="Admins", engine="openpyxl", dtype=str)
    print(f"[SKIP] Admins sheet already exists ({len(existing_admins)} row(s)). Nothing changed.")
    raise SystemExit(0)
except ValueError:
    pass  # sheet not present — continue

# Default admin account
admins_df = pd.DataFrame({
    "Username": ["admin"],
    "Password": ["Admin@TEG2026"],
})

with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
    students_df.to_excel(writer, sheet_name="Students", index=False)
    grades_df.to_excel(writer,   sheet_name="Grades",   index=False)
    teachers_df.to_excel(writer, sheet_name="Teachers", index=False)
    admins_df.to_excel(writer,   sheet_name="Admins",   index=False)

print("[OK] Admins sheet added to grades.xlsx")
print("     Default account: username=admin  password=Admin@TEG2026")
print("     Change the password directly in grades.xlsx > Admins sheet.")
