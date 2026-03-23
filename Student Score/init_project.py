"""
init_project.py
---------------
Run this script ONCE to bootstrap the project.
It creates the required folder structure and a starter grades.xlsx file
with two sheets: Students and Grades.

IMPORTANT: If you are upgrading from a previous version, delete the old
grades.xlsx first — the schema now uses two separate sheets.

Required libraries:
    pip install pandas openpyxl flask
"""

import os
import pandas as pd

# ── 1. Create folders ──────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

folders = [
    os.path.join(BASE_DIR, "templates"),
    os.path.join(BASE_DIR, "static"),
]

for folder in folders:
    os.makedirs(folder, exist_ok=True)
    print(f"[OK] Folder ready: {folder}")

# ── 2. Create grades.xlsx (two sheets) ────────────────────────────────────────
EXCEL_PATH = os.path.join(BASE_DIR, "grades.xlsx")

# ── Sheet 1: Students ─────────────────────────────────────────────────────────
# One row per student; ClassLabel is stored as a plain string (handles parens).
students_data = {
    "StudentID":      ["S001",          "S002",          "S003",          "S004"],
    "Name":           ["Alice Johnson", "Bob Martinez",  "Carol Smith",   "David Lee"],
    "ClassLabel":     ["L6T2",          "L6T2",          "L6T2(2)",       "L10T4"],
    "ParentPassword": ["parent123",     "parent456",     "parent789",     "parentabc"],
}

students_df = pd.DataFrame(students_data, columns=["StudentID", "Name", "ClassLabel", "ParentPassword"])

# ── Sheet 2: Grades ───────────────────────────────────────────────────────────
# One row per (StudentID, Term) — up to 4 rows per student.
# Name and ParentPassword are NOT duplicated here; looked up from Students sheet.
# Weighted formula: Conduct*0.05 + CP*0.05 + HW_ASS*0.15 + QUIZ*0.15 + MidTerm*0.25 + Final*0.35
grades_data = {
    "StudentID":   ["S001",  "S001",  "S002",  "S003"],
    "Term":        [1,       2,       1,       1],
    "Conduct":     [88.0,    90.0,    78.0,    87.0],
    "CP":          [90.0,    88.0,    80.0,    85.0],
    "HW_ASS":      [85.0,    87.0,    72.0,    92.0],
    "QUIZ":        [80.0,    83.0,    74.0,    89.0],
    "MidTerm":     [78.0,    82.0,    70.0,    88.0],
    "Final":       [82.0,    85.0,    75.0,    90.0],
    "FinalReport": [81.85,   84.65,   73.55,   89.25],
    # S001-T1: (88*.05)+(90*.05)+(85*.15)+(80*.15)+(78*.25)+(82*.35) = 81.85
    # S001-T2: (90*.05)+(88*.05)+(87*.15)+(83*.15)+(82*.25)+(85*.35) = 84.65
    # S002-T1: (78*.05)+(80*.05)+(72*.15)+(74*.15)+(70*.25)+(75*.35) = 73.55
    # S003-T1: (87*.05)+(85*.05)+(92*.15)+(89*.15)+(88*.25)+(90*.35) = 89.25
}

grades_df = pd.DataFrame(
    grades_data,
    columns=["StudentID", "Term", "Conduct", "CP", "HW_ASS",
             "QUIZ", "MidTerm", "Final", "FinalReport"],
)

# ── Sheet 3: Teachers ─────────────────────────────────────────────────────────
# One row per teacher account.  Add more rows here or directly in Excel.
teachers_data = {
    "Username": ["admin",   "teacher1"],
    "Password": ["admin123", "pass456"],
}

teachers_df = pd.DataFrame(teachers_data, columns=["Username", "Password"])

if os.path.exists(EXCEL_PATH):
    print(f"[SKIP] grades.xlsx already exists — not overwriting.")
    print(f"       Delete it and re-run this script to apply the new schema.")
else:
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        students_df.to_excel(writer, sheet_name="Students", index=False)
        grades_df.to_excel(writer,   sheet_name="Grades",   index=False)
        teachers_df.to_excel(writer, sheet_name="Teachers", index=False)
    print(f"[OK] Created: {EXCEL_PATH}")
    print(f"     Sheet 'Students': 4 sample students across classes L6T2, L6T2(2), L10T4")
    print(f"     Sheet 'Grades':   Term data for S001 (T1+T2), S002 (T1), S003 (T1)")
    print(f"     Sheet 'Teachers': accounts — admin / admin123, teacher1 / pass456")

print("\nProject initialised successfully.")
print("Next step:  python app.py")
