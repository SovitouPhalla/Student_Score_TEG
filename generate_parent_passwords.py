"""
generate_parent_passwords.py
---------------------------
Utility script to generate random 6-digit passwords for students with blank ParentPassword.

Run this ONCE to populate missing passwords:
    python generate_parent_passwords.py

This will:
1. Load the Students sheet from grades.xlsx
2. Find rows where ParentPassword is blank
3. Generate a random 6-digit integer (100000-999999) for each
4. Save the updated sheet back to grades.xlsx
5. Print a report showing which students got new passwords

IMPORTANT: Make sure grades.xlsx is NOT open in Excel before running this!
"""

import os
import random
import pandas as pd

# ── Configuration ──────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "grades.xlsx")

def generate_parent_passwords():
    """Generate and assign random 6-digit passwords to students with blank ParentPassword."""

    if not os.path.exists(EXCEL_PATH):
        print(f"ERROR: {EXCEL_PATH} not found.")
        print("Please make sure you're running this script from the project directory.")
        return

    print(f"Loading: {EXCEL_PATH}")

    try:
        # Load all sheets
        excel_file = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")

        students_df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="Students",
            engine="openpyxl",
            dtype={"StudentID": str, "Name": str, "ClassLabel": str, "ParentPassword": str},
        )

        grades_df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="Grades",
            engine="openpyxl",
            dtype={"StudentID": str},
        )

        teachers_df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="Teachers",
            engine="openpyxl",
            dtype={"Username": str, "Password": str, "Role": str},
        )

        admins_df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="Admins",
            engine="openpyxl",
            dtype={"Username": str, "Password": str},
        )

        approval_df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="ApprovalStatus",
            engine="openpyxl",
        )

    except Exception as e:
        print(f"ERROR reading Excel file: {e}")
        print("Make sure the file is not open in Excel and has the correct sheet names.")
        return

    # ── Find blank passwords ────────────────────────────────────────────────
    blank_mask = students_df["ParentPassword"].isna() | (students_df["ParentPassword"].astype(str).str.strip() == "")
    blank_count = blank_mask.sum()

    if blank_count == 0:
        print("\n✓ All students already have passwords. No changes needed.")
        return

    print(f"\nFound {blank_count} student(s) with blank ParentPassword.")
    print("Generating random 6-digit passwords...\n")

    # ── Generate passwords for blank entries ────────────────────────────────
    generated = []
    for idx in students_df[blank_mask].index:
        random_password = str(random.randint(100000, 999999))
        students_df.at[idx, "ParentPassword"] = random_password

        student_id = students_df.at[idx, "StudentID"]
        student_name = students_df.at[idx, "Name"]
        class_label = students_df.at[idx, "ClassLabel"]

        generated.append({
            "StudentID": student_id if student_id else "(blank)",
            "Name": student_name,
            "Class": class_label,
            "Password": random_password,
        })

    # ── Save back to Excel ──────────────────────────────────────────────────
    try:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            students_df.to_excel(writer, sheet_name="Students", index=False)
            grades_df.to_excel(writer, sheet_name="Grades", index=False)
            teachers_df.to_excel(writer, sheet_name="Teachers", index=False)
            admins_df.to_excel(writer, sheet_name="Admins", index=False)
            approval_df.to_excel(writer, sheet_name="ApprovalStatus", index=False)

        print("=" * 70)
        print("PASSWORDS GENERATED AND SAVED")
        print("=" * 70)
        for item in generated:
            print(f"\nStudent:     {item['Name']}")
            print(f"Student ID:  {item['StudentID']}")
            print(f"Class:       {item['Class']}")
            print(f"New Password: {item['Password']}")

        print("\n" + "=" * 70)
        print(f"✓ Successfully updated {len(generated)} student record(s)")
        print("=" * 70)
        print("\nNEXT STEPS:")
        print("1. Share these passwords with the school's administrative staff")
        print("2. Provide parents their child's name and password via secure channel")
        print("3. Parents can now log in using their child's FULL NAME + password\n")

    except PermissionError:
        print("ERROR: Cannot write to grades.xlsx - the file is open in Excel.")
        print("Please close Excel and try again.")
    except Exception as e:
        print(f"ERROR saving file: {e}")

if __name__ == "__main__":
    generate_parent_passwords()
