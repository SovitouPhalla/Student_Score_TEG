"""
fill_missing_passwords.py
-------------------------
Utility script to fill missing ParentPasswords with random 6-digit integers.

Usage:
    python fill_missing_passwords.py

This will:
  1. Load the Students sheet from grades.xlsx
  2. For any row where ParentPassword is blank/NaN, generate a 6-digit random integer
  3. Save the updated sheet back to Excel
  4. Print a report showing which students received new passwords

CRITICAL: This script leaves blank StudentIDs as blank (does NOT auto-generate IDs).
"""

import os
import pandas as pd
import random
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "grades.xlsx")


def generate_password() -> str:
    """Generate a random 6-digit integer as a string."""
    return str(random.randint(100000, 999999))


def fill_missing_passwords() -> None:
    """
    Load Students sheet, fill missing ParentPasswords, and save back to Excel.
    Prints a summary of changes made.
    """
    if not os.path.exists(EXCEL_PATH):
        print(f"ERROR: {EXCEL_PATH} not found. Please run init_project.py first.")
        return

    try:
        students_df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="Students",
            engine="openpyxl",
            dtype={"StudentID": str, "ClassLabel": str, "ParentPassword": str},
        )
    except PermissionError:
        print(
            "ERROR: grades.xlsx appears to be open in another program (e.g. Excel).\n"
            "       Please close it and try again."
        )
        return
    except ValueError:
        print(
            "ERROR: Could not find 'Students' sheet in grades.xlsx.\n"
            "       Please verify the file structure."
        )
        return

    # Identify rows with missing ParentPassword
    missing_mask = students_df["ParentPassword"].isna() | (
        students_df["ParentPassword"].astype(str).str.strip() == ""
    )
    missing_count = missing_mask.sum()

    if missing_count == 0:
        print("[OK] All students have ParentPasswords. No changes needed.")
        return

    print(f"[INFO] Found {missing_count} student(s) with missing ParentPassword.")
    print("       Generating random 6-digit passwords...\n")

    # Generate passwords for missing rows
    changes = []
    for idx in students_df[missing_mask].index:
        student_id = str(students_df.at[idx, "StudentID"]).strip()
        name = str(students_df.at[idx, "Name"]).strip()
        new_password = generate_password()
        students_df.at[idx, "ParentPassword"] = new_password
        changes.append(
            {
                "StudentID": student_id or "(blank)",
                "Name": name,
                "NewPassword": new_password,
            }
        )

    # Save back to Excel
    try:
        # Load all sheets to preserve them
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
            dtype={"StudentID": str},
        )
    except (ValueError, PermissionError) as e:
        print(
            f"ERROR: Could not load all sheets from grades.xlsx: {e}\n"
            "       Please verify the file is valid and not open."
        )
        return

    try:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            students_df.to_excel(writer, sheet_name="Students", index=False)
            grades_df.to_excel(writer, sheet_name="Grades", index=False)
            teachers_df.to_excel(writer, sheet_name="Teachers", index=False)
            admins_df.to_excel(writer, sheet_name="Admins", index=False)
            approval_df.to_excel(writer, sheet_name="ApprovalStatus", index=False)
    except PermissionError:
        print(
            "ERROR: Cannot save grades.xlsx — the file is open in another program.\n"
            "       Please close it and try again."
        )
        return

    # Print summary
    print("[OK] Updated {0} student(s) with new passwords:\n".format(missing_count))
    for change in changes:
        print(
            f"  • ID: {change['StudentID']:10s} | Name: {change['Name']:25s} | Password: {change['NewPassword']}"
        )

    print(f"\n[OK] Successfully saved {EXCEL_PATH}")
    print(
        "\nIMPORTANT: Share these passwords with parents via a secure channel.\n"
        "           Do NOT send them in plain text via email or SMS."
    )


if __name__ == "__main__":
    fill_missing_passwords()
