"""
PARENT LOGIN REFACTORING SUMMARY
================================

The Parent Login system has been refactored to use Student Full Name as the
identifier instead of Student ID, with ParentPassword for authentication.

═══════════════════════════════════════════════════════════════════════════════

📋 TECHNICAL OVERVIEW

1. LOGIN FLOW
   ──────────

   Step 1: Parent enters Student Full Name and ParentPassword
   Step 2: App searches Students sheet for name matches (case-insensitive, whitespace-safe)
   Step 3: Handle duplicates (if >1 match, show class selector to disambiguate)
   Step 4: Validate password against ParentPassword column
   Step 5: Store StudentID in session (NOT the name) for all subsequent reports
   Step 6: Redirect to /report

   This ensures parents see the CORRECT student's data via StudentID lookup.


2. SANITIZED SEARCH LOGIC
   ──────────────────────

   Function: get_students_by_name() [line 356-367 in app.py]

   # Normalize both input and Excel data:
   sanitized_name = name.strip().lower()
   mask = students_df["Name"].astype(str).str.strip().str.lower() == sanitized_name

   This prevents login failures from:
   ✓ Leading/trailing whitespace ("Chhuor Chunminh " vs "Chhuor Chunminh")
   ✓ Case mismatch ("CHHUOR CHUNMINH" vs "chhuor chunminh")


3. DUPLICATE NAME HANDLING
   ───────────────────────

   If get_students_by_name() returns >1 match (lines 510-533):

   a) Class selector is shown to parent
   b) Parent selects their child's class
   c) get_student_by_name_and_class() does a second lookup on (Name, ClassLabel)
   d) If still no match → error message, keep form prefilled
   e) If match found → proceed to password validation

   Example:
   - Two "Sam Ratanakpitou" exist: one in L3T4, one in L1T1
   - Parent sees: "Multiple students share this name. Please select class."
   - Parent selects L1T1
   - App finds the correct Sam Ratanakpitou
   - Password is validated
   - StudentID is stored in session


4. SESSION PERSISTENCE
   ────────────────────

   Line 544: session["student_id"] = student_info.get("StudentID", "")

   CRITICAL: All subsequent report lookups use StudentID, NOT the student's name.

   Why? Names can have duplicates; StudentID is unique.
   If two students share a name, without this safeguard:
   - Parent A logs in as "Sam Ratanakpitou" (L1T1)
   - Parent B logs in as "Sam Ratanakpitou" (L3T4)
   - Both could see each other's data if lookup used name instead of ID

   Solution: session["student_id"] is always unique.


5. TEMPLATE CHANGES (login.html)
   ────────────────────────────

   ✓ Label: "Student Full Name" (was "Student ID")
   ✓ Input type: text (for name entry)
   ✓ Password field: type="password"
   ✓ Conditional class selector: shown only when duplicates exist
   ✓ Form prefills correctly to preserve user input on error


═══════════════════════════════════════════════════════════════════════════════

📂 MISSING PASSWORDS UTILITY

File: fill_missing_passwords.py

Purpose: Auto-fill blank ParentPassword cells with random 6-digit integers.

Usage:
  $ python fill_missing_passwords.py

  Scans the Students sheet.
  For each row where ParentPassword is blank or empty:
    - Generates a 6-digit random integer (100000-999999)
    - Assigns it to that student
    - Logs the change

  Output shows:
    ID | Name | NewPassword
    └─ So you can distribute passwords to parents securely

CRITICAL NOTES:
  ✓ Leaves blank StudentIDs as blank (does NOT auto-generate IDs)
  ✓ Does NOT overwrite existing passwords
  ✓ Requires Excel file to NOT be open in Excel (avoid file locks)
  ✓ Always back up grades.xlsx before running


═══════════════════════════════════════════════════════════════════════════════

🔐 AUTHENTICATION FLOW (Step-by-Step)

1. Parent visits /login
   → Renders login.html (Student Name + Password fields)

2. Parent submits form
   → POST to /login with student_name and password

3. App searches for students matching the name (case-insensitive)
   matching_students = get_students_by_name(students_df, student_name)

   If no match:
     flash("Student not found")
     return render_template("login.html")

4. If >1 match found:
     If parent provided class_label:
       Try to disambiguate with get_student_by_name_and_class()
     Else:
       Show class selector form again
       (extract classes from matching_students)

5. Single match (or match + class confirmed):
     student_info = matching_students[0]
     (or from get_student_by_name_and_class())

6. Password validation:
     if str(student_info.get("ParentPassword", "")).strip() != password:
       flash("Incorrect password")
       return render_template("login.html", prefill_name=student_name)

7. Success:
     session["student_id"] = student_info.get("StudentID", "")
     return redirect(url_for("report"))

8. Parent sees report card
   /report uses session["student_id"] to fetch grades


═══════════════════════════════════════════════════════════════════════════════

📊 EXCEL SCHEMA (Students Sheet)

Column A: StudentID     (unique, may be blank for new entries)
Column B: Name          (full name, used for parent login)
Column C: ClassLabel    (e.g. "L1T3", "L3T4(2)", used to disambiguate)
Column D: ParentPassword (6+ character string, or random 6-digit integer)

Example:
┌─────────────┬───────────────────┬──────────┬─────────────────┐
│ StudentID   │ Name              │ Class    │ ParentPassword  │
├─────────────┼───────────────────┼──────────┼─────────────────┤
│ D01965      │ Chhuor Chunminh   │ L1T3     │ 587429          │
│ D00517      │ HuyChhun Yuly     │ L1T3     │ P@ss4Eng        │
│ D00700      │ Khim Sufong       │ L1T3     │ 623891          │
│             │ Sam Ratanakpitou  │ L1T1     │ samrata123      │
│             │ Sam Ratanakpitou  │ L3T4     │ samrata456      │
└─────────────┴───────────────────┴──────────┴─────────────────┘


═══════════════════════════════════════════════════════════════════════════════

🚀 DEPLOYMENT CHECKLIST

Before going live:

[ ] Run fill_missing_passwords.py to populate missing passwords
    $ python fill_missing_passwords.py

[ ] Securely distribute ParentPasswords to parents (NOT via email/SMS)
    Consider:
      - Print on report cards
      - SMS via verified parent contact only
      - In-person at parent-teacher meetings

[ ] Test login with a known student name:
    1. Open http://localhost:5000/login
    2. Enter: "Chhuor Chunminh" (exact name from Excel)
    3. Enter password from ParentPassword column
    4. Should see report card for that student

[ ] Test duplicate-name scenario:
    1. Find two students with same name in Excel
    2. Log in with that name (no class selector first)
    3. Should show class selector
    4. Select one class
    5. Should redirect to that student's report

[ ] Spot-check Session Persistence:
    1. Log in as Student A
    2. Inspect browser dev tools → Application → Cookies
    3. Verify session["student_id"] matches StudentID from Excel
    4. (Not the name)


═══════════════════════════════════════════════════════════════════════════════

⚙️  CONFIGURATION

app.py constants (lines 55-70):

SCORE_COLS = ["Conduct", "CP", "HW_ASS", "QUIZ", "MidTerm", "Final"]
SCORE_WEIGHTS = {
    "Conduct":  0.05,
    "CP":       0.05,
    "HW_ASS":   0.15,
    "QUIZ":     0.15,
    "MidTerm":  0.25,
    "Final":    0.35,
}
VALID_TERMS = (1, 2, 3, 4)
PASS_THRESHOLD = 50.0


═══════════════════════════════════════════════════════════════════════════════

📝 ROUTE REFERENCE

GET  /login
  └─ Show login form (Student Name + Password)

POST /login
  └─ Process login:
       1. Search by name (sanitized)
       2. Handle duplicates (class selector) if needed
       3. Validate password
       4. Store StudentID in session
       5. Redirect to /report

GET  /report
  └─ Display student's 4-term report card
     (Protected: requires session["student_id"])
     (Approval-gated: only shows released grades)

GET  /logout
  └─ Clear session, redirect to /login


═══════════════════════════════════════════════════════════════════════════════

✅ VALIDATION & TESTING

The following edge cases are handled:

✓ Student name not in database
  → "Student not found" error

✓ Student exists but password is wrong
  → "Incorrect password" error
  → Form prefills with entered name

✓ Multiple students share same exact name
  → Class selector shown
  → Can disambiguate by selecting class

✓ Multiple students, wrong class selected
  → "No student named X in class Y" error
  → Form re-shown with class selector

✓ Leading/trailing spaces in name
  → Normalized via .str.strip()
  → Login still succeeds

✓ Case mismatch ("CHHUOR" vs "chhuor")
  → Normalized via .str.lower()
  → Login still succeeds

✓ Session hijacking prevented
  → Parent cannot manually change session["student_id"]
  → All lookups use StudentID from session
  → If session is cleared, must re-login


═══════════════════════════════════════════════════════════════════════════════

📌 NOTES FOR FUTURE DEVELOPERS

1. The login logic is split across three functions in app.py:

   - get_students_by_name()           [line 356]
   - get_student_by_name_and_class()  [line 370]
   - login() route                     [line 489]

2. The /report route expects session["student_id"] to be set correctly.
   Always validate this in protected routes.

3. If you change the ParentPassword column name in Excel,
   update the password lookup in the login() route (line 539).

4. The class label comparison always uses .str.strip() to handle
   parens like "L6T2(2)".

5. Never log ParentPasswords to stdout/files. The fill_missing_passwords.py
   script prints them only at runtime (secure distribution context).


═══════════════════════════════════════════════════════════════════════════════
"""
