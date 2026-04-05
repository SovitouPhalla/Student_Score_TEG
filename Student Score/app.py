"""
app.py
------
Student Term Report Portal  —  4-Term + ClassLabel Edition
Flask + pandas/openpyxl (Excel as database)

Schema (grades.xlsx)
--------------------
  Sheet "Students":
    StudentID | Name | ClassLabel | ParentPassword

  Sheet "Grades":
    StudentID | Term | Midterm | Final | Participation |
    Homework  | Behavior | FinalReport

  Each student has exactly one row in Students.
  Each student has up to 4 rows in Grades (one per term).
  Name and ParentPassword live only in Students.

Routes
------
GET  /                  → redirect to /login
GET  /login             → parent login form
POST /login             → authenticate, redirect to /report
GET  /report            → 4-term summary report card (session-protected)
GET  /logout            → clear session, redirect to /login
GET  /update            → teacher: search form
POST /update/search     → teacher: look up student + term, show score form
POST /update/save       → teacher: validate, calculate, upsert row
"""

import os
import pandas as pd
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    flash,
)

# ── App setup ──────────────────────────────────────────────────────────────────
app = Flask(__name__)

# Read secret key from environment variable.
# On PythonAnywhere set this in the WSGI file:
#   os.environ['SECRET_KEY'] = 'your-random-secret-here'
# Generate a good value once with: python -c "import secrets; print(secrets.token_hex(32))"
app.secret_key = os.environ.get("SECRET_KEY", "change-me-before-deploying")

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "grades.xlsx")

SCORE_COLS = ["Conduct", "CP", "HW_ASS", "QUIZ", "MidTerm", "Final"]

# Weights must sum to 1.0
# Conduct 5% | CP 5% | HW/Assignments 15% | Quiz 15% | MidTerm 25% | Final 35%
SCORE_WEIGHTS = {
    "Conduct": 0.05,
    "CP":      0.05,
    "HW_ASS":  0.15,
    "QUIZ":    0.15,
    "MidTerm": 0.25,
    "Final":   0.35,
}

VALID_TERMS    = (1, 2, 3, 4)
PASS_THRESHOLD = 50.0

# ── Helpers: I/O ───────────────────────────────────────────────────────────────

def load_sheets():
    """
    Load Students and Grades sheets from grades.xlsx.
    Returns (students_df, grades_df).
    Raises FileNotFoundError or OSError (file locked / wrong format).
    """
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(
            "grades.xlsx not found. Please run init_project.py first."
        )
    try:
        students_df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="Students",
            engine="openpyxl",
            dtype={"StudentID": str, "ClassLabel": str, "ParentPassword": str},
        )
        grades_df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="Grades",
            engine="openpyxl",
            dtype={"StudentID": str},
        )
        return students_df, grades_df
    except PermissionError:
        raise OSError(
            "grades.xlsx appears to be open in another program (e.g. Excel). "
            "Please close it and try again."
        )
    except ValueError:
        raise OSError(
            "grades.xlsx is using the old single-sheet format. "
            "Please delete grades.xlsx and run init_project.py again to "
            "create the updated two-sheet version."
        )


def load_teachers() -> pd.DataFrame:
    """
    Load only the Teachers sheet.
    Returns a DataFrame with columns Username, Password, and Role.
    Returns an empty DataFrame if the sheet is missing (graceful fallback).
    """
    if not os.path.exists(EXCEL_PATH):
        return pd.DataFrame(columns=["Username", "Password", "Role"])
    try:
        return pd.read_excel(
            EXCEL_PATH,
            sheet_name="Teachers",
            engine="openpyxl",
            dtype={"Username": str, "Password": str, "Role": str},
        )
    except (ValueError, PermissionError):
        return pd.DataFrame(columns=["Username", "Password", "Role"])


def load_admins() -> pd.DataFrame:
    """Load only the Admins sheet. Returns empty DataFrame if missing."""
    if not os.path.exists(EXCEL_PATH):
        return pd.DataFrame(columns=["Username", "Password"])
    try:
        return pd.read_excel(
            EXCEL_PATH,
            sheet_name="Admins",
            engine="openpyxl",
            dtype={"Username": str, "Password": str},
        )
    except (ValueError, PermissionError):
        return pd.DataFrame(columns=["Username", "Password"])


def load_approval() -> pd.DataFrame:
    """
    Load the ApprovalStatus sheet.
    Columns: StudentID (str) | Term (int) | Approved (bool) | RequestNote (str)
    Migration: old schema used ClassLabel — returns empty DataFrame so it is
    rebuilt on the next save_approval() call.
    Returns an empty DataFrame if the sheet is missing.
    """
    if not os.path.exists(EXCEL_PATH):
        return pd.DataFrame(columns=["StudentID", "Term", "Approved", "RequestNote"])
    try:
        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="ApprovalStatus",
            engine="openpyxl",
            dtype={"StudentID": str, "RequestNote": str},
        )
        # Migration guard: discard old ClassLabel-keyed data
        if "StudentID" not in df.columns:
            return pd.DataFrame(columns=["StudentID", "Term", "Approved", "RequestNote"])
        if "RequestNote" not in df.columns:
            df["RequestNote"] = ""
        df["RequestNote"] = df["RequestNote"].fillna("")
        return df
    except (ValueError, PermissionError):
        return pd.DataFrame(columns=["StudentID", "Term", "Approved", "RequestNote"])


def is_approved(approval_df: pd.DataFrame, student_id: str, term: int) -> bool:
    """Return True only if the StudentID+Term row exists and Approved == True."""
    try:
        mask = (
            (approval_df["StudentID"].astype(str).str.strip() == str(student_id).strip()) &
            (approval_df["Term"].astype(int) == int(term))
        )
    except (KeyError, ValueError):
        return False
    match = approval_df[mask]
    if match.empty:
        return False
    return bool(match.iloc[0]["Approved"])


def get_approval_row(approval_df: pd.DataFrame, student_id: str, term: int):
    """
    Return the full approval row as a dict for a StudentID+Term, or None.
    Guarantees RequestNote is always a stripped string (never NaN).
    """
    try:
        mask = (
            (approval_df["StudentID"].astype(str).str.strip() == str(student_id).strip()) &
            (approval_df["Term"].astype(int) == int(term))
        )
    except (KeyError, ValueError):
        return None
    match = approval_df[mask]
    if match.empty:
        return None
    row = match.iloc[0].to_dict()
    row["RequestNote"] = str(row.get("RequestNote", "") or "").strip()
    return row


def term_review_status(approval_df: pd.DataFrame, student_id: str, term: int) -> str:
    """
    Returns the review status string for a StudentID+Term:
      'approved'           — Approved == True
      'changes_requested'  — Approved == False and RequestNote is non-empty
      'pending'            — everything else (no row, or row with no note)
    """
    row = get_approval_row(approval_df, student_id, term)
    if row is None:
        return "pending"
    if bool(row.get("Approved")):
        return "approved"
    if row.get("RequestNote", "").strip():
        return "changes_requested"
    return "pending"


def _upsert_approval(approval_df: pd.DataFrame, student_id: str, term: int,
                     approved: bool, note: str) -> pd.DataFrame:
    """
    Insert or update a single StudentID+Term row in approval_df.
    Returns the modified DataFrame.
    """
    try:
        mask = (
            (approval_df["StudentID"].astype(str).str.strip() == str(student_id).strip()) &
            (approval_df["Term"].astype(int) == int(term))
        )
    except (KeyError, ValueError):
        mask = pd.Series([], dtype=bool)

    if approval_df[mask].empty:
        new_row = pd.DataFrame([{
            "StudentID":   str(student_id),
            "Term":        int(term),
            "Approved":    approved,
            "RequestNote": note,
        }])
        approval_df = pd.concat([approval_df, new_row], ignore_index=True)
    else:
        approval_df.loc[mask, "Approved"]    = approved
        approval_df.loc[mask, "RequestNote"] = note
    return approval_df


def save_sheets(students_df: pd.DataFrame, grades_df: pd.DataFrame) -> None:
    """
    Persist Students and Grades back to grades.xlsx.
    Teachers, Admins, and ApprovalStatus sheets are loaded first and re-written unchanged.
    """
    teachers_df  = load_teachers()
    admins_df    = load_admins()
    approval_df  = load_approval()
    try:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            students_df.to_excel(writer,  sheet_name="Students",       index=False)
            grades_df.to_excel(writer,    sheet_name="Grades",         index=False)
            teachers_df.to_excel(writer,  sheet_name="Teachers",       index=False)
            admins_df.to_excel(writer,    sheet_name="Admins",         index=False)
            approval_df.to_excel(writer,  sheet_name="ApprovalStatus", index=False)
    except PermissionError:
        raise OSError(
            "Cannot save grades.xlsx — the file is open in another program. "
            "Please close it and try again."
        )


def save_approval(approval_df: pd.DataFrame) -> None:
    """
    Persist only the ApprovalStatus sheet, preserving all other sheets unchanged.
    Raises OSError if the file is locked.
    """
    try:
        students_df, grades_df = load_sheets()
    except (FileNotFoundError, OSError):
        students_df = pd.DataFrame(columns=["StudentID", "Name", "ClassLabel", "ParentPassword"])
        grades_df   = pd.DataFrame(columns=["StudentID", "Term", "Conduct", "CP",
                                            "HW_ASS", "QUIZ", "MidTerm", "Final", "FinalReport"])
    teachers_df = load_teachers()
    admins_df   = load_admins()
    try:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            students_df.to_excel(writer,  sheet_name="Students",       index=False)
            grades_df.to_excel(writer,    sheet_name="Grades",         index=False)
            teachers_df.to_excel(writer,  sheet_name="Teachers",       index=False)
            admins_df.to_excel(writer,    sheet_name="Admins",         index=False)
            approval_df.to_excel(writer,  sheet_name="ApprovalStatus", index=False)
    except PermissionError:
        raise OSError(
            "Cannot save grades.xlsx — the file is open in another program. "
            "Please close it and try again."
        )


# ── Helpers: calculations ──────────────────────────────────────────────────────

def calc_final(row: pd.Series) -> float:
    """
    Weighted final grade:
      Conduct 5% | CP 5% | HW_ASS 15% | QUIZ 15% | MidTerm 25% | Final 35%
    """
    return round(
        sum(float(row[col]) * weight for col, weight in SCORE_WEIGHTS.items()),
        2,
    )


# ── Helpers: Students sheet ────────────────────────────────────────────────────

def get_class_labels(students_df: pd.DataFrame) -> list:
    """Return a sorted list of unique ClassLabel strings."""
    labels = students_df["ClassLabel"].dropna().astype(str).unique().tolist()
    return sorted(labels)


def get_students_by_class(students_df: pd.DataFrame, class_label: str) -> list:
    """
    Return a list of {StudentID, Name} dicts for a given ClassLabel.
    ClassLabel comparison is always string-based (safe with parens like 'L6T2(2)').
    """
    mask = students_df["ClassLabel"].astype(str) == str(class_label)
    subset = students_df[mask][["StudentID", "Name"]].copy()
    return subset.to_dict(orient="records")


def get_class_students_map(students_df: pd.DataFrame) -> dict:
    """
    Build a mapping of ClassLabel -> [{StudentID, Name}, ...]
    for embedding as JSON in the teacher UI.
    """
    result = {}
    for label in get_class_labels(students_df):
        result[label] = get_students_by_class(students_df, label)
    return result


def get_student_info(students_df: pd.DataFrame, student_id: str):
    """
    Return the Students sheet row for a given StudentID as a dict, or None.
    Also accepts '__row:{idx}' for students with blank StudentIDs.
    """
    student_id_str = str(student_id).strip()
    if student_id_str.startswith("__row:"):
        try:
            row_idx = int(student_id_str.split(":")[1])
            if row_idx in students_df.index:
                return students_df.loc[row_idx].to_dict()
        except (ValueError, IndexError):
            pass
        return None
    mask = students_df["StudentID"].astype(str).str.strip() == str(student_id).strip()
    match = students_df[mask]
    if match.empty:
        return None
    return match.iloc[0].to_dict()


def get_students_by_name(students_df: pd.DataFrame, name: str) -> list:
    """
    Search for students by name with case-insensitive, whitespace-tolerant matching.
    Returns a list of dicts (one or more matches), or empty list if no match.
    CRITICAL: Uses .str.strip() and .str.lower() for robust matching.
    """
    sanitized_name = name.strip().lower()
    mask = students_df["Name"].astype(str).str.strip().str.lower() == sanitized_name
    matches = students_df[mask]
    if matches.empty:
        return []
    return matches.reset_index().rename(columns={"index": "_row_idx"}).to_dict(orient="records")


def get_student_by_name_and_class(students_df: pd.DataFrame, name: str, class_label: str):
    """
    Search for a specific student by both name AND class label (to disambiguate duplicates).
    Returns a single dict, or None if not found.
    """
    sanitized_name = name.strip().lower()
    mask = (
        (students_df["Name"].astype(str).str.strip().str.lower() == sanitized_name) &
        (students_df["ClassLabel"].astype(str).str.strip() == class_label.strip())
    )
    match = students_df[mask]
    if match.empty:
        return None
    return match.reset_index().rename(columns={"index": "_row_idx"}).iloc[0].to_dict()


def get_teacher(username: str):
    """
    Return the Teachers sheet row for a given username as a dict, or None.
    Username comparison is case-insensitive.
    """
    teachers_df = load_teachers()
    mask = teachers_df["Username"].astype(str).str.strip().str.lower() == username.strip().lower()
    match = teachers_df[mask]
    if match.empty:
        return None
    return match.iloc[0].to_dict()


def get_admin(username: str):
    """
    Return the Admins sheet row for a given username as a dict, or None.
    Username comparison is case-insensitive.
    """
    admins_df = load_admins()
    mask = admins_df["Username"].astype(str).str.strip().str.lower() == username.strip().lower()
    match = admins_df[mask]
    if match.empty:
        return None
    return match.iloc[0].to_dict()


def _admin_required():
    """Return a redirect Response if user is not an admin, else None."""
    if session.get("role") != "admin":
        flash("Admin access required. Please log in.", "warning")
        return redirect(url_for("admin_login"))
    return None


def _hod_required():
    """Return a redirect Response if user is not an HOD teacher, else None."""
    if session.get("is_hod") != True:
        flash("HOD access required. Please log in with an HOD account.", "warning")
        return redirect(url_for("teacher_login"))
    return None


# ── Helpers: Grades sheet ──────────────────────────────────────────────────────

def get_student_term(grades_df: pd.DataFrame, student_id: str, term: int):
    """Return the row matching (StudentID, Term) as a dict, or None."""
    try:
        mask = (
            (grades_df["StudentID"].astype(str).str.strip() == student_id.strip()) &
            (grades_df["Term"].astype(int) == int(term))
        )
    except (KeyError, ValueError):
        return None
    match = grades_df[mask]
    if match.empty:
        return None
    return match.iloc[0].to_dict()


def get_all_terms(grades_df: pd.DataFrame, student_id: str) -> dict:
    """
    Return {1: row_dict_or_None, 2: ..., 3: ..., 4: ...} for a student.
    Terms with no data are None — displayed as 'Not Yet Released'.
    """
    return {t: get_student_term(grades_df, student_id, t) for t in VALID_TERMS}


@app.route("/debug/check-teacher", methods=["GET"])
def debug_check_teacher():
    """Debug endpoint - show teachers data."""
    teachers_df = load_teachers()
    html = "<h1>Teachers Sheet</h1>"
    html += f"<p>Columns: {list(teachers_df.columns)}</p>"
    html += "<table border='1' cellpadding='5'>"
    html += "<tr>" + "".join(f"<th>{col}</th>" for col in teachers_df.columns) + "</tr>"
    for _, row in teachers_df.iterrows():
        html += "<tr>" + "".join(f"<td>{row[col]}</td>" for col in teachers_df.columns) + "</tr>"
    html += "</table>"
    html += f"<hr><h2>Current Session:</h2>"
    html += f"<p>is_hod: <strong>{session.get('is_hod')}</strong></p>"
    html += f"<p>teacher_role: <strong>{session.get('teacher_role')}</strong></p>"
    html += f"<p>teacher_user: <strong>{session.get('teacher_user')}</strong></p>"
    return html


@app.route("/test-session")
def test_session():
    return f"""
    <h1>Session State</h1>
    <p><strong>teacher_user:</strong> {session.get('teacher_user')}</p>
    <p><strong>teacher_role:</strong> {session.get('teacher_role')}</p>
    <p><strong>is_hod:</strong> {session.get('is_hod')}</p>
    <hr>
    <a href="/update">Back to Update</a> | <a href="/teacher/logout">Logout</a>
    """

@app.route("/")
def index():
    return redirect(url_for("login"))


# ── Parent: login / report / logout ───────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        student_name = request.form.get("student_name", "").strip()
        password     = request.form.get("password", "").strip()
        class_label  = request.form.get("class_label", "").strip()

        try:
            students_df, _ = load_sheets()
        except (FileNotFoundError, OSError) as exc:
            flash(str(exc), "error")
            return render_template("login.html")

        # ── Step 1: Search for students by name ──────────────────────────────
        matching_students = get_students_by_name(students_df, student_name)

        if not matching_students:
            flash("Student not found. Please check the name and try again.", "error")
            return render_template("login.html")

        # ── Step 2: Handle multiple matches ──────────────────────────────────
        if len(matching_students) > 1:
            # If class_label was provided, try to disambiguate
            if class_label:
                student_info = get_student_by_name_and_class(students_df, student_name, class_label)
                if student_info is None:
                    flash(f"No student named '{student_name}' found in class '{class_label}'.", "error")
                    # Re-show the form with the class selector populated
                    matching_classes = [s.get("ClassLabel", "") for s in matching_students]
                    return render_template(
                        "login.html",
                        prefill_name=student_name,
                        prefill_class=class_label,
                        show_class_selector=True,
                        matching_classes=sorted(set(matching_classes)),
                    )
            else:
                # Show the clarification form with class selector
                matching_classes = [s.get("ClassLabel", "") for s in matching_students]
                return render_template(
                    "login.html",
                    prefill_name=student_name,
                    show_class_selector=True,
                    matching_classes=sorted(set(matching_classes)),
                )
        else:
            # Single match found
            student_info = matching_students[0]

        # ── Step 3: Validate password ───────────────────────────────────────
        if str(student_info.get("ParentPassword", "")).strip() != password:
            flash("Incorrect password.", "error")
            return render_template("login.html", prefill_name=student_name)

        # ── Step 4: Authenticate and store session ──────────────────────────
        raw_id = str(student_info.get("StudentID", "") or "").strip()
        if not raw_id or raw_id == "nan":
            session["student_id"] = f"__row:{student_info.get('_row_idx', 0)}"
        else:
            session["student_id"] = raw_id
        return redirect(url_for("report"))

    return render_template("login.html")


@app.route("/report")
def report():
    if "student_id" not in session:
        flash("Please log in to view the report.", "warning")
        return redirect(url_for("login"))

    try:
        students_df, grades_df = load_sheets()
    except (FileNotFoundError, OSError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("login"))

    student_info = get_student_info(students_df, session["student_id"])
    if student_info is None:
        flash("Student record not found.", "error")
        session.clear()
        return redirect(url_for("login"))

    # Build per-term data (None = not yet released)
    all_terms = get_all_terms(grades_df, session["student_id"])

    # ── Approval gate ──────────────────────────────────────────────────────────
    # Even if scores exist in Grades, parents only see them once the
    # Head of Department has approved that specific StudentID + Term.
    approval_df    = load_approval()
    student_id_key = session["student_id"]
    all_terms = {
        t: (data if (data is not None and is_approved(approval_df, student_id_key, t)) else None)
        for t, data in all_terms.items()
    }
    # ──────────────────────────────────────────────────────────────────────────

    # Year-to-date average: average FinalReport of all released terms only
    completed = [t for t in all_terms.values() if t is not None]
    if completed:
        ytd_avg    = round(
            sum(float(t["FinalReport"]) for t in completed) / len(completed), 2
        )
        ytd_passed = ytd_avg >= PASS_THRESHOLD
    else:
        ytd_avg    = None
        ytd_passed = False

    return render_template(
        "report.html",
        student      = student_info,
        all_terms    = all_terms,
        score_cols   = SCORE_COLS,
        score_weights = SCORE_WEIGHTS,
        ytd_avg      = ytd_avg,
        ytd_passed   = ytd_passed,
        threshold    = PASS_THRESHOLD,
        valid_terms  = VALID_TERMS,
    )


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


# ── Teacher: authentication ────────────────────────────────────────────────────

@app.route("/teacher/login", methods=["GET", "POST"])
def teacher_login():
    # Already logged in — go straight to the portal
    if session.get("teacher_user"):
        return redirect(url_for("update"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        teacher = get_teacher(username)

        if teacher is None or str(teacher["Password"]).strip() != password:
            flash("Invalid username or password.", "error")
            return render_template("teacher_login.html")

        session["teacher_user"] = teacher["Username"]
        session["teacher_role"] = teacher.get("Role", "Teacher")
        session["is_hod"] = (teacher.get("Role", "").strip().lower() == "hod")

        # DEBUG: Print to console
        print(f"\n=== LOGIN DEBUG ===")
        print(f"Username: {teacher['Username']}")
        print(f"Role column value: {repr(teacher.get('Role'))}")
        print(f"teacher_role set to: {session['teacher_role']}")
        print(f"is_hod set to: {session['is_hod']}")
        print(f"==================\n")

        flash(f"Welcome, {teacher['Username']}!", "success")
        return redirect(url_for("update"))

    return render_template("teacher_login.html")


@app.route("/teacher/logout")
def teacher_logout():
    session.pop("teacher_user", None)
    session.pop("teacher_role", None)
    session.pop("is_hod", None)
    flash("Teacher session ended.", "info")
    return redirect(url_for("teacher_login"))


# ── HOD: Score approval ─────────────────────────────────────────────────────

@app.route("/hod/review", methods=["GET", "POST"])
def hod_review():
    """
    HOD (Head of Department) approval dashboard — per-student, per-term.
    GET  → class/term selector; ?class_label=X&term=Y shows the student list.
    POST → approve / revoke / request_changes (single) or batch_approve (multiple).
    Security: Only accessible to teachers with Role='HOD'; guided by _hod_required().
    """
    guard = _hod_required()
    if guard:
        return guard

    approval_df = load_approval()

    # ── Handle actions ──────────────────────────────────────────────────────────
    if request.method == "POST":
        class_label  = request.form.get("class_label",  "").strip()
        term_raw     = request.form.get("term",          "").strip()
        action       = request.form.get("action",        "").strip()
        request_note = request.form.get("request_note",  "").strip()

        try:
            term = int(term_raw)
            if term not in VALID_TERMS:
                raise ValueError
        except (ValueError, TypeError):
            flash("Invalid term value.", "error")
            return redirect(url_for("hod_review", class_label=class_label))

        redirect_back = redirect(url_for(
            "hod_review", class_label=class_label, term=term
        ))

        # ── Batch approve ──────────────────────────────────────────────────────
        if action == "batch_approve":
            student_ids = request.form.getlist("student_ids")
            if not student_ids:
                flash("No students selected for batch approval.", "warning")
                return redirect_back
            for sid in student_ids:
                approval_df = _upsert_approval(approval_df, sid, term, True, "")
            try:
                save_approval(approval_df)
            except OSError as exc:
                flash(str(exc), "error")
                return redirect_back
            flash(
                f"✔ Approved {len(student_ids)} student(s) for "
                f"{class_label} — Term {term}. Scores are now visible to parents.",
                "success",
            )
            return redirect_back

        # ── Single-student action ──────────────────────────────────────────────
        student_id = request.form.get("student_id", "").strip()
        if not student_id:
            flash("Missing student ID.", "error")
            return redirect_back

        if action == "approve":
            new_approved, new_note = True, ""
            msg      = f"✔ Approved: {student_id} — Term {term}. Score is now visible to parents."
            msg_type = "success"
        elif action == "revoke":
            new_approved, new_note = False, ""
            msg      = f"Revoked: {student_id} — Term {term}. Score is now hidden."
            msg_type = "info"
        elif action == "request_changes":
            new_approved, new_note = False, request_note
            msg      = f"Changes requested for {student_id} — Term {term}. Teacher has been notified."
            msg_type = "warning"
        else:
            flash("Unknown action.", "error")
            return redirect_back

        approval_df = _upsert_approval(approval_df, student_id, term, new_approved, new_note)
        try:
            save_approval(approval_df)
        except OSError as exc:
            flash(str(exc), "error")
            return redirect_back

        flash(msg, msg_type)
        return redirect_back

    # ── Build data for template ────────────────────────────────────────────────
    try:
        students_df, grades_df = load_sheets()
    except (FileNotFoundError, OSError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("teacher_login"))

    class_labels = get_class_labels(students_df)
    sel_class    = request.args.get("class_label", "").strip()
    try:
        sel_term = int(request.args.get("term", "1"))
        if sel_term not in VALID_TERMS:
            sel_term = 1
    except (ValueError, TypeError):
        sel_term = 1

    student_rows = []
    if sel_class:
        for s in get_students_by_class(students_df, sel_class):
            sid          = str(s.get("StudentID", "")).strip()
            grade_row    = get_student_term(grades_df, sid, sel_term) if sid else None
            approval_row = get_approval_row(approval_df, sid, sel_term) if sid else None
            student_rows.append({
                "student_id": sid,
                "name":       s.get("Name", ""),
                "grades":     grade_row,
                "status":     term_review_status(approval_df, sid, sel_term) if sid else "pending",
                "note":       approval_row["RequestNote"] if approval_row else "",
            })

    approved_count = sum(1 for r in student_rows if r["status"] == "approved")
    changes_count  = sum(1 for r in student_rows if r["status"] == "changes_requested")
    pending_count  = sum(1 for r in student_rows
                         if r["status"] == "pending" and r["grades"] is not None)
    nodata_count   = sum(1 for r in student_rows if r["grades"] is None)

    return render_template(
        "hod_dashboard.html",
        class_labels   = class_labels,
        sel_class      = sel_class,
        sel_term       = sel_term,
        student_rows   = student_rows,
        valid_terms    = VALID_TERMS,
        score_cols     = SCORE_COLS,
        approved_count = approved_count,
        changes_count  = changes_count,
        pending_count  = pending_count,
        nodata_count   = nodata_count,
    )


@app.route("/hod/student_preview/<student_id>")
def hod_student_preview(student_id):
    """
    HOD preview of a student's report card.
    Shows all available scores regardless of approval status.
    Includes "HOD DRAFT" watermark to clarify this is a preview before formal approval.
    Security: Only accessible to HOD teachers.
    """
    guard = _hod_required()
    if guard:
        return guard

    try:
        students_df, grades_df = load_sheets()
    except (FileNotFoundError, OSError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("hod_review"))

    student_info = get_student_info(students_df, student_id)
    if student_info is None:
        flash(f'Student "{student_id}" not found.', "error")
        return redirect(url_for("hod_review"))

    # No approval gate — show all available term data
    all_terms = get_all_terms(grades_df, student_id)

    completed = [t for t in all_terms.values() if t is not None]
    if completed:
        ytd_avg    = round(
            sum(float(t["FinalReport"]) for t in completed) / len(completed), 2
        )
        ytd_passed = ytd_avg >= PASS_THRESHOLD
    else:
        ytd_avg    = None
        ytd_passed = False

    return render_template(
        "report.html",
        student          = student_info,
        all_terms        = all_terms,
        score_cols       = SCORE_COLS,
        score_weights    = SCORE_WEIGHTS,
        ytd_avg          = ytd_avg,
        ytd_passed       = ytd_passed,
        threshold        = PASS_THRESHOLD,
        valid_terms      = VALID_TERMS,
        preview_mode     = True,
        hod_preview      = True,
    )


# ── Teacher: update scores ─────────────────────────────────────────────────────

def _render_update(student=None, term=None, class_label=None,
                   class_labels=None, class_students_map=None, error=None,
                   changes_requested=None):
    """Central render helper — keeps all three route functions DRY."""
    if error:
        flash(error, "error")
    return render_template(
        "update.html",
        student              = student,
        term                 = term,
        class_label          = class_label,
        class_labels         = class_labels or [],
        class_students_map   = class_students_map or {},
        score_cols           = SCORE_COLS,
        score_weights        = SCORE_WEIGHTS,
        valid_terms          = VALID_TERMS,
        changes_requested    = changes_requested or [],
    )


def _load_for_update():
    """
    Load both sheets and build the class metadata needed by the teacher
    page.  Returns (students_df, grades_df, class_labels, class_students_map)
    or raises FileNotFoundError / OSError.
    """
    students_df, grades_df = load_sheets()
    class_labels       = get_class_labels(students_df)
    class_students_map = get_class_students_map(students_df)
    return students_df, grades_df, class_labels, class_students_map


def _validate_term(raw: str):
    """
    Return (int_term, None) on success or (None, error_message) on failure.
    """
    try:
        t = int(raw)
        if t not in VALID_TERMS:
            raise ValueError
        return t, None
    except (ValueError, TypeError):
        return None, f"Term must be 1, 2, 3, or 4. Received: '{raw}'."


@app.route("/update", methods=["GET"])
def update():
    if "teacher_user" not in session:
        flash("Please log in as a teacher to access this page.", "warning")
        return redirect(url_for("teacher_login"))
    prefill_id    = request.args.get("student_id",  "").strip()
    prefill_term  = request.args.get("term",         "").strip()
    prefill_class = request.args.get("class_label",  "").strip()

    try:
        students_df, grades_df, class_labels, class_students_map = _load_for_update()
    except (FileNotFoundError, OSError) as exc:
        return _render_update(error=str(exc))

    # Build list of student+term combinations with "Changes Requested" status
    approval_df = load_approval()
    changes_requested = []
    for cl in class_labels:
        for s in get_students_by_class(students_df, cl):
            sid = str(s.get("StudentID", "")).strip()
            if not sid:
                continue
            for t in VALID_TERMS:
                if term_review_status(approval_df, sid, t) == "changes_requested":
                    row = get_approval_row(approval_df, sid, t)
                    changes_requested.append({
                        "student_id":   sid,
                        "student_name": s.get("Name", ""),
                        "class_label":  cl,
                        "term":         t,
                        "note":         row["RequestNote"] if row else "",
                    })

    if prefill_id and prefill_term:
        term, err = _validate_term(prefill_term)
        if err:
            return _render_update(
                class_labels=class_labels,
                class_students_map=class_students_map,
                changes_requested=changes_requested,
                error=err,
            )
        student = get_student_term(grades_df, prefill_id, term)
        if student is not None:
            info = get_student_info(students_df, prefill_id)
            if info:
                student["Name"]       = info["Name"]
                student["ClassLabel"] = info["ClassLabel"]
        return _render_update(
            student=student, term=term,
            class_label=prefill_class,
            class_labels=class_labels,
            class_students_map=class_students_map,
            changes_requested=changes_requested,
        )

    return _render_update(
        class_labels=class_labels,
        class_students_map=class_students_map,
        changes_requested=changes_requested,
    )


@app.route("/update/search", methods=["POST"])
def update_search():
    """
    Phase 1 → Phase 2 transition.
    Validates ClassLabel + StudentID + Term, then either loads the existing
    grade row or prepares a blank entry for a new term.
    """
    if "teacher_user" not in session:
        flash("Please log in as a teacher to access this page.", "warning")
        return redirect(url_for("teacher_login"))
    class_label = request.form.get("class_label", "").strip()
    student_id  = request.form.get("student_id",  "").strip()
    term_raw    = request.form.get("term",         "").strip()

    try:
        students_df, grades_df, class_labels, class_students_map = _load_for_update()
    except (FileNotFoundError, OSError) as exc:
        return _render_update(error=str(exc))

    if not student_id:
        return _render_update(
            class_label=class_label,
            class_labels=class_labels,
            class_students_map=class_students_map,
            error="Please select a student before searching.",
        )

    term, err = _validate_term(term_raw)
    if err:
        return _render_update(
            class_label=class_label,
            class_labels=class_labels,
            class_students_map=class_students_map,
            error=err,
        )

    # Verify the student exists in the Students sheet
    student_info = get_student_info(students_df, student_id)
    if student_info is None:
        return _render_update(
            class_label=class_label,
            class_labels=class_labels,
            class_students_map=class_students_map,
            error=f'No student found with ID "{student_id}". '
                  f"Please check the selection and try again.",
            term=term,
        )

    # Find the specific term grade row
    student = get_student_term(grades_df, student_id, term)

    if student is None:
        # Student exists but this term hasn't been entered yet.
        student = {
            "StudentID":     student_info["StudentID"],
            "Term":          term,
            "Name":          student_info["Name"],
            "ClassLabel":    student_info["ClassLabel"],
            "Midterm":       "",
            "Final":         "",
            "Participation": "",
            "Homework":      "",
            "Behavior":      "",
            "FinalReport":   "",
            "_is_new_term":  True,
        }
    else:
        # Merge Name and ClassLabel from Students sheet
        student["Name"]       = student_info["Name"]
        student["ClassLabel"] = student_info["ClassLabel"]

    return _render_update(
        student=student, term=term,
        class_label=class_label,
        class_labels=class_labels,
        class_students_map=class_students_map,
    )


@app.route("/update/save", methods=["POST"])
def update_save():
    """
    Phase 2 submission.
    Validates all inputs, then either updates the existing (StudentID, Term)
    row in Grades or inserts a new one.  Recalculates FinalReport before saving.
    """
    if "teacher_user" not in session:
        flash("Please log in as a teacher to access this page.", "warning")
        return redirect(url_for("teacher_login"))
    class_label = request.form.get("class_label", "").strip()
    student_id  = request.form.get("student_id",  "").strip()
    term_raw    = request.form.get("term",         "").strip()

    # ── Step 1: validate term ──────────────────────────────────────────────────
    term, term_err = _validate_term(term_raw)
    if term_err:
        return _render_update(error=term_err)

    # ── Step 2: validate all five score inputs ─────────────────────────────────
    scores = {}
    for col in SCORE_COLS:
        raw = request.form.get(col, "").strip()
        try:
            value = float(raw)
            if not (0.0 <= value <= 100.0):
                raise ValueError(f"out of range: {value}")
        except ValueError:
            try:
                students_df, grades_df, class_labels, class_students_map = _load_for_update()
                student = get_student_term(grades_df, student_id, term)
                if student is not None:
                    info = get_student_info(students_df, student_id)
                    if info:
                        student["Name"]       = info["Name"]
                        student["ClassLabel"] = info["ClassLabel"]
            except (FileNotFoundError, OSError):
                student        = None
                class_labels   = []
                class_students_map = {}
            flash(
                f'"{col}" must be a number between 0 and 100. '
                f"Received: '{raw}'",
                "error",
            )
            return render_template(
                "update.html",
                student=student, term=term,
                class_label=class_label,
                class_labels=class_labels,
                class_students_map=class_students_map,
                score_cols=SCORE_COLS, score_weights=SCORE_WEIGHTS, valid_terms=VALID_TERMS,
            )
        scores[col] = value

    # ── Step 3: load workbook ──────────────────────────────────────────────────
    try:
        students_df, grades_df, class_labels, class_students_map = _load_for_update()
    except (FileNotFoundError, OSError) as exc:
        return _render_update(error=str(exc), term=term)

    student_info = get_student_info(students_df, student_id)
    if student_info is None:
        return _render_update(
            class_labels=class_labels,
            class_students_map=class_students_map,
            error=f'Student "{student_id}" not found.',
            term=term,
        )

    student_name = student_info["Name"]

    # ── Step 4: update existing row OR insert new row in Grades ───────────────
    try:
        mask = (
            (grades_df["StudentID"].astype(str).str.strip() == student_id) &
            (grades_df["Term"].astype(int) == term)
        )
    except KeyError:
        mask = pd.Series([False] * len(grades_df))

    idx = grades_df.index[mask]

    if not idx.empty:
        # UPDATE existing term row
        row_idx = idx[0]
        for col, value in scores.items():
            grades_df.at[row_idx, col] = value
        new_final = calc_final(grades_df.loc[row_idx])
        grades_df.at[row_idx, "FinalReport"] = new_final

    else:
        # INSERT new term row
        new_row = {
            "StudentID": student_info["StudentID"],
            "Term":      term,
            **scores,
            "FinalReport": 0.0,  # placeholder; recalculated below
        }
        temp = pd.Series(new_row)
        new_row["FinalReport"] = calc_final(temp)
        new_final = new_row["FinalReport"]

        grades_df = pd.concat(
            [grades_df, pd.DataFrame([new_row])], ignore_index=True
        )
        grades_df = grades_df.sort_values(["StudentID", "Term"]).reset_index(drop=True)

    # ── Step 5: persist ────────────────────────────────────────────────────────
    try:
        save_sheets(students_df, grades_df)
    except OSError as exc:
        student = get_student_term(grades_df, student_id, term)
        if student is not None:
            student["Name"]       = student_name
            student["ClassLabel"] = student_info["ClassLabel"]
        flash(str(exc), "error")
        return render_template(
            "update.html",
            student=student, term=term,
            class_label=class_label,
            class_labels=class_labels,
            class_students_map=class_students_map,
            score_cols=SCORE_COLS, score_weights=SCORE_WEIGHTS, valid_terms=VALID_TERMS,
        )

    flash(
        f"✓ Term {term} scores saved for {student_name} "
        f"(Final Report: {new_final}). "
        f"Results are now pending review — they will become visible to parents "
        f"once approved by the Head of Department.",
        "success",
    )
    return redirect(url_for(
        "update",
        student_id=student_id,
        term=term,
        class_label=class_label,
    ))


# ── Admin: authentication ──────────────────────────────────────────────────────

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    # If HOD is somehow logged in as teacher, redirect them to HOD portal
    if session.get("is_hod") and session.get("teacher_user"):
        return redirect(url_for("hod_review"))

    if session.get("role") == "admin":
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        admin = get_admin(username)

        if admin is None or str(admin["Password"]).strip() != password:
            flash("Invalid username or password.", "error")
            return render_template("admin_login.html")

        session["role"]       = "admin"
        session["admin_user"] = admin["Username"]
        flash(f"Welcome, {admin['Username']}!", "success")
        return redirect(url_for("admin_dashboard"))

    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.pop("role", None)
    session.pop("admin_user", None)
    flash("Admin session ended.", "info")
    return redirect(url_for("admin_login"))


# ── Admin: dashboard (roster + filter + search + inline edit form) ─────────────

@app.route("/admin/dashboard")
def admin_dashboard():
    guard = _admin_required()
    if guard:
        return guard

    try:
        students_df, _ = load_sheets()
    except (FileNotFoundError, OSError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("admin_login"))

    filter_class = request.args.get("filter_class", "").strip()
    q            = request.args.get("q", "").strip().lower()
    edit_idx_raw = request.args.get("edit", "").strip()

    class_labels = get_class_labels(students_df)

    # Attach integer row index for CRUD operations (handles blank StudentIDs)
    df = students_df.copy()
    df["_row_idx"] = df.index

    if filter_class:
        df = df[df["ClassLabel"].astype(str) == filter_class]
    if q:
        name_match = df["Name"].astype(str).str.lower().str.contains(q, na=False)
        id_match   = df["StudentID"].astype(str).str.lower().str.contains(q, na=False)
        df = df[name_match | id_match]

    students = df.fillna("").to_dict(orient="records")

    # Resolve the row to edit, if requested
    edit_student = None
    if edit_idx_raw:
        try:
            ridx = int(edit_idx_raw)
            if ridx in students_df.index:
                row = students_df.loc[ridx].fillna("")
                edit_student = {
                    "_row_idx":      ridx,
                    "StudentID":     str(row["StudentID"]),
                    "Name":          str(row["Name"]),
                    "ClassLabel":    str(row["ClassLabel"]),
                    "ParentPassword": str(row["ParentPassword"]),
                }
        except (ValueError, KeyError):
            pass

    return render_template(
        "admin_dashboard.html",
        students     = students,
        class_labels = class_labels,
        filter_class = filter_class,
        q            = q,
        edit_student = edit_student,
        total        = len(students_df),
        filtered     = len(students),
    )


# ── Admin: Add student ─────────────────────────────────────────────────────────

@app.route("/admin/student/add", methods=["POST"])
def admin_student_add():
    guard = _admin_required()
    if guard:
        return guard

    student_id  = request.form.get("student_id",  "").strip()
    name        = request.form.get("name",         "").strip()
    class_label = request.form.get("class_label",  "").strip()
    password    = request.form.get("password",     "").strip()
    # Preserve current filter for redirect
    fc = request.form.get("filter_class", "")
    q  = request.form.get("q", "")

    if not name:
        flash("Name is required.", "error")
        return redirect(url_for("admin_dashboard", filter_class=fc, q=q))
    if not class_label:
        flash("Class is required.", "error")
        return redirect(url_for("admin_dashboard", filter_class=fc, q=q))
    if not password:
        flash("Password is required.", "error")
        return redirect(url_for("admin_dashboard", filter_class=fc, q=q))

    try:
        students_df, grades_df = load_sheets()
    except (FileNotFoundError, OSError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("admin_dashboard"))

    # Duplicate StudentID check (only if ID was provided)
    if student_id:
        existing = students_df[
            students_df["StudentID"].astype(str).str.strip() == student_id
        ]
        if not existing.empty:
            flash(f'Student ID "{student_id}" already exists.', "error")
            return redirect(url_for("admin_dashboard", filter_class=fc, q=q))

    new_row = pd.DataFrame([{
        "StudentID":     student_id,
        "Name":         name,
        "ClassLabel":   class_label,
        "ParentPassword": password,
    }])
    students_df = pd.concat([students_df, new_row], ignore_index=True)

    try:
        save_sheets(students_df, grades_df)
    except OSError as exc:
        flash(str(exc), "error")
        return redirect(url_for("admin_dashboard", filter_class=fc, q=q))

    flash(f'Student "{name}" added to {class_label}.', "success")
    return redirect(url_for("admin_dashboard", filter_class=class_label))


# ── Admin: Edit student ────────────────────────────────────────────────────────

@app.route("/admin/student/edit", methods=["POST"])
def admin_student_edit():
    guard = _admin_required()
    if guard:
        return guard

    try:
        row_idx = int(request.form.get("row_idx", ""))
    except (ValueError, TypeError):
        flash("Invalid student reference.", "error")
        return redirect(url_for("admin_dashboard"))

    name        = request.form.get("name",        "").strip()
    class_label = request.form.get("class_label", "").strip()
    student_id  = request.form.get("student_id",  "").strip()
    password    = request.form.get("password",    "").strip()

    if not name:
        flash("Name is required.", "error")
        return redirect(url_for("admin_dashboard", edit=row_idx))
    if not class_label:
        flash("Class is required.", "error")
        return redirect(url_for("admin_dashboard", edit=row_idx))
    if not password:
        flash("Password is required.", "error")
        return redirect(url_for("admin_dashboard", edit=row_idx))

    try:
        students_df, grades_df = load_sheets()
    except (FileNotFoundError, OSError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("admin_dashboard"))

    if row_idx not in students_df.index:
        flash("Student record not found.", "error")
        return redirect(url_for("admin_dashboard"))

    old_id = str(students_df.at[row_idx, "StudentID"]).strip()

    # If StudentID changed, propagate to Grades as well
    if student_id != old_id:
        if student_id:
            col    = students_df["StudentID"].astype(str).str.strip()
            others = col[students_df.index != row_idx]
            if student_id in others.values:
                flash(f'Student ID "{student_id}" is already in use.', "error")
                return redirect(url_for("admin_dashboard", edit=row_idx))
        grades_mask = grades_df["StudentID"].astype(str).str.strip() == old_id
        grades_df.loc[grades_mask, "StudentID"] = student_id
        students_df.at[row_idx, "StudentID"] = student_id

    students_df.at[row_idx, "Name"]           = name
    students_df.at[row_idx, "ClassLabel"]     = class_label
    students_df.at[row_idx, "ParentPassword"] = password

    try:
        save_sheets(students_df, grades_df)
    except OSError as exc:
        flash(str(exc), "error")
        return redirect(url_for("admin_dashboard", edit=row_idx))

    flash(f'Student "{name}" updated.', "success")
    return redirect(url_for("admin_dashboard", filter_class=class_label))


# ── Admin: Delete student ──────────────────────────────────────────────────────

@app.route("/admin/student/delete", methods=["POST"])
def admin_student_delete():
    guard = _admin_required()
    if guard:
        return guard

    try:
        row_idx = int(request.form.get("row_idx", ""))
    except (ValueError, TypeError):
        flash("Invalid student reference.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        students_df, grades_df = load_sheets()
    except (FileNotFoundError, OSError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("admin_dashboard"))

    if row_idx not in students_df.index:
        flash("Student record not found.", "error")
        return redirect(url_for("admin_dashboard"))

    student_name    = str(students_df.at[row_idx, "Name"])
    student_id_val  = str(students_df.at[row_idx, "StudentID"]).strip()
    filter_class    = str(students_df.at[row_idx, "ClassLabel"])

    # Remove from Students
    students_df = students_df.drop(index=row_idx).reset_index(drop=True)

    # Remove all grade rows for this student
    if student_id_val:
        grades_mask = grades_df["StudentID"].astype(str).str.strip() == student_id_val
        grades_df   = grades_df[~grades_mask].reset_index(drop=True)

    try:
        save_sheets(students_df, grades_df)
    except OSError as exc:
        flash(str(exc), "error")
        return redirect(url_for("admin_dashboard"))

    flash(f'Student "{student_name}" deleted.', "success")
    return redirect(url_for("admin_dashboard", filter_class=filter_class))


# ── Admin: Approval dashboard ──────────────────────────────────────────────────

@app.route("/admin/approve_scores", methods=["GET", "POST"])
def approve_scores():
    """
    Redirect old admin approval URL to HOD review (migration endpoint).
    Preserves GET parameters for seamless transition.
    """
    return redirect(url_for("hod_review",
                           class_label=request.args.get("class_label"),
                           term=request.args.get("term")))


@app.route("/admin/student_preview/<student_id>")
def admin_student_preview(student_id):
    """
    Admin preview of a student's report card.
    Shows all available scores regardless of approval status so the HOD can
    review grades before deciding whether to approve or request changes.
    """
    guard = _admin_required()
    if guard:
        return guard

    try:
        students_df, grades_df = load_sheets()
    except (FileNotFoundError, OSError) as exc:
        flash(str(exc), "error")
        return redirect(url_for("approve_scores"))

    student_info = get_student_info(students_df, student_id)
    if student_info is None:
        flash(f'Student "{student_id}" not found.', "error")
        return redirect(url_for("approve_scores"))

    # No approval gate — show all available term data
    all_terms = get_all_terms(grades_df, student_id)

    completed = [t for t in all_terms.values() if t is not None]
    if completed:
        ytd_avg    = round(
            sum(float(t["FinalReport"]) for t in completed) / len(completed), 2
        )
        ytd_passed = ytd_avg >= PASS_THRESHOLD
    else:
        ytd_avg    = None
        ytd_passed = False

    return render_template(
        "report.html",
        student       = student_info,
        all_terms     = all_terms,
        score_cols    = SCORE_COLS,
        score_weights = SCORE_WEIGHTS,
        ytd_avg       = ytd_avg,
        ytd_passed    = ytd_passed,
        threshold     = PASS_THRESHOLD,
        valid_terms   = VALID_TERMS,
        preview_mode  = True,
    )



# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(debug=False)
