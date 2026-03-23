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
import secrets
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
app.secret_key = secrets.token_hex(32)

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
            dtype={"StudentID": str, "ClassLabel": str},
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
    Returns a DataFrame with columns Username and Password.
    Returns an empty DataFrame if the sheet is missing (graceful fallback).
    """
    if not os.path.exists(EXCEL_PATH):
        return pd.DataFrame(columns=["Username", "Password"])
    try:
        return pd.read_excel(
            EXCEL_PATH,
            sheet_name="Teachers",
            engine="openpyxl",
            dtype={"Username": str, "Password": str},
        )
    except (ValueError, PermissionError):
        return pd.DataFrame(columns=["Username", "Password"])


def save_sheets(students_df: pd.DataFrame, grades_df: pd.DataFrame) -> None:
    """
    Persist Students and Grades back to grades.xlsx.
    The Teachers sheet is loaded first and re-written unchanged so it is
    never accidentally deleted when saving grade data.
    """
    teachers_df = load_teachers()
    try:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            students_df.to_excel(writer,  sheet_name="Students", index=False)
            grades_df.to_excel(writer,    sheet_name="Grades",   index=False)
            teachers_df.to_excel(writer,  sheet_name="Teachers", index=False)
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
    Used for login auth, Name/Password lookup, and ClassLabel display.
    """
    mask = students_df["StudentID"].astype(str).str.strip() == student_id.strip()
    match = students_df[mask]
    if match.empty:
        return None
    return match.iloc[0].to_dict()


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


# ── Root ───────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return redirect(url_for("login"))


# ── Parent: login / report / logout ───────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        student_id = request.form.get("student_id", "").strip()
        password   = request.form.get("password", "").strip()

        try:
            students_df, _ = load_sheets()
        except (FileNotFoundError, OSError) as exc:
            flash(str(exc), "error")
            return render_template("login.html")

        student_info = get_student_info(students_df, student_id)

        if student_info is None:
            flash("Student not found. Please check the Student ID.", "error")
            return render_template("login.html")

        if str(student_info["ParentPassword"]).strip() != password:
            flash("Incorrect password.", "error")
            return render_template("login.html")

        session["student_id"] = student_id
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
        flash(f"Welcome, {teacher['Username']}!", "success")
        return redirect(url_for("update"))

    return render_template("teacher_login.html")


@app.route("/teacher/logout")
def teacher_logout():
    session.pop("teacher_user", None)
    flash("Teacher session ended.", "info")
    return redirect(url_for("teacher_login"))


# ── Teacher: update scores ─────────────────────────────────────────────────────

def _render_update(student=None, term=None, class_label=None,
                   class_labels=None, class_students_map=None, error=None):
    """Central render helper — keeps all three route functions DRY."""
    if error:
        flash(error, "error")
    return render_template(
        "update.html",
        student            = student,
        term               = term,
        class_label        = class_label,
        class_labels       = class_labels or [],
        class_students_map = class_students_map or {},
        score_cols         = SCORE_COLS,
        score_weights      = SCORE_WEIGHTS,
        valid_terms        = VALID_TERMS,
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
    """
    Show the teacher search form.
    After a successful save the redirect includes ?student_id=&term=&class_label=
    so the teacher lands back on the same student/term (PRG pattern).
    """
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

    if prefill_id and prefill_term:
        term, err = _validate_term(prefill_term)
        if err:
            return _render_update(
                class_labels=class_labels,
                class_students_map=class_students_map,
                error=err,
            )
        student = get_student_term(grades_df, prefill_id, term)
        # Merge student info (Name, ClassLabel) from Students sheet
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
        )

    return _render_update(
        class_labels=class_labels,
        class_students_map=class_students_map,
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
        f"✓ Term {term} scores saved for {student_name}. "
        f"Final Report: {new_final}",
        "success",
    )
    return redirect(url_for(
        "update",
        student_id=student_id,
        term=term,
        class_label=class_label,
    ))


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(debug=True)
