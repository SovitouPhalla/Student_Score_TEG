# Student Term Report Portal

A Flask web application for **Tribal Education Group** that lets parents view student term report cards and allows teachers to enter grades per department.

---

## Features

- **Two departments** — English and Chinese, each with their own student database and scoring rules
- **Parent login** — Students log in by name + password to view their 4-term report card
- **Teacher portal** — Department-specific login; teachers enter scores which are held pending approval
- **HOD approval workflow** — Head of Department reviews and approves/rejects scores before they become visible to parents
- **Admin dashboard** — Add, edit, and delete student records
- **i18n support** — English and Khmer (ភាសាខ្មែរ) via Flask-Babel
- **Print-ready report cards** — Styled for A4 with a print button

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3, Flask |
| Data | pandas + openpyxl (Excel as database) |
| Templates | Jinja2 |
| i18n | Flask-Babel |
| Frontend | Plain HTML/CSS (no framework) |

---

## Project Structure

```
├── app.py                  # Main Flask application
├── grades.xlsx             # English dept database (git-ignored)
├── chinese_grades.xlsx     # Chinese dept database (git-ignored)
├── babel.cfg               # Babel extraction config
├── static/
│   ├── style.css
│   ├── teg-logo.png
│   └── ...
├── templates/
│   ├── layout.html         # Base template / nav bar
│   ├── landing.html        # Department selection page
│   ├── login.html          # English parent login
│   ├── login_chinese.html  # Chinese parent login
│   ├── report.html         # English report card
│   ├── report_chinese.html # Chinese report card
│   ├── teacher_login.html  # Teacher login (dept-aware)
│   ├── teacher_dept_select.html
│   ├── update.html         # Grade entry portal
│   ├── hod_dashboard.html  # HOD approval dashboard
│   ├── admin_login.html
│   └── admin_dashboard.html
└── translations/
    └── km/LC_MESSAGES/     # Khmer translations
```

---

## Setup

### 1. Create and activate a virtual environment

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
```

### 2. Install dependencies

```bash
pip install flask flask-babel pandas openpyxl
```

### 3. Initialise the Excel databases

```bash
python init_project.py
```

This creates `grades.xlsx` (English) and `chinese_grades.xlsx` (Chinese) with the required sheets.

### 4. Set a secret key

Set the `SECRET_KEY` environment variable before running in production:

```bash
# Generate a key
python -c "import secrets; print(secrets.token_hex(32))"

# Windows
set SECRET_KEY=your-generated-key-here
```

### 5. Run

```bash
python app.py
```

Open `http://127.0.0.1:5000` in a browser.

---

## Excel Database Schema

### `grades.xlsx` (English Department)

| Sheet | Columns |
|---|---|
| Students | StudentID, Name, ClassLabel, ParentPassword |
| Grades | StudentID, Term, Conduct, CP, HW_ASS, QUIZ, MidTerm, Final, FinalReport |
| Teachers | Username, Password, Role |
| Admins | Username, Password |
| ApprovalStatus | StudentID, Term, Approved, RequestNote |

### `chinese_grades.xlsx` (Chinese Department)

| Sheet | Columns |
|---|---|
| Students | No, Name, Class, Password |
| Grades | No, Term, Behavior, CP, Homework, Quiz, FinalTest, TotalGrade, Status |
| Teachers | Username, Password, Role |

---

## Scoring Weights

### English Department

| Category | Weight |
|---|---|
| Conduct | 5% |
| Class Participation | 5% |
| Homework & Assignments | 15% |
| Quiz | 15% |
| Mid-Term Exam | 25% |
| Final Exam | 35% |

Passing threshold: **50%**

### Chinese Department

| Category | Weight |
|---|---|
| Behavior | 10% |
| Class Participation | 10% |
| Homework | 20% |
| Quiz | 20% |
| Final Test | 40% |

Passing threshold: **60%**

---

## URL Reference

| URL | Description |
|---|---|
| `/` | Landing page — department selection |
| `/login` | English parent login |
| `/cn/login` | Chinese parent login |
| `/teacher/login` | Teacher department selection |
| `/teacher/login/en` | English teacher login |
| `/teacher/login/cn` | Chinese teacher login |
| `/update` | Teacher grade entry portal |
| `/hod/review` | HOD approval dashboard |
| `/admin/login` | Admin login |
| `/admin/dashboard` | Admin student management |

---

## Updating Translations

```bash
# Extract new strings
pybabel extract -F babel.cfg -o translations/messages.pot .

# Update existing .po files
pybabel update -i translations/messages.pot -d translations

# Compile .po to .mo
pybabel compile -d translations
```

---

## Important Notes

- **Do not open `grades.xlsx` in Excel while the app is running.** Excel places a lock file (`~$grades.xlsx`) that prevents Flask from saving changes.
- Excel files are git-ignored to protect live student data.
- Scores entered by teachers are **not visible to parents** until approved by the HOD.
