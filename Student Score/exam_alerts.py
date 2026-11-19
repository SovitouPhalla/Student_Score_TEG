"""
exam_alerts.py
--------------
Final-exam score-entry alerts.

Admins record the final exam date of each (department, term) because it
changes every term. From that date the teachers have SCORE_GRACE_DAYS days to
enter every student's Final score. While scores are missing, an alert is shown
in-app and (via send_pending_reminders) posted to a shared Telegram group.

Storage is a small JSON file instead of an Excel sheet: save_sheets() and
save_approval() in app.py rewrite the whole workbook and would drop unknown
sheets.

State file layout:
    {"exams": {"en:2": "2026-12-18"}, "sent": ["en:2:2026-12-18:open", ...]}

Configuration (environment variables):
    TELEGRAM_BOT_TOKEN   bot token from @BotFather
    TELEGRAM_CHAT_ID     id of the shared teachers' group (negative number)
    PORTAL_URL           optional link appended to reminders
    EXAM_ALERTS_PATH     optional override of the JSON state file location
"""

import json
import os
import tempfile
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATE_PATH = os.environ.get("EXAM_ALERTS_PATH") or os.path.join(BASE_DIR, "exam_alerts.json")

TIMEZONE = ZoneInfo("Asia/Phnom_Penh")
SCORE_GRACE_DAYS = 7          # scores are due this many days after the final exam
DUE_SOON_DAYS = 2             # "due soon" stage starts this many days before the deadline
STALE_AFTER_DAYS = 30         # stop alerting this long after the deadline
OVERDUE_REMINDER_DAYS = 7     # daily Telegram reminders while overdue, then stop

TELEGRAM_MAX_CHARS = 3900     # Telegram's hard limit is 4096

DEPT_NAMES = {"en": "English", "cn": "Chinese"}
STAGE_ORDER = {"overdue": 0, "due_today": 1, "due_soon": 2, "open": 3}


def today_local() -> date:
    return datetime.now(TIMEZONE).date()


# ── State file ─────────────────────────────────────────────────────────────────

def load_state() -> dict:
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}
    exams = data.get("exams") if isinstance(data.get("exams"), dict) else {}
    sent = data.get("sent") if isinstance(data.get("sent"), list) else []
    return {"exams": exams, "sent": sent}


def save_state(state: dict) -> None:
    """Write atomically so a crash never leaves a half-written file."""
    directory = os.path.dirname(STATE_PATH) or "."
    fd, tmp_path = tempfile.mkstemp(dir=directory, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(state, fh, indent=2, sort_keys=True)
        os.replace(tmp_path, STATE_PATH)
    except OSError:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def _exam_key(dept: str, term: int) -> str:
    return f"{dept}:{int(term)}"


def set_final_date(dept: str, term: int, final_date: date) -> None:
    state = load_state()
    state["exams"][_exam_key(dept, term)] = final_date.isoformat()
    save_state(state)


def delete_final_date(dept: str, term: int) -> None:
    state = load_state()
    state["exams"].pop(_exam_key(dept, term), None)
    save_state(state)


def list_exams(state: dict | None = None) -> list:
    """Return [{dept, term, final_date, due_date}] sorted by date. Bad entries are skipped."""
    state = state or load_state()
    exams = []
    for key, raw in state["exams"].items():
        try:
            dept, term = key.split(":")
            final_date = date.fromisoformat(raw)
            exams.append({
                "dept": dept,
                "term": int(term),
                "final_date": final_date,
                "due_date": final_date + timedelta(days=SCORE_GRACE_DAYS),
            })
        except (ValueError, TypeError):
            continue
    return sorted(exams, key=lambda e: (e["final_date"], e["dept"], e["term"]))


# ── Alert calculation ──────────────────────────────────────────────────────────

def _stage(days_left: int) -> str:
    if days_left < 0:
        return "overdue"
    if days_left == 0:
        return "due_today"
    if days_left <= DUE_SOON_DAYS:
        return "due_soon"
    return "open"


def compute_alerts(dept: str, students: list, scored: set, today: date | None = None,
                   state: dict | None = None) -> list:
    """
    Return one alert per (term, class) that still has Final scores missing.

    students: [{"id": str, "class": str, "terms": iterable of int}]
    scored:   {(student_id, term)} that already have a Final score
    An alert exists from the day after the final exam until the scores are all
    in, or STALE_AFTER_DAYS after the deadline, whichever comes first.
    """
    today = today or today_local()
    alerts = []
    for exam in list_exams(state):
        if exam["dept"] != dept:
            continue
        term, final_date, due_date = exam["term"], exam["final_date"], exam["due_date"]
        if today <= final_date or today > due_date + timedelta(days=STALE_AFTER_DAYS):
            continue
        days_left = (due_date - today).days
        missing_by_class: dict = {}
        for s in students:
            if term not in s["terms"] or (s["id"], term) in scored:
                continue
            missing_by_class[s["class"]] = missing_by_class.get(s["class"], 0) + 1
        for class_label in sorted(missing_by_class):
            alerts.append({
                "dept": dept,
                "term": term,
                "class": class_label,
                "missing": missing_by_class[class_label],
                "final_date": final_date,
                "due_date": due_date,
                "days_left": days_left,
                "stage": _stage(days_left),
            })
    alerts.sort(key=lambda a: (STAGE_ORDER[a["stage"]], a["dept"], a["term"], a["class"]))
    return alerts


# ── Telegram ───────────────────────────────────────────────────────────────────

def telegram_configured() -> bool:
    return bool(os.environ.get("TELEGRAM_BOT_TOKEN") and os.environ.get("TELEGRAM_CHAT_ID"))


def send_telegram(text: str) -> tuple:
    """Post text to the shared group. Returns (ok, error_message)."""
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return False, "TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID are not set."
    try:
        resp = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat_id, "text": text, "disable_web_page_preview": True},
            timeout=10,
        )
    except requests.RequestException as exc:
        # Never echo str(exc): it contains the request URL, which includes the bot token.
        return False, f"Could not reach Telegram ({type(exc).__name__})."
    if resp.status_code != 200:
        try:
            detail = resp.json().get("description", "")
        except ValueError:
            detail = ""
        return False, f"Telegram rejected the message (HTTP {resp.status_code}) {detail}".strip()
    return True, ""


_STAGE_HEADINGS = {
    "open": "Final scores are due",
    "due_soon": "Final scores are due SOON",
    "due_today": "Final scores are due TODAY",
    "overdue": "Final scores are OVERDUE",
}


def _deadline_text(alert: dict) -> str:
    due = alert["due_date"].strftime("%d %b %Y")
    d = alert["days_left"]
    if d < 0:
        return f"was due {due} ({-d} day{'s' if d != -1 else ''} ago)"
    if d == 0:
        return f"due today ({due})"
    return f"due {due} ({d} day{'s' if d != 1 else ''} left)"


def format_message(alerts: list) -> str:
    """Build one group message from alerts, grouped by (dept, term)."""
    lines = ["⏰ Reminder: please enter your class Final scores"]
    groups: dict = {}
    for a in alerts:
        groups.setdefault((a["dept"], a["term"]), []).append(a)
    for (dept, term), items in sorted(groups.items()):
        first = items[0]
        lines.append("")
        lines.append(
            f"{DEPT_NAMES.get(dept, dept)} – Term {term}: "
            f"{_STAGE_HEADINGS[first['stage']]}, {_deadline_text(first)}"
        )
        for a in items:
            lines.append(f"  • {a['class']}: {a['missing']} student(s) missing")
    portal = os.environ.get("PORTAL_URL")
    if portal:
        lines += ["", f"Enter scores: {portal}"]
    text = "\n".join(lines)
    if len(text) > TELEGRAM_MAX_CHARS:
        text = text[:TELEGRAM_MAX_CHARS].rsplit("\n", 1)[0] + "\n…(list truncated, see the portal)"
    return text


def _reminder_key(alert: dict, today: date) -> str | None:
    """Dedup key for one alert's current stage, or None if no reminder is due."""
    base = f"{alert['dept']}:{alert['term']}:{alert['final_date'].isoformat()}"
    if alert["stage"] == "overdue":
        if -alert["days_left"] > OVERDUE_REMINDER_DAYS:
            return None
        return f"{base}:overdue:{today.isoformat()}"
    return f"{base}:{alert['stage']}"


def send_pending_reminders(alerts: list, today: date | None = None, dry_run: bool = False,
                           force: bool = False) -> dict:
    """
    Send ONE message to the shared group covering every alert whose current
    stage has not been announced yet. Safe to run as often as you like (cron):
    each stage is announced once, overdue once per day. force=True ignores that
    history and resends every current alert (used by the admin "resend" button).

    Returns {"sent": bool, "count": int, "message": str, "error": str}.
    """
    today = today or today_local()
    state = load_state()
    already = set(state["sent"])

    fresh, keys = [], set()
    for alert in alerts:
        key = _reminder_key(alert, today)
        if key and (force or key not in already):
            fresh.append(alert)
            keys.add(key)

    result = {"sent": False, "count": len(fresh), "message": "", "error": ""}
    if not fresh:
        return result
    result["message"] = format_message(fresh)
    if dry_run:
        return result

    ok, error = send_telegram(result["message"])
    if not ok:
        result["error"] = error
        return result

    state["sent"] = sorted(already | keys)
    save_state(state)
    result["sent"] = True
    return result
