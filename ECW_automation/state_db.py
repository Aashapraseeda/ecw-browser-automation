"""
state_db.py
-----------
Persistent tracking of which patients have already had a PediForms form
and PCareLink message sent, so repeated runs (e.g. 9 AM / 1 PM / 5 PM cron)
never resend to the same patient-visit twice.

Identity key: (acct_no, appointment_date) - not acct_no alone, since the
same patient will have a brand new visit months later that must still
be processed as "new".
"""

import os
import sqlite3
from datetime import datetime, timedelta

DB_PATH = os.getenv(
    "STATE_DB_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "patients_state.db"),
)


def _connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS patients (
            acct_no           TEXT NOT NULL,
            appointment_date  TEXT NOT NULL,
            last_name         TEXT,
            first_name        TEXT,
            visit_type        TEXT,
            form_name         TEXT,
            form_filename     TEXT,
            folder_name       TEXT,
            search_name       TEXT,
            status            TEXT NOT NULL DEFAULT 'form_sent',
            form_sent_at      TEXT,
            completed_at      TEXT,
            reminder_count    INTEGER NOT NULL DEFAULT 0,
            last_reminder_at  TEXT,
            PRIMARY KEY (acct_no, appointment_date)
        )
    """)
    return conn


def normalize_date(value):
    """Excel may hand back a datetime, date, or plain string - normalize to YYYY-MM-DD."""
    if hasattr(value, "date"):
        return value.date().isoformat()
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value).strip()


def is_known(acct_no, appointment_date):
    conn = _connect()
    row = conn.execute(
        "SELECT 1 FROM patients WHERE acct_no=? AND appointment_date=?",
        (acct_no, appointment_date),
    ).fetchone()
    conn.close()
    return row is not None


def insert_form_sent(patient):
    conn = _connect()
    conn.execute(
        """
        INSERT OR IGNORE INTO patients
        (acct_no, appointment_date, last_name, first_name, visit_type,
         form_name, form_filename, folder_name, search_name, status, form_sent_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'form_sent', ?)
        """,
        (
            patient["acct_no"], patient["appointment_date"], patient["last_name"],
            patient["first_name"], patient["visit_type"], patient["form_name"],
            patient["form_filename"], patient["folder_name"], patient["search_name"],
            datetime.utcnow().isoformat(),
        ),
    )
    conn.commit()
    conn.close()


def get_pending_patients():
    """Patients still waiting on form completion - includes ones from earlier runs/days."""
    conn = _connect()
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM patients WHERE status = 'form_sent'").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def mark_downloaded(acct_no, appointment_date):
    conn = _connect()
    conn.execute(
        "UPDATE patients SET status='downloaded' WHERE acct_no=? AND appointment_date=?",
        (acct_no, appointment_date),
    )
    conn.commit()
    conn.close()


def get_patients_needing_upload():
    """Have a downloaded PDF but aren't confirmed uploaded yet (includes retry of failed uploads)."""
    conn = _connect()
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM patients WHERE status = 'downloaded'").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def mark_completed(acct_no, appointment_date):
    conn = _connect()
    conn.execute(
        "UPDATE patients SET status='completed', completed_at=? WHERE acct_no=? AND appointment_date=?",
        (datetime.utcnow().isoformat(), acct_no, appointment_date),
    )
    conn.commit()
    conn.close()


def get_patients_needing_reminder(interval_hours, max_reminders):
    conn = _connect()
    conn.row_factory = sqlite3.Row
    cutoff = (datetime.utcnow() - timedelta(hours=interval_hours)).isoformat()
    rows = conn.execute(
        """
        SELECT * FROM patients
        WHERE status = 'form_sent'
        AND reminder_count < ?
        AND (
            (last_reminder_at IS NULL AND form_sent_at <= ?)
            OR last_reminder_at <= ?
        )
        """,
        (max_reminders, cutoff, cutoff),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def record_reminder_sent(acct_no, appointment_date):
    conn = _connect()
    conn.execute(
        """
        UPDATE patients SET reminder_count = reminder_count + 1, last_reminder_at = ?
        WHERE acct_no=? AND appointment_date=?
        """,
        (datetime.utcnow().isoformat(), acct_no, appointment_date),
    )
    conn.commit()
    conn.close()


def cleanup_old_completed(retention_days=30):
    """Delete completed records past the retention window. Returns count deleted."""
    conn = _connect()
    cutoff = (datetime.utcnow() - timedelta(days=retention_days)).isoformat()
    cur = conn.execute(
        "DELETE FROM patients WHERE status='completed' AND completed_at <= ?",
        (cutoff,),
    )
    conn.commit()
    deleted = cur.rowcount
    conn.close()
    return deleted
