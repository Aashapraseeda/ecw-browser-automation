"""
database/state_db.py
---------------------
Persistent tracking of which patients have already had a Patient Forms Now
form sent, so repeated runs (e.g. 9 AM / 1 PM / 5 PM cron) never resend to
the same patient-visit twice.

Identity key: (acct_no, appointment_date) - not acct_no alone, since the
same patient will have a brand new visit months later that must still
be processed as "new".

Ported verbatim from the reference project (ECW_automation/state_db.py) -
only the DB path now comes from config.settings instead of a bare os.getenv.
"""

import os
import sqlite3
from datetime import datetime, timedelta

from config import settings

DB_PATH = settings.STATE_DB_PATH


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


def delete_by_acct_no(acct_numbers):
    """
    Delete ALL records (any status) for the given account numbers,
    regardless of appointment_date. Used by demo/test pipelines to reset
    known test patients before each run, so repeated testing against the
    same patients isn't blocked by "already processed" bookkeeping from a
    prior run. Not used by the production pipeline.
    """
    acct_numbers = list(acct_numbers)
    if not acct_numbers:
        return 0
    conn = _connect()
    placeholders = ",".join("?" for _ in acct_numbers)
    cur = conn.execute(
        f"DELETE FROM patients WHERE acct_no IN ({placeholders})",
        acct_numbers,
    )
    conn.commit()
    deleted = cur.rowcount
    conn.close()
    return deleted


if __name__ == "__main__":
    # Smoke test: insert -> is_known -> pending -> downloaded -> needing_upload -> completed -> cleanup
    test_patient = {
        "acct_no": "TEST001",
        "appointment_date": "2026-07-16",
        "last_name": "Test",
        "first_name": "Patient",
        "visit_type": "9 MONTH WC",
        "form_name": "ASQ9Mos",
        "form_filename": "ASQ9Mos",
        "folder_name": "Test Patient_doc",
        "search_name": "Test,Patient",
    }

    print(f"DB path: {DB_PATH}")
    assert not is_known(test_patient["acct_no"], test_patient["appointment_date"])
    insert_form_sent(test_patient)
    assert is_known(test_patient["acct_no"], test_patient["appointment_date"])
    pending = get_pending_patients()
    assert any(p["acct_no"] == "TEST001" for p in pending)
    print(f"Pending: {len(pending)} (includes TEST001: OK)")

    mark_downloaded(test_patient["acct_no"], test_patient["appointment_date"])
    needing_upload = get_patients_needing_upload()
    assert any(p["acct_no"] == "TEST001" for p in needing_upload)
    print(f"Needing upload: {len(needing_upload)} (includes TEST001: OK)")

    mark_completed(test_patient["acct_no"], test_patient["appointment_date"])

    # Clean up the test row itself (retention_days=0 deletes anything completed now-or-before-now)
    deleted = cleanup_old_completed(retention_days=0)
    assert deleted >= 1
    assert not is_known(test_patient["acct_no"], test_patient["appointment_date"])
    print(f"Cleanup deleted {deleted} record(s), TEST001 no longer known: OK")
    print("\nstate_db.py smoke test PASSED")
