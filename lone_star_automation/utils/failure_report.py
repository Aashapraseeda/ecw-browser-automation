"""
utils/failure_report.py
------------------------
Pure reporting/visibility helper for the ReachMyDr messaging step - NOT
part of the messaging flow itself and touches no database state. Builds
a plain-text "which patients still need a manual ReachMyDr message"
report, in the exact format requested, and saves it to a timestamped file
under logs/. Shared by pcarelink/messenger.py (production + demo) so the
format only needs to be defined once.

Each failure record is a plain dict, independent of the original patient
dict object (never the same object reference) - built via record_failure()
below - so nothing here can ever mutate a patient dict that's still going
to be used elsewhere (e.g. for insert_form_sent()).
"""

import os
from datetime import datetime


def record_failure(failed_list, patient, reason, practice=None):
    """
    Appends a NEW, independent dict describing this failure - never the
    original patient dict - so later mutation/reuse of patient dicts
    elsewhere in the pipeline can't affect an already-recorded failure
    (and vice versa).
    """
    failed_list.append({
        "acct_no": patient.get("acct_no", ""),
        "last_name": patient.get("last_name", ""),
        "first_name": patient.get("first_name", ""),
        "practice": practice or "",
        "dob": patient.get("dob"),
        "appointment_date": patient.get("appointment_date"),
        "reason": reason,
    })


def _format_date_for_report(iso_date_str):
    if not iso_date_str:
        return "N/A"
    try:
        y, m, d = str(iso_date_str).split("-")[:3]
        d = d.split("T")[0].split(" ")[0]
        return f"{m}/{d}/{y}"
    except Exception:
        return str(iso_date_str)


def build_failure_report_text(failed):
    lines = ["=" * 50, "REACHMYDR MESSAGE FAILURES", "=" * 50, ""]
    if not failed:
        lines.append("No failed ReachMyDr messages.")
        return "\n".join(lines) + "\n"

    lines.append(f"Total Failed: {len(failed)}")
    lines.append("")
    for i, f in enumerate(failed, start=1):
        lines.append(f"{i}.")
        lines.append(f"Account #: {f.get('acct_no', '')}")
        lines.append(f"Patient : {f.get('last_name', '')}, {f.get('first_name', '')}")
        lines.append(f"Practice: {f.get('practice') or 'N/A'}")
        lines.append(f"DOB     : {_format_date_for_report(f.get('dob'))}")
        lines.append(f"Visit   : {_format_date_for_report(f.get('appointment_date'))}")
        lines.append(f"Reason  : {f.get('reason', 'Unknown')}")
        lines.append("")
    return "\n".join(lines) + "\n"


def save_failure_report(failed, log_dir):
    """Writes the report to logs/failed_reachmydr_messages_<timestamp>.txt
    regardless of whether there were any failures (an empty run still
    produces a "No failed ReachMyDr messages." file, so the absence of a
    file is never mistaken for "nobody checked"). Returns the file path."""
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(log_dir, f"failed_reachmydr_messages_{timestamp}.txt")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(build_failure_report_text(failed))
    return path


def print_and_save_failure_report(failed, log_dir):
    text = build_failure_report_text(failed)
    print("\n" + text)
    path = save_failure_report(failed, log_dir)
    print(f"Failure report saved to: {path}")
    return path
