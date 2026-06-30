# modules/pediform/selectors.py
# All confirmed PediForm selectors in one place.
# Update here when the UI changes; nothing else needs editing.

# ── Login page ─────────────────────────────────────────────────────────────────
LOGIN_ORG          = "#admin-practice"
LOGIN_EMAIL        = "#admin-email"
LOGIN_PASSWORD     = "#admin-password"
LOGIN_SUBMIT       = "button.patient-portal-submit"

# ── Navigation ─────────────────────────────────────────────────────────────────
NAV_TODAYS_PATIENTS = "Today's Patients"           # used with get_by_role("link")
NAV_SUBMISSIONS     = 'a[href="/staff/submissions"]'

# ── Schedule import (Today's Patients page) ────────────────────────────────────
SCHEDULE_FILE_INPUT  = "input[type='file']"        # standard HTML file input
SCHEDULE_IMPORT_BTN  = "Import schedule"           # used with get_by_role("button")

# ── Patient table ──────────────────────────────────────────────────────────────
PATIENT_TABLE_ROWS = "table tbody tr"
PATIENT_VIEW_BTN   = "a.btn-ghost.btn-sm"

# ── Patient view page ──────────────────────────────────────────────────────────
# "Auto-assign (well visit)" has class "secondary btn-inline" — must NOT match it.
# Use button text to be unambiguous.
SEND_FORM_OPEN_BTN   = "+ Send a form"   # used with get_by_role("button", name=...)

# ── Send form panel (appears after clicking + Send a form) ─────────────────────
# These selectors are NOT yet confirmed — will update after panel inspect.
SEND_FORM_SUBMIT_BTN = "button.btn"      # placeholder — verify after panel opens

# ── Submissions / download page ────────────────────────────────────────────────
SUBMISSION_TABLE_ROWS = "table tbody tr"
STATUS_COMPLETED      = 'span:has-text("Completed")'
DOWNLOAD_PDF_BTN      = 'button:has-text("Download PDF")'
