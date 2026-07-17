"""
utils/date_utils.py
---------------------
DOB -> ASQ age-bracket helpers.

Needed because Lone Star's Patient Forms Now table does NOT encode age in
the Visit Type text (unlike the reference project's "9 MONTH WC" style
values) - Visit Type here is generic (e.g. "New patient", presumably
something like "Well Check" for WC visits). Eligibility must instead be
computed from the DOB and Appointment columns shown in the same table row.
"""

from datetime import date, datetime


def parse_date_flexible(value):
    """
    Parses the date text formats seen in PFN table cells (e.g. "2025-07-14"
    for DOB, "2026-07-16 14:00" for Appointment). Returns a date, or None
    if unparseable. UNVERIFIED against the live table beyond the one
    example screenshot seen so far - add formats here if real data uses
    something different.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in (
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%B %d, %Y",
        "%b %d, %Y",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def age_in_months(dob, reference_date):
    """Whole months between dob and reference_date (e.g. the appointment date)."""
    months = (reference_date.year - dob.year) * 12 + (reference_date.month - dob.month)
    if reference_date.day < dob.day:
        months -= 1
    return months


def match_asq_bracket(age_months):
    """
    Maps age in months to one of the 8 supported ASQ brackets
    (9/12/15/18/24/30/36/48), using the exact same boundary thresholds as
    the first project's age_bucket_label() (automation_pd_forms/utils/
    date_utils.py), restricted to the brackets we have an ASQ form for.

    Explicit ranges (not an arbitrary +/- tolerance), matching the first
    project's cascading if-checks:
        8-10 months  -> 9
        11-12 months -> 12
        13-16 months -> 15
        17-20 months -> 18
        21-26 months -> 24
        27-32 months -> 30
        33-41 months -> 36
        42-53 months -> 48
        <8 or >=54 months -> None (skip - not a supported ASQ age; the
            first project's own next bucket past 53 months is "5-6 YEAR",
            which isn't one of our forms, so this is a deliberate upper
            cutoff, not an oversight)
    """
    if age_months < 8:
        return None
    if age_months < 11:
        return 9
    if age_months < 13:
        return 12
    if age_months < 17:
        return 15
    if age_months < 21:
        return 18
    if age_months < 27:
        return 24
    if age_months < 33:
        return 30
    if age_months < 42:
        return 36
    if age_months < 54:
        return 48
    return None
