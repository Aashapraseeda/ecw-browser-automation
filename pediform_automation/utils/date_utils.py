"""
utils/date_utils.py
-------------------
Age and date helper functions.
"""

from datetime import date, datetime
from typing import Optional


def calculate_age_in_months(dob: date, reference: Optional[date] = None) -> int:
    """Return age in whole months (floor)."""
    ref = reference or date.today()
    months = (ref.year - dob.year) * 12 + (ref.month - dob.month)
    if ref.day < dob.day:
        months -= 1
    return max(months, 0)


def calculate_age(dob: date, reference: Optional[date] = None) -> float:
    """Return age in fractional years."""
    return calculate_age_in_months(dob, reference) / 12.0


def appointment_within_days(appt_date: date, window_days: int) -> bool:
    """True if appt_date is today or within the next window_days days."""
    today = date.today()
    delta = (appt_date - today).days
    return 0 <= delta <= window_days


def age_bucket_label(dob: date, reference: Optional[date] = None) -> Optional[str]:
    """
    Map a date-of-birth to one of the canonical well-child age-bucket labels.
    Returns None if the age doesn't fall into any defined bucket.

    Bucket boundaries (inclusive unless noted):
      NEWBORN   : 0–6 days
      1 WEEK    : 7–27 days
      2 MONTH   : 1–3 months
      4 MONTH   : 3–5 months
      6 MONTH   : 5–8 months
      9 MONTH   : 8–11 months
      12 MONTH  : 11–13 months
      15 MONTH  : 13–17 months
      18 MONTH  : 17–21 months
      24 MONTH  : 21–27 months
      30 MONTH  : 27–33 months
      3 YEAR    : 33–42 months
      4 YEAR    : 42–54 months
      5-6 YEAR  : 54–78 months (4y6m–6y6m)
      7-11 YEAR : 78–138 months
      12-18 YEAR: 138–228 months
    """
    ref = reference or date.today()
    months = calculate_age_in_months(dob, ref)

    # Use days for newborn precision
    days = (ref - dob).days

    if days < 7:
        return "NEWBORN"
    if days < 28:
        return "1 WEEK"
    if months < 3:
        return "2 MONTH"
    if months < 5:
        return "4 MONTH"
    if months < 8:
        return "6 MONTH"
    if months < 11:
        return "9 MONTH"
    if months < 13:
        return "12 MONTH"
    if months < 17:
        return "15 MONTH"
    if months < 21:
        return "18 MONTH"
    if months < 27:
        return "24 MONTH"
    if months < 33:
        return "30 MONTH"
    if months < 42:
        return "3 YEAR"
    if months < 54:
        return "4 YEAR"
    if months < 78:
        return "5-6 YEAR"
    if months < 138:
        return "7-11 YEAR"
    if months < 228:
        return "12-18 YEAR"
    return None


def to_date(value) -> Optional[date]:
    """Coerce a datetime / date / string to date, or return None."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    # Try common string formats
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            pass
    return None
