"""
modules/reachmydr/messenger.py
-------------------------------
Placeholder ReachMyDr integration.
Logs a warning for every call — no actual messages are sent until
credentials are configured and this module is implemented.
"""

from typing import List, Dict
from utils.logger import get_logger

logger = get_logger(__name__)


class ReachMyDrMessenger:
    """
    Placeholder class. Replace method bodies with real Playwright automation
    once ReachMyDr credentials are available.
    """

    def __init__(self):
        logger.warning(
            "[ReachMyDr] No credentials configured — all reminders will be SKIPPED."
        )

    async def login(self, page) -> None:
        logger.warning("[ReachMyDr] login() not implemented — skipping.")

    async def search_patient(self, page, patient_name: str) -> bool:
        logger.warning(f"[ReachMyDr] search_patient({patient_name}) — skipping.")
        return False

    async def send_message(self, page, message: str) -> bool:
        logger.warning("[ReachMyDr] send_message() — skipping.")
        return False

    async def send_reminder_to_patient(self, page, patient: Dict) -> bool:
        logger.warning(
            f"[ReachMyDr] Would send reminder to {patient.get('full_name')} — skipping."
        )
        return False

    async def send_reminders(self, patients: List[Dict]) -> List[Dict]:
        """
        Attempt to send reminders to all patients.
        Currently a no-op placeholder.
        """
        logger.warning(
            f"[ReachMyDr] send_reminders() called for {len(patients)} patient(s) — "
            "all skipped (placeholder)."
        )
        return []
