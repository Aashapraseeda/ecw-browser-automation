"""
phase1_reference_clinic_export_only.py
------------------------------------------
Phase 1 parallel-execution test: runs ONLY the reference clinic's eCW
schedule export (login -> eBO Reports -> date range -> generate ->
download Excel), then stops. Does NOT invoke Patient Forms Now, PCareLink,
or eCW chart upload - those are later phases in the incremental rollout.

Does NOT modify ECW_automation (that project stays untouched, per
instruction) - imports its existing ecw_export_schedule() function from
outside the project and calls it directly, bypassing the rest of
main()'s pipeline (Pediforms sending, PCareLink messaging) entirely.
Importing main.py only executes its module-level code (constants,
function defs) - the `if __name__ == "__main__":` guard means the full
pipeline never runs just from importing it.
"""

import sys
import os
import asyncio

ECW_AUTOMATION_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ECW_automation")
sys.path.insert(0, ECW_AUTOMATION_DIR)

import main as ecw_automation_main  # noqa: E402 - must import after sys.path insert


if __name__ == "__main__":
    print("=" * 50)
    print("PHASE 1 - REFERENCE CLINIC: ECW EXPORT ONLY")
    print("=" * 50)
    asyncio.run(ecw_automation_main.ecw_export_schedule())
