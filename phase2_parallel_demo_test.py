"""
phase2_parallel_demo_test.py
--------------------------------
Phase 2 of the incremental concurrent-execution rollout: runs BOTH
clinics' full DEMO pipelines simultaneously, as separate OS processes.

  - Reference clinic: `python main_1.py` inside ECW_automation/ - its own
    existing demo pipeline. Now identifies its 2 River Ridge test patients
    via facility exclusion instead of Visit Reason (see
    REQUIRE_VISIT_REASON_FOR_DEMO in that file - Lone Star Midlothian's
    patient is excluded there).
  - Lone Star: `python main_demo.py` inside lone_star_automation/ - its
    own demo pipeline. Identifies its 1 Lone Star Midlothian test patient
    via Appointment Facility Name (see DEMO_REQUIRE_VISIT_REASON in
    config/settings.py).

Both pipelines run their FULL flow: eCW export -> PFN import -> determine
eligibility -> send ASQ form -> PCareLink reminder message -> check
completed forms -> download PDFs -> upload to eCW chart -> update state
DB. This is the first live end-to-end test of Lone Star's PFN-table
DOB-based eligibility scraping and PCareLink wiring.

Neither project's code is modified by this script - it only launches each
one's own existing entry point as a separate process, each with its own
browser instance/session (same isolation pattern validated in Phase 1).
Each project has its own state DB, so there's no shared-state conflict.
"""

import asyncio
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REFERENCE_DIR = os.path.join(BASE_DIR, "ECW_automation")
LONE_STAR_DIR = os.path.join(BASE_DIR, "lone_star_automation")


async def _run_and_stream(name, cmd, cwd):
    print(f"[{name}] starting: {' '.join(cmd)} (cwd={cwd})")
    start = time.monotonic()
    proc = await asyncio.create_subprocess_exec(
        *cmd,
        cwd=cwd,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
    )

    async def _stream_output():
        async for line in proc.stdout:
            print(f"[{name}] {line.decode(errors='replace').rstrip()}")

    await asyncio.gather(_stream_output(), proc.wait())
    elapsed = time.monotonic() - start
    return name, proc.returncode, elapsed


async def run_reference_demo():
    return await _run_and_stream("reference_clinic_demo", ["python", "main_1.py"], REFERENCE_DIR)


async def run_lone_star_demo():
    return await _run_and_stream("lone_star_demo", ["python", "main_demo.py"], LONE_STAR_DIR)


async def main():
    print("=" * 60)
    print("PHASE 2: PARALLEL DEMO PIPELINE TEST (both clinics, full flow)")
    print("Both processes run their full demo pipeline simultaneously:")
    print("  export -> PFN import -> eligibility -> send form ->")
    print("  PCareLink message -> check completed -> download -> upload")
    print("=" * 60)

    overall_start = time.monotonic()
    results = await asyncio.gather(run_reference_demo(), run_lone_star_demo())
    overall_elapsed = time.monotonic() - overall_start

    print("\n" + "=" * 60)
    print("PHASE 2 COMPLETE")
    print("=" * 60)
    for name, code, elapsed in results:
        status = "OK" if code == 0 else f"FAILED (exit code {code})"
        print(f"  {name}: {status} - {elapsed:.1f}s")
    print(f"Total wall-clock time: {overall_elapsed:.1f}s")
    print("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
