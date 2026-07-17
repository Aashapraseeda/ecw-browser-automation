"""
phase1_parallel_export_test.py
----------------------------------
Phase 1 of the incremental concurrent-execution rollout.

Launches two COMPLETELY SEPARATE OS processes simultaneously:
  - Reference clinic: phase1_reference_clinic_export_only.py, which
    imports ECW_automation's existing, unmodified ecw_export_schedule()
    and calls it directly (no Pediforms/PCareLink involved).
  - Lone Star: `python -m ecw.schedule_export` inside lone_star_automation/
    (already-verified export + Facility filter step, unchanged).

Both use the SAME eCW login credentials but run as separate OS processes,
each launching its own independent Playwright browser instance - separate
cookie/session storage per process, matching the manual two-Chrome-profile
test that confirmed this does not cause a session conflict (unlike two
tabs sharing one profile's cookie jar, which does).

Each process is fully independent: own browser, own downloads folder, own
.env/config, own logs. Nothing is shared between them. This script does
NOT touch Patient Forms Now, PCareLink, or eCW chart upload - it stops
right after both Excel downloads complete, per the incremental rollout
plan (later phases extend to PFN/PCareLink/upload once this is confirmed
reliable).
"""

import asyncio
import os
import sys
import time

# Windows console defaults to cp1252, which can't represent every character
# the reference clinic's script prints (e.g. em-dashes). Reconfigure this
# script's own stdout to tolerate anything rather than crashing the whole
# orchestrator over a display-only encoding mismatch in relayed output.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REFERENCE_DIR = os.path.join(BASE_DIR, "ECW_automation")
LONE_STAR_DIR = os.path.join(BASE_DIR, "lone_star_automation")
WRAPPER_SCRIPT = os.path.join(BASE_DIR, "phase1_reference_clinic_export_only.py")


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


async def run_reference_export():
    return await _run_and_stream(
        "reference_clinic", ["python", WRAPPER_SCRIPT], REFERENCE_DIR
    )


async def run_lone_star_export():
    return await _run_and_stream(
        "lone_star", ["python", "-m", "ecw.schedule_export"], LONE_STAR_DIR
    )


async def main():
    print("=" * 60)
    print("PHASE 1: PARALLEL ECW EXPORT TEST (both clinics, export only)")
    print("Both processes log into eCW SIMULTANEOUSLY with the same")
    print("credentials, each in its own separate browser instance.")
    print("=" * 60)

    overall_start = time.monotonic()
    results = await asyncio.gather(run_reference_export(), run_lone_star_export())
    overall_elapsed = time.monotonic() - overall_start

    print("\n" + "=" * 60)
    print("PHASE 1 COMPLETE")
    print("=" * 60)
    for name, code, elapsed in results:
        status = "OK" if code == 0 else f"FAILED (exit code {code})"
        print(f"  {name}: {status} - {elapsed:.1f}s")
    print(f"Total wall-clock time: {overall_elapsed:.1f}s")

    ref_excel = os.path.join(REFERENCE_DIR, "ecw_schedule.xlsx")
    ls_excel = os.path.join(LONE_STAR_DIR, "ecw_schedule.xlsx")
    print(f"\nReference clinic Excel exists: {os.path.exists(ref_excel)} ({ref_excel})")
    print(f"Lone Star Excel exists: {os.path.exists(ls_excel)} ({ls_excel})")
    print("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
