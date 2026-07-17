"""
run_production_parallel.py
-------------------------------
Runs BOTH clinics' PRODUCTION pipelines simultaneously, each as a
separate OS process (-> separate browser instance -> separate cookie/
session storage). Same eCW login is used by both, which is safe run
this way - confirmed via manual two-Chrome-profile testing and a live
parallel export test earlier in this project's development (two tabs
sharing one profile's cookies DOES conflict; two separate processes/
browsers do not).

  - Reference clinic (Nurture Kids): ECW_automation/main.py
  - Lone Star Pediatrics: lone_star_automation/main.py

This sends REAL forms and REAL PCareLink messages to REAL patients on
both sides - only run this when you actually mean to.

Usage:
    python run_production_parallel.py

Each pipeline's own console output is prefixed with its clinic name so
you can tell them apart in one combined stream.
"""

import asyncio
import os
import sys
import time

# Windows console defaults to cp1252, which can't represent every
# character a pipeline prints (e.g. em-dashes) - tolerate anything
# rather than crashing the orchestrator over a display-only mismatch.
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


async def run_reference_production():
    return await _run_and_stream("reference_clinic_production", ["python", "main.py"], REFERENCE_DIR)


async def run_lone_star_production():
    return await _run_and_stream("lone_star_production", ["python", "main.py"], LONE_STAR_DIR)


async def main():
    print("=" * 60)
    print("RUNNING BOTH PRODUCTION PIPELINES SIMULTANEOUSLY")
    print("  reference_clinic_production -> ECW_automation/main.py")
    print("  lone_star_production        -> lone_star_automation/main.py")
    print("This sends REAL forms and REAL PCareLink messages to REAL patients.")
    print("=" * 60)

    overall_start = time.monotonic()
    results = await asyncio.gather(run_reference_production(), run_lone_star_production())
    overall_elapsed = time.monotonic() - overall_start

    print("\n" + "=" * 60)
    print("BOTH PRODUCTION PIPELINES FINISHED")
    print("=" * 60)
    for name, code, elapsed in results:
        status = "OK" if code == 0 else f"FAILED (exit code {code})"
        print(f"  {name}: {status} - {elapsed:.1f}s")
    print(f"Total wall-clock time: {overall_elapsed:.1f}s")
    print("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
