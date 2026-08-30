"""Run every standalone COM gate script in this directory and aggregate.

Each gate is a self-contained script (comments_readback_gate.py,
live_editing_gate.py, ...) that exits 0 on PASS or honest SKIP and nonzero
on FAIL, printing PASS/FAIL/SKIPPED lines as it goes. This runner executes
each gate in its own subprocess (a wedged PowerPoint can never take the
others down), classifies the outcome, prints a per-gate summary, and exits
nonzero if ANY gate failed.

Classification per gate:
- FAIL: nonzero exit code (or the subprocess timed out / crashed).
- SKIP: exit 0 with a SKIPPED marker in the output (the gate could not run
  here, e.g. the user's PowerPoint is open, and said so honestly).
- PASS: exit 0 without a SKIPPED marker.

Run:  .venv/Scripts/python.exe -X utf8 tests/com_gates/run_all.py
"""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
PER_GATE_TIMEOUT = 900  # seconds; each gate polls its own zombies


def discover() -> list[Path]:
    return sorted(
        p for p in HERE.glob("*_gate.py") if p.name != Path(__file__).name
    )


def run_gate(script: Path) -> tuple[str, float, str]:
    """Execute one gate; return (verdict, seconds, tail_of_output)."""
    t0 = time.monotonic()
    try:
        proc = subprocess.run(
            [sys.executable, "-X", "utf8", str(script)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=PER_GATE_TIMEOUT,
            cwd=str(HERE.parents[1]),  # repo root
        )
    except subprocess.TimeoutExpired:
        return "FAIL", time.monotonic() - t0, "timed out"
    seconds = time.monotonic() - t0
    out = (proc.stdout or "") + ("\n" + proc.stderr if proc.stderr else "")
    if proc.returncode != 0:
        return "FAIL", seconds, out
    if "SKIPPED" in out:
        return "SKIP", seconds, out
    return "PASS", seconds, out


def main() -> int:
    gates = discover()
    if not gates:
        print("no *_gate.py scripts found in", HERE)
        return 1
    results: list[tuple[str, str, float]] = []
    for script in gates:
        print(f"=== {script.name} ===", flush=True)
        verdict, seconds, out = run_gate(script)
        # Echo the gate's own PASS/FAIL/SKIPPED lines for the record.
        for ln in out.splitlines():
            if ln.startswith(("PASS", "FAIL", "SKIPPED", "VERDICT")):
                print("   ", ln)
        print(f"--- {script.name}: {verdict} ({seconds:.1f}s)\n", flush=True)
        results.append((script.name, verdict, seconds))

    counts = {"PASS": 0, "FAIL": 0, "SKIP": 0}
    for _name, verdict, _s in results:
        counts[verdict] += 1
    print("=" * 60)
    for name, verdict, seconds in results:
        print(f"{verdict:<5} {name} ({seconds:.1f}s)")
    print(
        f"TOTAL: {counts['PASS']} passed, {counts['FAIL']} failed, "
        f"{counts['SKIP']} skipped"
    )
    return 1 if counts["FAIL"] else 0


if __name__ == "__main__":
    sys.exit(main())
