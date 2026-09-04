"""Shared COM gate: tell a LINGERING SIBLING apart from the USER's PowerPoint.

V1.1_QUEUE item 2. Every COM-touching test file carries a gate that skips
when POWERPNT.EXE is running, because PowerPoint is a singleton COM server
and attaching to the user's instance is forbidden. The check was instant
and unconditional, which conflated two very different situations:

- The user really does have PowerPoint open. Skipping is correct and the
  skip message should say so.
- A COM round that JUST finished is still shutting down. POWERPNT lingers
  a few seconds after Quit (the bridge's own cleanup allows up to 15s for
  this), so the next test's gate saw a process, assumed it was the user's,
  and skipped honest coverage. Running the COM-heavy files back to back is
  exactly when this bites, and it produced the queue's "22 honest skips".

The fix is to WAIT before judging. If the process count drains to zero
within the window, it was our own sibling and the test proceeds with full
coverage. If it persists, it is genuinely the user's instance (or a real
orphan) and the skip stands, with a message that now distinguishes the two.

This does not serialize anything by itself: pytest already runs tests
sequentially here. It removes the false-positive skips that sequential
execution was CAUSING, which is what item 2 was actually about.
"""

from __future__ import annotations

import time

#: How long a lingering sibling gets to exit before we conclude the running
#: PowerPoint belongs to the user. The bridge's own post-Quit poll allows
#: 15s, so anything shorter here would re-introduce the race it fixed.
DRAIN_SECONDS = 20.0
_POLL = 0.5


def wait_for_powerpnt_drain(
    count_fn, timeout: float = DRAIN_SECONDS
) -> tuple[bool, float]:
    """Poll until no POWERPNT.EXE remains, or the timeout expires.

    Returns (drained, waited_seconds). count_fn is the caller's process
    counter (normally bridge.powerpnt_count) so this module needs no
    import of the package under test."""
    t0 = time.monotonic()
    try:
        if count_fn() == 0:
            return True, 0.0
    except Exception:
        return False, 0.0
    deadline = t0 + timeout
    while time.monotonic() < deadline:
        time.sleep(_POLL)
        try:
            if count_fn() == 0:
                return True, round(time.monotonic() - t0, 1)
        except Exception:
            break
    return False, round(time.monotonic() - t0, 1)


def powerpoint_blocks_com_tests(count_fn) -> str | None:
    """None when COM tests may proceed; otherwise the skip reason.

    Waits out a lingering sibling instance before declaring the running
    PowerPoint to be the user's."""
    drained, waited = wait_for_powerpnt_drain(count_fn)
    if drained:
        return None
    return (
        "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is still running after "
        f"waiting {waited:.0f}s for a lingering test instance to exit, so "
        "it is the user's instance (PowerPoint is a singleton COM server "
        "and this suite never attaches to it). COM coverage did NOT run."
    )
