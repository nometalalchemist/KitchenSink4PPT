"""Phase 5 tests: the COM bridge and the export layer.

COM testing rules (singleton contract, non-negotiable):
- Every COM scenario runs in an ISOLATED SUBPROCESS (one comprehensive
  scenario per subprocess, not many attach cycles) so a wedged PowerPoint
  cannot hang or pollute the pytest process.
- tasklist gate: if POWERPNT.EXE is already running (the user's instance;
  PowerPoint is a singleton COM server), COM tests SKIP honestly rather
  than attach to the user's PowerPoint. The user may open PowerPoint at any
  moment; the gate re-checks at each test start and the bridge's own
  pre-check plus refusal rule covers a mid-run open.
- Launched-vs-attached: only the LAUNCHED case is testable here (the gate
  guarantees no pre-existing instance, so the bridge launches and must quit
  cleanly, which the zombie counts assert). The ATTACHED case (pre-existing
  user instance: no Quit, close only our presentations, DocumentLocked
  refusal on the user's open file) cannot be exercised without opening a
  presentation in a user-owned instance, which the standing safety rules
  forbid in agent rounds; it is covered by code review and reserved for a
  supervised live round.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.ops import export as export_ops

REPO = Path(__file__).resolve().parents[2]
CORPUS = REPO / "tests" / "corpus"
ARTIFACTS = REPO / "tests" / "artifacts"
DELTA = ARTIFACTS / "delta_triangle.pptx"

IS_WIN = sys.platform == "win32"

try:
    import win32com.client  # noqa: F401

    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

if IS_WIN and HAS_PYWIN32:
    from kitchensink4ppt.com import bridge

    HAS_POWERPOINT = bridge.powerpoint_installed()
else:
    bridge = None
    HAS_POWERPOINT = False


def _powerpnt_running() -> bool:
    if bridge is None:
        return False
    return bridge.powerpnt_count() > 0


def _com_gate():
    """Runtime gate for COM tests; called at each test start because the
    user may open PowerPoint between tests."""
    if not IS_WIN:
        pytest.skip("COM bridge is Windows-only")
    if not HAS_PYWIN32:
        pytest.skip("pywin32 not installed")
    if not HAS_POWERPOINT:
        pytest.skip("PowerPoint is not installed on this machine")
    if _powerpnt_running():
        pytest.skip(
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running (the "
            "user's instance; PowerPoint is a singleton COM server, so the "
            "test would attach to it). COM coverage did NOT run."
        )


# One scenario per subprocess: the script below is the entire COM round for
# a mode, and the pytest process only parses its JSON verdict.
_SCENARIO = r"""
import json, sys, time
from pathlib import Path

from kitchensink4ppt.com import bridge
from kitchensink4ppt.ops import export as export_ops

mode = sys.argv[1]
out = {}
pre = bridge.powerpnt_count()
out["pre_powerpnt"] = pre
if pre > 0:
    out["skipped"] = "user PowerPoint opened mid-round; refusing to attach"
    print("RESULT " + json.dumps(out))
    sys.exit(0)

if mode == "exports":
    src = Path(sys.argv[2])
    art = Path(sys.argv[3])
    art.mkdir(parents=True, exist_ok=True)
    before_bytes = src.read_bytes()
    before_mtime = src.stat().st_mtime_ns

    out["pdf"] = bridge.com_export_pdf(str(src), str(art / "delta_triangle.pdf"))
    out["auto_pdf"] = export_ops.export_pdf(
        str(src), str(art / "delta_triangle_auto.pdf"), engine="auto"
    )
    out["images"] = bridge.com_export_slide_images(
        str(src), str(art / "delta_png"), width=1280
    )

    out["source_bytes_unchanged"] = src.read_bytes() == before_bytes
    out["source_mtime_unchanged"] = src.stat().st_mtime_ns == before_mtime

elif mode == "validate":
    good = Path(sys.argv[2])
    tmp = Path(sys.argv[3])
    donor = Path(sys.argv[4])
    out["good"] = bridge.com_validate_opens_clean(str(good))

    bad = tmp / "corrupt_truncated.pptx"
    data = donor.read_bytes()
    bad.write_bytes(data[: len(data) // 2])
    t0 = time.monotonic()
    out["bad"] = bridge.com_validate_opens_clean(str(bad))
    out["bad_seconds"] = round(time.monotonic() - t0, 1)

out["status"] = bridge.powerpoint_status()
out["post_powerpnt"] = bridge.powerpnt_count()
out["zombie"] = bridge.zombie_check()
print("RESULT " + json.dumps(out))
"""


def _run_scenario(tmp_path: Path, mode: str, *args: str) -> dict:
    script = tmp_path / f"scenario_{mode}.py"
    script.write_text(_SCENARIO, encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(script), mode, *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=480,
        cwd=str(REPO),
    )
    result_line = next(
        (
            ln
            for ln in reversed((proc.stdout or "").splitlines())
            if ln.startswith("RESULT ")
        ),
        None,
    )
    assert proc.returncode == 0 and result_line, (
        f"scenario subprocess failed (exit {proc.returncode})\n"
        f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
    )
    out = json.loads(result_line[len("RESULT "):])
    if "skipped" in out:
        pytest.skip(f"COM round self-skipped: {out['skipped']}")
    return out


# --------------------------------------------------------------- COM rounds


@pytest.mark.timeout(600)
def test_com_export_round(tmp_path):
    """PDF export (bridge + auto routing), PNG slide export, source file
    untouched, zombie-free exit of the launched instance."""
    _com_gate()
    assert DELTA.exists(), "Phase 4 artifact delta_triangle.pptx is missing"
    art = ARTIFACTS / "phase5"
    out = _run_scenario(tmp_path, "exports", str(DELTA), str(art))

    # PDF via the bridge
    pdf = Path(out["pdf"]["pdf"])
    assert pdf.exists() and pdf.stat().st_size > 0
    assert out["pdf"]["bytes"] == pdf.stat().st_size
    assert out["pdf"]["engine"] == "powerpoint-com"
    assert pdf.read_bytes()[:5] == b"%PDF-"

    # PDF via ops.export auto routing (PowerPoint installed => COM engine)
    assert out["auto_pdf"]["engine"] == "powerpoint-com"
    auto_pdf = Path(out["auto_pdf"]["pdf"])
    assert auto_pdf.exists() and auto_pdf.stat().st_size > 0

    # PNG slide export: files exist, are non-empty PNGs, aspect-derived height
    images = out["images"]["images"]
    assert len(images) >= 1
    for entry in images:
        png = Path(entry["file"])
        assert png.exists() and png.stat().st_size > 0
        assert png.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"
    assert out["images"]["width"] == 1280
    assert out["images"]["height"] > 0

    # source untouched: exports must never modify the deck
    assert out["source_bytes_unchanged"] is True
    assert out["source_mtime_unchanged"] is True

    # launched case: the instance we started is gone again
    assert out["pre_powerpnt"] == 0
    assert out["post_powerpnt"] == 0
    assert out["zombie"]["powerpnt_processes"] == 0

    # status diagnostic ran without attaching
    assert out["status"]["installed"] is True
    assert out["status"]["powerpoint_running"] is False


@pytest.mark.timeout(600)
def test_com_validate_round(tmp_path):
    """validate_opens_clean: clean verdict with full-load counts on a real
    corpus deck; typed non-clean verdict (no hang) on a truncated file;
    zombie-free exit."""
    _com_gate()
    good = CORPUS / "proposal_defense.pptx"
    assert good.exists(), "corpus deck missing (conftest generates stand-ins)"
    out = _run_scenario(tmp_path, "validate", str(good), str(tmp_path), str(DELTA))

    assert out["good"]["opens_clean"] is True
    assert out["good"]["slides"] > 0
    assert out["good"]["shapes"] > 0

    assert out["bad"]["opens_clean"] is False
    assert out["bad"]["error"]
    # DisplayAlerts=ppAlertsNone means the corrupt open RAISES instead of
    # hanging on a modal repair dialog; well under the subprocess timeout.
    assert out["bad_seconds"] < 120

    assert out["post_powerpnt"] == 0
    assert out["zombie"]["powerpnt_processes"] == 0


# ------------------------------------------------- diagnostics (no dispatch)


def test_powerpoint_status_and_zombie_check_inline():
    """powerpoint_status touches only the registry, the process table, and
    ROT display names; it never dispatches into an app, so it is safe to run
    inline regardless of what the user has open."""
    if not IS_WIN or not HAS_PYWIN32:
        pytest.skip("Windows + pywin32 only")
    status = bridge.powerpoint_status()
    assert isinstance(status["installed"], bool)
    assert isinstance(status["powerpoint_running"], bool)
    assert isinstance(status["open_presentations"], list)
    if not status["powerpoint_running"]:
        assert status["open_presentations"] == []
    z = bridge.zombie_check()
    assert isinstance(z["powerpnt_processes"], int)
    assert z["powerpnt_processes"] >= 0
    assert status["powerpoint_running"] == (z["powerpnt_processes"] > 0)


# ------------------------------------------------------- export layer (file)


def test_get_export_engines_shape():
    report = export_ops.get_export_engines()
    engines = report["engines"]
    for name in ("powerpoint_com", "libreoffice"):
        assert isinstance(engines[name]["available"], bool)
        assert isinstance(engines[name]["pdf"], bool)
        assert isinstance(engines[name]["slide_images"], bool)
    assert report["auto_pdf"] in ("com", "libreoffice", None)
    assert report["auto_images"] in ("com", "libreoffice", None)


def test_export_pdf_rejects_unknown_engine(make_deck):
    deck = make_deck("engine.pptx")
    with pytest.raises(PptMcpError, match="unknown engine"):
        export_ops.export_pdf(str(deck), engine="ghostscript")


def test_export_pdf_missing_source():
    from kitchensink4ppt.core.errors import DocumentNotFound

    with pytest.raises(DocumentNotFound):
        export_ops.export_pdf(str(REPO / "no_such_deck.pptx"))


@pytest.mark.timeout(600)
def test_libreoffice_pdf_fallback(make_deck, tmp_path):
    """The non-COM fallback: soffice headless PDF conversion. Runs whenever
    LibreOffice is installed (independent of PowerPoint/COM state)."""
    soffice = export_ops._find_soffice()
    if soffice is None:
        pytest.skip(
            "LibreOffice not found (checked KS4P_SOFFICE, "
            f"{export_ops.SOFFICE_WELL_KNOWN}, PATH); fallback untested here"
        )
    deck = make_deck("lo_deck.pptx")
    out = tmp_path / "lo_deck_export.pdf"
    before = deck.read_bytes()
    result = export_ops.export_pdf(str(deck), str(out), engine="libreoffice")
    assert result["engine"] == "libreoffice"
    assert out.exists() and out.stat().st_size > 0
    assert out.read_bytes()[:5] == b"%PDF-"
    assert deck.read_bytes() == before  # source untouched


def test_libreoffice_images_refuse_without_pdftoppm(make_deck):
    """soffice alone renders only the first slide, so the LO image path
    demands pdftoppm and must say so honestly."""
    if export_ops._find_soffice() is None:
        pytest.skip("LibreOffice not found")
    if export_ops._find_pdftoppm() is not None:
        pytest.skip("pdftoppm IS available here; the refusal path is moot")
    deck = make_deck("lo_img.pptx")
    with pytest.raises(PptMcpError, match="pdftoppm"):
        export_ops.export_slide_images(str(deck), engine="libreoffice")


def test_image_auto_routing_honest_when_only_soffice(monkeypatch, make_deck):
    """auto image routing: with COM unavailable and pdftoppm missing, the
    error must name what is missing rather than half-work."""
    monkeypatch.setattr(export_ops, "_com_available", lambda: False)
    monkeypatch.setattr(export_ops, "_find_pdftoppm", lambda: None)
    deck = make_deck("route.pptx")
    if export_ops._find_soffice() is not None:
        with pytest.raises(PptMcpError, match="pdftoppm"):
            export_ops.export_slide_images(str(deck), engine="auto")
    else:
        with pytest.raises(PptMcpError, match="PowerPoint"):
            export_ops.export_slide_images(str(deck), engine="auto")


def test_pdf_auto_routing_honest_when_no_engines(monkeypatch, make_deck):
    monkeypatch.setattr(export_ops, "_com_available", lambda: False)
    monkeypatch.setattr(export_ops, "_find_soffice", lambda: None)
    deck = make_deck("noeng.pptx")
    with pytest.raises(PptMcpError) as exc_info:
        export_ops.export_pdf(str(deck), engine="auto")
    msg = str(exc_info.value)
    assert "PowerPoint" in msg and "LibreOffice" in msg  # names both options


def test_sandbox_gates_export_paths(monkeypatch, tmp_path, make_deck):
    """check_path applies to export sources like every other path entry."""
    deck = make_deck("sandboxed.pptx")
    outside = tmp_path / "elsewhere"
    outside.mkdir()
    monkeypatch.setenv("KS4P_ALLOWED_ROOTS", str(outside))
    from kitchensink4ppt.core.sandbox import SandboxViolation

    with pytest.raises(SandboxViolation):
        export_ops.export_pdf(str(deck))
