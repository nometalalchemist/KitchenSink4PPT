"""compress_deck coverage (ops/optimize.py): offender report, dry-run
zero-mutation guarantee (verified by md5 over every part), image
downsampling via optional Pillow, the honest no-Pillow refusal, the
deregister-then-sweep purge, and spine performance on the 131-slide
military brief. COM validation of a compressed deck follows the repo's
subprocess + tasklist-gate rules.
"""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest

import make_corpus
from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import _traverse as tv
from kitchensink4ppt.ops import optimize
from kitchensink4ppt.ops.media import insert_image
from kitchensink4ppt.ops.optimize import compress_deck

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


def _copy(tmp_path: Path, name: str) -> Path:
    dest = tmp_path / name
    shutil.copy2(CORPUS / name, dest)
    return dest


def _part_md5s(pkg: PptxPackage) -> dict[str, str]:
    return {
        name: hashlib.md5(pkg.raw_part(name)).hexdigest()
        for name in pkg.part_names()
    }


def _has_pillow() -> bool:
    return optimize._pillow()[0] is not None


# ============================================================ parameter guard


class TestParams:
    def test_bounds(self, make_deck):
        pkg = PptxPackage(make_deck("params.pptx"))
        with pytest.raises(PptMcpError, match="max_dpi"):
            compress_deck(pkg, max_dpi=10)
        with pytest.raises(PptMcpError, match="max_dpi"):
            compress_deck(pkg, max_dpi=1000)
        with pytest.raises(PptMcpError, match="jpeg_quality"):
            compress_deck(pkg, jpeg_quality=0)
        with pytest.raises(PptMcpError, match="jpeg_quality"):
            compress_deck(pkg, jpeg_quality=101)


# ================================================================== dry run


class TestDryRun:
    def test_dry_run_mutates_nothing_md5(self, tmp_path):
        path = _copy(tmp_path, "military_brief.pptx")
        file_md5_before = hashlib.md5(path.read_bytes()).hexdigest()
        pkg = PptxPackage(path)
        before = _part_md5s(pkg)

        report = compress_deck(pkg, dry_run=True)

        assert report["dry_run"] is True
        assert not pkg._dirty, "dry run must not dirty any part"
        assert _part_md5s(pkg) == before, "dry run must not touch part bytes"
        assert (
            hashlib.md5(path.read_bytes()).hexdigest() == file_md5_before
        ), "dry run must not touch the file"
        # ...and the purge SIMULATION must not have saved the scratch copy
        assert (
            hashlib.md5(path.read_bytes()).hexdigest() == file_md5_before
        )

    def test_offender_report_shape(self, tmp_path):
        path = _copy(tmp_path, "military_brief.pptx")
        pkg = PptxPackage(path)
        report = compress_deck(pkg, dry_run=True)
        media = report["media"]
        assert report["media_count"] == len(media) > 0
        sizes = [e["bytes"] for e in media]
        assert sizes == sorted(sizes, reverse=True), "offenders first"
        for entry in media:
            assert entry["part"].startswith("ppt/media/")
            assert entry["action"] in (
                "resize", "recompress", "keep", "unreferenced",
                "skipped-no-pillow",
            )
            assert "referenced_by" in entry
            assert "_new_data" not in entry, "internal payloads must be stripped"
        assert report["media_bytes_total"] == sum(sizes)

    def test_spine_performance_on_131_slides(self, tmp_path):
        """The audit's performance bar: the full spine and the compress
        report must stay sane on the heaviest real deck."""
        path = _copy(tmp_path, "military_brief.pptx")
        pkg = PptxPackage(path)
        t0 = time.perf_counter()
        run_count = sum(1 for _ in tv.iter_runs(pkg))
        color_count = sum(1 for _ in tv.iter_colors(pkg))
        spine_s = time.perf_counter() - t0
        assert run_count > 0 and color_count >= 0
        assert spine_s < 30, f"spine took {spine_s:.1f}s on 131 slides"

        t0 = time.perf_counter()
        compress_deck(pkg, dry_run=True)
        report_s = time.perf_counter() - t0
        assert report_s < 90, f"compress dry-run took {report_s:.1f}s"


# ==================================================================== purge


class TestPurge:
    def _orphan_media_deck(self, make_deck, tmp_path) -> tuple[Path, str]:
        """A deck carrying one referenced image and one deliberately
        orphaned media part (bytes present, no relationship anywhere)."""
        img = make_corpus._png(tmp_path / "used.png", 40, 40, (0, 128, 0))
        path = make_deck("orphan.pptx")
        pkg = PptxPackage(path)
        insert_image(pkg, 0, str(img), 1, 1, w=1)
        orphan_bytes = make_corpus._png(
            tmp_path / "orphan.png", 32, 32, (128, 0, 128)
        ).read_bytes()
        orphan_part = "ppt/media/image999.png"
        pkg.set_raw_part(orphan_part, orphan_bytes)
        pkg.save()
        return path, orphan_part

    def test_orphan_media_reported_then_purged(self, make_deck, tmp_path):
        path, orphan_part = self._orphan_media_deck(make_deck, tmp_path)

        pkg = PptxPackage(path)
        assert pkg.has_part(orphan_part)
        dry = compress_deck(pkg, dry_run=True)
        entry = next(e for e in dry["media"] if e["part"] == orphan_part)
        assert entry["action"] == "unreferenced"
        assert orphan_part in dry["purge"].get("media", [])
        assert pkg.has_part(orphan_part), "dry run must not purge"

        report = compress_deck(pkg, dry_run=False)
        assert orphan_part in report["purge"]["media"]
        assert not pkg.has_part(orphan_part)
        assert report["purge"]["bytes_freed"] > 0
        pkg.save()  # atomic save validates the swept package

        reopened = PptxPackage(path)
        assert not reopened.has_part(orphan_part)
        # the referenced image survived
        assert tv.media_usage(reopened), "referenced media must remain"

    def test_unused_layout_purge_on_real_deck(self, tmp_path):
        path = _copy(tmp_path, "military_brief.pptx")
        pkg = PptxPackage(path)
        unused_before = tv.unused_layouts(pkg)
        slide_count = len(pkg.slide_parts())

        report = compress_deck(pkg, dry_run=False)
        purged_layouts = report["purge"].get("layouts", [])
        assert set(purged_layouts) == set(unused_before)
        pkg.save()

        reopened = PptxPackage(path)
        assert len(reopened.slide_parts()) == slide_count
        assert tv.unused_layouts(reopened) == []
        for layout in purged_layouts:
            assert not reopened.has_part(layout)
        # every slide still resolves its layout
        assert tv.used_layouts(reopened) <= set(reopened.part_names())

    def test_orphan_notes_slide_purged(self, make_deck):
        from kitchensink4ppt.ops.notes import set_notes

        path = make_deck("orphan_notes.pptx")
        pkg = PptxPackage(path)
        set_notes(pkg, 0, "these notes will be orphaned")
        # orphan it: cut the slide -> notesSlide relationship by hand
        from kitchensink4ppt.core.package import rels_name
        from kitchensink4ppt.ops.read import notes_part_for

        slide_part = pkg.slide_parts()[0]
        notes_part = notes_part_for(pkg, slide_part)
        if notes_part is None:
            pytest.skip("synthetic deck did not produce a notes part")
        rels = pkg.rels_for(slide_part)
        for rel in list(rels.getroot()):
            if rel.get("Type", "").endswith("/notesSlide"):
                rels.getroot().remove(rel)
        pkg.mark_dirty(rels_name(slide_part))
        pkg.save()

        pkg = PptxPackage(path)
        assert notes_part in tv.orphan_notes_slides(pkg)
        report = compress_deck(pkg, dry_run=False)
        assert notes_part in report["purge"].get("notes_slides", [])
        pkg.save()
        assert not PptxPackage(path).has_part(notes_part)

    def test_purge_can_be_disabled(self, make_deck, tmp_path):
        path, orphan_part = self._orphan_media_deck(make_deck, tmp_path)
        pkg = PptxPackage(path)
        report = compress_deck(pkg, purge_unused=False)
        assert report["purge"] is None
        assert pkg.has_part(orphan_part)


# ================================================================== images


@pytest.mark.skipif(not _has_pillow(), reason="Pillow not installed")
class TestImageCompression:
    def _oversized_deck(self, make_deck, tmp_path) -> tuple[Path, str]:
        """A 1200px-wide image displayed at 1 inch: wildly over any sane
        DPI, so compress must downsample it."""
        big = make_corpus._png(tmp_path / "big.png", 1200, 900, (90, 90, 90))
        path = make_deck("bigimg.pptx")
        pkg = PptxPackage(path)
        res = insert_image(pkg, 0, str(big), 1, 1, w=1.0)
        pkg.save()
        return path, res["media_part"]

    def test_oversized_image_downsampled(self, make_deck, tmp_path):
        path, media_part = self._oversized_deck(make_deck, tmp_path)
        pkg = PptxPackage(path)
        size_before = len(pkg.raw_part(media_part))

        dry = compress_deck(pkg, max_dpi=150, dry_run=True)
        entry = next(e for e in dry["media"] if e["part"] == media_part)
        assert entry["action"] == "resize"
        assert entry["needed_px"]["w"] <= 200  # 1in * 150dpi + ceiling slack
        expected_savings = entry["savings_bytes"]
        assert len(pkg.raw_part(media_part)) == size_before, "dry run mutated!"

        report = compress_deck(pkg, max_dpi=150, dry_run=False)
        entry = next(e for e in report["media"] if e["part"] == media_part)
        assert entry["action"] == "resize"
        assert entry["savings_bytes"] == expected_savings
        assert len(pkg.raw_part(media_part)) < size_before
        assert report["image_savings_bytes"] >= expected_savings
        pkg.save()

        reopened = PptxPackage(path)
        from kitchensink4ppt.ops.media import image_size_px, sniff_format

        data = reopened.raw_part(media_part)
        assert sniff_format(data) == "png", "format never changes in place"
        px = image_size_px(data, "png")
        assert px is not None and px[0] <= 200

    def test_military_brief_real_compress_shrinks_file(self, tmp_path):
        path = _copy(tmp_path, "military_brief.pptx")
        size_before = path.stat().st_size
        pkg = PptxPackage(path)
        slide_count = len(pkg.slide_parts())

        report = compress_deck(pkg)
        assert report["pillow_available"] is True
        pkg.save()

        assert path.stat().st_size < size_before
        reopened = PptxPackage(path)
        assert len(reopened.slide_parts()) == slide_count
        # every remaining media reference still resolves
        for name, sources in tv.media_usage(reopened).items():
            assert reopened.has_part(name)
            assert sources, f"{name} survived compression but is unreferenced"


class TestNoPillowPath:
    def test_refuses_recompression_honestly(self, make_deck, tmp_path, monkeypatch):
        big = make_corpus._png(tmp_path / "big2.png", 800, 600, (60, 60, 60))
        path = make_deck("nopillow.pptx")
        pkg = PptxPackage(path)
        res = insert_image(pkg, 0, str(big), 1, 1, w=1.0)
        pkg.save()

        monkeypatch.setattr(optimize, "_pillow", lambda: (None, None))
        pkg = PptxPackage(path)
        report = compress_deck(pkg, dry_run=True)
        assert report["pillow_available"] is False
        assert "note" in report and "[optimize]" in report["note"]
        entry = next(e for e in report["media"] if e["part"] == res["media_part"])
        assert entry["action"] == "skipped-no-pillow"
        assert entry.get("est_savings_bytes", 0) > 0
        # the purge half still ran (pure XML)
        assert report["purge"] is not None


# ======================================================= COM validation gate


def _powerpnt_running() -> bool:
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
    )
    return any("POWERPNT.EXE" in ln.upper() for ln in out.stdout.splitlines())


def _powerpnt_still_running_after_grace(seconds: float = 20.0) -> bool:
    """A POWERPNT seen here may be a previous validator's instance still
    tearing down (it takes a few seconds to exit after Quit), not the
    user's. Poll briefly before concluding it is the user's session."""
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if not _powerpnt_running():
            return False
        time.sleep(1.0)
    return True


@pytest.mark.timeout(600)
def test_com_validator_on_compressed_deck(tmp_path):
    """PowerPoint must open the compressed + purged military brief clean."""
    if sys.platform != "win32":
        pytest.skip("COM validation is Windows-only")
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        pytest.skip("pywin32 not installed")
    from kitchensink4ppt.com import bridge

    if not bridge.powerpoint_installed():
        pytest.skip("PowerPoint is not installed on this machine")
    if _powerpnt_still_running_after_grace():
        pytest.skip(
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running (the "
            "user's instance). Compressed-deck COM validation did NOT run."
        )

    path = _copy(tmp_path, "military_brief.pptx")
    pkg = PptxPackage(path)
    compress_deck(pkg)
    pkg.save()

    validator = Path(__file__).resolve().parents[1] / "ppt_validator.py"
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(validator), str(path)],
        capture_output=True,
        text=True,
        timeout=570,
    )
    if "SKIPPED-USER-POWERPOINT-OPEN" in proc.stdout:
        pytest.skip("user PowerPoint opened mid-run; validator refused honestly")
    assert proc.returncode == 0, (
        f"COM validator failed (exit {proc.returncode})\n"
        f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
    )
    assert "PASS" in proc.stdout and "FAIL" not in proc.stdout
