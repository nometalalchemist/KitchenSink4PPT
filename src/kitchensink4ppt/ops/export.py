"""Export layer: PDF and slide-image export with engine auto-routing.

Contract (differs from the pure-package ops modules by design):
- Functions here take FILE PATHS, not a PptxPackage; export is a
  representation transform of the saved file, never a package mutation, and
  the source file is never modified by any path through this module.
- Engine routing: "auto" prefers PowerPoint COM (Windows + PowerPoint
  installed + pywin32), falling back to LibreOffice headless
  (soffice --headless --convert-to). "com" and "libreoffice" force one
  engine. When neither engine exists the error names both options honestly.
- COM is ground truth for fidelity; LibreOffice output drifts on theme
  colors, some effects, and font substitution (feasibility doc section 7.2).
- Slide images via LibreOffice need a PDF rasterizer: soffice's own
  --convert-to png emits only the FIRST slide, so the LO image path converts
  to PDF then rasterizes with pdftoppm (poppler) when present, and refuses
  honestly when it is not. COM has no such dependency.
- Heavy imports (win32com via com.bridge) stay inside function bodies;
  this module imports cleanly on any platform.
- Every path parameter passes through the sandbox check_path at entry, and
  every produced artifact is verified to exist and be non-empty.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from ..core.errors import DocumentNotFound, PptMcpError
from ..core.sandbox import check_path

SOFFICE_WELL_KNOWN = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
SOFFICE_TIMEOUT = 300  # seconds; big decks convert slowly
PDFTOPPM_TIMEOUT = 300

DEFAULT_IMAGE_WIDTH = 1280


# ------------------------------------------------------------- detection


def _find_soffice() -> Path | None:
    env = os.environ.get("KS4P_SOFFICE")
    if env and Path(env).exists():
        return Path(env)
    if SOFFICE_WELL_KNOWN.exists():
        return SOFFICE_WELL_KNOWN
    hit = shutil.which("soffice")
    return Path(hit) if hit else None


def _find_pdftoppm() -> Path | None:
    hit = shutil.which("pdftoppm")
    return Path(hit) if hit else None


def _com_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    from ..com.bridge import powerpoint_installed

    return powerpoint_installed()


def get_export_engines() -> dict:
    """What can this machine export with, right now?"""
    soffice = _find_soffice()
    pdftoppm = _find_pdftoppm()
    com_ok = _com_available()
    engines = {
        "powerpoint_com": {
            "available": com_ok,
            "pdf": com_ok,
            "slide_images": com_ok,
            "note": "ground truth for fidelity; Windows + PowerPoint + pywin32",
        },
        "libreoffice": {
            "available": soffice is not None,
            "path": str(soffice) if soffice else None,
            "pdf": soffice is not None,
            "slide_images": soffice is not None and pdftoppm is not None,
            "note": (
                "headless fallback; fidelity drifts on theme colors, effects, "
                "fonts. Slide images additionally need pdftoppm (poppler): "
                + ("found" if pdftoppm else "NOT found")
            ),
        },
    }
    return {"engines": engines, "auto_pdf": _pick_engine_name(), "auto_images": _pick_image_engine_name()}


def _pick_engine_name() -> str | None:
    if _com_available():
        return "com"
    if _find_soffice() is not None:
        return "libreoffice"
    return None


def _pick_image_engine_name() -> str | None:
    if _com_available():
        return "com"
    if _find_soffice() is not None and _find_pdftoppm() is not None:
        return "libreoffice"
    return None


def _no_engine_error(what: str) -> PptMcpError:
    return PptMcpError(
        f"no {what} engine is available on this machine. Options: install "
        "Microsoft PowerPoint (COM path, Windows only, needs pywin32) or "
        "LibreOffice (headless soffice; set KS4P_SOFFICE if it is in a "
        "non-standard location)."
    )


# ------------------------------------------------------------- helpers


def _require_file(path: str, purpose: str) -> Path:
    path = check_path(path, purpose)
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    return p


def _verify_output(out: Path, what: str, engine: str) -> None:
    if not out.exists() or out.stat().st_size == 0:
        raise PptMcpError(
            f"{engine} reported success but produced no {what} at {out}"
        )


def _run_soffice(soffice: Path, args: list[str], workdir: Path) -> None:
    """Run soffice headless with an ISOLATED user profile (a running
    LibreOffice GUI otherwise makes headless conversion silently no-op)."""
    profile = workdir / "lo_profile"
    profile.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(soffice),
        "--headless",
        "--norestore",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        *args,
    ]
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=SOFFICE_TIMEOUT
        )
    except subprocess.TimeoutExpired as exc:
        raise PptMcpError(
            f"LibreOffice conversion timed out after {SOFFICE_TIMEOUT}s"
        ) from exc
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "").strip()[:500]
        raise PptMcpError(
            f"LibreOffice conversion failed (exit {result.returncode}): {detail}"
        )


def _libreoffice_pdf(p: Path, out: Path) -> None:
    soffice = _find_soffice()
    if soffice is None:  # callers pre-check; belt and braces
        raise _no_engine_error("PDF export")
    with tempfile.TemporaryDirectory(prefix="ks4p_lo_") as td:
        tdir = Path(td)
        _run_soffice(
            soffice,
            ["--convert-to", "pdf", "--outdir", str(tdir), str(p.resolve())],
            tdir,
        )
        produced = tdir / (p.stem + ".pdf")
        _verify_output(produced, "PDF", "LibreOffice")
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(out))
    _verify_output(out, "PDF", "LibreOffice")


# ------------------------------------------------------------- PDF


def export_pdf(pkg_path: str, output: str | None = None, engine: str = "auto") -> dict:
    """Export a presentation to PDF. engine: "auto" | "com" | "libreoffice".

    Auto prefers PowerPoint COM (fidelity ground truth) and falls back to
    LibreOffice headless. The source file is never modified (COM opens
    ReadOnly and exports a copy; LibreOffice reads only)."""
    p = _require_file(pkg_path, "PDF export source")
    if output:
        output = check_path(output, "PDF export output")
    out = Path(output) if output else p.with_suffix(".pdf")

    if engine not in ("auto", "com", "libreoffice"):
        raise PptMcpError(
            f"unknown engine '{engine}'; use auto, com, or libreoffice"
        )
    chosen = engine
    if engine == "auto":
        chosen = _pick_engine_name()
        if chosen is None:
            raise _no_engine_error("PDF export")

    if chosen == "com":
        if not _com_available():
            raise PptMcpError(
                "engine='com' but PowerPoint COM is not available here "
                "(needs Windows, PowerPoint, and pywin32); LibreOffice "
                + ("IS available, use engine='libreoffice'."
                   if _find_soffice() else "is not available either.")
            )
        from ..com.bridge import com_export_pdf

        return com_export_pdf(str(p), str(out))

    # libreoffice
    if _find_soffice() is None:
        raise PptMcpError(
            "engine='libreoffice' but soffice was not found (checked "
            f"KS4P_SOFFICE, {SOFFICE_WELL_KNOWN}, and PATH)."
        )
    _libreoffice_pdf(p, out)
    return {"pdf": str(out), "bytes": out.stat().st_size, "engine": "libreoffice"}


# ------------------------------------------------------------- slide images


def export_slide_images(
    pkg_path: str,
    output_dir: str | None = None,
    slides: list[int] | None = None,
    width: int = DEFAULT_IMAGE_WIDTH,
    height: int | None = None,
    engine: str = "auto",
) -> dict:
    """Export slides as PNG images (the render-to-verify primitive).

    `slides` takes 0-based presentation-order indices; None = all slides.
    engine "auto" prefers PowerPoint COM; the LibreOffice path converts to
    PDF then rasterizes with pdftoppm (poppler) and refuses honestly when
    pdftoppm is missing (soffice alone can only render the first slide)."""
    p = _require_file(pkg_path, "slide image export source")
    if output_dir:
        output_dir = check_path(output_dir, "slide image output directory")
    out_dir = Path(output_dir) if output_dir else p.with_name(p.stem + "_slides")

    if engine not in ("auto", "com", "libreoffice"):
        raise PptMcpError(
            f"unknown engine '{engine}'; use auto, com, or libreoffice"
        )
    chosen = engine
    if engine == "auto":
        chosen = _pick_image_engine_name()
        if chosen is None:
            hint = ""
            if _find_soffice() is not None and _find_pdftoppm() is None:
                hint = (
                    " LibreOffice is installed but slide images also need "
                    "pdftoppm (poppler) for PDF rasterization; install "
                    "poppler or use PowerPoint COM."
                )
            raise PptMcpError(
                "no slide-image engine is available on this machine. Options: "
                "Microsoft PowerPoint (COM) or LibreOffice + poppler "
                "(pdftoppm)." + hint
            )

    if chosen == "com":
        if not _com_available():
            raise PptMcpError(
                "engine='com' but PowerPoint COM is not available here "
                "(needs Windows, PowerPoint, and pywin32)."
            )
        from ..com.bridge import com_export_slide_images

        return com_export_slide_images(
            str(p),
            str(out_dir),
            slides=slides,
            width=width,
            height=height if height is not None else None,
        )

    # libreoffice: pptx -> pdf -> pdftoppm PNGs
    if _find_soffice() is None:
        raise PptMcpError(
            "engine='libreoffice' but soffice was not found (checked "
            f"KS4P_SOFFICE, {SOFFICE_WELL_KNOWN}, and PATH)."
        )
    pdftoppm = _find_pdftoppm()
    if pdftoppm is None:
        raise PptMcpError(
            "the LibreOffice slide-image path needs pdftoppm (poppler) to "
            "rasterize the intermediate PDF; soffice alone renders only the "
            "first slide. Install poppler or use the PowerPoint COM engine."
        )
    if width <= 0:
        raise PptMcpError(f"width must be positive, got {width}")
    out_dir.mkdir(parents=True, exist_ok=True)
    images: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="ks4p_lo_img_") as td:
        tdir = Path(td)
        pdf_tmp = tdir / (p.stem + ".pdf")
        _libreoffice_pdf(p, pdf_tmp)
        scale = ["-scale-to-x", str(width), "-scale-to-y", "-1"]
        if height is not None:
            scale = ["-scale-to-x", str(width), "-scale-to-y", str(height)]
        try:
            result = subprocess.run(
                [
                    str(pdftoppm), "-png", *scale,
                    str(pdf_tmp), str(tdir / "page"),
                ],
                capture_output=True,
                text=True,
                timeout=PDFTOPPM_TIMEOUT,
            )
        except subprocess.TimeoutExpired as exc:
            raise PptMcpError(
                f"pdftoppm timed out after {PDFTOPPM_TIMEOUT}s"
            ) from exc
        if result.returncode != 0:
            detail = (result.stderr or "").strip()[:500]
            raise PptMcpError(f"pdftoppm failed (exit {result.returncode}): {detail}")
        pages = sorted(tdir.glob("page-*.png"))
        if not pages:
            raise PptMcpError(
                "pdftoppm reported success but produced no PNG pages"
            )
        wanted = list(range(len(pages))) if slides is None else list(slides)
        for idx in wanted:
            if not isinstance(idx, int) or idx < 0 or idx >= len(pages):
                raise PptMcpError(
                    f"slide index {idx} out of range, presentation has "
                    f"{len(pages)} (indices are 0-based)"
                )
            dest = out_dir / f"slide{idx + 1}.png"
            shutil.move(str(pages[idx]), str(dest))
            _verify_output(dest, f"PNG for slide {idx + 1}", "LibreOffice")
            images.append({"slide": idx, "file": str(dest)})
    return {
        "images": images,
        "width": width,
        "height": height,
        "engine": "libreoffice",
        "fidelity_note": (
            "LibreOffice rendering; theme colors, some effects, and fonts "
            "may drift from PowerPoint output"
        ),
    }
