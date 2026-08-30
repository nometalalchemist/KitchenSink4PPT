"""Path sandboxing (KS4P_ALLOWED_ROOTS): opt-in containment for every path
the server touches.

Self-contained: decks are built on the fly with python-pptx. The env var is
manipulated through monkeypatch; the sandbox module re-parses whenever the
raw env value changes, so no reload tricks are needed.
"""

import os
import subprocess
import sys
from pathlib import Path

import pytest

from kitchensink4ppt.core import safesave, sandbox
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.core.sandbox import SandboxViolation, check_path

ENV = sandbox.ENV_VAR


# ------------------------------------------------------------------ fixtures


def _fresh_deck(where: Path, name: str = "deck.pptx") -> Path:
    import make_corpus

    where.mkdir(parents=True, exist_ok=True)
    return make_corpus.build_deck(where / name, seed=1, extra_slides=0)


@pytest.fixture()
def unsandboxed(monkeypatch):
    monkeypatch.delenv(ENV, raising=False)


@pytest.fixture()
def root(tmp_path, monkeypatch):
    """One allowed root at tmp_path/inside; a sibling escape dir beside it."""
    inside = tmp_path / "inside"
    inside.mkdir()
    (tmp_path / "outside").mkdir()
    monkeypatch.setenv(ENV, str(inside))
    return inside


# ------------------------------------------------------- unset = unrestricted


class TestUnset:
    def test_check_path_is_identity_when_unset(self, unsandboxed):
        weird = r"C:\definitely\not\a\real\place\..\thing.pptx"
        assert check_path(weird, "test") == weird
        assert sandbox.active() is False

    def test_empty_value_means_unrestricted(self, monkeypatch):
        monkeypatch.setenv(ENV, "")
        assert sandbox.active() is False
        monkeypatch.setenv(ENV, f" {os.pathsep} ")
        assert sandbox.active() is False

    def test_pptx_open_anywhere_when_unset(self, unsandboxed, tmp_path):
        f = _fresh_deck(tmp_path)
        pkg = PptxPackage(f)
        assert pkg.presentation() is not None


# --------------------------------------------------------------- containment


class TestContainment:
    def test_inside_root_allowed(self, root):
        p = check_path(root / "sub" / "a.pptx", "test")
        assert os.path.normcase(p).startswith(os.path.normcase(str(root)))

    def test_root_itself_allowed(self, root):
        assert check_path(root, "test")

    def test_outside_blocked(self, root, tmp_path):
        with pytest.raises(SandboxViolation):
            check_path(tmp_path / "outside" / "a.pptx", "test")

    def test_traversal_escape_blocked(self, root, tmp_path):
        sneaky = root / "sub" / ".." / ".." / "outside" / "a.pptx"
        with pytest.raises(SandboxViolation):
            check_path(sneaky, "test")

    def test_prefix_collision_blocked(self, tmp_path, monkeypatch):
        docs = tmp_path / "Documents"
        docs2 = tmp_path / "Documents2"
        docs.mkdir()
        docs2.mkdir()
        monkeypatch.setenv(ENV, str(docs))
        assert check_path(docs / "a.pptx", "test")
        with pytest.raises(SandboxViolation):
            check_path(docs2 / "a.pptx", "test")

    def test_case_differences_match(self, root):
        if os.path.normcase("A") != os.path.normcase("a"):
            pytest.skip("case-sensitive filesystem")
        assert check_path(str(root).upper() + os.sep + "a.pptx", "test")

    def test_multi_root(self, tmp_path, monkeypatch):
        r1 = tmp_path / "r1"
        r2 = tmp_path / "r2"
        r1.mkdir()
        r2.mkdir()
        monkeypatch.setenv(ENV, os.pathsep.join([str(r1), str(r2)]))
        assert check_path(r1 / "a.pptx", "test")
        assert check_path(r2 / "b.pptx", "test")
        with pytest.raises(SandboxViolation):
            check_path(tmp_path / "elsewhere.pptx", "test")

    def test_unc_refused_for_local_roots(self, root):
        with pytest.raises(SandboxViolation) as exc:
            check_path(r"\\some-server\share\deck.pptx", "test")
        assert "UNC" in str(exc.value)

    def test_extended_length_unc_refused(self, root):
        with pytest.raises(SandboxViolation):
            check_path(r"\\?\UNC\some-server\share\deck.pptx", "test")

    @pytest.mark.skipif(os.name != "nt", reason="\\\\?\\ paths are Windows-only")
    def test_extended_length_local_normalized(self, root):
        p = check_path("\\\\?\\" + str(root / "a.pptx"), "test")
        assert not p.startswith("\\\\?\\")

    def test_nonexistent_create_target_inside_allowed(self, root):
        target = root / "new" / "deeper" / "created.pptx"
        assert not target.parent.exists()
        got = check_path(target, "test")
        assert os.path.normcase(got).endswith("created.pptx")

    def test_nonexistent_create_target_outside_blocked(self, root, tmp_path):
        with pytest.raises(SandboxViolation):
            check_path(tmp_path / "outside" / "new" / "created.pptx", "test")

    def test_null_byte_refused(self, root):
        with pytest.raises(SandboxViolation):
            check_path(str(root) + os.sep + "a\x00b.pptx", "test")

    def test_error_names_roots_and_env_var(self, root, tmp_path):
        with pytest.raises(SandboxViolation) as exc:
            check_path(tmp_path / "outside" / "a.pptx", "open presentation")
        msg = str(exc.value)
        assert ENV in msg
        assert str(root) in msg
        assert "outside" in msg  # the offending path appears
        assert "open presentation" in msg  # the purpose appears

    def test_env_change_is_seen_without_reload(self, tmp_path, monkeypatch):
        a = tmp_path / "a"
        b = tmp_path / "b"
        a.mkdir()
        b.mkdir()
        monkeypatch.setenv(ENV, str(a))
        with pytest.raises(SandboxViolation):
            check_path(b / "x.pptx", "test")
        monkeypatch.setenv(ENV, str(b))
        assert check_path(b / "x.pptx", "test")


# ---------------------------------------------------------- junction escapes


class TestJunction:
    @pytest.mark.skipif(os.name != "nt", reason="junctions are Windows-only")
    def test_junction_escape_blocked(self, root, tmp_path):
        """A junction inside the root pointing outside must not smuggle the
        target back in: realpath resolves it before containment."""
        target = tmp_path / "outside" / "secret"
        target.mkdir()
        link = root / "jump"
        proc = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(link), str(target)],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0 or not link.exists():
            pytest.skip(f"cannot create junction here: {proc.stderr.strip()}")
        try:
            with pytest.raises(SandboxViolation):
                check_path(link / "a.pptx", "test")
        finally:
            link.rmdir()  # removes the junction, not the target

    @pytest.mark.skipif(sys.platform != "win32", reason="symlink test uses os.symlink")
    def test_symlink_escape_blocked(self, root, tmp_path):
        target = tmp_path / "outside" / "linked"
        target.mkdir()
        link = root / "sym"
        try:
            os.symlink(str(target), str(link), target_is_directory=True)
        except OSError:
            pytest.skip("symlink creation not permitted (no developer mode)")
        try:
            with pytest.raises(SandboxViolation):
                check_path(link / "a.pptx", "test")
        finally:
            link.rmdir()


# ------------------------------------------------------ enforcement plumbing


class TestEnforcement:
    def test_pptx_open_outside_blocked(self, root, tmp_path):
        f = _fresh_deck(tmp_path / "outside")
        with pytest.raises(SandboxViolation):
            PptxPackage(f)

    def test_pptx_open_inside_allowed_and_saves(self, root):
        f = _fresh_deck(root)
        pkg = PptxPackage(f)
        pkg.save()

    def test_save_dest_outside_blocked(self, root, tmp_path):
        f = _fresh_deck(root)
        pkg = PptxPackage(f)
        with pytest.raises(SandboxViolation):
            pkg.save(tmp_path / "outside" / "escape.pptx")

    def test_write_lock_outside_blocked(self, root, tmp_path):
        with pytest.raises(SandboxViolation):
            with safesave.write_lock(tmp_path / "outside" / "a.pptx"):
                pass

    def test_normal_edit_cycle_unaffected_inside_root(self, root):
        """A full open-modify-save cycle inside the root works with the
        sandbox on, including the backup slot rotation beside the deck."""
        f = _fresh_deck(root)
        pkg = PptxPackage(f)
        pkg.save()  # first save: rotates slots under root/.ks4p-backups
        pkg2 = PptxPackage(f)
        assert pkg2.presentation() is not None
