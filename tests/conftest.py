"""The test suite runs against a corpus of REAL presentations in tests/corpus/
(the author's own decks: dissertation defense, conference talks. Private, not
shipped).

Without the corpus, corpus-dependent tests are skipped with an explanation.
To run the full suite, drop your own .pptx/.potx files into tests/corpus/
using the expected names below; decks with real layouts, tables, pictures,
and speaker notes give the most meaningful coverage.
"""

from pathlib import Path

import pytest

CORPUS = Path(__file__).parent / "corpus"

# name -> features the tests expect it to contain
EXPECTED = {
    "proposal_defense.pptx": "26-slide templated defense deck (the key deck)",
    "nsu_pcsj.pptx": "conference talk deck",
    "unitar_final.pptx": "conference deck with images",
    "conference_template.potx": "PowerPoint template package (.potx)",
    "military_brief.pptx": "heavy deck with many layouts and shapes",
    "pmr_tables.pptx": "deck with tables",
}


def _present() -> set[str]:
    return {p.name for p in CORPUS.glob("*.pptx")} | {
        p.name for p in CORPUS.glob("*.potx")
    }


def pytest_configure(config):
    """Missing corpus files are GENERATED as structural stand-ins (see
    tests/make_corpus.py), so the full suite runs anywhere, CI included.
    Real local decks, when present, always take precedence."""
    if set(EXPECTED) - _present():
        import make_corpus  # noqa: F401  (lives beside this file)

        made = make_corpus.generate_missing(verbose=True)
        if made:
            print(f"conftest: generated {len(made)} synthetic corpus file(s)")


def pytest_collection_modifyitems(config, items):
    missing = set(EXPECTED) - _present()
    if not missing:
        return
    skip = pytest.mark.skip(
        reason=(
            f"test corpus incomplete (missing: {sorted(missing)}) and "
            "generation failed; run python tests/make_corpus.py for details."
        )
    )
    uses_corpus: dict[str, bool] = {}
    for item in items:
        fname = str(item.fspath)
        if fname not in uses_corpus:
            uses_corpus[fname] = "corpus" in Path(fname).read_text(
                encoding="utf-8", errors="ignore"
            )
        # Only corpus-dependent tests are skipped; corpus-free tests still run.
        if uses_corpus[fname]:
            item.add_marker(skip)


@pytest.fixture()
def make_deck(tmp_path):
    """Build a small synthetic deck in tmp_path and return its Path."""

    def _make(name: str = "deck.pptx", *, seed: int = 0, extra_slides: int = 2):
        import make_corpus

        return make_corpus.build_deck(
            tmp_path / name, seed=seed, extra_slides=extra_slides
        )

    return _make
