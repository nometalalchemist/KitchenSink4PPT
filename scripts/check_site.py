"""Copy guard for docs/index.html and docs/llms.txt.

Checks the things that break quietly on a multilingual single-page site:
every dictionary carries the same keys, every data-i18n key in the markup
resolves, the JSON payload parses, no em dashes reach a reader in any
language, and the stated counts agree with the pack registry.

Run:  .venv/Scripts/python.exe -X utf8 scripts/check_site.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

DOCS = ROOT / "docs"
EM_DASH = "—"
LANGS = ["en", "ko", "ja", "zh", "de", "fr", "es"]

failures: list[str] = []


def fail(msg: str) -> None:
    failures.append(msg)


def main() -> int:
    html = (DOCS / "index.html").read_text(encoding="utf-8")

    # 1. the dictionaries parse as JSON and carry identical key sets
    block = html[html.index("var I18N = {") + len("var I18N = "):
                 html.index("/*I18N-END*/")].strip().rstrip(";")
    block = re.sub(r"^(\w+): \{", r'"\1": {', block, flags=re.M)
    try:
        dicts = json.loads(block)
    except json.JSONDecodeError as exc:
        fail(f"I18N block is not valid JSON: {exc}")
        return report()

    if list(dicts) != LANGS:
        fail(f"languages are {list(dicts)}, expected {LANGS}")
    base = set(dicts["en"])
    for lang, d in dicts.items():
        missing = base - set(d)
        extra = set(d) - base
        if missing:
            fail(f"{lang} is missing keys: {sorted(missing)}")
        if extra:
            fail(f"{lang} has keys en does not: {sorted(extra)}")

    # 2. every key the markup asks for exists, and nothing is dead weight
    used = set(re.findall(r'data-i18n="([^"]+)"', html))
    used.add("pageTitle")
    used.add("tale.label")
    for key in sorted(used - base):
        fail(f"markup uses {key!r}, which no dictionary defines")
    for key in sorted(base - used):
        fail(f"dictionary defines {key!r}, which the markup never uses")

    # 3. no em dashes anywhere a reader can see, in any language
    for lang, d in dicts.items():
        for key, value in d.items():
            if EM_DASH in value:
                fail(f"em dash in {lang}/{key}")
    visible = re.sub(r"<script.*?</script>", "", html, flags=re.S)
    visible = re.sub(r"<style.*?</style>", "", visible, flags=re.S)
    visible = re.sub(r"<!--.*?-->", "", visible, flags=re.S)
    if EM_DASH in visible:
        fail("em dash in the page markup")

    llms = (DOCS / "llms.txt").read_text(encoding="utf-8")
    if EM_DASH in llms:
        fail("em dash in llms.txt")

    # 4. the numbers on the page agree with the live registry
    from kitchensink4ppt import packs, server  # noqa: F401

    names = packs.tool_names()
    total = sum(len(v) for v in names.values())
    lite = len(names["lite"])
    if f'<span class="n">{total}</span>' not in html:
        fail(f"specs strip does not state the real tool total ({total})")
    for lang, d in dicts.items():
        if str(total) not in d["m.catalog"]:
            fail(f"{lang}/m.catalog does not carry the tool total {total}")
        if str(lite) not in d["pk1.c"]:
            fail(f"{lang}/pk1.c does not carry the lite core count {lite}")

    # 5. the inventory shows one card per pack, lite included
    cards = len(re.findall(r'data-i18n="pk\d+\.t"', html))
    expected = len(packs.pack_names()) + 1
    if cards != expected:
        fail(f"{cards} inventory cards for {expected} packs (lite included)")

    # 6. retired pack names must not resurface as page copy
    for lang, d in dicts.items():
        for key, value in d.items():
            for dead in ("v1.0", "com-live", "transitions-animations"):
                if dead in value:
                    fail(f"{lang}/{key} still says {dead!r}")

    return report()


def report() -> int:
    if failures:
        print(f"FAIL ({len(failures)}):")
        for f in failures:
            print("  -", f)
        return 1
    print("site copy guard: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
