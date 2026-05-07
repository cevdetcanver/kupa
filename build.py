#!/usr/bin/env python3
"""build.py — page template'lerini partial'larla derler.

Her sayfanın kendi PATH_PREFIX'i var (relative URL'ler icin).
"""

import re
from pathlib import Path

ROOT = Path(__file__).parent
TPL = ROOT / "templates"
PARTIALS = TPL / "partials"
PAGES = TPL / "pages"

PAGES_CONFIG = {
    "index.html": {
        "out": ROOT / "index.html",
        "vars": {"PATH_PREFIX": ""},
    },
    "detay.html": {
        "out": ROOT / "detay" / "index.html",
        "vars": {"PATH_PREFIX": "../"},
    },
}

INCLUDE_RE = re.compile(r"\{\{INCLUDE\s+([^\s|}]+)\s*\}\}")


def expand_includes(text: str) -> str:
    prev = None
    cur = text
    while cur != prev:
        prev = cur
        cur = INCLUDE_RE.sub(
            lambda m: (PARTIALS / m.group(1)).read_text(encoding="utf-8").rstrip("\n"),
            cur,
        )
    return cur


def apply_vars(text: str, variables: dict) -> str:
    for k, v in variables.items():
        text = text.replace(f"{{{{{k}}}}}", v)
    return text


def main():
    for tpl_name, cfg in PAGES_CONFIG.items():
        tpl = (PAGES / tpl_name).read_text(encoding="utf-8")
        rendered = expand_includes(tpl)
        rendered = apply_vars(rendered, cfg["vars"])
        cfg["out"].parent.mkdir(parents=True, exist_ok=True)
        cfg["out"].write_text(rendered, encoding="utf-8")
        print(f"OK  {cfg['out'].relative_to(ROOT)}  ({len(rendered):,} bytes)")


if __name__ == "__main__":
    main()
