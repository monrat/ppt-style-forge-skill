#!/usr/bin/env python3
"""Extract a compact PowerPoint template inventory using only stdlib.

Thin, human-readable formatter over the shared `pptx_io` parser. The structured
data lives in pptx_io.inventory(); this script just pretty-prints it so a human
(or an LLM) can read the mechanical truth of a template at a glance. It does
not modify the deck.

    python3 scripts/extract_pptx_inventory.py path/to/template.pptx
"""
from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

# Allow running both as `python scripts/extract_pptx_inventory.py` (script dir is
# on sys.path automatically) and as a module.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import pptx_io  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path, help="Path to a .pptx template")
    args = parser.parse_args()

    with zipfile.ZipFile(args.pptx) as zf:
        inv = pptx_io.inventory(zf)

    w, h = inv["slide_size"]
    print(f"slide_size: {w} x {h} in")
    print(f"active_slides: {inv['active_slides']}")

    for theme in inv["themes"]:
        print(f"\nTHEME {theme['part']}")
        print("  major_font:", theme["major_font"])
        print("  minor_font:", theme["minor_font"])
        for name, hexval in theme["colors"].items():
            print(f"  {name}: {hexval}")

    print(f"\nLAYOUTS: {len(inv['layouts'])}")
    for layout in inv["layouts"]:
        print(f"\n- {layout['name']}")
        for ph in layout["placeholders"]:
            box = ph["box"]
            sample = ph["text"][:80]
            print(f"    {ph['ph_type']}#{ph['ph_idx']} box={box} text={sample!r}")

    print(f"\nSLIDE_PARTS: {len(inv['slides'])}")
    for s in inv["slides"]:
        print(f"  {s['part']}: layout={s['layout']!r} text={s['text']!r}")


if __name__ == "__main__":
    main()
