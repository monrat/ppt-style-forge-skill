#!/usr/bin/env python3
"""Create a starter brand PPT style skill directory.

This script creates the documentation skeleton only. Copy source templates and
brand references yourself, then run the inventory/validation workflow described in
SKILL.md.
"""
from __future__ import annotations

import argparse
from pathlib import Path

SKILL_MD = """---
name: {name}
description: Generate {brand}-native PowerPoint decks from supplied templates and brand guidance.
platforms: [linux, macos, windows]
---

# {brand} PPT Style

## Source of truth

- `templates/` contains source PowerPoint templates.
- `references/style-analysis.md` captures brand facts and guidance.
- `references/layouts.md` maps slide purposes to named PowerPoint layouts.
- `references/components.md` defines reusable components.
- `references/quality-checklist.md` defines validation gates.

## Workflow

1. Inventory templates.
2. Extract theme and brand tokens.
3. Write layout lock.
4. Generate from named layouts/placeholders.
5. Validate mechanically.
6. Render and inspect visually.
7. Iterate until clean.
"""

FILES = {
    "references/style-analysis.md": "# Style Analysis\n\nFill from template XML + brand guidance.\n",
    "references/layouts.md": "# Layout Library\n\nDocument exact named layouts, placeholders, and use cases.\n",
    "references/components.md": "# Components\n\nDocument typography, footer/logo band, charts, tables, images, icons, and diagrams.\n",
    "references/quality-checklist.md": "# Quality Checklist\n\nUse the checklist from ppt-style-forge as the starting point.\n",
    "references/color-tokens.json": "{}\n",
    "examples/sample-spec.json": "{\n  \"slides\": []\n}\n",
    "assets/README.md": "# Assets\n\nStore approved redistributable assets here. Do not store proprietary assets unless approved.\n",
    "templates/README.md": "# Templates\n\nPlace source .pptx templates here if redistribution is allowed.\n",
    "scripts/README.md": "# Scripts\n\nAdd generator and validator scripts here.\n",
}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("out_dir", type=Path)
    parser.add_argument("--brand", required=True)
    parser.add_argument("--name", help="Skill directory/name, defaults to slugified brand + -ppt-style")
    args = parser.parse_args()

    name = args.name or args.brand.lower().replace(" ", "-") + "-ppt-style"
    root = args.out_dir / name
    root.mkdir(parents=True, exist_ok=True)
    (root / "SKILL.md").write_text(SKILL_MD.format(name=name, brand=args.brand), encoding="utf-8")
    for rel, content in FILES.items():
        path = root / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
    print(root)


if __name__ == "__main__":
    main()
