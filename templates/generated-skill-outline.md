---
name: <brand>-ppt-style
description: Generate <brand>-native PowerPoint decks from supplied templates and brand guidance.
platforms: [linux, macos, windows]
---

# <Brand> PPT Style

## Source of truth

- `templates/` contains the actual PowerPoint source templates.
- `references/style-analysis.md` captures brand rules from the template XML and brand guidance.
- `references/layouts.md` maps slide purposes to named PowerPoint layouts.
- `references/components.md` defines reusable PPT-native components.
- `references/quality-checklist.md` defines P0/P1/P2 checks.

## Non-negotiables

1. Start from a supplied `.pptx` template.
2. Use named layouts and placeholders.
3. Preserve master, theme, footer, logos, and brand assets.
4. Do not freehand deck geometry unless a documented diagram grid is explicitly required.
5. Validate structure and render visually before delivery.

## Workflow

1. Inventory templates: `python3 scripts/extract_pptx_inventory.py templates/source.pptx`
2. Extract theme tokens (from the inventory output).
3. Build layout lock (prose) — and its structured form, a `deck-spec.yaml`.
4. Author the deck-spec (see `examples/deck-spec.example.yaml` for the schema).
5. Generate: `python3 scripts/generate_deck.py deck-spec.yaml --template templates/source.pptx`
6. Validate: `python3 scripts/validate_deck.py out/deck.pptx --spec deck-spec.yaml --template templates/source.pptx`
7. Render/inspect/fix: `python3 scripts/render_qa.py out/deck.pptx`
8. Package final `.pptx`.
