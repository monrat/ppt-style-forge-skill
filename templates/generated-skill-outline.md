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

1. Inventory templates.
2. Extract theme tokens.
3. Build layout lock.
4. Author JSON/YAML slide spec.
5. Generate deck.
6. Validate.
7. Render/inspect/fix.
8. Package final `.pptx`.
