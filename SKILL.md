---
name: ppt-style-forge
description: Build a reusable PowerPoint style-generation skill from a PPT template plus brand book/guidance. Extracts template layouts/theme, derives brand rules, writes validators, packages assets, and enforces visual QA.
version: 0.1.0
author: Hermes Agent
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [PowerPoint, brand, design-system, templates, slides, skill-generation]
    related_skills: [powerpoint, github]
---

# PPT Style Forge

Use this skill when the user asks to turn a **PowerPoint template, brand book, brand guidance site, or design-system deck** into a reusable PPT generation style skill.

The output is a new skill directory that can generate brand-native decks from the supplied template. The goal is not a pretty one-off deck; the goal is a repeatable, validated workflow.

## Core principle

**Do not recreate the brand from memory.** Start from the actual `.pptx` template, then extract and preserve its masters, layouts, placeholders, theme, footer/logo bands, fonts, colours, and built-in assets. Brand guidance calibrates the rules; the PowerPoint file supplies the mechanical truth.

## When to use

- A user provides a PPT template and asks for a deck-generation skill.
- A user provides a brand book and wants pixel/brand fidelity in generated PPTs.
- A user wants a Claude/Hermes design-system-like workflow for PowerPoint.
- A user wants to package a brand-specific PPT skill for reuse or GitHub release.

## Inputs to collect

1. Source `.pptx` template(s): report, live, pitch, divider, etc.
2. Brand book / PDF / guidance URLs.
3. Reference decks or screenshots, if available.
4. Intended deck types and audience.
5. Known hard rules: colours, fonts, logo/footer usage, chart/icon rules, accessibility, legal text.

## Required workflow

### 1. Inventory every template

Run:

```bash
python3 scripts/extract_pptx_inventory.py templates/source.pptx > references/template-inventory.txt
```

Capture:

- slide size
- layouts and placeholder types
- placeholder coordinates
- theme fonts and colours
- footer/logo/identity elements
- sample placeholder text that must never leak into output

### 2. Extract brand rules

Read all brand guidance. Separate:

- mechanical facts from template XML;
- brand policy from guidance;
- visual calibration from screenshots;
- assets that are available vs merely referenced.

If an official asset library is not available, do **not** pretend it is. Use no asset, or clearly label generated/custom assets as non-official.

### 3. Write the new skill structure

Create:

```text
<brand>-ppt-style/
├── SKILL.md
├── templates/
├── references/
│   ├── style-analysis.md
│   ├── layouts.md
│   ├── components.md
│   ├── quality-checklist.md
│   └── color-tokens.json
├── scripts/
│   ├── generate_deck.py
│   ├── validate_deck.py
│   └── extract_template_inventory.py
├── examples/
└── assets/
```

### 4. Build a layout library

For each named layout, document:

- exact layout name;
- placeholder list;
- intended use;
- fields the slide spec should provide;
- do/don't notes;
- whether pictures/charts/tables need special handling.

Before generation, create a **layout lock**: one line per slide mapping page → exact layout name → why → fields.

### 5. Generate from placeholders, not freehand geometry

Rules:

- Add fresh slides from named layouts.
- Fill placeholders; do not duplicate populated sample slides.
- Fill every text/title content placeholder or choose a simpler layout.
- Preserve masters, theme, footers, logos, and relationship parts.
- Use manual coordinates only for genuine diagrams, and document them.

### 6. Validate mechanically

Validator must check at least:

- slide size;
- active slide count vs intended count;
- theme fonts/colours;
- named layout usage;
- empty text/title placeholders;
- leftover sample text;
- table/chart data actually present;
- banned/off-brand colours;
- footer/logo/identity preservation.

### 7. Render and visually QA

A first render is almost never final. Convert to images or inspect in a canonical renderer and look for:

- overlap;
- text overflow;
- wrapped titles colliding with decorative elements;
- footer/logo collisions;
- inconsistent gaps;
- low contrast;
- bad placeholder ordering;
- leftover sample content.

Fix and re-run validation. Do not declare success without a fix-and-verify loop unless a real renderer is unavailable; if unavailable, state the limitation.

### 8. Package for reuse

Include:

- source templates if redistribution is allowed;
- scripts;
- references;
- examples;
- asset manifests;
- clear notes for unavailable/proprietary assets.

Do not publish proprietary brand books, fonts, templates, or assets unless the user explicitly owns/approves publication.

## Output quality bar

A good generated style skill should let a future agent produce a deck that:

- looks native to the supplied template;
- passes mechanical validation;
- has documented layout choices;
- avoids hallucinated assets;
- can be fixed by editing a spec and regenerating, not by manually patching slides.

## Pitfalls

- Brand book only is insufficient for pixel fidelity; the `.pptx` template is needed.
- Using brand colours on freehand slides is not enough.
- Empty placeholders and sample copy are common template-generation failures.
- Default Office chart colours often leak into decks.
- Icon libraries are often referenced in guidance but not actually available; do not crop website screenshots as production icons.
- Multi-column placeholder order can be wrong if inherited layout coordinates are ignored.

## Linked references

- `references/brand-extraction.md`
- `references/ppt-template-analysis.md`
- `references/quality-checklist.md`
- `templates/generated-skill-outline.md`
