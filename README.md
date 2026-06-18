# PPT Style Forge Skill

A Hermes skill for turning a PowerPoint template plus brand guidance / brand book into a reusable PPT generation style skill.

It captures a repeatable workflow for building brand-native PowerPoint generators: extract template structure, derive brand tokens, write layout rules, create validator scripts, package reference assets, and verify generated decks.

## What it is for

Use when you have:

- a `.pptx` template, master deck, or design system deck;
- a brand book / brand guidance pages / reference PDFs;
- a need to generate new decks that stay template-native, not merely "brand-coloured".

## Core principle

Do not ask an LLM to freehand a deck from brand colours. Build the new skill around the actual PowerPoint master, layouts, placeholders, theme XML, and visual QA loop.

## Install

```bash
pip install python-pptx          # required by generate_deck.py ONLY;
                                 # validate_deck.py is stdlib-only
# optional, for visual QA per-slide PNG rendering:
pip install pywin32              # Windows: PowerPoint COM (best per-slide output)
#   OR install LibreOffice + poppler:  https://www.libreoffice.org
#      (LibreOffice -> PDF -> pdftoppm gives per-slide PNG cross-platform)
```

Copy `SKILL.md` and the `references/`, `scripts/`, `templates/`, and `examples/` folders into a Hermes skill directory, or load this repository as a skill source.

## Quickstart

```bash
# 1. Inventory the template — see its real layouts, theme, and placeholders.
python3 scripts/extract_pptx_inventory.py templates/source.pptx > references/template-inventory.txt

# 2. Author a deck-spec that locks each slide to a real named layout.
#    See examples/deck-spec.example.yaml for the schema.

# 3. Generate the deck from the spec + template.
python3 scripts/generate_deck.py examples/deck-spec.example.yaml --template templates/source.pptx

# 4. Validate against the 9 P0 rules (stdlib-only; exit code non-zero on failure).
python3 scripts/validate_deck.py out/example.pptx \
  --spec examples/deck-spec.example.yaml --template templates/source.pptx

# 5. Render each slide to an image for visual QA.
python3 scripts/render_qa.py out/example.pptx
#    -> slide1.png, slide2.png, ...  (PowerPoint COM, or LibreOffice + pdftoppm)
#    Without a renderer, it prints a clear notice and exits 0 — the deck is
#    still valid; open it manually for visual QA.
```

> **About the two example specs:** `input-spec.example.yaml` describes the
> forge skill's *inputs* (templates + brand book → a new brand skill);
> `deck-spec.example.yaml` is the *generator's* input (per-slide layout +
> fields → one `.pptx`). They operate at different levels — see each file's
> header.

Fix issues by editing the deck-spec and re-running from step 3 — never by hand-patching the `.pptx`. See `SKILL.md` for the full workflow and the Scripts quick-reference table.

## Self-test

A brand-flavoured test template + an automated end-to-end smoke test live under
`tests/`. The smoke test builds the fixture, runs inventory → generate →
validate, and asserts each step succeeds (validate PASSES with python-pptx
blocked, proving it is truly stdlib-only):

```bash
bash tests/smoke_test.sh
```

To rebuild the fixture by hand:

```bash
python3 scripts/_make_test_template.py tests/fixtures/sample.pptx
```
