# PPT Style Forge Skill

A Hermes skill for turning a PowerPoint template plus brand guidance / brand book into a reusable PPT generation style skill.

It captures a repeatable workflow for building brand-native PowerPoint generators: extract template structure, derive brand tokens, write layout rules, create validator scripts, package reference assets, and verify generated decks.

## Install / use

Copy `SKILL.md` and the `references/`, `scripts/`, `templates/`, and `examples/` folders into a Hermes skill directory, or load this repository as a skill source.

## What it is for

Use when you have:

- a `.pptx` template, master deck, or design system deck;
- a brand book / brand guidance pages / reference PDFs;
- a need to generate new decks that stay template-native, not merely “brand-coloured”.

## Core principle

Do not ask an LLM to freehand a deck from brand colours. Build the new skill around the actual PowerPoint master, layouts, placeholders, theme XML, and visual QA loop.
