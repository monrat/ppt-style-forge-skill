# PowerPoint Template Analysis

## Why template-native matters

A brand-coloured slide is not necessarily a brand-native slide. Fidelity comes from:

1. PowerPoint master and theme
2. named layouts
3. title/body placeholder geometry
4. footer/logo/identity elements
5. typography and colour tokens
6. charts/tables/image placeholders

## Inventory command

Use the bundled script:

```bash
python3 scripts/extract_pptx_inventory.py path/to/template.pptx > inventory.txt
```

## Layout lock

Before generating a deck, write a layout lock:

```text
slide 1 -> Layout Name -> why this layout -> fields/placeholders to fill
slide 2 -> Layout Name -> why this layout -> fields/placeholders to fill
```

Never use an unconfirmed layout name.

## Placeholder fill order

For multi-frame layouts, sort text placeholders by visual reading order: top row first, then left-to-right. Beware inherited layout geometry: a slide placeholder may not carry explicit coordinates, so fallback to layout placeholder coordinates.

## Tables and charts

A fresh table/chart placeholder may not contain an actual table/chart object until inserted. Validate by extracting text/cells or opening/rendering the file.
