# Brand Extraction Workflow

Use this when converting a brand book / guidance site / template deck into a reusable PPT style skill.

## Inputs

- PowerPoint templates (`.pptx`) are the mechanical source of truth.
- Brand books, PDF guidance, and web pages are policy / visual intent sources.
- Reference screenshots are useful for visual calibration but should not replace template XML.

## Extract from templates

- slide size
- masters and layouts
- placeholder names, types, and coordinates
- theme fonts
- theme colours
- sample slide text and known placeholder copy
- chart/table/image placeholder behaviour
- footer/logo/identity band location

## Extract from brand guidance

- template family purposes
- primary and supporting palettes
- typography rules
- cover/image guidance
- icon rules and asset restrictions
- chart/diagram discipline
- sustainability/event/product-specific guidance
- hard do/don't rules

## Convert into skill rules

Write rules as operational constraints:

- Which template to choose when.
- Which named layout to choose for each content shape.
- Which placeholders a spec must fill.
- Which colours/fonts/assets are allowed.
- What the validator must reject.
- What visual QA must inspect.

## Do not overfit

If guidance says an asset must come from an approved library and the library is not available, do not fake it. Prefer no icon/asset, or label custom assets clearly.
