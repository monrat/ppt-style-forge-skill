#!/usr/bin/env python3
"""Generate a brand-native deck from a deck-spec and a source .pptx template.

This is the producer half of the PPT Style Forge workflow (SKILL.md step 5). It
is deliberately strict: it never invents geometry, never duplicates sample
slides, and refuses unknown layouts. The mechanical truth comes from the
template's named layouts; the deck-spec only says *which* layout each slide uses
and *what text* goes in each placeholder.

    python3 scripts/generate_deck.py deck-spec.yaml --template templates/src.pptx

Inputs (deck-spec), per slide:
    layout:   exact named layout from the template (REQUIRED, must exist)
    purpose:  free-text note carried into the report (optional)
    fields:   {placeholder_key: text}
              placeholder_key is matched by placeholder type ('title','body',
              'ctrTitle','subTitle') or by idx ('1','2'); first match wins.

Outputs:
    <output>.pptx                    the generated deck
    <output>.generation-report.json  per-slide record of layout + fields filled

Safety rules (enforced; violations abort generation):
    - layout must exist in the template
    - slides are created from layouts, never by copying sample slides
    - every title/body/ctrTitle/subTitle placeholder must be filled, else error
      (forces the author to pick a layout whose placeholders they actually use)
"""
from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pptx_io  # noqa: E402
from pptx_io import load_spec  # noqa: E402

# Placeholder types that carry visible text and therefore MUST be filled.
TEXT_PLACEHOLDER_TYPES = {"title", "body", "ctitle", "subtitle"}

# python-pptx exposes placeholder type as an enum member name (CENTER_TITLE,
# SUBTITLE, OBJECT, ...). OOXML and deck-specs use camelCase / lowercase names
# (ctrTitle, subTitle, body). Normalise both sides to a single canonical key so
# the spec author can write whichever spelling they prefer.
_CANONICAL = {
    # python-pptx enum name -> canonical
    "TITLE": "title",
    "CENTER_TITLE": "ctitle",
    "SUBTITLE": "subtitle",
    "OBJECT": "body",      # a "body"/content placeholder is typed OBJECT in pptx
    "BODY": "body",
    # OOXML ph @type spellings -> canonical
    "CTR_TITLE": "ctitle",
    "CTR": "ctitle",
}
_OOXML_TO_ENUM = {
    "title": "TITLE",
    "ctrTitle": "CENTER_TITLE",
    "subTitle": "SUBTITLE",
    "body": "OBJECT",
    "obj": "OBJECT",
    "objBody": "OBJECT",
}


# load_spec is imported from pptx_io near the top of this module. It lives in
# the dependency-free shared module so validate_deck.py can use it without
# pulling in python-pptx (which this module imports at module level).


# --- layout locking ---------------------------------------------------------

def index_layouts(prs: Presentation) -> dict[str, object]:
    """Map layout name -> slide_layout object. Errors on duplicate names."""
    by_name: dict[str, object] = {}
    for layout in prs.slide_layouts:
        name = layout.name.strip()  # tolerate padded layout names in templates
        if name in by_name:
            # Not fatal, but surface it: ambiguous layout names break locking.
            print(f"WARNING: layout name {name!r} is not unique; "
                  f"later definitions shadow earlier ones.", file=sys.stderr)
        by_name[name] = layout
    return by_name


def verify_layouts_exist(spec: dict, available: dict[str, object], zf_path: Path) -> None:
    """Cross-check spec layouts against the template via OOXML (source of truth).

    python-pptx exposes slide_layouts, but we double-check against the raw part
    names from pptx_io so the lock matches what validate_deck.py will see.
    """
    with zipfile.ZipFile(zf_path) as zf:
        tmpl_layouts = {ly["name"] for ly in pptx_io.inventory(zf)["layouts"]}
    spec_names = {s.get("layout") for s in spec.get("slides", []) if s.get("layout")}
    missing = spec_names - tmpl_layouts
    if missing:
        raise SystemExit(
            "ABORT: deck-spec references layouts not present in template:\n  "
            + "\n  ".join(sorted(missing))
            + "\nAvailable layouts:\n  "
            + "\n  ".join(sorted(tmpl_layouts))
        )


# --- placeholder matching ---------------------------------------------------

def _canonical_key(ph) -> str:
    """Canonical placeholder key: 'title' / 'ctitle' / 'subtitle' / 'body' / ...

    Normalises across python-pptx enum names (CENTER_TITLE) and OOXML ph @type
    values (ctrTitle) so the spec can use either spelling.
    """
    pf = ph.placeholder_format
    t = pf.type
    enum_name = t.name if t is not None else ""
    canon = _CANONICAL.get(enum_name, enum_name.lower())
    return canon


def _normalise_spec_key(key: str) -> str:
    """Normalise a key as written in the deck-spec to canonical form."""
    k = str(key).strip()
    if k in _CANONICAL:
        return _CANONICAL[k]
    if k in _OOXML_TO_ENUM:
        return _CANONICAL[_OOXML_TO_ENUM[k]]
    low = k.lower()
    if low in _CANONICAL:
        return _CANONICAL[low]
    # snake_case enum name like CENTER_TITLE
    if low.replace("_", "") in {c.replace("_", "") for c in _CANONICAL}:
        for orig, canon in _CANONICAL.items():
            if orig.lower().replace("_", "") == low.replace("_", ""):
                return canon
    return k  # leave as-is (e.g. an explicit idx key handled separately)


def fill_slide(slide, fields: dict, slide_index: int) -> dict:
    """Fill a slide's placeholders from `fields`. Returns a per-slide report.

    Field keys are normalised, so 'title'/'Title'/'TITLE'/'ctrTitle' all match
    the right placeholder. A field may also target a placeholder by index using
    'idx<N>'. Any required text placeholder (title/ctitle/subtitle/body) left
    empty is a hard error — pick a simpler layout instead.
    """
    report = {"filled": {}, "unmatched_field_keys": [], "empty_required": [], "warnings": []}
    placeholders = list(slide.placeholders)

    # Build lookup by canonical key and by idx. A canon key may map to MORE than
    # one placeholder (e.g. a "Two Content" layout has two body placeholders);
    # remember all of them so we can warn when a spec uses the bare canon key
    # and would silently fill every match with the same text.
    canon_to_phs: dict[str, list[object]] = {}
    by_idx: dict[int, object] = {}
    for ph in placeholders:
        canon_to_phs.setdefault(_canonical_key(ph), []).append(ph)
        by_idx[ph.placeholder_format.idx] = ph

    # Normalise the spec's field keys once.
    norm_fields: dict[tuple[str, str | None], str] = {}
    for k, v in (fields or {}).items():
        ks = str(k)
        if ks.lower().startswith("idx") and ks[3:].isdigit():
            norm_fields[("idx", int(ks[3:]))] = str(v)
        else:
            norm_fields[("canon", _normalise_spec_key(ks))] = str(v)

    for ph in placeholders:
        canon = _canonical_key(ph)
        idx = ph.placeholder_format.idx
        is_text_required = canon in TEXT_PLACEHOLDER_TYPES
        match = None
        if ("canon", canon) in norm_fields:
            match = ("canon", canon)
        elif ("idx", idx) in norm_fields:
            match = ("idx", idx)
        if match is not None:
            value = norm_fields[match]
            if hasattr(ph, "text"):
                ph.text = value
            report["filled"][f"{canon}#{idx}"] = value[:80]
            # Warn if a bare canon key matched several placeholders — they all
            # got the same text, which is almost never intended. Use idx<N>.
            if match[0] == "canon" and len(canon_to_phs[canon]) > 1:
                idxs = ",".join(str(p.placeholder_format.idx) for p in canon_to_phs[canon])
                w = (f"field {canon!r} filled {len(canon_to_phs[canon])} placeholders "
                     f"(idx {idxs}) with the same text. Target each with idx1/idx2 "
                     f"if they should differ.")
                if w not in report["warnings"]:
                    report["warnings"].append(w)
        elif is_text_required:
            report["empty_required"].append(f"{canon}#{idx}")

    # Track which field keys actually matched a placeholder, so the report can
    # flag typos (e.g. a field key that targets a placeholder this layout lacks).
    consumed_signatures: set[tuple[str, object]] = set()
    for ph in placeholders:
        canon = _canonical_key(ph)
        idx = ph.placeholder_format.idx
        if ("canon", canon) in norm_fields:
            consumed_signatures.add(("canon", canon))
        if ("idx", idx) in norm_fields:
            consumed_signatures.add(("idx", idx))
    for (sig, keyval) in norm_fields:
        if (sig, keyval) not in consumed_signatures:
            report["unmatched_field_keys"].append(
                keyval if sig == "canon" else f"idx{keyval}"
            )

    if report["empty_required"]:
        raise ValueError(
            f"slide {slide_index}: required text placeholders left empty: "
            f"{report['empty_required']}. Either fill them in the spec or choose "
            f"a simpler layout that has no such placeholder."
        )
    return report


def _drop_slide(prs: Presentation, sld_id_el) -> None:
    """Fully remove a slide from the package: sldId entry, rel, and the part.

    python-pptx's `Slides` collection has no public delete; only removing the
    <p:sldId> element leaves the slide part and its relationships orphaned in
    the package (inflate, leak sample content). This drops all three.
    `sld_id_el` is a <p:sldId> element from prs.slides._sldIdLst.
    """
    R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    rId = sld_id_el.attrib.get(R_NS)
    sld_id_el.getparent().remove(sld_id_el)
    if rId is not None:
        slide_part = prs.part.related_part(rId)
        prs.part.drop_rel(rId)
        # Also detach the part from the package so it isn't serialised.
        try:
            prs.part.package._parts.pop(slide_part.partname, None)  # noqa: SLF001
        except AttributeError:
            # Older/newer python-pptx may hold parts differently; the rel drop
            # above already makes the part unreachable, which is the key fix.
            pass


# --- main -------------------------------------------------------------------

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("spec", type=Path, help="deck-spec.yaml or .json")
    p.add_argument("--template", type=Path, required=True, help="source .pptx template")
    p.add_argument("--out", type=Path, default=None, help="output path (default: ./out/<spec stem>.pptx)")
    args = p.parse_args()

    spec = load_spec(args.spec)
    if "slides" not in spec or not isinstance(spec["slides"], list):
        raise SystemExit("ABORT: deck-spec must contain a top-level 'slides' list")

    # Lock: verify every named layout exists in the template before touching it.
    verify_layouts_exist(spec, {}, args.template)

    prs = Presentation(str(args.template))
    layouts = index_layouts(prs)

    # The stock Presentation(template) still contains the template's own sample
    # slides. Per SKILL.md we must NOT ship those, so we drop them all and
    # rebuild purely from layouts. We add our generated slides first, then
    # remove the original `original_count` slides (which sit at the head of
    # sldIdLst). _drop_slide removes the sldId entry, the presentation->slide
    # relationship, AND the slide part itself, so no orphan parts leak into the
    # output package.
    original_count = len(prs.slides._sldIdLst)  # noqa: SLF001

    generated_reports = []
    for n, slide_spec in enumerate(spec["slides"], start=1):
        name = slide_spec.get("layout")
        if not name:
            raise SystemExit(f"ABORT: slide {n} has no 'layout' field")
        if name not in layouts:
            raise SystemExit(f"ABORT: slide {n}: layout {name!r} not found in template")
        slide = prs.slides.add_slide(layouts[name])
        try:
            r = fill_slide(slide, slide_spec.get("fields", {}) or {}, n)
        except ValueError as e:
            raise SystemExit(f"ABORT: {e}")
        r["slide"] = n
        r["layout"] = name
        r["purpose"] = slide_spec.get("purpose", "")
        generated_reports.append(r)

    # Remove the ORIGINAL template slides (samples), keeping the generated ones
    # appended at the tail of sldIdLst.
    sld_id_lst = prs.slides._sldIdLst  # noqa: SLF001
    for _ in range(original_count):
        _drop_slide(prs, sld_id_lst[0])

    # Output path precedence: --out flag > spec 'output' field > default.
    spec_output = spec.get("output") if isinstance(spec, dict) else None
    out_path = args.out or (Path(spec_output) if spec_output else Path("out") / f"{args.spec.stem}.pptx")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))

    report = {
        "template": str(args.template),
        "output": str(out_path),
        "slide_count": len(generated_reports),
        "slides": generated_reports,
    }
    report_path = out_path.with_suffix(".generation-report.json")
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

    # Surface per-slide warnings to stderr so they aren't buried in the JSON.
    for s in generated_reports:
        for w in s.get("warnings", []):
            print(f"WARNING (slide {s['slide']}): {w}", file=sys.stderr)

    print(f"generated {len(generated_reports)} slides -> {out_path}")
    print(f"report -> {report_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
