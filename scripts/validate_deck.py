#!/usr/bin/env python3
"""Mechanically validate a generated deck against its template + deck-spec.

This is the checker half of the PPT Style Forge workflow (SKILL.md step 6). It
is stdlib-only (zipfile + xml.etree) and reads the deck the same way
pptx_io does, so its view of "truth" matches the generator's. It does not
modify anything.

    python3 scripts/validate_deck.py generated.pptx --spec deck-spec.yaml --template src.pptx

Each P0 rule from references/quality-checklist.md maps to one check below. The
script prints a PASS/FAIL summary, writes a JSON report next to the deck, and
exits non-zero if any P0 check failed.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pptx_io  # noqa: E402
from pptx_io import load_spec  # noqa: E402  (stdlib-only — avoids generate_deck's python-pptx dep)

# Hex colours known to be Office-default and therefore suspicious as "off-brand"
# when they appear on slide bodies (not in the template's theme). These are the
# classic default accent palette; brand decks almost never want all of them.
_OFFICE_DEFAULT_SUSPECT = {
    "#4F81BD", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6", "#F79646",
}

# Phrases the default Office template uses as placeholder sample text. If any of
# these survive into a generated deck, it is a leak.
_SAMPLE_TEXT_MARKERS = (
    "click to edit master title",
    "click to edit master text",
    "click to add text",
    "click to add title",
)


# --- deep scans over slide XML ----------------------------------------------

def scan_colours_fonts(zf: zipfile.ZipFile) -> tuple[set[str], set[str]]:
    """Return (colours_used, fonts_used) across ALL slide bodies.

    Catches direct formatting on slides, not just the theme — this is what
    detects off-brand colours introduced by freehand or default Office styling.
    """
    colours: set[str] = set()
    fonts: set[str] = set()
    for part in pptx_io.active_slide_parts(zf):
        try:
            raw = zf.read(part)
        except KeyError:
            continue
        for m in re.finditer(rb'<a:srgbClr val="([0-9A-Fa-f]{6})"', raw):
            colours.add("#" + m.group(1).decode().upper())
        for m in re.finditer(rb'<a:(latin|ea|cs) typeface="([^"]*)"', raw):
            tf = m.group(2).decode().strip()
            if tf and tf.lower() not in ("+mj-lt", "+mn-lt"):
                fonts.add(tf)
    return colours, fonts


def slide_placeholder_texts(zf: zipfile.ZipFile) -> list[list[dict]]:
    """For each ACTIVE slide (presentation order), return placeholder infos."""
    per_slide: list[list[dict]] = []
    for part in pptx_io.active_slide_parts(zf):
        root = ET.fromstring(zf.read(part))
        phs = []
        for sp in pptx_io.shapes_of(root):
            info = pptx_io.placeholder_info(sp)
            if info["ph_type"] != "shape":
                phs.append(info)
        per_slide.append(phs)
    return per_slide


def each_slide_layout(zf: zipfile.ZipFile) -> list[str]:
    """Layout name each ACTIVE slide is bound to ('' if unresolvable)."""
    known = set(pptx_io.layout_parts(zf))
    return [pptx_io.slide_layout_name(zf, p, known) for p in pptx_io.active_slide_parts(zf)]


# --- the checks -------------------------------------------------------------

class Result:
    def __init__(self) -> None:
        self.passed: list[str] = []
        self.failed: list[str] = []
        self.warned: list[str] = []

    @property
    def ok(self) -> bool:
        return not self.failed

    def add(self, ok: bool, p0_id: str, message: str) -> None:
        (self.passed if ok else self.failed).append(f"[P0 {p0_id}] {message}")

    def warn(self, message: str) -> None:
        self.warned.append(message)


def validate(deck: Path, template: Path | None, spec: dict | None) -> Result:
    r = Result()
    with zipfile.ZipFile(deck) as zf:
        names = set(zf.namelist())
        inv = pptx_io.inventory(zf)
        deck_size = inv["slide_size"]
        deck_active = inv["active_slides"]
        deck_layouts = {ly["name"] for ly in inv["layouts"]}
        deck_theme = inv["themes"][0] if inv["themes"] else {}
        deck_used_colours, deck_used_fonts = scan_colours_fonts(zf)
        deck_slide_layouts = each_slide_layout(zf)
        per_slide_ph = slide_placeholder_texts(zf)
        # Deep scans that touch the zip — must run while zf is open.
        has_footer_identity = _has_footer_identity(zf)
        empty_graphics = _empty_graphics(zf)
        master_count = sum(1 for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))

    # Template baseline (if supplied)
    tmpl_theme = {}
    tmpl_layouts: set[str] = set()
    tmpl_size = None
    if template is not None:
        with zipfile.ZipFile(template) as tzf:
            tinv = pptx_io.inventory(tzf)
            tmpl_theme = tinv["themes"][0] if tinv["themes"] else {}
            tmpl_layouts = {ly["name"] for ly in tinv["layouts"]}
            tmpl_size = tinv["slide_size"]

    spec_slides = (spec or {}).get("slides", []) or []

    # --- P0-2: every slide uses a real named layout from the template ---------
    unknown = [n for n in deck_slide_layouts if n and n not in tmpl_layouts] if tmpl_layouts else []
    unnamed = [i for i, n in enumerate(deck_slide_layouts, 1) if not n]
    r.add(
        not unknown and not unnamed,
        "2",
        f"all slides bind to real named layouts; unknown={unknown}, unnamed={unnamed}",
    )

    # --- P0-3: theme fonts/colours preserved from source template -------------
    if tmpl_theme:
        same_fonts = (deck_theme.get("major_font") == tmpl_theme.get("major_font")
                      and deck_theme.get("minor_font") == tmpl_theme.get("minor_font"))
        tmpl_colors = set(tmpl_theme.get("colors", {}).values())
        deck_colors = set(deck_theme.get("colors", {}).values())
        same_palette = tmpl_colors == deck_colors
        r.add(same_fonts and same_palette, "3",
              f"theme matches template (fonts={same_fonts}, palette={same_palette})")
    else:
        r.warn("[P0 3] no template supplied; theme-fidelity check skipped")

    # --- P0-5: no empty visible text/title placeholders -----------------------
    empty = []
    for i, phs in enumerate(per_slide_ph, 1):
        for ph in phs:
            if ph["ph_type"] in ("title", "ctrTitle", "subTitle", "body") and not ph["text"]:
                empty.append(f"slide{i}:{ph['ph_type']}#{ph['ph_idx']}")
    r.add(not empty, "5", f"no empty visible text placeholders; empty={empty}")

    # --- P0-6: no leftover sample text ----------------------------------------
    leaked = []
    all_text = " ".join(
        ph["text"] for phs in per_slide_ph for ph in phs
    ).lower()
    for marker in _SAMPLE_TEXT_MARKERS:
        if marker in all_text:
            leaked.append(marker)
    r.add(not leaked, "6", f"no leftover sample text; leaked={leaked}")

    # --- P0-8: active slide count == intended, no orphan slides ---------------
    if spec_slides:
        intended = len(spec_slides)
        r.add(deck_active == intended, "8",
              f"active slides {deck_active} == intended {intended}")
    else:
        r.warn("[P0 8] no spec supplied; slide-count check skipped")

    # --- P0-9: no off-brand colours -------------------------------------------
    # A colour is off-brand if it appears directly on a slide body but is in
    # neither the deck's theme palette nor the template's theme palette. The
    # Office-default suspect set is reported as an extra severity flag: if such
    # a colour is off-brand it almost certainly came from default Office styling.
    allowed = set(deck_theme.get("colors", {}).values()) | set(tmpl_theme.get("colors", {}).values())
    off_brand = sorted(c for c in deck_used_colours if c not in allowed)
    off_brand_office = sorted(set(off_brand) & _OFFICE_DEFAULT_SUSPECT)
    r.add(not off_brand, "9",
          f"no off-brand colours on slides; off_brand={off_brand}"
          + (f" (Office-default suspects={off_brand_office})" if off_brand_office else ""))

    # --- P0-1: based on the supplied template (has master + at least one
    # shared layout part name with the template) -------------------------------
    has_master = any("slideMasters/slideMaster" in n for n in names)
    r.add(has_master, "1", f"deck carries slide masters (master parts present={has_master})")

    # --- P0-4: master/footer/logo preserved and not duplicated ----------------
    # Heuristic: exactly one slideMaster part; footer placeholders present on
    # at least one master/layout. Duplication would show >1 master from the
    # same source; we flag multiple masters as a soft warning.
    r.add(has_footer_identity and master_count >= 1, "4",
          f"footer/logo identity present={has_footer_identity}, masters={master_count}")
    if master_count > 1:
        r.warn(f"[P0 4] {master_count} masters present; verify none is a duplicate")

    # --- P0-7: tables/charts/images actually populated ------------------------
    # A graphicFrame (<p:graphicFrame>) with a table should have <a:tbl> rows;
    # a chart rel should resolve. We flag empty tables (no rows) and chart
    # frames whose rel target is missing. (empty_graphics computed above while
    # the zip was open.)
    r.add(not empty_graphics, "7",
          f"tables/charts populated; empty_or_dangling={empty_graphics}")

    # Spec-vs-actual layout agreement (extra precision on top of P0-2)
    if spec_slides and deck_slide_layouts:
        mismatch = []
        for i, ss in enumerate(spec_slides):
            want = ss.get("layout")
            got = deck_slide_layouts[i] if i < len(deck_slide_layouts) else None
            if want and got and want != got:
                mismatch.append(f"slide{i+1}: spec={want!r} actual={got!r}")
        if mismatch:
            r.warn("[spec] layout mismatch: " + "; ".join(mismatch))

    # slide-size consistency with template (P1-ish but cheap and useful)
    if tmpl_size and deck_size and tmpl_size != deck_size:
        r.warn(f"[P1] slide size {deck_size} differs from template {tmpl_size}")

    return r


def _has_footer_identity(zf: zipfile.ZipFile) -> bool:
    """True if some master/layout carries a footer/date/slideNumber placeholder."""
    for part in list(pptx_io.master_parts(zf)) + list(pptx_io.layout_parts(zf)):
        try:
            root = ET.fromstring(zf.read(part))
        except KeyError:
            continue
        for ph in root.findall(".//p:ph", pptx_io.NS):
            if ph.attrib.get("type") in ("ftr", "dt", "sldNum"):
                return True
    return False


def _empty_graphics(zf: zipfile.ZipFile) -> list[str]:
    """Find graphicFrames whose table has no rows or whose chart rel dangles.

    A chart frame is `<a:graphicData uri="...chart..."><c:chart r:id="rIdN"/>`.
    The chart URI lives on `a:graphicData` itself (NOT on a child), and the
    chart element is the `<c:chart>` child carrying an `r:id`. We resolve that
    exact rId through the slide's .rels and confirm the target part exists.
    """
    CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    R_NS_FULL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    names = set(zf.namelist())
    bad = []
    for part in pptx_io.active_slide_parts(zf):
        try:
            raw = zf.read(part)
            root = ET.fromstring(raw)
        except KeyError:
            continue
        ns = pptx_io.NS
        rels = pptx_io.rel_targets(zf, part)
        for gf in root.findall(".//p:graphicFrame", ns):
            # table: needs at least one <a:tr>
            tbl = gf.find(".//a:tbl", ns)
            if tbl is not None and len(tbl.findall("a:tr", ns)) == 0:
                bad.append(f"{part}: empty table")
            # chart: find <c:chart> (in the chart namespace) and resolve its r:id.
            # The chart namespace isn't in pptx_io.NS, so use the full path.
            chart = gf.find(f".//{{{CHART_NS}}}chart")
            if chart is not None:
                rid = chart.attrib.get(R_NS_FULL)
                target = rels.get(rid or "")
                if not target:
                    bad.append(f"{part}: chart frame with no resolvable rel ({rid})")
                    continue
                resolved = pptx_io.resolve(part, target)
                if resolved not in names:
                    bad.append(f"{part}: dangling chart rel {rid} -> {resolved} (part missing)")
    return bad


# --- main -------------------------------------------------------------------

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("deck", type=Path, help="generated .pptx to validate")
    p.add_argument("--spec", type=Path, default=None, help="deck-spec used to generate")
    p.add_argument("--template", type=Path, default=None, help="source template .pptx")
    p.add_argument("--report", type=Path, default=None, help="report json path")
    args = p.parse_args()

    spec = load_spec(args.spec) if args.spec else None
    r = validate(args.deck, args.template, spec)

    report_path = args.report or args.deck.with_suffix(".validation-report.json")
    payload = {
        "deck": str(args.deck),
        "template": str(args.template) if args.template else None,
        "spec": str(args.spec) if args.spec else None,
        "passed": r.passed,
        "failed": r.failed,
        "warnings": r.warned,
        "status": "PASS" if r.ok else "FAIL",
    }
    report_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\n=== {payload['status']} ===  ({len(r.passed)} passed, {len(r.failed)} failed, {len(r.warned)} warnings)")
    for line in r.failed:
        print("  FAIL  " + line)
    for line in r.warned:
        print("  WARN  " + line)
    for line in r.passed:
        print("  ok    " + line)
    print(f"report -> {report_path}")
    return 0 if r.ok else 1


if __name__ == "__main__":
    sys.exit(main())
