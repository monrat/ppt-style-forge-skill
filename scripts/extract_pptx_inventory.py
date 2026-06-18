#!/usr/bin/env python3
"""Extract a compact PowerPoint template inventory using only stdlib.

This is intentionally conservative: it reads the pptx zip/XML and reports slide
size, theme font/colour tokens, slide layout names, and placeholders. It does not
modify the deck.
"""
from __future__ import annotations

import argparse
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
EMU_PER_INCH = 914400


def inches(value: str | int | None) -> float | None:
    if value is None:
        return None
    return round(int(value) / EMU_PER_INCH, 3)


def text(el: ET.Element) -> str:
    return "".join((t.text or "") for t in el.findall(".//a:t", NS)).strip()


def rel_targets(zf: zipfile.ZipFile, part: str) -> dict[str, str]:
    rel_part = posixpath.dirname(part) + "/_rels/" + posixpath.basename(part) + ".rels"
    if rel_part not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rel_part))
    return {rel.attrib.get("Id", ""): rel.attrib.get("Target", "") for rel in root}


def resolve(part: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def placeholder_info(shape: ET.Element) -> tuple[str, str, tuple[float | None, ...] | None, str]:
    ph = shape.find(".//p:ph", NS)
    ph_type = ph.attrib.get("type", "body") if ph is not None else "shape"
    ph_idx = ph.attrib.get("idx", "") if ph is not None else ""
    xfrm = shape.find(".//a:xfrm", NS)
    box = None
    if xfrm is not None:
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is not None and ext is not None:
            box = (
                inches(off.attrib.get("x")),
                inches(off.attrib.get("y")),
                inches(ext.attrib.get("cx")),
                inches(ext.attrib.get("cy")),
            )
    return ph_type, ph_idx, box, text(shape)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", type=Path)
    args = parser.parse_args()

    with zipfile.ZipFile(args.pptx) as zf:
        names = set(zf.namelist())
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        size = pres.find("p:sldSz", NS)
        if size is not None:
            print(f"slide_size: {inches(size.attrib.get('cx'))} x {inches(size.attrib.get('cy'))} in")
        print(f"active_slides: {len(pres.findall('.//p:sldId', NS))}")

        theme_parts = sorted(n for n in names if n.startswith("ppt/theme/theme") and n.endswith(".xml"))
        for theme in theme_parts[:3]:
            root = ET.fromstring(zf.read(theme))
            major = root.find(".//a:majorFont/a:latin", NS)
            minor = root.find(".//a:minorFont/a:latin", NS)
            print(f"\nTHEME {theme}")
            print("  major_font:", major.attrib.get("typeface") if major is not None else "")
            print("  minor_font:", minor.attrib.get("typeface") if minor is not None else "")
            for clr in root.findall(".//a:clrScheme/*", NS):
                srgb = clr.find("a:srgbClr", NS)
                if srgb is not None:
                    print(f"  {clr.tag.rsplit('}', 1)[-1]}: #{srgb.attrib.get('val')}")

        layout_parts = sorted(
            (n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)),
            key=lambda n: int(re.search(r"slideLayout(\d+)", n).group(1)),
        )
        print(f"\nLAYOUTS: {len(layout_parts)}")
        for part in layout_parts:
            root = ET.fromstring(zf.read(part))
            csld = root.find("p:cSld", NS)
            layout_name = csld.attrib.get("name", part) if csld is not None else part
            print(f"\n- {layout_name}")
            for shape in root.findall(".//p:sp", NS):
                ph_type, ph_idx, box, sample = placeholder_info(shape)
                if ph_type == "shape" and not sample:
                    continue
                print(f"    {ph_type}#{ph_idx} box={box} text={sample[:80]!r}")

        slide_parts = sorted(
            (n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=lambda n: int(re.search(r"slide(\d+)", n).group(1)),
        )
        print(f"\nSLIDE_PARTS: {len(slide_parts)}")
        for part in slide_parts[:20]:
            root = ET.fromstring(zf.read(part))
            rels = rel_targets(zf, part)
            layout = ""
            for target in rels.values():
                if "slideLayouts/" in target:
                    layout_part = resolve(part, target)
                    if layout_part in names:
                        layout_root = ET.fromstring(zf.read(layout_part))
                        csld = layout_root.find("p:cSld", NS)
                        layout = csld.attrib.get("name", layout_part) if csld is not None else layout_part
            first_text = " ".join(text(sp) for sp in root.findall(".//p:sp", NS))[:120]
            print(f"  {part}: layout={layout!r} text={first_text!r}")


if __name__ == "__main__":
    main()
