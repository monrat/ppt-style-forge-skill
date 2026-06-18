"""Shared, dependency-free helpers for reading a .pptx as OOXML.

Both `extract_pptx_inventory.py` (human-readable inventory) and
`validate_deck.py` (mechanical checks) build on these functions so the parsing
logic lives in exactly one place. Everything here is stdlib-only (zipfile +
xml.etree) and read-only: nothing in this module mutates a deck.

A .pptx is an OPC zip. The parts we care about:

    ppt/presentation.xml      slide size, active slide ids
    ppt/theme/themeN.xml      major/minor font, colour scheme
    ppt/slideLayouts/slideLayoutN.xml   named layouts + placeholders
    ppt/slides/slideN.xml     actual slides; their _rels point at a layout
    ppt/slideMasters/slideMasterN.xml   masters (carry footer/logo identity)

The functions return plain Python data (dicts/lists) so callers can either
pretty-print them (inventory) or assert on them (validator).
"""
from __future__ import annotations

import json
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable

# DrawingML / PresentationML namespaces. ElementTree needs these explicitly.
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

EMU_PER_INCH = 914400


# --- small primitives -------------------------------------------------------

def inches(value: str | int | None) -> float | None:
    """Convert an EMU string/int to inches, rounded to 3 dp. None stays None."""
    if value is None:
        return None
    return round(int(value) / EMU_PER_INCH, 3)


def local(tag: str) -> str:
    """Strip the XML namespace from an ElementTree tag: '{ns}name' -> 'name'."""
    return tag.rsplit("}", 1)[-1]


def text(el: ET.Element) -> str:
    """Concatenate all <a:t> run text under an element, trimmed."""
    return "".join((t.text or "") for t in el.findall(".//a:t", NS)).strip()


def rel_targets(zf: zipfile.ZipFile, part: str) -> dict[str, str]:
    """Return {rId: target} for a part's .rels, or {} if it has none."""
    rel_part = posixpath.dirname(part) + "/_rels/" + posixpath.basename(part) + ".rels"
    if rel_part not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rel_part))
    return {rel.attrib.get("Id", ""): rel.attrib.get("Target", "") for rel in root}


def resolve(part: str, target: str) -> str:
    """Resolve a relationship target (absolute '/x' or relative) to a zip part."""
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


# --- shape / placeholder parsing -------------------------------------------

def placeholder_info(shape: ET.Element) -> dict:
    """Describe one <p:sp> shape as a dict.

    Keys: ph_type, ph_idx, box (x,y,cx,cy in inches or None), text.
    A shape with no <p:ph> is reported as ph_type='shape' (a free drawing).
    """
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
    return {"ph_type": ph_type, "ph_idx": ph_idx, "box": box, "text": text(shape)}


def shapes_of(root: ET.Element) -> list[ET.Element]:
    """All <p:sp> shapes under an element (layout, master, or slide)."""
    return root.findall(".//p:sp", NS)


# --- part discovery ---------------------------------------------------------

def slide_size(zf: zipfile.ZipFile) -> tuple[float | None, float | None]:
    """Return (width_in, height_in) from presentation.xml."""
    pres = ET.fromstring(zf.read("ppt/presentation.xml"))
    size = pres.find("p:sldSz", NS)
    if size is None:
        return (None, None)
    return (inches(size.attrib.get("cx")), inches(size.attrib.get("cy")))


def active_slide_count(zf: zipfile.ZipFile) -> int:
    """Count of <p:sldId> entries = active slides in the deck."""
    pres = ET.fromstring(zf.read("ppt/presentation.xml"))
    return len(pres.findall(".//p:sldId", NS))


# Relationship namespace (used to resolve rId -> target in .rels)
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def active_slide_parts(zf: zipfile.ZipFile) -> list[str]:
    """Slide parts in PRESENTATION ORDER — only those the deck actually shows.

    presentation.xml's <p:sldIdLst> lists active slides via rId references; the
    slide parts on disk may also include orphaned parts from deleted slides.
    Always use this (not slide_parts()) when you want to inspect what the viewer
    will actually render.
    """
    pres = ET.fromstring(zf.read("ppt/presentation.xml"))
    rels = rel_targets(zf, "ppt/presentation.xml")
    parts: list[str] = []
    for sldId in pres.findall(".//p:sldId", NS):
        rid = sldId.attrib.get(R_NS + "id", "")
        target = rels.get(rid, "")
        if not target:
            continue
        # presentation.xml lives at ppt/, so its rels resolve to ppt/slides/...
        parts.append(resolve("ppt/presentation.xml", target))
    return parts


def _ordered(parts: Iterable[str], pattern: str) -> list[str]:
    """Sort 'slideN.xml' / 'slideLayoutN.xml' parts by their N.

    `pattern` must contain one capturing group around the index digits.
    """
    rx = re.compile(pattern)

    def key(n: str) -> int:
        m = rx.search(n)
        return int(m.group(1)) if m else 0

    return sorted(parts, key=key)


def layout_parts(zf: zipfile.ZipFile) -> list[str]:
    """All slideLayout parts, ordered by index."""
    parts = (n for n in zf.namelist() if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n))
    return _ordered(parts, r"slideLayout(\d+)")


def slide_parts(zf: zipfile.ZipFile) -> list[str]:
    """All slide parts, ordered by index."""
    parts = (n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n))
    return _ordered(parts, r"slide(\d+)")


def master_parts(zf: zipfile.ZipFile) -> list[str]:
    """All slideMaster parts, ordered by index."""
    parts = (n for n in zf.namelist() if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))
    return _ordered(parts, r"slideMaster(\d+)")


def theme_parts(zf: zipfile.ZipFile) -> list[str]:
    """All theme parts (ppt/theme/themeN.xml), ordered by index."""
    parts = (n for n in zf.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n))
    return _ordered(parts, r"theme(\d+)")


# --- structured extraction --------------------------------------------------

def extract_theme(zf: zipfile.ZipFile, theme_part: str) -> dict:
    """Parse one theme part into {major_font, minor_font, colors: {name: #hex}}."""
    root = ET.fromstring(zf.read(theme_part))
    major = root.find(".//a:majorFont/a:latin", NS)
    minor = root.find(".//a:minorFont/a:latin", NS)
    colors: dict[str, str] = {}
    for clr in root.findall(".//a:clrScheme/*", NS):
        srgb = clr.find("a:srgbClr", NS)
        if srgb is not None:
            colors[local(clr.tag)] = "#" + srgb.attrib.get("val", "").upper()
    return {
        "part": theme_part,
        "major_font": major.attrib.get("typeface", "") if major is not None else "",
        "minor_font": minor.attrib.get("typeface", "") if minor is not None else "",
        "colors": colors,
    }


def layout_summary(zf: zipfile.ZipFile, layout_part: str) -> dict:
    """Parse one layout part into {name, part, placeholders: [placeholder_info...]}.

    Includes only placeholders that carry type/idx or have sample text, so the
    inventory stays readable; the validator uses the raw shapes via shapes_of().
    """
    root = ET.fromstring(zf.read(layout_part))
    csld = root.find("p:cSld", NS)
    name = csld.attrib.get("name", layout_part) if csld is not None else layout_part
    name = name.strip()  # template authors sometimes pad layout names with whitespace
    phs = []
    for shape in shapes_of(root):
        info = placeholder_info(shape)
        if info["ph_type"] == "shape" and not info["text"]:
            continue
        phs.append(info)
    return {"name": name, "part": layout_part, "placeholders": phs}


def slide_layout_name(zf: zipfile.ZipFile, slide_part: str, names: set[str]) -> str:
    """Resolve the layout NAME a slide is bound to, via its .rels.

    `names` is the set of known layout part names (for membership checks).
    Returns '' if the slide has no layout relationship we can resolve.
    """
    for target in rel_targets(zf, slide_part).values():
        if "slideLayouts/" in target:
            layout_part = resolve(slide_part, target)
            if layout_part in names:
                root = ET.fromstring(zf.read(layout_part))
                csld = root.find("p:cSld", NS)
                if csld is not None:
                    return csld.attrib.get("name", layout_part).strip()
    return ""


def inventory(zf: zipfile.ZipFile) -> dict:
    """Full structured inventory of a deck.

    Returns:
        slide_size: (w, h) inches
        active_slides: int
        themes: [extract_theme, ...]   (up to first 3)
        layouts: [layout_summary, ...]
        slides: [{part, layout, text}] (first 20, for a quick peek)
        known_layout_parts: set consumed internally; also surfaces count
    """
    names = set(zf.namelist())
    known_layout_parts = set(layout_parts(zf))

    slides = []
    for part in slide_parts(zf)[:20]:
        root = ET.fromstring(zf.read(part))
        layout = slide_layout_name(zf, part, known_layout_parts)
        first_text = " ".join(text(sp) for sp in shapes_of(root))[:120]
        slides.append({"part": part, "layout": layout, "text": first_text})

    return {
        "slide_size": slide_size(zf),
        "active_slides": active_slide_count(zf),
        "themes": [extract_theme(zf, t) for t in theme_parts(zf)[:3]],
        "layouts": [layout_summary(zf, p) for p in layout_parts(zf)],
        "slides": slides,
        "known_layout_parts": known_layout_parts,
    }


# --- deck-spec loading (stdlib only; no pyyaml dependency) ------------------
#
# `load_spec` reads a deck-spec (.json or .yaml). It lives here — in the
# dependency-free shared module — so that both `generate_deck.py` (which uses
# python-pptx) and `validate_deck.py` (stdlib-only) can import it WITHOUT
# validate_deck.py acquiring an implicit python-pptx dependency through
# generate_deck's module-level imports.

def load_spec(path: Path) -> dict:
    """Load a deck-spec from .json or .yaml.

    JSON uses stdlib json. For YAML we accept the subset this skill actually
    emits (top-level `key: value`, nested via indentation, lists via `- `);
    this avoids a hard pyyaml dependency. If real YAML is needed, install pyyaml
    and this function will detect and use it.
    """
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() == ".json":
        return json.loads(text)
    try:
        import yaml  # type: ignore
        return yaml.safe_load(text)
    except ModuleNotFoundError:
        return _mini_yaml(text)


def _mini_yaml(text: str) -> dict:
    """Tiny YAML subset parser for deck-specs. Not a general YAML parser.

    Supports: `key: value`, `key:` then indented children, `- item` lists,
    inline lists `[a, b]`, quoted strings. Sufficient for deck-spec.example.yaml.
    """
    lines = [ln.rstrip() for ln in text.splitlines()]
    root, _ = _parse_block(lines, 0, 0)
    return root


def _parse_block(lines: list[str], start: int, indent: int) -> tuple[dict, int]:
    obj: dict = {}
    i = start
    while i < len(lines):
        ln = lines[i]
        if not ln.strip() or ln.lstrip().startswith("#"):
            i += 1
            continue
        cur = len(ln) - len(ln.lstrip())
        if cur < indent:
            break
        if cur > indent:
            i += 1
            continue
        stripped = ln.strip()
        if stripped.startswith("- "):
            # list at this indent — collect siblings
            items, i = _parse_list(lines, i, indent)
            return items, i  # type: ignore[return-value]
        m = re.match(r'([\w-]+):\s*(.*)$', stripped)
        if not m:
            i += 1
            continue
        key, val = m.group(1), m.group(2).strip()
        if val in ("|", "|-", "|+", ">", ">-", ">+"):
            # block scalar: gather following indented lines as one string.
            block, i = _read_block_scalar(lines, i + 1, indent + 2, val)
            obj[key] = block
        elif val == "":
            child, i = _parse_block(lines, i + 1, indent + 2)
            obj[key] = child
        else:
            obj[key] = _scalar(val)
            i += 1
    return obj, i


def _read_block_scalar(lines: list[str], start: int, indent: int, style: str) -> tuple[str, int]:
    """Read a `|` (literal) or `>` (folded) block scalar into a string.

    Literal `|` preserves newlines; folded `>` joins lines with spaces. Trailing
    blank lines are folded to a single trailing newline for literal style.
    """
    raw: list[str] = []
    i = start
    while i < len(lines):
        ln = lines[i]
        if ln.strip() == "":
            raw.append("")
            i += 1
            continue
        cur = len(ln) - len(ln.lstrip())
        if cur < indent:
            break
        raw.append(ln[indent:] if len(ln) >= indent else ln.lstrip())
        i += 1
    # trim trailing blanks that belong to the surrounding indentation
    while raw and raw[-1] == "":
        raw.pop()
    if style.startswith(">"):
        text = " ".join(part for part in raw)
    else:
        text = "\n".join(raw)
    return text, i


def _parse_list(lines: list[str], start: int, indent: int) -> tuple[list, int]:
    items: list = []
    i = start
    while i < len(lines):
        ln = lines[i]
        if not ln.strip() or ln.lstrip().startswith("#"):
            i += 1
            continue
        cur = len(ln) - len(ln.lstrip())
        if cur < indent:
            break
        stripped = ln.strip()
        if not stripped.startswith("- "):
            break
        content = stripped[2:].strip()
        if re.match(r'[\w-]+:\s*', content):
            # inline first key of a mapping item
            child = {}
            m = re.match(r'([\w-]+):\s*(.*)$', content)
            k, v = m.group(1), m.group(2).strip()
            if v in ("|", "|-", "|+", ">", ">-", ">+"):
                block, i = _read_block_scalar(lines, i + 1, indent + 4, v)
                child[k] = block
            elif v == "":
                sub, i = _parse_block(lines, i + 1, indent + 4)
                child[k] = sub
            else:
                child[k] = _scalar(v)
                # consume further mapping keys at item-indent + 2
                j = i + 1
                item_indent = indent + 2
                while j < len(lines):
                    nxt = lines[j]
                    if not nxt.strip() or nxt.lstrip().startswith("#"):
                        j += 1
                        continue
                    ni = len(nxt) - len(nxt.lstrip())
                    if ni < item_indent:
                        break
                    if ni > item_indent:
                        # deeper block under previous key
                        break
                    ns = nxt.strip()
                    mm = re.match(r'([\w-]+):\s*(.*)$', ns)
                    if not mm:
                        break
                    kk, vv = mm.group(1), mm.group(2).strip()
                    if vv in ("|", "|-", "|+", ">", ">-", ">+"):
                        block, j = _read_block_scalar(lines, j + 1, item_indent + 2, vv)
                        child[kk] = block
                    elif vv == "":
                        sub, j = _parse_block(lines, j + 1, item_indent + 2)
                        child[kk] = sub
                    else:
                        child[kk] = _scalar(vv)
                        j += 1
                i = j
            items.append(child)
        else:
            items.append(_scalar(content))
            i += 1
    return items, i


def _scalar(val: str):
    if val.startswith("[") and val.endswith("]"):
        inner = val[1:-1].strip()
        if not inner:
            return []
        return [_scalar(x.strip()) for x in re.split(r',(?![^"]*")', inner)]
    if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
        return val[1:-1]
    if val.lower() in ("true", "false"):
        return val.lower() == "true"
    try:
        return int(val)
    except ValueError:
        pass
    try:
        return float(val)
    except ValueError:
        pass
    return val
