#!/usr/bin/env python3
"""Render a .pptx to per-slide images for visual QA (SKILL.md step 7).

A first render is almost never final; this script exists so the agent (or a
human) can eyeball the deck mechanically before declaring success. It is
strictly optional: SKILL.md says if no renderer is available, state the
limitation and continue. So this script never fails the pipeline — it either
renders, or prints a clear "renderer unavailable" notice and exits 0.

Renderer priority:
  1. LibreOffice / OpenOffice `soffice`  (cross-platform, headless)
  2. Microsoft PowerPoint via COM          (Windows only; requires pywin32)

    python3 scripts/render_qa.py out/deck.pptx
    python3 scripts/render_qa.py out/deck.pptx --out out/deck_pages --format png

Output: one image per slide under <out_dir>/ (slide1.png, slide2.png, ...).
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


# --- renderer discovery -----------------------------------------------------

def find_soffice() -> str | None:
    """Locate a LibreOffice/OpenOffice executable on PATH or common install dirs."""
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def find_powerpoint() -> str | None:
    """Locate POWERPNT.EXE on Windows (COM path)."""
    if os.name != "nt":
        return None
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def find_pdftoppm() -> str | None:
    """Locate poppler's `pdftoppm` (used for PDF -> per-page PNG)."""
    return shutil.which("pdftoppm")


# --- renderers --------------------------------------------------------------

def _soffice_to_pdf(exe: str, deck: Path, out_dir: Path) -> Path | None:
    """Convert a .pptx to PDF with LibreOffice headless. Returns the PDF path."""
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [exe, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(deck)]
    print("  running:", " ".join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print("  soffice stderr:", r.stderr.strip(), file=sys.stderr)
        return None
    pdf = out_dir / (deck.stem + ".pdf")
    if not pdf.exists():
        print(f"  soffice did not produce {pdf}", file=sys.stderr)
        return None
    return pdf


def _pdf_to_png(pdftoppm: str, pdf: Path, out_dir: Path, dpi: int = 110) -> int:
    """Split a PDF into one PNG per page with pdftoppm. Returns page count.

    Output files are slide-<n>.png in the same dir. 110 DPI is roughly the
    on-screen size of a 16:9 slide at 1280px wide.
    """
    # pdftoppm names pages "<prefix>-<n>.png"; we use prefix "slide" and a
    # leading-zero-free counter, then normalise names.
    prefix = str(out_dir / "slide")
    cmd = [pdftoppm, "-png", "-r", str(dpi), str(pdf), prefix]
    print("  running:", " ".join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print("  pdftoppm stderr:", r.stderr.strip(), file=sys.stderr)
        return 0
    # pdftoppm emits slide-1.png / slide-2.png ... (or slide-01.png with -l);
    # normalise to slide<N>.png.
    pngs = sorted(out_dir.glob("slide*.png"), key=_slide_number)
    count = 0
    for i, f in enumerate(pngs, start=1):
        target = out_dir / f"slide{i}.png"
        if f.resolve() != target.resolve():
            f.rename(target)
        count += 1
    return count


def render_with_soffice(exe: str, deck: Path, out_dir: Path, fmt: str) -> bool:
    """Render via LibreOffice. PNG goes deck->PDF->pdftoppm; PDF is direct.

    Returns True if the requested format was produced. If PNG was requested but
    pdftoppm is unavailable, falls back to producing a PDF and prints a clear
    note so the caller still gets a viewable artifact.
    """
    pdf = _soffice_to_pdf(exe, deck, out_dir)
    if pdf is None:
        return False
    if fmt == "pdf":
        print(f"  -> PDF: {pdf}")
        return True
    # fmt == "png": convert the PDF to per-page PNG via poppler's pdftoppm.
    pdftoppm = find_pdftoppm()
    if pdftoppm is None:
        print("  NOTE: PNG requested but `pdftoppm` (poppler) is not installed; "
              "produced a PDF instead. Install poppler for per-slide PNG, or use "
              "--format pdf. The full deck is viewable at: " + str(pdf))
        return True  # we produced *a* viewable render (the PDF)
    n = _pdf_to_png(pdftoppm, pdf, out_dir)
    if n == 0:
        print("  pdftoppm produced no pages; PDF is still available at " + str(pdf),
              file=sys.stderr)
        return False
    print(f"  -> {n} slide image(s) in {out_dir}")
    return True


def render_with_powerpoint(deck: Path, out_dir: Path, fmt: str) -> bool:
    """Use PowerPoint COM to export every slide. Windows + pywin32 required."""
    try:
        # pywin32 is optional; import lazily so the script is usable without it.
        import win32com.client  # type: ignore
    except ModuleNotFoundError:
        print("  NOTE: PowerPoint is installed but pywin32 is not. "
              "Install with:  python -m pip install pywin32")
        return False
    fmt_code = 18 if fmt == "png" else 1  # 18 = ppSaveAsPNG (per-slide); 1 = pptx
    if fmt != "png":
        print(f"  NOTE: PowerPoint COM supports per-slide export only as PNG "
              f"(requested {fmt!r}); rendering as PNG.")
        fmt = "png"
        fmt_code = 18
    out_dir.mkdir(parents=True, exist_ok=True)
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    # Dispatch may start hidden; some builds need Visible.
    try:
        ppt.Visible = 1
    except Exception:
        pass
    try:
        pres = ppt.Presentations.Open(str(deck.resolve()), WithWindow=False)
        # SaveAs PNG dumps one image per slide into the folder.
        pres.SaveAs(str(out_dir.resolve()), fmt_code)
        pres.Close()
    finally:
        ppt.Quit()
    # PowerPoint SaveAs-as-PNG drops one image per slide, but the filenames are
    # LOCALIZED ("Slide1.PNG" on English Office, "幻灯片1.PNG" on Chinese Office,
    # etc.). Don't match on the prefix — rename every PNG in the folder to a
    # stable slide<N>.png by sorted order.
    pngs = sorted(out_dir.glob("*.png"), key=_slide_number)
    renamed = 0
    for i, f in enumerate(pngs, start=1):
        target = out_dir / f"slide{i}.png"
        if f.resolve() != target.resolve():
            f.rename(target)
        renamed += 1
    print(f"  -> {renamed} slide image(s) in {out_dir}")
    return renamed > 0


def _slide_number(path: Path) -> int:
    """Extract the trailing integer from a localized slide filename for sorting."""
    import re
    m = re.search(r"(\d+)", path.stem)
    return int(m.group(1)) if m else 0


# --- main -------------------------------------------------------------------

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("deck", type=Path, help=".pptx to render")
    p.add_argument("--out", type=Path, default=None,
                   help="output dir (default: <deck>_pages)")
    p.add_argument("--format", default="png", choices=["png", "pdf"],
                   help="image format (PowerPoint COM always does per-slide PNG)")
    args = p.parse_args()

    if not args.deck.exists():
        print(f"ABORT: deck not found: {args.deck}", file=sys.stderr)
        return 2

    out_dir = args.out or args.deck.parent / (args.deck.stem + "_pages")

    print(f"render QA: {args.deck}")

    # Try LibreOffice first (most portable), then PowerPoint COM.
    soffice = find_soffice()
    if soffice:
        print(f"  renderer: LibreOffice ({soffice})")
        if render_with_soffice(soffice, args.deck, out_dir, args.format):
            print(f"done -> {out_dir}")
            return 0
        print("  LibreOffice render failed; falling back.", file=sys.stderr)

    pp = find_powerpoint()
    if pp:
        print(f"  renderer: PowerPoint COM ({pp})")
        if render_with_powerpoint(args.deck, out_dir, args.format):
            print(f"done -> {out_dir}")
            return 0
        print("  PowerPoint render failed.", file=sys.stderr)

    # No renderer worked — per SKILL.md, state the limitation and exit 0.
    print(
        "\n[render QA skipped] No usable renderer was found.\n"
        "SKILL.md step 7 requires at least one fix-and-verify visual pass. "
        "Open the deck manually in PowerPoint or LibreOffice and check for:\n"
        "  - overlap / text overflow / wrapped-title collisions\n"
        "  - footer/logo collisions\n"
        "  - inconsistent gaps, low contrast, leftover sample content\n"
        "To enable automatic rendering, install LibreOffice "
        "(https://www.libreoffice.org) or, on Windows, `pip install pywin32`."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
