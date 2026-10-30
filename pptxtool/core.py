# SPDX-FileCopyrightText: PPTXTool | Copyright (c) 2026 Travis L. Seymour, PhD <nogard@ucsc.edu>
# SPDX-License-Identifier: GPL-3.0-or-later

"""Presentation tooling (pre-notes, post-notes, standardize, pdf, watermark, compress).

GUI-free port of pptxtool.py: every operation reports progress through a `log`
callable instead of printing, and raises PptxToolError instead of exiting.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Callable
from datetime import date
from io import BytesIO
import math
import os
from pathlib import Path
import random
import shutil
import string
import subprocess
import tempfile

from lxml import etree
from PIL import Image, ImageChops
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.shapes.group import GroupShape
import pymupdf
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

Log = Callable[[str], None]

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

WATERMARK_TEXT = f"Copyright {date.today().year} Travis L. Seymour, PhD"
PDF_SETTINGS = ("/screen", "/ebook", "/printer", "/prepress", "/default")
DEFAULT_PDF_SETTINGS = "/ebook"


class PptxToolError(Exception):
    """An operation could not be carried out; the message is meant for the user."""


def _require_file(path: Path, suffix: str, command: str) -> None:
    """Raise unless path exists and has the given suffix."""
    if path.suffix.lower() != suffix:
        raise PptxToolError(f"{command} requires a {suffix} file.")
    if not path.exists():
        raise PptxToolError(f"File not found: {path}")


def find_tool(*names: str) -> str | None:
    """Return the first of names found on PATH, or None."""
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


def _require_tool(*names: str) -> str:
    """Return the first of names found on PATH, or raise."""
    found = find_tool(*names)
    if found is None:
        raise PptxToolError(f"'{names[0]}' not found on PATH.")
    return found


def _run(cmd: list[str], log: Log) -> None:
    """Run cmd, forwarding its output to log; raise with its output on failure."""
    result = subprocess.run(cmd, check=False, capture_output=True, text=True)
    for line in (result.stdout + result.stderr).splitlines():
        if line.strip():
            log(f"  {line.strip()}")
    if result.returncode != 0:
        raise PptxToolError(f"{Path(cmd[0]).name} failed with exit code {result.returncode}.")


# ----------------------------------------------------------------------
# SHAPE / TEXT HELPERS (PPTX)
# ----------------------------------------------------------------------

TARGET_FONT = "Arial"
STANDARD_BULLET_CHAR = "\u2022"  # Unicode bullet •
BLACK = RGBColor(0, 0, 0)

# Runs in these fonts encode symbols (e.g. Symbol "a" = α), so their font is kept.
SYMBOL_FONTS = frozenset(
    {
        "symbol",
        "wingdings",
        "wingdings 2",
        "wingdings 3",
        "webdings",
        "marlett",
        "mt extra",
        "zapf dingbats",
    }
)

LINE_PRESETS = frozenset(
    {
        "line",
        "straightConnector1",
        "bentConnector2",
        "bentConnector3",
        "bentConnector4",
        "bentConnector5",
        "curvedConnector2",
        "curvedConnector3",
        "curvedConnector4",
        "curvedConnector5",
        "rightArrow",
        "leftArrow",
        "upArrow",
        "downArrow",
        "leftRightArrow",
        "upDownArrow",
        "quadArrow",
        "notchedRightArrow",
        "bentArrow",
        "uturnArrow",
        "circularArrow",
        "leftCircularArrow",
        "leftRightCircularArrow",
        "curvedRightArrow",
        "curvedLeftArrow",
        "curvedUpArrow",
        "curvedDownArrow",
        "stripedRightArrow",
        "bentUpArrow",
        "leftUpArrow",
        "leftRightUpArrow",
    }
)


def _iter_shapes(shapes):
    """Yield every shape, descending into group shapes."""
    for shape in shapes:
        yield shape
        if isinstance(shape, GroupShape):
            yield from _iter_shapes(shape.shapes)


def _iter_text_frames(shapes):
    """Yield every text frame, including those in groups and table cells."""
    for shape in _iter_shapes(shapes):
        if shape.has_text_frame:
            yield shape.text_frame
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    yield cell.text_frame


def _is_symbol_run(run) -> bool:
    """True if the run explicitly uses a symbol font such as Symbol or Wingdings."""
    rPr = run._r.rPr
    if rPr is None:
        return False
    tags = (f"{A_NS}latin", f"{A_NS}ea", f"{A_NS}cs", f"{A_NS}sym")
    return any((el.get("typeface") or "").lower() in SYMBOL_FONTS for el in rPr.iterchildren(*tags))


def _set_line_black(shape) -> bool:
    """Set a line, connector or arrow outline to black (not other shape borders)."""
    is_connector = shape.element.tag == f"{P_NS}cxnSp"
    prst = shape.element.xpath("./p:spPr/a:prstGeom/@prst")
    if is_connector or (prst and prst[0] in LINE_PRESETS):
        shape.line.color.rgb = BLACK
        return True
    return False


# ----------------------------------------------------------------------
# BULLETS (PPTX)
# ----------------------------------------------------------------------

# Child order of a:pPr / a:lvlNpPr required by the DrawingML schema.
_PPR_CHILD_ORDER = [
    "lnSpc", "spcBef", "spcAft",
    "buClrTx", "buClr",
    "buSzTx", "buSzPct", "buSzPts",
    "buFontTx", "buFont",
    "buNone", "buAutoNum", "buChar", "buBlip",
    "tabLst", "defRPr", "extLst",
]  # fmt: skip


def _replace_ppr_child(pPr, new_child, group: list[str]) -> None:
    """Remove any children of pPr in `group`, then insert new_child in schema order."""
    for tag in group:
        for old in pPr.findall(f"{A_NS}{tag}"):
            pPr.remove(old)
    rank = _PPR_CHILD_ORDER.index(group[0])
    later = {f"{A_NS}{t}" for t in _PPR_CHILD_ORDER[rank + 1 :]}
    for i, child in enumerate(pPr):
        if child.tag in later:
            pPr.insert(i, new_child)
            return
    pPr.append(new_child)


def _make_bullets_black_circles(element) -> int:
    """
    Make every bullet under element a solid black round bullet (Arial '•').

    Covers master text styles, layout/placeholder list styles and per-paragraph
    overrides. Paragraphs with no bullet or with numbering are left alone; a
    paragraph that only overrides the bullet colour gets that colour set to black.
    """
    count = 0
    pPrs = element.iter(f"{A_NS}pPr", *(f"{A_NS}lvl{n}pPr" for n in range(1, 10)))
    for pPr in pPrs:
        if pPr.find(f"{A_NS}buNone") is not None:
            continue
        if pPr.find(f"{A_NS}buAutoNum") is not None:
            continue
        has_bullet = pPr.find(f"{A_NS}buChar") is not None or pPr.find(f"{A_NS}buBlip") is not None
        has_color = pPr.find(f"{A_NS}buClr") is not None or pPr.find(f"{A_NS}buClrTx") is not None
        if not (has_bullet or has_color):
            continue

        bu_clr = etree.Element(f"{A_NS}buClr")
        etree.SubElement(bu_clr, f"{A_NS}srgbClr").set("val", "000000")
        _replace_ppr_child(pPr, bu_clr, ["buClrTx", "buClr"])

        if has_bullet:
            bu_font = etree.Element(f"{A_NS}buFont")
            bu_font.set("typeface", TARGET_FONT)
            _replace_ppr_child(pPr, bu_font, ["buFontTx", "buFont"])

            bu_char = etree.Element(f"{A_NS}buChar")
            bu_char.set("char", STANDARD_BULLET_CHAR)
            _replace_ppr_child(pPr, bu_char, ["buChar", "buBlip"])

        count += 1
    return count


# ----------------------------------------------------------------------
# PRE-NOTES (PPTX) - printer-friendly version
# ----------------------------------------------------------------------


def _set_white_background(element) -> bool:
    """Replace any background on a slide, layout or master with solid white."""
    bg = element.find(f".//{P_NS}bg")
    if bg is not None:
        bg.getparent().remove(bg)

    cSld = element.find(f".//{P_NS}cSld")
    if cSld is None:
        return False
    bg = etree.Element(f"{P_NS}bg")
    bgPr = etree.SubElement(bg, f"{P_NS}bgPr")
    solidFill = etree.SubElement(bgPr, f"{A_NS}solidFill")
    etree.SubElement(solidFill, f"{A_NS}srgbClr").set("val", "FFFFFF")
    etree.SubElement(bgPr, f"{A_NS}effectLst")
    cSld.insert(0, bg)
    return True


def _delete_hidden_slides(prs) -> int:
    """Delete all hidden slides (show="0" on the slide) from the presentation."""
    sldIdLst = prs.slides._sldIdLst
    deleted_count = 0
    for sldId in list(sldIdLst):
        slide = prs.part.related_slide(sldId.rId)
        if slide._element.get("show") in ("0", "false"):
            sldIdLst.remove(sldId)
            prs.part.rels.pop(sldId.rId)
            deleted_count += 1
    return deleted_count


def _remove_decorations(master_or_layout) -> int:
    """Remove pictures, groups and textless non-placeholder shapes from a master/layout."""
    spTree = master_or_layout.element.find(f".//{P_NS}cSld/{P_NS}spTree")
    if spTree is None:
        return 0

    to_remove = []
    for elem in spTree:
        if elem.tag in (f"{P_NS}pic", f"{P_NS}grpSp"):
            to_remove.append(elem)
        elif elem.tag == f"{P_NS}sp":
            is_placeholder = elem.find(f"{P_NS}nvSpPr/{P_NS}nvPr/{P_NS}ph") is not None
            has_text = any(t.text and t.text.strip() for t in elem.iter(f"{A_NS}t"))
            if not (is_placeholder or has_text):
                to_remove.append(elem)

    for elem in to_remove:
        spTree.remove(elem)
    return len(to_remove)


def _standardize_text(part, stats: Counter, black_text: bool) -> None:
    """Set Arial on every non-symbol run and black round bullets on one part."""
    stats["bullets"] += _make_bullets_black_circles(part.element)

    for text_frame in _iter_text_frames(part.shapes):
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if black_text:
                    run.font.color.rgb = BLACK
                if _is_symbol_run(run):
                    stats["symbol_runs"] += 1
                else:
                    run.font.name = TARGET_FONT
                stats["runs"] += 1


def _make_printable(part, stats: Counter, strip_decorations: bool) -> None:
    """Apply all pre-notes changes to one master, layout or slide."""
    if strip_decorations:
        stats["decorations"] += _remove_decorations(part)
    stats["backgrounds"] += _set_white_background(part.element)
    _standardize_text(part, stats, black_text=True)
    stats["lines"] += sum(_set_line_black(s) for s in _iter_shapes(part.shapes))


def pre_notes_output(pptx: Path) -> Path:
    return pptx.with_name(pptx.stem + "_notes.pptx")


def pre_notes(pptx: Path, output: Path | None = None, log: Log = print) -> Path:
    """Convert a PPTX to printable notes: Arial black text, black round bullets,
    white backgrounds, no master decorations, no hidden slides.
    """
    _require_file(pptx, ".pptx", "pre-notes")
    out_path = output or pre_notes_output(pptx)

    log(f"Opening {pptx.name}...")
    prs = Presentation(pptx)
    stats: Counter = Counter()

    hidden_deleted = _delete_hidden_slides(prs)

    log("Cleaning slide masters and layouts...")
    for master in prs.slide_masters:
        _make_printable(master, stats, strip_decorations=True)
        for layout in master.slide_layouts:
            _make_printable(layout, stats, strip_decorations=True)

    log("Cleaning slides...")
    for slide in prs.slides:
        _make_printable(slide, stats, strip_decorations=False)

    log(f"Deleted {hidden_deleted} hidden slides")
    log(f"Removed {stats['decorations']} images and objects from masters/layouts")
    log(f"Set {stats['backgrounds']} backgrounds to white")
    log(
        f"Changed {stats['runs']} text runs to black '{TARGET_FONT}' "
        f"(kept font on {stats['symbol_runs']} symbol-font runs)"
    )
    log(f"Standardized {stats['bullets']} bullets to solid black circles")
    log(f"Changed {stats['lines']} lines/arrows to black")

    prs.save(out_path)
    log(f"Saved: {out_path}")
    return out_path


# ----------------------------------------------------------------------
# STANDARDIZE (PPTX, in place) - Arial fonts, black round bullets, colours kept
# ----------------------------------------------------------------------


def standardize(pptx: Path, log: Log = print) -> Path:
    """Set Arial fonts and black round bullets, editing the PPTX in place.

    Text colours, backgrounds and everything else are left unchanged.
    """
    _require_file(pptx, ".pptx", "standardize")

    log(f"Opening {pptx.name}...")
    prs = Presentation(pptx)
    stats: Counter = Counter()
    for master in prs.slide_masters:
        _standardize_text(master, stats, black_text=False)
        for layout in master.slide_layouts:
            _standardize_text(layout, stats, black_text=False)
    for slide in prs.slides:
        _standardize_text(slide, stats, black_text=False)

    log(f"Changed {stats['runs']} text runs to '{TARGET_FONT}' (kept font on {stats['symbol_runs']} symbol-font runs)")
    log(f"Standardized {stats['bullets']} bullets to solid black circles")

    prs.save(pptx)
    log(f"Saved: {pptx}")
    return pptx


# ----------------------------------------------------------------------
# PDF (LibreOffice)
# ----------------------------------------------------------------------


def _pptx_to_pdf(pptx: Path, out_dir: Path, log: Log) -> Path:
    """Convert pptx to <out_dir>/<stem>.pdf with headless LibreOffice and return it.

    Uses a throwaway LibreOffice profile so conversion works even while a
    LibreOffice window is open (otherwise it can silently do nothing).
    """
    soffice = _require_tool("libreoffice", "soffice")
    produced = out_dir / (pptx.stem + ".pdf")
    produced.unlink(missing_ok=True)

    log("Running LibreOffice (this can take a while)...")
    with tempfile.TemporaryDirectory(prefix="pptxtool_lo_") as profile:
        _run(
            [
                soffice,
                f"-env:UserInstallation={Path(profile).as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(out_dir),
                str(pptx),
            ],
            log,
        )

    if not produced.exists():
        raise PptxToolError(f"LibreOffice did not produce {produced}")
    return produced


def pdf(pptx: Path, output: Path | None = None, log: Log = print) -> Path:
    """Convert PPTX to PDF using LibreOffice (headless)."""
    _require_file(pptx, ".pptx", "pdf")
    out_path = output or pptx.with_suffix(".pdf")

    produced = _pptx_to_pdf(pptx, out_path.parent, log)
    if produced != out_path:
        shutil.move(produced, out_path)

    log(f"PDF written to: {out_path}")
    return out_path


# ----------------------------------------------------------------------
# WATERMARK
# ----------------------------------------------------------------------


def try_register() -> str:
    candidates = [
        (
            "Liberation Sans",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ),
        ("DejaVu Sans", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ("Arial", "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf"),
        ("Arial", "/usr/share/fonts/truetype/msttcorefonts/arial.ttf"),
        ("Arial", "/usr/share/fonts/truetype/msttcorefonts/ARIAL.TTF"),
        ("Arial", "/Library/Fonts/Arial.ttf"),
        ("Arial", "C:\\Windows\\Fonts\\arial.ttf"),
    ]
    for desired_name, path in candidates:
        if desired_name in pdfmetrics.getRegisteredFontNames():
            return desired_name

        try:
            if os.path.isfile(path):
                pdfmetrics.registerFont(TTFont(desired_name, path))
                return desired_name
        except Exception:  # noqa: BLE001, S110 - unusable font file; try the next one
            pass
    return "Helvetica"  # fallback


def _ensure_reportlab_font_metrics_work(font_name: str) -> str:
    """
    Minimal safety guard: if ReportLab can't produce sane metrics for the chosen font name,
    fall back to Helvetica. This prevents the 'giant watermark' failure mode where
    stringWidth() is effectively 0 and the fitter drives size to the upper bound.
    """
    try:
        w = pdfmetrics.stringWidth("TEST", font_name, 12)
        if w and w > 0:
            return font_name
    except Exception:  # noqa: BLE001, S110 - fall through to Helvetica
        pass
    return "Helvetica"


def rotated_aabb(w: float, h: float, theta_rad: float) -> tuple[float, float]:
    c = abs(math.cos(theta_rad))
    s = abs(math.sin(theta_rad))
    return (w * c + h * s, w * s + h * c)


def fit_font_size_for_box(
    text: str,
    font_name: str,
    page_w: float,
    page_h: float,
    theta_rad: float,
    margin: float,
) -> float:
    avail_w = max(page_w - 2 * margin, 1.0)
    avail_h = max(page_h - 2 * margin, 1.0)

    def ok(size: float) -> bool:
        adv = pdfmetrics.stringWidth(text, font_name, size)
        try:
            asc = pdfmetrics.getAscent(font_name, size)
            desc = pdfmetrics.getDescent(font_name, size)  # negative
        except Exception:  # noqa: BLE001
            asc, desc = 0.8 * size, -0.2 * size
        text_h = asc - desc
        aabb_w, aabb_h = rotated_aabb(adv, text_h, theta_rad)
        return (aabb_w <= avail_w) and (aabb_h <= avail_h)

    lo, hi = 1.0, 4000.0
    for _ in range(60):
        mid = (lo + hi) / 2.0
        if ok(mid):
            lo = mid
        else:
            hi = mid
    return lo


def make_watermark_pdf_bytes(
    page_w: float,
    page_h: float,
    text: str,
    angle_deg: float | None,
    font_name: str,
    gray: float,
    alpha: float | None = None,
    margin_frac: float = 0.04,
    tile_alpha: float | None = None,
    rng: random.Random | None = None,
) -> bytes:
    """Draw the main diagonal watermark, optionally over a faint tiled layer.

    With rng, the angle, size, position, gray and alpha are nudged per call so
    no two pages carry an identical watermark (defeats multi-page averaging).
    """
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_w, page_h))

    theta = math.radians(angle_deg) if angle_deg is not None else math.atan2(page_h, page_w)
    margin = math.hypot(page_w, page_h) * margin_frac

    size = fit_font_size_for_box(text, font_name, page_w, page_h, theta, margin)

    cx, cy = page_w / 2.0, page_h / 2.0
    if rng is not None:
        # Shrink a little first so the rotated, shifted text still fits the page.
        size *= rng.uniform(0.85, 0.95)
        theta += math.radians(rng.uniform(-4.0, 4.0))
        cx += rng.uniform(-0.03, 0.03) * page_w
        cy += rng.uniform(-0.03, 0.03) * page_h
        gray = min(1.0, max(0.0, gray + rng.uniform(-0.05, 0.05)))
        if alpha is not None:
            alpha = min(1.0, max(0.0, alpha + rng.uniform(-0.05, 0.05)))

    if tile_alpha:
        _draw_watermark_tiles(c, page_w, page_h, text, font_name, theta, tile_alpha, rng)

    adv = pdfmetrics.stringWidth(text, font_name, size)
    try:
        asc = pdfmetrics.getAscent(font_name, size)
        desc = pdfmetrics.getDescent(font_name, size)  # negative
    except Exception:  # noqa: BLE001
        asc, desc = 0.8 * size, -0.2 * size
    baseline_y = -(asc + desc) / 2.0

    c.saveState()
    c.translate(cx, cy)
    c.rotate(math.degrees(theta))
    c.setFillGray(gray)
    if alpha is not None:
        c.setFillAlpha(alpha)
    c.setFont(font_name, size)
    c.drawString(-adv / 2.0, baseline_y, text)
    c.restoreState()
    c.showPage()
    c.save()

    packet.seek(0)
    return packet.getvalue()


def _draw_watermark_tiles(
    c: canvas.Canvas,
    page_w: float,
    page_h: float,
    text: str,
    font_name: str,
    theta: float,
    tile_alpha: float,
    rng: random.Random | None,
) -> None:
    """Cover the page with small, faint, staggered copies of the watermark text."""
    size = max(page_h / 45.0, 6.0)
    adv = pdfmetrics.stringWidth(text, font_name, size)
    step_x = adv * 1.3
    step_y = size * 4.0
    # Random grid phase per page, so tiles never line up across pages.
    ox = rng.uniform(0.0, step_x) if rng else 0.0
    oy = rng.uniform(0.0, step_y) if rng else 0.0
    half = math.hypot(page_w, page_h) / 2.0 + step_x

    c.saveState()
    c.translate(page_w / 2.0, page_h / 2.0)
    c.rotate(math.degrees(theta))
    c.setFillGray(0.45)
    c.setFillAlpha(tile_alpha)
    # The rotated grid overshoots the page; keep only tiles that can touch it.
    cos_t, sin_t = math.cos(theta), math.sin(theta)
    reach = adv / 2.0 + size
    origins = []
    row = 0
    y = -half + oy
    while y < half:
        x = -half + ox + (step_x / 2.0 if row % 2 else 0.0)
        while x < half:
            mid_x = x + adv / 2.0
            px = page_w / 2.0 + mid_x * cos_t - y * sin_t
            py = page_h / 2.0 + mid_x * sin_t + y * cos_t
            if -reach < px < page_w + reach and -reach < py < page_h + reach:
                origins.append((x, y))
            x += step_x
        y += step_y
        row += 1
    c.setFont(font_name, size)
    for x, y in origins:
        c.drawString(x, y, text)
    c.restoreState()


def _report_page(log: Log, verb: str, i: int, n: int) -> None:
    """Log page progress every 10 pages and on the last one."""
    if i == n or i % 10 == 0:
        log(f"  {verb} page {i}/{n}")


def watermark_pdf(
    input_path: str,
    output_path: str,
    text: str,
    angle: float | None,
    gray: float,
    alpha: float | None,
    layer: str,
    tile_alpha: float | None = None,
    jitter: bool = False,
    seed: int | None = None,
    log: Log = print,
) -> None:
    font_name = try_register()
    font_name = _ensure_reportlab_font_metrics_work(font_name)

    reader = PdfReader(input_path)
    writer = PdfWriter()
    rng = random.Random(seed) if jitter else None

    wm_cache: dict[tuple[float, float, float | None, float, float | None], bytes] = {}

    n_pages = len(reader.pages)
    for i, page in enumerate(reader.pages, start=1):
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        key = (w, h, angle, gray, alpha)
        # Jittered watermarks are unique per page, so they are never cached.
        if rng is not None or key not in wm_cache:
            wm_cache[key] = make_watermark_pdf_bytes(
                page_w=w,
                page_h=h,
                text=text,
                angle_deg=angle,
                font_name=font_name,
                gray=gray,
                alpha=alpha,
                tile_alpha=tile_alpha,
                rng=rng,
            )

        wm_reader = PdfReader(BytesIO(wm_cache[key]))
        wm_page = wm_reader.pages[0]

        if layer == "bottom":
            wm_page.merge_page(page)
            writer.add_page(wm_page)
        else:
            page.merge_page(wm_page)
            writer.add_page(page)
        writer.pages[-1].compress_content_streams()
        _report_page(log, "Watermarked", i, n_pages)

    with open(output_path, "wb") as f:
        writer.write(f)


RASTER_WIDTH_PX = 1600
RASTER_JPEG_QUALITY = 70


def _is_grayscale(im: Image.Image, tolerance: int = 8) -> bool:
    r, g, b = im.split()
    return all(ImageChops.difference(x, y).getextrema()[1] <= tolerance for x, y in ((r, g), (g, b)))


def rasterize_pdf(
    src: Path,
    dst: Path,
    width_px: int = RASTER_WIDTH_PX,
    quality: int = RASTER_JPEG_QUALITY,
    log: Log = print,
) -> None:
    """Rebuild a PDF from one JPEG per page, so watermark and content share pixels.

    Pages are rendered to a fixed pixel width rather than a dpi, since slide
    PDFs vary widely in point size. Colourless pages are stored as grayscale.
    JPEG beats PNG and JPEG 2000 here: it is small and keeps the faint tiles,
    which JPEG 2000 smooths away.
    """
    doc = pymupdf.open(src)
    out = pymupdf.open()
    n_pages = len(doc)
    for i, page in enumerate(doc, start=1):
        zoom = width_px / page.rect.width
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        if _is_grayscale(im):
            im = im.convert("L")
        buf = BytesIO()
        im.save(buf, "JPEG", quality=quality, optimize=True)
        new_page = out.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, stream=buf.getvalue())
        _report_page(log, "Rasterized", i, n_pages)
    out.save(dst, garbage=4, deflate=True)


def watermark_outputs(base: Path) -> tuple[Path, Path]:
    """The (reader, raster) paths watermark_reader_and_raster writes for base."""
    return base.with_name(base.name + "_reader.pdf"), base.with_name(base.name + ".pdf")


def watermark_reader_and_raster(
    src: Path, base: Path, tile_alpha: float | None = None, log: Log = print, **watermark_args
) -> Path:
    """Write <base>_reader.pdf (live text, text watermark) and <base>.pdf (raster).

    The reader copy keeps text selectable for screen readers and skips the
    tiles, which a screen reader would otherwise read aloud dozens of times per
    slide. The raster copy gets the tiles and is pixels only, so the watermark
    can't be lifted out as a separate object. The raster copy may overwrite
    src: src is fully read before the raster copy is written. Returns its path.
    """
    reader_path, raster_path = watermark_outputs(base)
    log("Writing screen-reader copy...")
    watermark_pdf(input_path=str(src), output_path=str(reader_path), tile_alpha=None, log=log, **watermark_args)
    log(f"Reader PDF: {reader_path}")
    with tempfile.TemporaryDirectory(prefix="pptxtool_raster_") as tmp:
        tiled = Path(tmp) / "tiled.pdf"
        log("Writing tiled watermark copy...")
        watermark_pdf(input_path=str(src), output_path=str(tiled), tile_alpha=tile_alpha, log=log, **watermark_args)
        log("Rasterizing pages...")
        rasterize_pdf(tiled, raster_path, log=log)
    log(f"Raster PDF: {raster_path}")
    return raster_path


def watermark(
    pdf_file: Path,
    text: str = WATERMARK_TEXT,
    output: Path | None = None,
    angle: float | None = None,
    gray: float = 0.7,
    alpha: float | None = 0.5,
    layer: str = "top",
    tile_alpha: float | None = 0.05,
    jitter: bool = True,
    seed: int | None = None,
    log: Log = print,
) -> Path:
    """Watermark a PDF, writing a screen-reader copy and a raster-only copy.

    Writes <base>_reader.pdf and the raster copy as <base>.pdf, where base is
    output (or the input) without its suffix. angle=None means diagonal.
    """
    _require_file(pdf_file, ".pdf", "watermark")
    base = (output or pdf_file).with_suffix("")
    if layer not in {"bottom", "top"}:
        raise PptxToolError('Invalid layer. Use "bottom" or "top".')

    return watermark_reader_and_raster(
        pdf_file,
        base,
        text=text,
        angle=angle,
        gray=gray,
        alpha=alpha,
        layer=layer,
        tile_alpha=tile_alpha,
        jitter=jitter,
        seed=seed,
        log=log,
    )


# ----------------------------------------------------------------------
# COMPRESS (Ghostscript, safe temp-file strategy)
# ----------------------------------------------------------------------


def _compress_pdf(src: Path, dst: Path, pdfsettings: str, log: Log) -> None:
    """Compress src into dst with Ghostscript via a temp file (src may equal dst)."""
    gs = _require_tool("gs")

    # Write to temp file first (never read+write same file)
    with tempfile.NamedTemporaryFile(prefix="presentation_tool_", suffix=".pdf", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        before = src.stat().st_size
        _run(
            [
                gs,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={pdfsettings}",
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                f"-sOutputFile={tmp_path}",
                str(src),
            ],
            log,
        )
        after = tmp_path.stat().st_size
        shutil.move(tmp_path, dst)
        log(f"  Size: {before / 1e6:.2f} MB -> {after / 1e6:.2f} MB")
    finally:
        tmp_path.unlink(missing_ok=True)


def compress(
    pdf_file: Path, output: Path | None = None, pdfsettings: str = DEFAULT_PDF_SETTINGS, log: Log = print
) -> Path:
    """Compress a PDF using Ghostscript (writes via temp file, then replaces)."""
    _require_file(pdf_file, ".pdf", "compress")
    out_path = output or pdf_file
    log(f"Compressing with Ghostscript ({pdfsettings})...")
    _compress_pdf(pdf_file, out_path, pdfsettings, log)
    log(f"Compressed PDF written to: {out_path}")
    return out_path


# ----------------------------------------------------------------------
# POST-NOTES (pipeline: pdf -> compress -> watermark -> reader + raster copies)
# ----------------------------------------------------------------------


def _sanitize_filename(name: str) -> str:
    """Replace punctuation with spaces, then multiple spaces/underscores with single underscore."""
    # Replace all punctuation (except period for extension) with spaces
    result = name
    for char in string.punctuation:
        if char != ".":
            result = result.replace(char, " ")
    # Replace multiple spaces with single space, then spaces with underscores
    result = " ".join(result.split())
    result = result.replace(" ", "_")
    return result


def post_notes_outputs(pptx: Path) -> list[Path]:
    """Every file post_notes writes (or clobbers along the way), without duplicates."""
    base = pptx.with_name(_sanitize_filename(pptx.stem))
    paths = [pptx.with_suffix(".pdf"), *watermark_outputs(base)]
    return list(dict.fromkeys(paths))


def post_notes(
    pptx: Path,
    watermark_text: str = WATERMARK_TEXT,
    pdfsettings: str = DEFAULT_PDF_SETTINGS,
    log: Log = print,
) -> Path:
    """Convert PPTX to PDF, compress, watermark; write <stem>_reader.pdf and raster <stem>.pdf."""
    _require_file(pptx, ".pptx", "post-notes")
    _require_tool("gs")  # fail before the slow conversion, not after

    log("Step 1/3: Converting to PDF...")
    pdf_path = _pptx_to_pdf(pptx, pptx.parent, log)
    log(f"PDF created: {pdf_path}")

    log(f"Step 2/3: Compressing PDF ({pdfsettings})...")
    _compress_pdf(pdf_path, pdf_path, pdfsettings, log)
    log("PDF compressed")

    log("Step 3/3: Adding watermark...")
    raster_path = watermark_reader_and_raster(
        pdf_path,
        pdf_path.with_name(_sanitize_filename(pptx.stem)),
        text=watermark_text,
        angle=None,
        gray=0.7,
        alpha=0.5,
        layer="top",
        tile_alpha=0.05,
        jitter=True,
        log=log,
    )
    if pdf_path != raster_path:
        pdf_path.unlink()
    return raster_path


# Same pipeline, named for full-colour decks that skip pre-notes.
pdf_watermark_compress = post_notes
