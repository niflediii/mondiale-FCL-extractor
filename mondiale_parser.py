"""Extract FCL rate rows from a Mondiale VGL Import Seafreight FCL Quotation PDF."""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import pdfplumber
import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c

# ---------------------------------------------------------------------------
# Table geometry
#
# The quotation is an Excel export, so every page carries the same ruled grid.
# These are the vertical cell borders (PDF points) that separate the columns.
# ---------------------------------------------------------------------------

COLUMNS: list[tuple[str, float, float]] = [
    ("country", 14.0, 74.0),
    ("origin", 74.0, 149.4),
    ("destination", 149.4, 196.2),
    ("ship_line", 196.2, 235.6),
    ("currency", 235.6, 255.1),
    ("gp20", 255.1, 276.4),
    ("nor20", 276.4, 297.6),
    ("gp40", 297.6, 319.0),
    ("nor40", 319.0, 340.2),
    ("hc40", 340.2, 361.5),
    ("baf20", 361.5, 382.8),
    ("baf40", 382.8, 403.9),
    ("lss20", 403.9, 425.3),
    ("lss40", 425.3, 446.6),
    ("pss20", 446.6, 467.8),
    ("pss40", 467.8, 487.4),
    ("psc20", 529.3, 561.7),
    ("psc40", 561.7, 594.3),
    ("service_details", 594.3, 741.2),
    ("valid_from", 741.2, 783.6),
    ("valid_to", 783.6, 826.0),
]

# The Security Fee cell holds two overlapping runs: a bold currency label and a
# regular-weight amount, so it is read from glyphs rather than from the columns
# above. Excel's hidden spacer columns also bleed stray hyphens into the
# currency band, which is why only the bold run counts as the label.
SEC_FEE_CCY = (487.4, 507.3)
SEC_FEE_VAL = (507.3, 529.3)

HEADER_BOTTOM = 124.0  # below the two-line column header
FOOTER_TOP = 466.0  # above the "Rates are subject to variation..." boilerplate
SECTION_BAND = (85.0, 102.0)  # the region title, e.g. "North Asia"
ROW_TOLERANCE = 5.0  # points; rows are ~14pt apart
NOTES_INDENT = 30.0  # the "Notes" label sits left of this
BOLD_WEIGHT = 600  # pdfium reports 400 for regular and 700 for bold

# Glyphs closer than this are one word. Calibrated against pdfplumber's own
# word extraction: at 1.0 the two agree exactly on the rate tables.
WORD_GAP = 1.0

BLANK = {"", "-", "--", "---", "----"}
DATE = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{2}$")


def _clean(text: str) -> str:
    """Normalise the quotation's smart quotes."""
    # U+FFFD appears when a decoder cannot map the source's curly apostrophe.
    text = text.replace("�", "’")
    # A quote right after a digit is a foot mark (20' container), not an apostrophe.
    return re.sub(r"(?<=\d)[‘’]", "'", text)


def _cell(text: str) -> str:
    text = _clean(text).strip()
    return "" if text in BLANK else text


@dataclass
class Region:
    name: str
    rows: list[dict] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def note_text(self) -> str:
        return "\n".join(self.notes).strip()


# ---------------------------------------------------------------------------
# Rate tables are read with pdfium, which decodes glyphs in C. pdfminer (via
# pdfplumber) allocates a Python object per glyph and costs ~20x more on a file
# this dense. The two agree exactly on the ruled tables.
#
# The notes blocks are prose, and there pdfium's synthetic spaces land mid-word
# often enough to corrupt the text, so those few pages are read with pdfplumber
# instead. Notes appear on roughly 8 of the 86 pages, so the cost is small.
# ---------------------------------------------------------------------------


def _page_glyphs(textpage, height: float) -> list[dict]:
    """Positioned glyphs for one page, in pdfplumber's top-left coordinates."""
    handle = textpage.raw
    glyphs: list[dict] = []
    for index in range(textpage.count_chars()):
        text = chr(pdfium_c.FPDFText_GetUnicode(handle, index))
        if not text.strip():
            continue  # pdfium synthesises spaces from gaps; use the gaps instead
        # loose=True gives the glyph's advance box rather than its inked bounds,
        # which is what the column and word-gap maths expect.
        left, bottom, right, top = textpage.get_charbox(index, loose=True)
        glyphs.append(
            {
                "text": text,
                "x0": left,
                "x1": right,
                "top": height - top,
                "bottom": height - bottom,
                "bold": pdfium_c.FPDFText_GetFontWeight(handle, index) >= BOLD_WEIGHT,
            }
        )
    return glyphs


def _group_rows(items: list[dict]) -> list[list[dict]]:
    """Cluster glyphs or words into visual rows by their vertical position."""
    rows: list[list[dict]] = []
    for item in sorted(items, key=lambda i: (i["top"], i["x0"])):
        if rows and abs(item["top"] - rows[-1][0]["top"]) <= ROW_TOLERANCE:
            rows[-1].append(item)
        else:
            rows.append([item])
    return rows


def _words(line: list[dict]) -> list[dict]:
    """Join one row's glyphs into words wherever they sit closer than WORD_GAP."""
    words: list[dict] = []
    for glyph in line:
        if words and glyph["x0"] - words[-1]["x1"] <= WORD_GAP:
            word = words[-1]
            word["text"] += glyph["text"]
            word["x1"] = max(word["x1"], glyph["x1"])
        else:
            words.append({"text": glyph["text"], "x0": glyph["x0"], "x1": glyph["x1"]})
    return words


def _section_title(glyphs: list[dict]) -> str:
    """The region heading printed above each page's table, e.g. "North Asia"."""
    band = sorted(
        (g for g in glyphs if SECTION_BAND[0] < g["top"] < SECTION_BAND[1]),
        key=lambda g: g["x0"],
    )
    return " ".join(w["text"] for w in _words(band)).strip()


def _security_fee(line: list[dict]) -> tuple[str, str]:
    """Split the Security Fee cell into its currency label and its amount."""
    ccy_glyphs, val_glyphs = [], []
    for glyph in line:
        mid = (glyph["x0"] + glyph["x1"]) / 2
        if SEC_FEE_CCY[0] <= mid < SEC_FEE_CCY[1]:
            if glyph["bold"]:
                ccy_glyphs.append(glyph)
        elif SEC_FEE_VAL[0] <= mid < SEC_FEE_VAL[1]:
            val_glyphs.append(glyph)

    def join(items):
        return "".join(i["text"] for i in sorted(items, key=lambda i: i["x0"]))

    return _cell(join(ccy_glyphs)), _cell(join(val_glyphs))


def _rate_rows(glyphs: list[dict], in_notes: bool) -> tuple[list[dict], str | None, bool]:
    """Rate rows on one page, plus how its notes block (if any) begins.

    The second element is "label" when this page opens a notes block, "all"
    when the whole body continues one from an earlier page, else None.
    """
    body = [g for g in glyphs if HEADER_BOTTOM < g["top"] < FOOTER_TOP]
    if not body:
        return [], ("all" if in_notes else None), in_notes

    notes_start = "all" if in_notes else None
    rows: list[dict] = []

    for line in _group_rows(body):
        line.sort(key=lambda g: g["x0"])
        words = _words(line)
        if not words:
            continue

        # The notes block opens with a "Notes" label in the far-left column and
        # runs to the end of the region, sometimes spilling onto later pages.
        if not in_notes and words[0]["text"] == "Notes" and words[0]["x0"] < NOTES_INDENT:
            in_notes = True
            notes_start = "label"
        if in_notes:
            continue

        cells: dict[str, list[str]] = {name: [] for name, _, _ in COLUMNS}
        for word in words:
            mid = (word["x0"] + word["x1"]) / 2
            for name, left, right in COLUMNS:
                if left <= mid < right:
                    cells[name].append(word["text"])
                    break

        record = {name: _cell(" ".join(parts)) for name, parts in cells.items()}
        # A real rate row always carries a ship line and both validity dates.
        if not (record["origin"] and record["destination"] and record["ship_line"]):
            continue
        if not (DATE.match(record["valid_from"]) and DATE.match(record["valid_to"])):
            continue

        record["sec_fee_currency"], record["sec_fee"] = _security_fee(line)
        rows.append(record)

    return rows, notes_start, in_notes


def _note_lines(page, starts_at_label: bool) -> list[str]:
    """The note lines on one page, read with pdfplumber for faithful spacing."""
    body = [w for w in page.extract_words() if HEADER_BOTTOM < w["top"] < FOOTER_TOP]
    lines: list[str] = []
    reached = not starts_at_label

    for line in _group_rows(body):
        line.sort(key=lambda w: w["x0"])
        if not reached:
            if line[0]["text"] != "Notes" or line[0]["x0"] >= NOTES_INDENT:
                continue
            reached = True
            line = line[1:]
        text = _clean(" ".join(w["text"] for w in line)).strip()
        if text:
            lines.append(text)
    return lines


def parse_pdf(source: str | bytes | io.BytesIO) -> list[Region]:
    """Parse the quotation into regions, each with its rate rows and notes."""
    if isinstance(source, io.BytesIO):
        source = source.getvalue()

    regions: list[Region] = []
    # region index -> [(page number, whether the block starts at a "Notes" label)]
    pending_notes: list[tuple[int, int, bool]] = []
    in_notes = False

    document = pdfium.PdfDocument(source)
    try:
        for number, page in enumerate(document):
            textpage = page.get_textpage()
            try:
                glyphs = _page_glyphs(textpage, page.get_height())
            finally:
                textpage.close()
                page.close()

            section = _section_title(glyphs)
            if not section:
                continue
            if not regions or regions[-1].name != section:
                regions.append(Region(name=section))
                in_notes = False  # a new region restarts with its rate table

            rows, notes_start, in_notes = _rate_rows(glyphs, in_notes)
            regions[-1].rows.extend(rows)
            if notes_start:
                pending_notes.append((len(regions) - 1, number, notes_start == "label"))
    finally:
        document.close()

    if pending_notes:
        with pdfplumber.open(io.BytesIO(source) if isinstance(source, bytes) else source) as pdf:
            for region_index, number, at_label in pending_notes:
                page = pdf.pages[number]
                regions[region_index].notes.extend(_note_lines(page, at_label))
                page.flush_cache()

    return [r for r in regions if r.rows]
