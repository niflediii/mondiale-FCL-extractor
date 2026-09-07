"""Extract FCL rate rows from a Mondiale VGL Import Seafreight FCL Quotation PDF."""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import pdfplumber

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
# regular-weight amount, so it is read from chars rather than from the columns
# above. Excel's hidden spacer columns also bleed stray hyphens into the
# currency band, which is why only the bold run counts as the label.
SEC_FEE_CCY = (487.4, 507.3)
SEC_FEE_VAL = (507.3, 529.3)

HEADER_BOTTOM = 124.0  # below the two-line column header
FOOTER_TOP = 466.0  # above the "Rates are subject to variation..." boilerplate
SECTION_BAND = (85.0, 102.0)  # the region title, e.g. "North Asia"
ROW_TOLERANCE = 5.0  # points; rows are ~14pt apart

BLANK = {"", "-", "--", "---", "----"}
DATE = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{2}$")


def _clean(text: str) -> str:
    """Repair the mojibake left by the PDF's smart quotes."""
    text = text.replace("�", "’")
    # A quote right after a digit is a foot mark (20' container), not an apostrophe.
    return re.sub(r"(?<=\d)’", "'", text)


def _cell(text: str) -> str:
    text = _clean(text).strip()
    return "" if text in BLANK else text


def _section_title(words: list[dict]) -> str:
    """The region heading printed above each page's table, e.g. "North Asia"."""
    band = [w for w in words if SECTION_BAND[0] < w["top"] < SECTION_BAND[1]]
    return " ".join(w["text"] for w in sorted(band, key=lambda w: w["x0"])).strip()


@dataclass
class Region:
    name: str
    rows: list[dict] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def note_text(self) -> str:
        return "\n".join(self.notes).strip()


def _group_rows(words: list[dict]) -> list[list[dict]]:
    """Cluster words into visual rows by their vertical position."""
    rows: list[list[dict]] = []
    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if rows and abs(word["top"] - rows[-1][0]["top"]) <= ROW_TOLERANCE:
            rows[-1].append(word)
        else:
            rows.append([word])
    return rows


def _security_fee_chars(page) -> list[tuple[dict, bool]]:
    """The chars in the two Security Fee bands, flagged currency or amount.

    Scanned once per page rather than once per row, which is what makes the
    per-row lookup below cheap.
    """
    keep: list[tuple[dict, bool]] = []
    for char in page.chars:
        mid = (char["x0"] + char["x1"]) / 2
        if SEC_FEE_CCY[0] <= mid < SEC_FEE_CCY[1]:
            if "Bold" in char["fontname"]:
                keep.append((char, True))
        elif SEC_FEE_VAL[0] <= mid < SEC_FEE_VAL[1]:
            keep.append((char, False))
    return keep


def _security_fee(chars: list[tuple[dict, bool]], top: float, bottom: float) -> tuple[str, str]:
    ccy_chars, val_chars = [], []
    for char, is_currency in chars:
        if top <= char["top"] <= bottom:
            (ccy_chars if is_currency else val_chars).append(char)

    def join(items):
        return "".join(c["text"] for c in sorted(items, key=lambda c: c["x0"]))

    return _cell(join(ccy_chars)), _cell(join(val_chars))


def _parse_page(page, words: list[dict], in_notes: bool) -> tuple[list[dict], list[str], bool]:
    """Return (rate rows, note lines, still-in-notes) for one page."""
    body = [w for w in words if HEADER_BOTTOM < w["top"] < FOOTER_TOP]
    rows: list[dict] = []
    notes: list[str] = []

    fee_chars = _security_fee_chars(page) if body else []

    for line in _group_rows(body):
        line.sort(key=lambda w: w["x0"])

        # The notes block opens with a "Notes" label in the far-left column and
        # runs to the end of the region, sometimes spilling onto later pages.
        if not in_notes and line[0]["text"] == "Notes" and line[0]["x0"] < 30:
            in_notes = True
            line = line[1:]

        if in_notes:
            text = _clean(" ".join(w["text"] for w in line)).strip()
            if text:
                notes.append(text)
            continue

        cells: dict[str, list[str]] = {name: [] for name, _, _ in COLUMNS}
        for word in line:
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

        top = min(w["top"] for w in line)
        bottom = max(w["bottom"] for w in line)
        record["sec_fee_currency"], record["sec_fee"] = _security_fee(fee_chars, top, bottom)
        rows.append(record)

    return rows, notes, in_notes


def parse_pdf(source: str | bytes | io.BytesIO) -> list[Region]:
    """Parse the quotation into regions, each with its rate rows and notes."""
    if isinstance(source, bytes):
        source = io.BytesIO(source)

    regions: list[Region] = []
    in_notes = False
    with pdfplumber.open(source) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            section = _section_title(words)
            if section:
                if not regions or regions[-1].name != section:
                    regions.append(Region(name=section))
                    in_notes = False  # a new region restarts with its rate table
                rows, notes, in_notes = _parse_page(page, words, in_notes)
                regions[-1].rows.extend(rows)
                regions[-1].notes.extend(notes)
            page.flush_cache()

    return [r for r in regions if r.rows]
