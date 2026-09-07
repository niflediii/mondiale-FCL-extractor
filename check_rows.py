"""Cross-check the parser against an independent text scan, page by page.

The parser finds rate rows by grid geometry with pdfium. This counts them a
second way — scanning pdfplumber's text for the pair of validity dates that
ends every rate row — and reports any page where the two disagree.
"""

import re
import sys

import pdfplumber
import pypdfium2 as pdfium

import mondiale_parser as parser

DATE_PAIR = re.compile(r"\d{1,2}-[A-Za-z]{3}-\d{2}\s+\d{1,2}-[A-Za-z]{3}-\d{2}")


def parsed_rows_per_page(path):
    """Rate-row counts per page, straight from the parser's own machinery."""
    counts = []
    in_notes = False
    last_section = None
    document = pdfium.PdfDocument(path)
    try:
        for page in document:
            textpage = page.get_textpage()
            try:
                glyphs = parser._page_glyphs(textpage, page.get_height())
            finally:
                textpage.close()
                page.close()

            section = parser._section_title(glyphs)
            if not section:
                counts.append(0)
                continue
            if section != last_section:
                in_notes = False
                last_section = section
            rows, _, in_notes = parser._rate_rows(glyphs, in_notes)
            counts.append(len(rows))
    finally:
        document.close()
    return counts


def main(path):
    parsed = parsed_rows_per_page(path)

    mismatched = 0
    with pdfplumber.open(path) as pdf:
        if len(pdf.pages) != len(parsed):
            print(f"page count mismatch: pdfplumber {len(pdf.pages)}, pdfium {len(parsed)}")
            return 1
        for number, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            page.flush_cache()
            naive = sum(1 for line in text.split("\n") if DATE_PAIR.search(line))
            if naive != parsed[number]:
                mismatched += 1
                print(f"page {number + 1}: parsed {parsed[number]}, naive {naive}")

    print(f"total parsed {sum(parsed)}, mismatched pages {mismatched}")
    return 1 if mismatched else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
