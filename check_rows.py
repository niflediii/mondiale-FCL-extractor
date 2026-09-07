"""Cross-check the geometric parser against a naive text scan, page by page."""
import re
import sys

import pdfplumber

import mondiale_parser as P

DATE_PAIR = re.compile(r"\d{1,2}-[A-Za-z]{3}-\d{2}\s+\d{1,2}-[A-Za-z]{3}-\d{2}")

pdf = pdfplumber.open(sys.argv[1])
total = mismatched = 0
in_notes = False
last_section = None

for i, page in enumerate(pdf.pages, start=1):
    words = page.extract_words()
    section = P._section_title(words)
    if section != last_section:
        in_notes = False
        last_section = section
    rows, _, in_notes = P._parse_page(page, words, in_notes)
    naive = sum(1 for ln in (page.extract_text() or "").split("\n") if DATE_PAIR.search(ln))
    total += len(rows)
    if naive != len(rows):
        mismatched += 1
        print(f"page {i}: parsed {len(rows)}, naive {naive}")

print(f"total parsed {total}, mismatched pages {mismatched}")
