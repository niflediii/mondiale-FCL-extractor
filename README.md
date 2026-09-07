# Mondiale FCL Automation

Upload the Mondiale VGL **Import Seafreight FCL Quotation** PDF, get back an Excel
workbook with every rate row flattened into a single `FCL Rates` sheet.

## Run it

```bash
pip install -r requirements.txt
python app.py
```

Then open http://localhost:5000, drop the quotation PDF on the dropzone, and press
**Process File**. The workbook downloads automatically as
`Mondiale_FCL_YYYY-MM-DD.xlsx` (today's date).

An 86-page quotation takes about 5 seconds locally (7s end to end over HTTP).
The page shows a progress overlay while it works. gunicorn still runs with
`--timeout 600` so that a much larger quotation, or a slow instance, cannot be
cut off mid-parse.

## Deploying to Render

`render.yaml` is a Blueprint: in the Render dashboard choose **New > Blueprint**
and point it at this repo. Or create a Web Service by hand with:

- Language: **Python 3**
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app --workers 1 --threads 4 --timeout 600 --bind 0.0.0.0:$PORT`

`.python-version` pins Python 3.14 (the version this was built and tested on).
If a build ever fails fetching a wheel for `Pillow` or `pypdfium2`, drop that
file to `3.13`.

The free instance type (512 MB RAM, 0.1 CPU) is enough: peak memory is ~90 MB,
and since the parse is only a few CPU-seconds, even a tenth of a core finishes
in well under a minute. Free instances do spin down after 15 minutes idle and
cold-start in 30-60s, so the first request after a quiet spell is slow.

## Output

One sheet named `FCL Rates`, columns A–W:

| Col | Header | Source in the PDF |
| --- | --- | --- |
| A | Country | `Country` |
| B | Origin | `Port` |
| C | Destination | `Arrive` |
| D | Ship Line | `Ship Line` |
| E | Service Details | `Service Details` |
| F | Valid From | `Valid From` |
| G | Valid To | `Valid To` |
| H | Notes | the `Notes` block printed under the row's table |
| I | Currency | `Currency` |
| J–L | 20'GP, 40'GP, 40'HC | freight rates |
| M–N | 20'NOR, 40'NOR | freight rates |
| O–P | 20'BAF, 40'BAF | bunker adjustment |
| Q–R | 20'LSS, 40'LSS | low sulphur surcharge |
| S–T | 20'PSS/EBS, 40'PSS/EBS | peak season / emergency bunker |
| U | Security Fee_*CCY* | the currency is read from the PDF and appended to the header |
| V–W | 20'PSC (NZD), 40'PSC (NZD) | NZ port service charge |

Rates and fees are written as numbers, validity as real dates. A dash in the PDF
means "not offered" and is left blank.

Column order is set by `LAYOUT` in `excel_writer.py` — one entry per column, so
reordering or renaming a column is a one-line change there.

## How the PDF is read

The quotation is an Excel export, so every page carries the same ruled grid.
`mondiale_parser.py` assigns each word to a column by its x-position against
that grid rather than by splitting text, which keeps sparse rows (lots of
dashes) aligned.

Two PDF libraries are used deliberately:

- **pdfium** (`pypdfium2`) reads the rate tables. It decodes glyphs in C, where
  pdfminer allocates a Python object per glyph — about 20x slower on 324,000
  glyphs. Word boundaries come from a 1.0pt gap threshold, calibrated so that
  pdfium and pdfplumber agree exactly on the tables.
- **pdfplumber** reads the notes blocks. These are prose, and pdfium
  synthesises spaces from glyph gaps, which lands mid-word often enough to
  corrupt the text (`per` becoming `pe r`). Notes appear on roughly 8 of the 86
  pages, so the cost is small.

Swapping the tables onto pdfium cut a full quotation from 56s to 5s with
byte-identical output — verified cell-for-cell against the previous result.

Two details are worth knowing:

- **Security Fee** is one cell holding two overlapping runs — a bold currency
  label and a regular-weight amount — plus stray hyphens bleeding in from
  Excel's hidden spacer columns. The currency is therefore read from bold
  characters only.
- **Notes** are printed once per region (Australia, North Asia, South East Asia,
  Europe, Mediterranean, Middle East) beneath that region's last rate row, and
  can spill onto following pages. Each region's notes are applied to all of its
  rate rows.

## Files

| File | Purpose |
| --- | --- |
| `app.py` | Flask app: `/` serves the UI, `POST /api/process` does the work |
| `mondiale_parser.py` | PDF → rate rows + per-region notes |
| `excel_writer.py` | rate rows → the `FCL Rates` workbook |
| `check_rows.py` | cross-checks the parser against a naive text scan, page by page |
| `templates/`, `static/` | the upload interface |
| `Procfile` | gunicorn command for deployment (Render, Heroku) |

## Verifying a new quotation

```bash
python check_rows.py "path/to/quotation.pdf"
```

This counts rate rows two independent ways — by grid geometry and by scanning
text for validity-date pairs — and reports any page where they disagree. On the
September 2026 quotation both agree on all 86 pages (1,689 rows).
