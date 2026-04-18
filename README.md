# Palisades Fire Incident Reports — Data Extraction

This folder contains a Python script that scrapes all 344 Palisades Fire incident updates and compiles them into a single spreadsheet.

- Source: https://www.fire.ca.gov/incidents/2025/1/7/palisades-fire/updates
- Output: `3. palisades_fire_compiled.xlsx`

## Tools and libraries

- Python 3.10+ (https://www.python.org/)
- Selenium + Chrome WebDriver (https://www.selenium.dev/) — renders the CAL FIRE pages and exports each report as a PDF via Chrome DevTools `Page.printToPDF`.
- BeautifulSoup 4 (https://www.crummy.com/software/BeautifulSoup/) — parses the rendered HTML.
- pypdf (https://pypdf.readthedocs.io/) — reads PDF page counts.
- pandas (https://pandas.pydata.org/) and openpyxl (https://openpyxl.readthedocs.io/) — build the DataFrame and write it into the Excel spreadsheet.

## How the scraper works

1. Open the `/updates` index with Chrome, wait for the list of reports to render, and collect every report-detail URL.
2. For each report:
   - Load the page, wait for `main.main-content`, and parse the rendered HTML with BeautifulSoup.
   - Walk `<dl>` key/value sections and every `<div class="p-3">` heading block, mapping each section title to its template column via a normalised label table. Inline sub-labels (e.g. `Westside Location:`, `Livestream:`) are handled by a separate inline-section mapping.
   - Apply targeted text-range fallbacks for columns whose content isn't wrapped in a clean heading.
   - Trim each field at known downstream section markers so content never leaks across columns.
   - Export the rendered page as a PDF (`Page.printToPDF`), compute its page count with `pypdf`, and store the PDF file name and page count alongside the report data.
3. After all reports are parsed, sort them chronologically by `Report_Date` + `Report_Time`, assign `Report_Sequence` as `001`..`344`, and write the rows into `3. palisades_fire_compiled.xlsx`.

## Extra credit (Rule 4)

The following mechanism is implemented: any section heading without a mapping in the schema is converted into a new column using its title, and each such occurrence is appended to `NEW_COLUMNS_README.md` together with its source URL and section title (see `map_or_track_new_column` in `project3.py`). After running the script, no new columns were added — every section across all 344 reports maps onto a column that already exists in the template, so `NEW_COLUMNS_README.md` is not produced.

## Prerequisites

- Python 3.10+
- Google Chrome (recent stable version)

## Install and run

```bash
pip3 install -r requirements.txt
python3 project3.py
```

On success the script writes `3. palisades_fire_compiled.xlsx` with 344 rows.

## Screenshots

See the `screenshots/` folder for captures of the scraping process.
