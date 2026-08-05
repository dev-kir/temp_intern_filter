# Super Filter — SARI Organisation Statistics Generator

Reads a SARI survey Excel export (Answers + Scores sheets) and produces a **10-sheet interactive workbook** with dashboard, organisation summary, section/question breakdowns, and priority analysis.

## Download

**Ready-to-run builds** (no Python install needed):

- **macOS:** [Super Filter.zip](https://github.com/dev-kir/temp_intern_filter/raw/main/Super%20Filter.zip) — download, unzip, double-click `Super Filter.app`
- **Windows:** Run `build.sh` on a Mac to produce the `.app`, or build locally on Windows with PyInstaller:

  ```bash
  pip install pyinstaller
  pyinstaller --windowed --onedir --name "Super Filter" --icon icon.png --add-data "report_template.xlsx:." --hidden-import=tkinterdnd2 app.py
  ```

## Quick Start (from source)

```bash
python3 -m venv venv
source venv/bin/activate        # macOS / Linux
# venv\Scripts\activate         # Windows
pip install -r requirements.txt

python app.py                   # GUI
python super_filter.py          # CLI (needs input/output paths)
```

For the CLI, edit `INPUT_FILE` and `OUTPUT_FILE` at the top of `super_filter.py`.

## GUI Features

- **Drag & drop** an `.xlsx` file onto the window (requires `tkinterdnd2`)
- **Full Organisation Summary table** — scrollable, sortable by clicking any column header
- **Search** by organisation name
- **Colour-coded Overall score** — green (≥75%), amber (≥50%), red (<50%); all other columns stay default text colour
- **Filter by organisation** — or "All organisations" to see everything
- **Export Full Report** — generates the complete 10-sheet workbook

## Output Sheets

| # | Sheet | Purpose |
|---|---|---|
| 1 | **Read Me** | Info & methodology |
| 2 | **Lists** | Organisation names (hidden, drives dropdowns) |
| 3 | **Dashboard** | Interactive org selector with KPI cards, section table, bar chart |
| 4 | **Organisation Report** | Printable per-org report with priority questions |
| 5 | **Organisation Summary** | One row per org — scores, maturity tier, agreement |
| 6 | **Section Summary** | Per org × section — avg/median/min/max/normalised |
| 7 | **Question Summary** | Per org × question — consensus, std dev, review flags |
| 8 | **Answer Distribution** | Per org × question × answer option — counts & % |
| 9 | **Raw Answers** | Raw data with Standard section (BM→EN merged), email excluded |
| 10 | **Priority Detail** | Priority-ranked questions per org (hidden, drives formulas) |

## Maturity Tiers

| Overall Score | Tier |
|---|---|
| 0.0 – 0.2 | AI Aware - 0 |
| 0.2 – 0.4 | AI Explorer - 1 |
| 0.4 – 0.6 | AI Follower - 2 |
| 0.6 – 0.8 | AI Leader - 3 |
| 0.8 – 1.0 | AI Pioneer - 4 |

## Requirements

- Python 3.8+
- openpyxl, pandas, tkinterdnd2 (for drag-drop)

## Project Structure

```
ammar_super_filter/
├── super_filter.py           # CLI — generates 10-sheet workbook
├── app.py                    # GUI — organisation summary viewer
├── report_template.xlsx      # Excel template (presentation layer)
├── icon.svg                  # App icon (vector)
├── icon.png                  # App icon (raster, used by PyInstaller)
├── build.sh                  # Build standalone .app
├── requirements.txt
├── README.md
├── Super Filter.zip          # Built macOS app (download & run)
├── Super Filter.app/         # Built macOS app (unzipped)
├── venv/
├── SARI_Results_*.xlsx       # Input files
└── SARI_Organisation.xlsx    # Output file
```
