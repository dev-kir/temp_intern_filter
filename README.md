# Super Filter — SARI Organisation Statistics Generator

Reads a SARI survey Excel export (Answers + Scores sheets) and produces a **10-sheet interactive workbook** with dashboard, organisation summary, section/question breakdowns, and priority analysis.

## Quick Start

```bash
# Setup
python3 -m venv venv
source venv/bin/activate        # macOS / Linux
# venv\Scripts\activate         # Windows
pip install -r requirements.txt

# CLI — generate the full workbook
python super_filter.py

# GUI — interactive organisation summary viewer
python app.py
```

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

## GUI Features

- **Open Excel** — file picker button
- **Full Organisation Summary table** — scrollable, sortable by any column
- **Filter** — type to search by organisation name
- **Color-coded rows** — green (high score), orange (mid), red (low)
- **Export** — generates the complete 10-sheet workbook

## Maturity Tiers

| Overall Score | Tier |
|---|---|
| 0.0 – 0.2 | AI Aware - 0 |
| 0.2 – 0.4 | AI Explorer - 1 |
| 0.4 – 0.6 | AI Follower - 2 |
| 0.6 – 0.8 | AI Leader - 3 |
| 0.8 – 1.0 | AI Pioneer - 4 |

## BM→EN Merge

Bahasa Malaysia sections are mapped to English via Question ID prefix:

| BM Section | → | EN Section |
|---|---|---|
| Latar Belakang | → | Background |
| Strategi & Kepimpinan | → | Strategy & Leadership |
| Bakat & Budaya Organisasi | → | Talent & Organisational Culture |
| Pengurusan Data & Kesiapsiagaan | → | Data Management & Readiness |
| Infrastruktur & Teknologi | → | Infrastructure & Technology |
| Tadbir Urus, Dasar & Etika | → | Governance, Policy & Ethics |
| Pelaburan | → | Investment |
| Pelaksanaan AI & Impak | → | AI Implementation & Potential Impact |

## Building a Standalone App (no Python required)

To create a double-clickable `.app` (macOS) or `.exe` (Windows) that runs without Python:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "Super Filter" --icon icon.svg app.py
```

The output will be in `dist/Super Filter.app` (macOS) or `dist/Super Filter.exe` (Windows).

## Requirements

- Python 3.8+
- openpyxl, pandas, tkinterdnd2 (for drag-drop)

## Project Structure

```
ammar_super_filter/
├── super_filter.py           # CLI — generates 10-sheet workbook
├── app.py                    # GUI — organisation summary viewer
├── report_template.xlsx      # Excel template (presentation layer)
├── icon.svg                  # App icon
├── requirements.txt          # Python dependencies
├── README.md                 # This file
├── ALGORITHM.md              # Full algorithm docs
├── Super_Filter.md           # Original design brief
├── .gitignore
├── venv/
├── SARI_Results_*.xlsx       # Input files
└── SARI_Organisation.xlsx    # Output file
```
