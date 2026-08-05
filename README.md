# Super Filter — SARI Organisation Statistics Generator

Reads a SARI survey Excel export and produces a **10-sheet interactive workbook** matching the `SARI_Organisation.xlsx` format.

## Output Sheets

| # | Sheet | Purpose |
|---|---|---|
| 1 | **Read Me** | Info & purpose |
| 2 | **Lists** | All organisation names (for dropdowns) |
| 3 | **Dashboard** | Interactive org selector with live formulas |
| 4 | **Organisation Report** | Per-org report with priority questions |
| 5 | **Organisation Summary** | One row per org — scores, maturity tier, agreement |
| 6 | **Section Summary** | Per org × section — avg/median/min/max/normalised |
| 7 | **Question Summary** | Per org × question — consensus, std dev, review flags |
| 8 | **Answer Distribution** | Per org × question × answer option — counts & % |
| 9 | **Raw Answers** | Raw data with Standard section (BM→EN merged) |
| 10 | **Priority Detail** | Priority-ranked questions per org (low score first) |

## Quick Start

```bash
# Setup
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Run
python super_filter.py
```

**Config:** Edit `INPUT_FILE` and `OUTPUT_FILE` at the top of `super_filter.py`.

## Key Metrics

| Metric | Description | Where |
|---|---|---|
| **Average score** | Raw 0–4 average across all scored questions | Org Summary col I |
| **Overall score** | Normalised 0–1 (avg_score / 4.0) | Org Summary col J |
| **Maturity tier** | AI Aware → Explorer → Follower → Leader → Pioneer | Org Summary col Q |
| **Consensus** | Proportion agreeing on most common answer | Question Summary col I |
| **Agreement** | High (≥0.8) / Moderate (≥0.6) / Low / Not measurable | Multiple sheets |
| **Review flag** | Low consensus or high std dev — needs attention | Question Summary col R |
| **Priority rank** | Questions sorted by lowest normalised score first | Priority Detail |

## Maturity Tiers

| Overall Score | Tier |
|---|---|
| 0.0 – 0.2 | AI Aware - 0 |
| 0.2 – 0.4 | AI Explorer - 1 |
| 0.4 – 0.6 | AI Follower - 2 |
| 0.6 – 0.8 | AI Leader - 3 |
| 0.8 – 1.0 | AI Pioneer - 4 |

## Interactive Features

- **Dashboard:** Change cell B4 to any org name — all stats update via formulas
- **Organisation Report:** Change cell B3 to any org name — full report updates
- **All data sheets** have auto-filters and frozen headers
- **Score columns** have red→yellow→green color scales
- **Maturity tier & distance** are Excel formulas (recalculate on org change)

## BM→EN Merge

Bahasa Malaysia sections are mapped to English equivalents:

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

> For full algorithm details, formulas, and data flow diagrams, see **[ALGORITHM.md](ALGORITHM.md)**.

## Requirements

- Python 3.8+
- openpyxl (`pip install openpyxl`)

## Project Structure

```
ammar_super_filter/
├── super_filter.py                          # CLI processing script (10-sheet output)
├── app.py                                   # Desktop GUI app (Tkinter)
├── requirements.txt                         # Python dependencies
├── README.md                                # This file
├── ALGORITHM.md                             # Full algorithm docs
├── Super_Filter.md                          # Original design brief
├── .gitignore
├── venv/
├── SARI_Results_*.xlsx                       # Input files
└── SARI_Organisation.xlsx                   # Output file
```
