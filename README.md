# Super Filter — SARI Organisation Statistics Generator

Reads a SARI survey Excel export (Answers + Scores sheets) and produces a **10-sheet interactive workbook** with dashboard, organisation summary, section/question breakdowns, and priority analysis.

## Quick Start

**No terminal needed** — double-click `run_app.command` (macOS) or `run_app.bat` (Windows). The first run sets up its own environment in about a minute; every run after that opens straight away. Python 3 must be installed once, from [python.org](https://python.org).

From a terminal instead:

```bash
python3 -m venv venv
source venv/bin/activate        # macOS / Linux
# venv\Scripts\activate         # Windows
pip install -r requirements.txt

python app.py                   # GUI
python super_filter.py IN.xlsx -o OUT.xlsx   # CLI
```

## How it is put together

Two stages, deliberately separate, so the table on screen and the sheet in the file can never disagree:

| | | |
|---|---|---|
| `compute(input)` | ~0.5 s | reads the export, returns every derived table as plain Python rows |
| `build(input, template, output)` | ~70 s | writes those rows into `report_template.xlsx` |

The GUI calls `compute()` to draw the Organisation Summary and hands the *same* result to `build()` on export. Nothing is calculated twice.

`report_template.xlsx` owns all presentation — fonts, fills, column widths, page setup, the Dashboard chart and every formula. The script only replaces data and repairs the organisation dropdowns. **To change how the report looks, edit the template in Excel, not the Python.**

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

- **Drag and drop** a `.xlsx` onto the window, or click the drop zone to browse
- **Organisation Summary table** — every column, sortable by clicking any heading
- **Organisation dropdown** — jump to one organisation
- **Search** — free-text filter on organisation name
- **Colour-coded rows** — green (overall ≥ 75%), orange (≥ 50%), red (below)
- **Export** — writes the full 10-sheet workbook on a background thread, with progress, so the window stays responsive

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

---

# Building a standalone app

For users who should not have to install Python at all.

## The one rule

**A build only ever produces an app for the machine it was built on.** PyInstaller does not cross-compile. A Mac makes a `.app`; a Windows PC makes a `.exe`; neither can make the other. One codebase, two builds.

| Build machine | Command | You get |
|---|---|---|
| Windows | `pyinstaller app.spec --noconfirm` | `dist\SuperFilter\SuperFilter.exe` |
| macOS | `pyinstaller app.spec --noconfirm` | `dist/Super Filter.app` |

If you have no Windows machine, use the GitHub Actions recipe further down — it builds the `.exe` on a free Windows runner.

## Creating the `.exe` on Windows

```bat
git clone https://github.com/dev-kir/temp_intern_filter.git
cd temp_intern_filter

py -3 -m venv venv
venv\Scripts\python -m pip install --upgrade pip
venv\Scripts\python -m pip install -r requirements.txt pyinstaller

venv\Scripts\pyinstaller app.spec --noconfirm
```

The result is the folder `dist\SuperFilter\`. Inside it, `SuperFilter.exe` is what people double-click.

**Ship the whole folder, not just the .exe.** The `.exe` needs the DLLs and `report_template.xlsx` sitting next to it. Zip `dist\SuperFilter\` and send that.

### What `app.spec` is doing, and why it matters

```python
datas = [('report_template.xlsx', '.')]     # ← without this, export fails at runtime
hiddenimports = ['super_filter']
```

`report_template.xlsx` is data, not code, so PyInstaller cannot discover it by following imports. Leave that line out and the app builds perfectly, launches perfectly, and then throws `FileNotFoundError` the moment someone clicks Export. `super_filter.resource_path()` looks in `sys._MEIPASS` first for exactly this reason.

The spec uses **onedir**, not `--onefile`. A onefile build unpacks ~25 MB to a temp directory on every launch, which looks like the app is hanging. A folder starts instantly.

Do **not** pass `--icon icon.svg`. PyInstaller needs `.ico` on Windows and `.icns` on macOS; an SVG makes the build fail. Convert it first, then set `icon=` in `app.spec`.

### First-run warning

The build is unsigned, so:

- **Windows** — SmartScreen says "Windows protected your PC". Click *More info* → *Run anyway*.
- **macOS** — Gatekeeper refuses a plain double-click. Right-click the app → *Open* → *Open*. Once only.

Signing removes this and costs money (an Authenticode certificate on Windows, an Apple Developer account on macOS). Not worth it for an internal tool.

## Creating the `.exe` without a Windows machine

Commit this as `.github/workflows/build.yml`. Push a tag like `v1.0.0` and GitHub builds both halves and attaches them to a release.

```yaml
name: build

on:
  push:
    tags: ['v*']
  workflow_dispatch:

jobs:
  build:
    strategy:
      matrix:
        include:
          - os: windows-latest
            artifact: SuperFilter-windows
          - os: macos-latest
            artifact: SuperFilter-macos
    runs-on: ${{ matrix.os }}
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'
      - run: pip install -r requirements.txt pyinstaller
      - run: pyinstaller app.spec --noconfirm
      - uses: actions/upload-artifact@v4
        with:
          name: ${{ matrix.artifact }}
          path: dist/
```

Free for public repositories. The Windows `.exe` appears under the workflow run's Artifacts.

## Which option to choose

| | Folder + launcher | Built app | 
|---|---|---|
| Ready now | yes | one build away |
| User installs Python | yes, once | no |
| You can produce it on a Mac | both halves | Mac half only |
| Maintenance | none | rebuild per release |

Start with the launchers. Build real apps only when the tool leaves the people who wrote it.

---

## Requirements

- Python 3.8+
- `openpyxl` — reading and writing workbooks
- `tkinterdnd2` — drag-and-drop; **optional**, the app falls back to click-to-browse without it

pandas is no longer used.

## Project Structure

```
ammar_super_filter/
├── super_filter.py           # compute() + build() — all the logic
├── app.py                    # Tkinter GUI
├── app.spec                  # PyInstaller build recipe
├── report_template.xlsx      # presentation layer — styling, charts, formulas
├── run_app.command           # macOS double-click launcher
├── run_app.bat               # Windows double-click launcher
├── icon.svg                  # app icon (convert to .ico/.icns to use in a build)
├── requirements.txt
├── README.md
├── ALGORITHM.md              # full algorithm docs
└── Super_Filter.md           # original design brief
```

Survey exports and generated workbooks are **not** committed. They contain respondent names and job titles; keep them out of the repository.
