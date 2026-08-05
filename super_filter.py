"""
Super Filter — SARI Survey Results Processor (Pivot Edition)
=============================================================
Reads the raw SARI survey Excel export and produces a pivot table where:
  - Each ROW = one organisation (126 rows)
  - Each QUESTION becomes its own column, grouped by section
  - When multiple participants give different answers, they are joined with " | "

Usage:
    python super_filter.py

Configuration:
    Edit the CONFIG section below to change input/output paths and toggle columns.
    Comment out any column in OUTPUT_COLUMNS to exclude it from the output.
"""

import sys
from pathlib import Path
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG — change these to suit your needs
# ═══════════════════════════════════════════════════════════════════════════

INPUT_FILE = "SARI_Results_2026-08-05-01-54-15.xlsx"
OUTPUT_FILE = "SARI_Results_Processed.xlsx"

# What to display in each question cell.
# Options: "answer" (the text), "answer_value" (the code), "answer_score" (the number)
QUESTION_CELL_CONTENT = "answer"

# Column indices in the SOURCE Excel (1-based, matching the raw file).
RAW_COL_MAP = {
    1:  "respondent_id",
    2:  "submitted_at",
    3:  "section",
    4:  "question_num",
    5:  "question_id",
    6:  "question",
    7:  "answer",
    8:  "answer_value",
    9:  "answer_score",
    10: "max_score",
    11: "participant_name",
    12: "email",
    13: "job_title",
    14: "organisation_name",
    15: "organisation_type",
    16: "organisation_size",
    17: "stakeholder_category",
    18: "pcds_sector",
    19: "district",
    20: "role_level",
    21: "department",
    22: "age_band",
    23: "part_of_group",
    24: "parent_company",
}

# ── Output column definitions (org-level, left side of pivot) ──────────────
# Each entry: (header_label, internal_key, aggregator)
#
# aggregator types:
#   "single" — one value per org (taken from first row; all rows should match)
#   "list"   — multiple values per org, joined with " | "
#
# To EXCLUDE a column, comment out its line with a leading #.

OUTPUT_COLUMNS = [
    ("Organisation Name",     "organisation_name",   "single"),
    ("Parent Company",        "parent_company",      "single"),
    ("Organisation Type",     "organisation_type",   "single"),
    ("Organisation Size",     "organisation_size",   "single"),
    ("Stakeholder Category",  "stakeholder_category","single"),
    ("PDCS Sector",           "pcds_sector",         "single"),
    ("District",              "district",            "single"),
    ("Part of Group",         "part_of_group",       "single"),
    ("Role Level",            "role_level",          "list"),
    ("Department",            "department",          "list"),
    ("Age Band",              "age_band",            "list"),
    ("Job Title",             "job_title",           "list"),
]

# ── Section ordering (defines the left-to-right order of question columns) ─
# Sections not listed here will appear at the end in alphabetical order.
SECTION_ORDER = [
    "Background",
    "Strategy & Leadership",
    "Talent & Organisational Culture",
    "Data Management & Readiness",
    "Infrastructure & Technology",
    "Governance, Policy & Ethics",
    "Investment",
    "AI Implementation & Potential Impact",
    "Latar Belakang",
    "Strategi & Kepimpinan",
    "Bakat & Budaya Organisasi",
    "Pengurusan Data & Kesiapsiagaan",
    "Infrastruktur & Teknologi",
    "Tadbir Urus, Dasar & Etika",
    "Pelaburan",
    "Pelaksanaan AI & Impak",
]

# ═══════════════════════════════════════════════════════════════════════════
# PROCESSING LOGIC — you shouldn't need to edit below this line
# ═══════════════════════════════════════════════════════════════════════════


def read_raw_data(filepath: str) -> list[dict]:
    """Read the source Excel and return a list of row dicts."""
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active

    rows = []
    for row_cells in ws.iter_rows(min_row=2, values_only=True):
        record = {}
        for col_idx, key in RAW_COL_MAP.items():
            val = row_cells[col_idx - 1]
            record[key] = str(val).strip() if val is not None else ""
        rows.append(record)

    wb.close()
    return rows


def build_question_order(rows: list[dict]) -> list[tuple]:
    """
    Determine the ordered list of (section, question_id, question_num, question_text).
    Uses (section, question_id) as the unique key since the same qid can appear
    in both English and BM sections.
    """
    seen = set()
    questions = []

    for row in rows:
        sec = row["section"]
        qid = row["question_id"]
        key = (sec, qid)
        if key not in seen:
            seen.add(key)
            questions.append((
                sec,
                qid,
                int(row["question_num"]) if row["question_num"].isdigit() else 0,
                row["question"],
            ))

    def sort_key(item):
        sec, qid, qnum, _ = item
        sec_idx = SECTION_ORDER.index(sec) if sec in SECTION_ORDER else len(SECTION_ORDER)
        return (sec_idx, qnum)

    questions.sort(key=sort_key)
    return questions


def build_org_data(rows: list[dict]) -> dict:
    """
    Group all rows by organisation_name.
    Returns: dict[org_name] -> {
        "single": {key: value},
        "list":   {key: set()},
        "answers": {question_id: set of answer strings},
    }
    """
    orgs: dict[str, dict] = defaultdict(lambda: {
        "single": {},
        "list": defaultdict(set),
        "answers": defaultdict(set),
    })

    for row in rows:
        org = row["organisation_name"]
        if not org:
            continue

        o = orgs[org]

        # Single-value fields
        for key in ["organisation_name", "parent_company", "organisation_type",
                     "organisation_size", "stakeholder_category", "pcds_sector",
                     "district", "part_of_group"]:
            if key not in o["single"]:
                o["single"][key] = row.get(key, "")

        # List fields
        for key in ["role_level", "department", "age_band", "job_title"]:
            val = row.get(key, "")
            if val:
                o["list"][key].add(val)

        # Per-question answers (keyed by section + question_id)
        sec = row["section"]
        qid = row["question_id"]
        content = row.get(QUESTION_CELL_CONTENT, "")
        if content:
            o["answers"][(sec, qid)].add(content)

    return orgs


def style_sheet(ws, num_cols: int, num_rows: int, section_spans: list[tuple]):
    """Apply styling: header, section group headers, alternating rows, borders."""
    header_font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    section_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    light_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_align = Alignment(vertical="top", wrap_text=True)

    # Row 1: Section group headers (merged cells)
    for sec_name, start_col, end_col in section_spans:
        if start_col == end_col:
            cell = ws.cell(row=1, column=start_col)
            cell.value = sec_name
        else:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
            cell = ws.cell(row=1, column=start_col)
            cell.value = sec_name
        cell.font = section_font
        cell.fill = section_fill
        cell.alignment = header_align
        cell.border = thin_border
        # Apply border to all cells in the merged range
        for c in range(start_col, end_col + 1):
            ws.cell(row=1, column=c).border = thin_border

    # Row 2: Question ID headers
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # Data rows (starting row 3)
    for r in range(3, num_rows + 3):
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.alignment = data_align
            if (r - 3) % 2 == 1:
                cell.fill = light_fill


def write_output(orgs: dict, questions: list[tuple], output_path: str):
    """Write the pivot table to a new Excel file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pivot"

    # ── Build column layout ──
    # Left side: org-level columns
    org_headers = [col[0] for col in OUTPUT_COLUMNS]
    org_col_count = len(OUTPUT_COLUMNS)

    # Right side: one column per question, grouped by section
    # section_spans: list of (section_name, start_col, end_col)
    section_spans = []
    current_section = None
    section_start = None
    question_columns = []  # list of (section, qid, question_text)

    for sec, qid, qnum, qtext in questions:
        question_columns.append((sec, qid, qtext))
        col_idx = org_col_count + len(question_columns)  # 1-based

        if sec != current_section:
            if current_section is not None:
                section_spans.append((current_section, section_start, col_idx - 1))
            current_section = sec
            section_start = col_idx

    # Don't forget the last section
    if current_section is not None:
        section_spans.append((current_section, section_start, org_col_count + len(question_columns)))

    total_cols = org_col_count + len(question_columns)

    # ── Write Row 1: Section group headers ──
    # For org-level columns, merge them into one "Organisation Info" header
    if org_col_count > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=org_col_count)
        cell = ws.cell(row=1, column=1)
        cell.value = "Organisation Info"
    elif org_col_count == 1:
        cell = ws.cell(row=1, column=1)
        cell.value = org_headers[0]

    # Section headers for question columns
    for sec_name, start_col, end_col in section_spans:
        if start_col == end_col:
            cell = ws.cell(row=1, column=start_col)
            cell.value = sec_name
        else:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
            cell = ws.cell(row=1, column=start_col)
            cell.value = sec_name

    # ── Write Row 2: Column headers ──
    for col_idx, header in enumerate(org_headers, 1):
        ws.cell(row=2, column=col_idx, value=header)

    for i, (sec, qid, qtext) in enumerate(question_columns):
        col_idx = org_col_count + i + 1
        ws.cell(row=2, column=col_idx, value=qid)

    # ── Write Row 3: Question text (as a reference row, hidden by default) ──
    # Actually, let's put question text as comments or a separate sheet.
    # For now, Row 3+ is data.

    # ── Write data rows (starting row 3) ──
    row_num = 3
    for org_name in sorted(orgs.keys()):
        o = orgs[org_name]

        # Build list values
        list_values = {}
        for key in ["role_level", "department", "age_band", "job_title"]:
            vals = sorted(o["list"].get(key, set()))
            list_values[key] = " | ".join(vals) if vals else ""

        # Write org-level columns
        for col_idx, (_, key, agg_type) in enumerate(OUTPUT_COLUMNS, 1):
            if agg_type == "single":
                val = o["single"].get(key, "")
            elif agg_type == "list":
                val = list_values.get(key, "")
            else:
                val = ""
            ws.cell(row=row_num, column=col_idx, value=val)

        # Write question columns
        for i, (sec, qid, qtext) in enumerate(question_columns):
            col_idx = org_col_count + i + 1
            answers = o["answers"].get((sec, qid), set())
            val = " | ".join(sorted(answers)) if answers else ""
            ws.cell(row=row_num, column=col_idx, value=val)

        row_num += 1

    # ── Style ──
    style_sheet(ws, total_cols, len(orgs), section_spans)

    # ── Freeze panes (freeze org columns + header rows) ──
    ws.freeze_panes = ws.cell(row=3, column=org_col_count + 1)

    # ── Auto-filter ──
    ws.auto_filter.ref = f"A2:{get_column_letter(total_cols)}{row_num - 1}"

    # ── Auto-fit column widths (sample first 200 rows) ──
    for col_idx in range(1, total_cols + 1):
        max_width = len(str(ws.cell(row=2, column=col_idx).value or ""))
        for r in range(3, min(row_num, 203)):
            cell_val = str(ws.cell(row=r, column=col_idx).value or "")
            max_width = max(max_width, len(cell_val))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_width + 3, 50)

    # ── Create a "Questions" reference sheet ──
    ws2 = wb.create_sheet("Question Reference")
    ws2.append(["Section", "Question ID", "Question #", "Question Text"])
    for sec, qid, qnum, qtext in questions:
        ws2.append([sec, qid, qnum, qtext])
    # Style the reference sheet
    for col in range(1, 5):
        cell = ws2.cell(row=1, column=col)
        cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
        cell.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2.column_dimensions["A"].width = 40
    ws2.column_dimensions["B"].width = 22
    ws2.column_dimensions["C"].width = 12
    ws2.column_dimensions["D"].width = 100
    ws2.freeze_panes = "A2"

    wb.save(output_path)
    print(f"✅ Done! Output written to: {output_path}")
    print(f"   {len(orgs)} organisations × {len(question_columns)} questions")
    print(f"   {len(section_spans)} sections as column groups")
    print(f"   See 'Question Reference' sheet for question_id → question text mapping")


def main():
    input_path = Path(INPUT_FILE)
    if not input_path.exists():
        print(f"❌ Input file not found: {INPUT_FILE}")
        print("   Update INPUT_FILE in the CONFIG section of this script.")
        sys.exit(1)

    print(f"📂 Reading: {INPUT_FILE}")
    rows = read_raw_data(str(input_path))
    print(f"   {len(rows)} raw rows loaded")

    print("🔧 Building question order...")
    questions = build_question_order(rows)
    print(f"   {len(questions)} unique questions across {len(set(q[0] for q in questions))} sections")

    print("🔧 Grouping by organisation...")
    orgs = build_org_data(rows)
    print(f"   {len(orgs)} organisations found")

    write_output(orgs, questions, OUTPUT_FILE)


if __name__ == "__main__":
    main()
