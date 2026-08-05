"""
Super Filter — SARI Survey Results Processor
=============================================
Reads the raw SARI survey Excel export and produces a clean, grouped output
where each row represents one organisation × section × question × answer.

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

# Column indices in the SOURCE Excel (1-based, matching the raw file).
# These map the raw column positions to internal keys.
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

# ── Output column definitions ─────────────────────────────────────────────
# Each entry: (header_label, internal_key, aggregator)
#
# aggregator types:
#   "single"   — one value per org (taken from first row; all rows should match)
#   "list"     — multiple values per org, joined with " | "
#   "per_answer"— one value per (org, section, question, answer) row
#
# To EXCLUDE a column, comment out its line with a leading #.

OUTPUT_COLUMNS = [
    # ── Org-level (single value) ──
    ("Organisation Name",     "organisation_name",   "single"),
    ("Parent Company",        "parent_company",      "single"),
    ("Organisation Type",     "organisation_type",   "single"),
    ("Organisation Size",     "organisation_size",   "single"),
    ("Stakeholder Category",  "stakeholder_category","single"),
    ("PDCS Sector",           "pcds_sector",         "single"),
    ("District",              "district",            "single"),
    ("Part of Group",         "part_of_group",       "single"),

    # ── Org-level (aggregated list) ──
    ("Role Level",            "role_level",          "list"),
    ("Department",            "department",          "list"),
    ("Age Band",              "age_band",            "list"),
    ("Job Title",             "job_title",           "list"),

    # ── Per-answer columns ──
    ("Section",               "section",             "per_answer"),
    ("Question ID",           "question_id",         "per_answer"),
    ("Question",              "question",            "per_answer"),
    ("Answer",                "answer",              "per_answer"),
    ("Answer Value",          "answer_value",        "per_answer"),
    ("Answer Score",          "answer_score",        "per_answer"),
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


def build_org_data(rows: list[dict]) -> dict:
    """
    Group all rows by organisation_name.
    Returns: dict[org_name] -> {
        "single": {key: value},       # org-level single-value fields
        "list":   {key: set()},       # org-level multi-value fields
        "answers": [(section, qid, question, answer, answer_value, answer_score), ...]
    }
    """
    orgs: dict[str, dict] = defaultdict(lambda: {
        "single": {},
        "list": defaultdict(set),
        "answers": [],
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

        # Per-answer data
        o["answers"].append((
            row.get("section", ""),
            row.get("question_id", ""),
            row.get("question", ""),
            row.get("answer", ""),
            row.get("answer_value", ""),
            row.get("answer_score", ""),
        ))

    return orgs


def style_header(ws, num_cols: int):
    """Apply styling to the header row."""
    header_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border


def style_data_rows(ws, start_row: int, end_row: int, num_cols: int):
    """Apply alternating row colors and borders to data rows."""
    light_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    data_align = Alignment(vertical="top", wrap_text=True)

    for r in range(start_row, end_row + 1):
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.alignment = data_align
            if (r - start_row) % 2 == 1:
                cell.fill = light_fill


def write_output(orgs: dict, output_path: str):
    """Write the processed data to a new Excel file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Processed Results"

    # ── Write header ──
    headers = [col[0] for col in OUTPUT_COLUMNS]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)

    # ── Write data rows ──
    row_num = 2
    for org_name in sorted(orgs.keys()):
        o = orgs[org_name]

        # Build the "list" values as joined strings
        list_values = {}
        for key in ["role_level", "department", "age_band", "job_title"]:
            vals = sorted(o["list"].get(key, set()))
            list_values[key] = " | ".join(vals) if vals else ""

        # Write one row per answer
        for ans in o["answers"]:
            section, qid, question, answer, answer_value, answer_score = ans

            for col_idx, (_, key, agg_type) in enumerate(OUTPUT_COLUMNS, 1):
                if agg_type == "single":
                    val = o["single"].get(key, "")
                elif agg_type == "list":
                    val = list_values.get(key, "")
                elif agg_type == "per_answer":
                    mapping = {
                        "section": section,
                        "question_id": qid,
                        "question": question,
                        "answer": answer,
                        "answer_value": answer_value,
                        "answer_score": answer_score,
                    }
                    val = mapping.get(key, "")
                else:
                    val = ""

                ws.cell(row=row_num, column=col_idx, value=val)

            row_num += 1

    # ── Style ──
    num_cols = len(OUTPUT_COLUMNS)
    style_header(ws, num_cols)
    if row_num > 2:
        style_data_rows(ws, 2, row_num - 1, num_cols)

    # ── Auto-fit column widths ──
    for col_idx in range(1, num_cols + 1):
        max_width = len(str(ws.cell(row=1, column=col_idx).value or ""))
        for r in range(2, min(row_num, 200)):  # sample first 200 rows for speed
            cell_val = str(ws.cell(row=r, column=col_idx).value or "")
            max_width = max(max_width, len(cell_val))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_width + 4, 60)

    # ── Freeze top row ──
    ws.freeze_panes = "A2"

    # ── Add auto-filter ──
    ws.auto_filter.ref = f"A1:{get_column_letter(num_cols)}{row_num - 1}"

    wb.save(output_path)
    print(f"✅ Done! Output written to: {output_path}")
    print(f"   {row_num - 2} data rows across {len(orgs)} organisations")


def main():
    input_path = Path(INPUT_FILE)
    if not input_path.exists():
        print(f"❌ Input file not found: {INPUT_FILE}")
        print("   Update INPUT_FILE in the CONFIG section of this script.")
        sys.exit(1)

    print(f"📂 Reading: {INPUT_FILE}")
    rows = read_raw_data(str(input_path))
    print(f"   {len(rows)} raw rows loaded")

    print("🔧 Processing...")
    orgs = build_org_data(rows)
    print(f"   {len(orgs)} organisations found")

    write_output(orgs, OUTPUT_FILE)


if __name__ == "__main__":
    main()
