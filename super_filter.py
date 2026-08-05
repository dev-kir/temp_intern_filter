"""
Super Filter — SARI Survey Results Processor (Pivot + Scorecard)
==================================================================
Reads the raw SARI survey Excel export and produces:

  Sheet 1 "Pivot"      — One row per org, questions as columns (text answers)
  Sheet 2 "Scorecard"  — One row per org, per-section average scores + overall
  Sheet 3 "Question Ref" — Maps question IDs to full question text

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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG — change these to suit your needs
# ═══════════════════════════════════════════════════════════════════════════

INPUT_FILE = "SARI_Results_2026-08-05-01-54-15.xlsx"
OUTPUT_FILE = "SARI_Results_Processed.xlsx"

# What to display in Pivot question cells.
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
# Only English sections — BM answers are merged into their EN counterparts.
SECTION_ORDER = [
    "Background",
    "Strategy & Leadership",
    "Talent & Organisational Culture",
    "Data Management & Readiness",
    "Infrastructure & Technology",
    "Governance, Policy & Ethics",
    "Investment",
    "AI Implementation & Potential Impact",
]

# ── BM → EN section mapping (merge BM answers into EN columns) ─
BM_TO_EN_SECTION = {
    "Latar Belakang":                  "Background",
    "Strategi & Kepimpinan":           "Strategy & Leadership",
    "Bakat & Budaya Organisasi":       "Talent & Organisational Culture",
    "Pengurusan Data & Kesiapsiagaan": "Data Management & Readiness",
    "Infrastruktur & Teknologi":       "Infrastructure & Technology",
    "Tadbir Urus, Dasar & Etika":      "Governance, Policy & Ethics",
    "Pelaburan":                       "Investment",
    "Pelaksanaan AI & Impak":          "AI Implementation & Potential Impact",
}

# ── Scorecard: which sections to include (English only by default) ──
# Comment out sections you don't want in the scorecard.
SCORECARD_SECTIONS = [
    # "Background",                        # NOT SCORED — demographic questions only
    "Strategy & Leadership",
    "Talent & Organisational Culture",
    "Data Management & Readiness",
    "Infrastructure & Technology",
    "Governance, Policy & Ethics",
    "Investment",
    "AI Implementation & Potential Impact",
]

# ═══════════════════════════════════════════════════════════════════════════
# PROCESSING LOGIC
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
    """Return ordered list of (section, question_id, question_num, question_text).
    Only includes English sections — BM answers are merged into EN columns."""
    seen = set()
    questions = []

    for row in rows:
        sec = row["section"]
        # Only collect from English sections (BM sections are merged)
        if sec not in SECTION_ORDER:
            continue
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
        "single":  {key: value},
        "list":    {key: set()},
        "answers": {(section, qid): set of answer strings},
        "scores":  {(section, qid): list of (score, max_score)},
    }
    """
    orgs: dict[str, dict] = defaultdict(lambda: {
        "single": {},
        "list": defaultdict(set),
        "answers": defaultdict(set),
        "scores": defaultdict(list),
    })

    for row in rows:
        org = row["organisation_name"]
        if not org:
            continue

        o = orgs[org]

        for key in ["organisation_name", "parent_company", "organisation_type",
                     "organisation_size", "stakeholder_category", "pcds_sector",
                     "district", "part_of_group"]:
            if key not in o["single"]:
                o["single"][key] = row.get(key, "")

        for key in ["role_level", "department", "age_band", "job_title"]:
            val = row.get(key, "")
            if val:
                o["list"][key].add(val)

        sec = row["section"]
        qid = row["question_id"]

        # Merge BM sections into their EN counterparts
        en_sec = BM_TO_EN_SECTION.get(sec, sec)

        content = row.get(QUESTION_CELL_CONTENT, "")
        if content:
            o["answers"][(en_sec, qid)].add(content)

        score_str = row.get("answer_score", "")
        max_str = row.get("max_score", "")
        if score_str:
            try:
                score = float(score_str)
                max_s = float(max_str) if max_str else 4.0
                o["scores"][(en_sec, qid)].append((score, max_s))
            except ValueError:
                pass

    return orgs


# ═══════════════════════════════════════════════════════════════════════════
# STYLING HELPERS
# ═══════════════════════════════════════════════════════════════════════════

HEADER_FONT = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
SECTION_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
SECTION_FONT = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
LIGHT_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATA_ALIGN = Alignment(vertical="top", wrap_text=True)


def apply_header_style(ws, row, num_cols):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER


def apply_data_style(ws, start_row, end_row, num_cols):
    for r in range(start_row, end_row + 1):
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            cell.alignment = DATA_ALIGN
            if (r - start_row) % 2 == 1:
                cell.fill = LIGHT_FILL


def auto_fit_cols(ws, num_cols, org_col_count=0, max_sample=200):
    """Auto-fit column widths. Org columns get wider max; question columns get narrower."""
    for col_idx in range(1, num_cols + 1):
        header_len = len(str(ws.cell(row=1, column=col_idx).value or ""))
        header2_len = len(str(ws.cell(row=2, column=col_idx).value or ""))
        max_width = max(header_len, header2_len)

        # Sample data rows for content width
        for r in range(3, min(ws.max_row + 1, max_sample + 3)):
            cell_val = str(ws.cell(row=r, column=col_idx).value or "")
            # Cap per-cell contribution to avoid one huge cell blowing up the column
            max_width = max(max_width, min(len(cell_val), 60))

        # Different max widths for org columns vs question columns
        if col_idx <= org_col_count:
            max_allowed = 40  # org columns can be wider
        else:
            max_allowed = 28  # question columns stay narrow, text wraps

        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_width + 2, max_allowed)


# ═══════════════════════════════════════════════════════════════════════════
# PIVOT SHEET
# ═══════════════════════════════════════════════════════════════════════════

def write_pivot_sheet(wb, orgs: dict, questions: list[tuple]):
    ws = wb.active
    ws.title = "Pivot"

    org_headers = [col[0] for col in OUTPUT_COLUMNS]
    org_col_count = len(OUTPUT_COLUMNS)

    section_spans = []
    current_section = None
    section_start = None
    question_columns = []

    for sec, qid, qnum, qtext in questions:
        question_columns.append((sec, qid, qtext))
        col_idx = org_col_count + len(question_columns)

        if sec != current_section:
            if current_section is not None:
                section_spans.append((current_section, section_start, col_idx - 1))
            current_section = sec
            section_start = col_idx

    if current_section is not None:
        section_spans.append((current_section, section_start, org_col_count + len(question_columns)))

    total_cols = org_col_count + len(question_columns)

    # Row 1: Section group headers
    if org_col_count > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=org_col_count)
        cell = ws.cell(row=1, column=1)
        cell.value = "Organisation Info"
    elif org_col_count == 1:
        ws.cell(row=1, column=1).value = org_headers[0]

    for sec_name, start_col, end_col in section_spans:
        if start_col == end_col:
            cell = ws.cell(row=1, column=start_col)
            cell.value = sec_name
        else:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
            cell = ws.cell(row=1, column=start_col)
            cell.value = sec_name
        cell.font = SECTION_FONT
        cell.fill = SECTION_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER
        for c in range(start_col, end_col + 1):
            ws.cell(row=1, column=c).border = THIN_BORDER

    # Row 2: Column headers
    for col_idx, header in enumerate(org_headers, 1):
        ws.cell(row=2, column=col_idx, value=header)
    for i, (sec, qid, qtext) in enumerate(question_columns):
        ws.cell(row=2, column=org_col_count + i + 1, value=qid)

    apply_header_style(ws, 2, total_cols)

    # Data rows (starting row 3)
    row_num = 3
    for org_name in sorted(orgs.keys()):
        o = orgs[org_name]

        list_values = {}
        for key in ["role_level", "department", "age_band", "job_title"]:
            vals = sorted(o["list"].get(key, set()))
            list_values[key] = " | ".join(vals) if vals else ""

        for col_idx, (_, key, agg_type) in enumerate(OUTPUT_COLUMNS, 1):
            if agg_type == "single":
                val = o["single"].get(key, "")
            elif agg_type == "list":
                val = list_values.get(key, "")
            else:
                val = ""
            ws.cell(row=row_num, column=col_idx, value=val)

        for i, (sec, qid, qtext) in enumerate(question_columns):
            col_idx = org_col_count + i + 1
            answers = o["answers"].get((sec, qid), set())
            val = " | ".join(sorted(answers)) if answers else ""
            ws.cell(row=row_num, column=col_idx, value=val)

        row_num += 1

    apply_data_style(ws, 3, row_num - 1, total_cols)
    # Freeze only the org name column + header rows so user can scroll through questions
    ws.freeze_panes = "B3"
    ws.auto_filter.ref = f"A2:{get_column_letter(total_cols)}{row_num - 1}"
    auto_fit_cols(ws, total_cols, org_col_count)

    return question_columns


# ═══════════════════════════════════════════════════════════════════════════
# SCORECARD SHEET
# ═══════════════════════════════════════════════════════════════════════════

def compute_section_scores(orgs: dict, questions: list[tuple]) -> dict:
    """
    For each org, compute the average score per section and overall.
    Returns: dict[org_name] -> {
        "section_scores": {section_name: average_score},
        "overall_score": float,
    }
    """
    # Build mapping: section -> list of question_ids
    section_qids = defaultdict(list)
    for sec, qid, qnum, qtext in questions:
        if sec in SCORECARD_SECTIONS:
            section_qids[sec].append(qid)

    result = {}
    for org_name, o in orgs.items():
        section_scores = {}
        all_scores = []

        for sec in SCORECARD_SECTIONS:
            sec_total = 0.0
            sec_count = 0
            for qid in section_qids.get(sec, []):
                score_list = o["scores"].get((sec, qid), [])
                if score_list:
                    # Average across all participants for this question
                    q_avg = sum(s for s, _ in score_list) / len(score_list)
                    sec_total += q_avg
                    sec_count += 1
                    all_scores.extend(s for s, _ in score_list)

            section_scores[sec] = round(sec_total / sec_count, 2) if sec_count > 0 else None

        overall = round(sum(all_scores) / len(all_scores), 2) if all_scores else None
        result[org_name] = {
            "section_scores": section_scores,
            "overall_score": overall,
        }

    return result


def write_scorecard_sheet(wb, orgs: dict, questions: list[tuple]):
    ws = wb.create_sheet("Scorecard")

    score_data = compute_section_scores(orgs, questions)

    # ── Headers ──
    org_headers = [col[0] for col in OUTPUT_COLUMNS]
    section_headers = SCORECARD_SECTIONS
    all_headers = org_headers + section_headers + ["OVERALL"]

    for col_idx, header in enumerate(all_headers, 1):
        ws.cell(row=1, column=col_idx, value=header)

    apply_header_style(ws, 1, len(all_headers))

    # ── Data ──
    row_num = 2
    for org_name in sorted(orgs.keys()):
        o = orgs[org_name]
        sd = score_data.get(org_name, {})

        list_values = {}
        for key in ["role_level", "department", "age_band", "job_title"]:
            vals = sorted(o["list"].get(key, set()))
            list_values[key] = " | ".join(vals) if vals else ""

        # Org-level columns
        for col_idx, (_, key, agg_type) in enumerate(OUTPUT_COLUMNS, 1):
            if agg_type == "single":
                val = o["single"].get(key, "")
            elif agg_type == "list":
                val = list_values.get(key, "")
            else:
                val = ""
            ws.cell(row=row_num, column=col_idx, value=val)

        # Section scores
        base_col = len(OUTPUT_COLUMNS)
        for i, sec in enumerate(SCORECARD_SECTIONS):
            score = sd.get("section_scores", {}).get(sec)
            cell = ws.cell(row=row_num, column=base_col + i + 1)
            if score is not None:
                cell.value = score
                cell.number_format = '0.00'

        # Overall score
        overall_cell = ws.cell(row=row_num, column=base_col + len(SCORECARD_SECTIONS) + 1)
        overall = sd.get("overall_score")
        if overall is not None:
            overall_cell.value = overall
            overall_cell.number_format = '0.00'
            overall_cell.font = Font(name="Calibri", bold=True, size=10)

        row_num += 1

    total_cols = len(all_headers)
    apply_data_style(ws, 2, row_num - 1, total_cols)

    # ── Color scale on score columns (green = high, red = low) ──
    score_start_col = len(OUTPUT_COLUMNS) + 1
    score_end_col = total_cols
    score_range = f"{get_column_letter(score_start_col)}2:{get_column_letter(score_end_col)}{row_num - 1}"
    ws.conditional_formatting.add(
        score_range,
        ColorScaleRule(
            start_type="min", start_color="F8696B",   # red
            mid_type="percentile", mid_value=50, mid_color="FFEB84",  # yellow
            end_type="max", end_color="63BE7B",        # green
        ),
    )

    # Freeze only org name column + header so user can scroll through sections
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(total_cols)}{row_num - 1}"
    auto_fit_cols(ws, total_cols, len(OUTPUT_COLUMNS))


# ═══════════════════════════════════════════════════════════════════════════
# QUESTION REFERENCE SHEET
# ═══════════════════════════════════════════════════════════════════════════

def write_question_ref_sheet(wb, questions: list[tuple]):
    ws = wb.create_sheet("Question Reference")
    ws.append(["Section", "Question ID", "Question #", "Question Text"])
    for sec, qid, qnum, qtext in questions:
        ws.append([sec, qid, qnum, qtext])

    apply_header_style(ws, 1, 4)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 100
    ws.freeze_panes = "A2"


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    input_path = Path(INPUT_FILE)
    if not input_path.exists():
        print(f"❌ Input file not found: {INPUT_FILE}")
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

    wb = openpyxl.Workbook()

    print("📊 Writing Pivot sheet...")
    write_pivot_sheet(wb, orgs, questions)

    print("📊 Writing Scorecard sheet...")
    write_scorecard_sheet(wb, orgs, questions)

    print("📊 Writing Question Reference sheet...")
    write_question_ref_sheet(wb, questions)

    wb.save(OUTPUT_FILE)
    print(f"\n✅ Done! Output written to: {OUTPUT_FILE}")
    print(f"   Pivot:        {len(orgs)} orgs × {len(questions)} questions")
    print(f"   Scorecard:    {len(orgs)} orgs × {len(SCORECARD_SECTIONS)} sections + overall")
    print(f"   Question Ref: {len(questions)} questions")


if __name__ == "__main__":
    main()
