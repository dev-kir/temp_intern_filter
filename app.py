"""
Super Filter GUI — Streamlit App
==================================
Upload an Excel file, view the pivot table & scorecard, apply filters,
set custom question weightage, and export results.

Usage:
    streamlit run app.py
"""

import io
import sys
from pathlib import Path
from collections import defaultdict

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG (same as super_filter.py — editable here too)
# ═══════════════════════════════════════════════════════════════════════════

RAW_COL_MAP = {
    1:  "respondent_id",    2:  "submitted_at",      3:  "section",
    4:  "question_num",     5:  "question_id",       6:  "question",
    7:  "answer",           8:  "answer_value",       9:  "answer_score",
    10: "max_score",        11: "participant_name",  12: "email",
    13: "job_title",        14: "organisation_name", 15: "organisation_type",
    16: "organisation_size",17: "stakeholder_category",18:"pcds_sector",
    19: "district",        20: "role_level",         21: "department",
    22: "age_band",         23: "part_of_group",      24: "parent_company",
}

ALL_OUTPUT_COLUMNS = [
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

SECTION_ORDER = [
    "Background", "Strategy & Leadership", "Talent & Organisational Culture",
    "Data Management & Readiness", "Infrastructure & Technology",
    "Governance, Policy & Ethics", "Investment", "AI Implementation & Potential Impact",
]

BM_TO_EN_SECTION = {
    "Latar Belakang": "Background", "Strategi & Kepimpinan": "Strategy & Leadership",
    "Bakat & Budaya Organisasi": "Talent & Organisational Culture",
    "Pengurusan Data & Kesiapsiagaan": "Data Management & Readiness",
    "Infrastruktur & Teknologi": "Infrastructure & Technology",
    "Tadbir Urus, Dasar & Etika": "Governance, Policy & Ethics",
    "Pelaburan": "Investment", "Pelaksanaan AI & Impak": "AI Implementation & Potential Impact",
}

ALL_SCORECARD_SECTIONS = [
    "Strategy & Leadership", "Talent & Organisational Culture",
    "Data Management & Readiness", "Infrastructure & Technology",
    "Governance, Policy & Ethics", "Investment", "AI Implementation & Potential Impact",
]

# ═══════════════════════════════════════════════════════════════════════════
# PROCESSING FUNCTIONS (same logic as super_filter.py)
# ═══════════════════════════════════════════════════════════════════════════

def read_raw_data(filepath_or_buffer) -> list[dict]:
    wb = openpyxl.load_workbook(filepath_or_buffer, read_only=True)
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
    seen = set()
    questions = []
    for row in rows:
        sec = row["section"]
        if sec not in SECTION_ORDER:
            continue
        qid = row["question_id"]
        key = (sec, qid)
        if key not in seen:
            seen.add(key)
            questions.append((sec, qid,
                int(row["question_num"]) if row["question_num"].isdigit() else 0,
                row["question"]))
    questions.sort(key=lambda x: (
        SECTION_ORDER.index(x[0]) if x[0] in SECTION_ORDER else 99, x[2]))
    return questions


def build_org_data(rows: list[dict]) -> dict:
    orgs = defaultdict(lambda: {
        "single": {}, "list": defaultdict(set),
        "answers": defaultdict(set), "scores": defaultdict(list),
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
        en_sec = BM_TO_EN_SECTION.get(sec, sec)
        content = row.get("answer", "")
        if content:
            o["answers"][(en_sec, qid)].add(content)
        score_str = row.get("answer_score", "")
        if score_str:
            try:
                score = float(score_str)
                o["scores"][(en_sec, qid)].append(score)
            except ValueError:
                pass
    return orgs


def compute_scores(orgs: dict, questions: list[tuple], scorecard_sections: list,
                   weights: dict = None) -> dict:
    section_qids = defaultdict(list)
    for sec, qid, qnum, qtext in questions:
        if sec in scorecard_sections:
            section_qids[sec].append(qid)

    result = {}
    for org_name, o in orgs.items():
        section_scores = {}
        all_weighted = 0.0
        total_weight = 0.0

        for sec in scorecard_sections:
            sec_total = 0.0
            sec_count = 0
            for qid in section_qids.get(sec, []):
                score_list = o["scores"].get((sec, qid), [])
                if score_list:
                    q_avg = sum(score_list) / len(score_list)
                    w = weights.get(qid, 1.0) if weights else 1.0
                    sec_total += q_avg * w
                    sec_count += w
                    all_weighted += q_avg * w
                    total_weight += w
            section_scores[sec] = round(sec_total / sec_count, 2) if sec_count > 0 else None

        overall = round(all_weighted / total_weight, 2) if total_weight > 0 else None
        result[org_name] = {"section_scores": section_scores, "overall_score": overall}
    return result


# ═══════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════════

def export_to_excel(orgs: dict, questions: list[tuple], scorecard_sections: list,
                    output_columns: list, weights: dict = None) -> bytes:
    wb = openpyxl.Workbook()

    # ── Pivot sheet ──
    ws = wb.active
    ws.title = "Pivot"
    org_headers = [c[0] for c in output_columns]
    org_col_count = len(output_columns)

    question_columns = [(s, q, t) for s, q, n, t in questions]
    section_spans = []
    cur_sec, sec_start = None, None
    for i, (sec, qid, qtext) in enumerate(question_columns):
        col_idx = org_col_count + i + 1
        if sec != cur_sec:
            if cur_sec is not None:
                section_spans.append((cur_sec, sec_start, col_idx - 1))
            cur_sec, sec_start = sec, col_idx
    if cur_sec is not None:
        section_spans.append((cur_sec, sec_start, org_col_count + len(question_columns)))

    total_cols = org_col_count + len(question_columns)

    # Row 1: section headers
    if org_col_count > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=org_col_count)
        ws.cell(1, 1).value = "Organisation Info"
    for sec_name, sc, ec in section_spans:
        if sc == ec:
            ws.cell(1, sc).value = sec_name
        else:
            ws.merge_cells(start_row=1, start_column=sc, end_row=1, end_column=ec)
            ws.cell(1, sc).value = sec_name

    # Row 2: column headers
    for ci, h in enumerate(org_headers, 1):
        ws.cell(2, ci).value = h
    for i, (sec, qid, qtext) in enumerate(question_columns):
        ws.cell(2, org_col_count + i + 1).value = qid

    # Data
    row_num = 3
    for org_name in sorted(orgs.keys()):
        o = orgs[org_name]
        list_vals = {}
        for key in ["role_level", "department", "age_band", "job_title"]:
            vals = sorted(o["list"].get(key, set()))
            list_vals[key] = "\n".join(f"{i}. {v}" for i, v in enumerate(vals, 1)) if vals else ""

        for ci, (_, key, agg) in enumerate(output_columns, 1):
            if agg == "single":
                ws.cell(row_num, ci).value = o["single"].get(key, "")
            elif agg == "list":
                ws.cell(row_num, ci).value = list_vals.get(key, "")

        for i, (sec, qid, qtext) in enumerate(question_columns):
            answers = o["answers"].get((sec, qid), set())
            ws.cell(row_num, org_col_count + i + 1).value = " | ".join(sorted(answers)) if answers else ""
        row_num += 1

    # ── Scorecard sheet ──
    ws2 = wb.create_sheet("Scorecard")
    score_data = compute_scores(orgs, questions, scorecard_sections, weights)
    all_headers = org_headers + scorecard_sections + ["OVERALL"]
    for ci, h in enumerate(all_headers, 1):
        ws2.cell(1, ci).value = h

    row_num = 2
    for org_name in sorted(orgs.keys()):
        o = orgs[org_name]
        sd = score_data.get(org_name, {})
        list_vals = {}
        for key in ["role_level", "department", "age_band", "job_title"]:
            vals = sorted(o["list"].get(key, set()))
            list_vals[key] = "\n".join(f"{i}. {v}" for i, v in enumerate(vals, 1)) if vals else ""

        for ci, (_, key, agg) in enumerate(output_columns, 1):
            if agg == "single":
                ws2.cell(row_num, ci).value = o["single"].get(key, "")
            elif agg == "list":
                ws2.cell(row_num, ci).value = list_vals.get(key, "")

        base = len(output_columns)
        for i, sec in enumerate(scorecard_sections):
            s = sd.get("section_scores", {}).get(sec)
            if s is not None:
                ws2.cell(row_num, base + i + 1).value = s
                ws2.cell(row_num, base + i + 1).number_format = '0.00'
        ov = sd.get("overall_score")
        if ov is not None:
            ws2.cell(row_num, base + len(scorecard_sections) + 1).value = ov
            ws2.cell(row_num, base + len(scorecard_sections) + 1).number_format = '0.00'
        row_num += 1

    # ── Question Reference sheet ──
    ws3 = wb.create_sheet("Question Reference")
    ws3.append(["Section", "Question ID", "Question #", "Question Text"])
    for sec, qid, qnum, qtext in questions:
        ws3.append([sec, qid, qnum, qtext])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Super Filter — SARI Survey", layout="wide")
st.title("Super Filter — SARI Survey Results Processor")

# ── Sidebar: Upload & Config ──
with st.sidebar:
    st.header("Upload")
    uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx"])

    if uploaded_file:
        st.success(f"Loaded: {uploaded_file.name}")

    st.header("Org Columns")
    selected_org_cols = []
    for label, key, agg in ALL_OUTPUT_COLUMNS:
        if st.checkbox(label, value=True, key=f"col_{key}"):
            selected_org_cols.append((label, key, agg))

    st.header("Scorecard Sections")
    selected_sections = []
    for sec in ALL_SCORECARD_SECTIONS:
        if st.checkbox(sec, value=True, key=f"sec_{sec}"):
            selected_sections.append(sec)

    st.header("Question Weightage")
    st.caption("Default = 1.0 for all. Set custom weights below.")
    use_weights = st.checkbox("Enable custom weightage", value=False)

# ── Main area ──
if uploaded_file is None:
    st.info("Upload an Excel file in the sidebar to get started.")
    st.markdown("""
    ### What this does:
    1. **Upload** your SARI survey Excel export
    2. **View** the pivot table (one row per organisation)
    3. **View** the scorecard with per-section averages
    4. **Filter** columns and sections in the sidebar
    5. **Set custom weightage** per question
    6. **Export** the results as Excel
    """)
    st.stop()

# Process data
with st.spinner("Processing..."):
    rows = read_raw_data(uploaded_file)
    questions = build_question_order(rows)
    orgs = build_org_data(rows)

st.success(f"Loaded {len(rows)} rows, {len(orgs)} organisations, {len(questions)} questions")

# ── Weightage input ──
weights = {}
if use_weights:
    st.subheader("Question Weightage")
    st.caption("Set weight for each question (default = 1.0). Higher = more important.")
    wcols = st.columns(4)
    for i, (sec, qid, qnum, qtext) in enumerate(questions):
        if sec in selected_sections:
            with wcols[i % 4]:
                weights[qid] = st.number_input(
                    f"{qid}", min_value=0.0, max_value=10.0, value=1.0, step=0.5,
                    key=f"w_{qid}", help=qtext[:80])

# ── Tabs ──
tab1, tab2, tab3 = st.tabs(["Pivot Table", "Scorecard", "Question Reference"])

with tab1:
    st.subheader("Pivot Table — One row per organisation")

    # Build pivot dataframe
    org_headers = [c[0] for c in selected_org_cols]
    qcols = [(s, q, t) for s, q, n, t in questions]
    all_headers = org_headers + [qid for _, qid, _ in qcols]

    data_rows = []
    for org_name in sorted(orgs.keys()):
        o = orgs[org_name]
        row_data = {}

        list_vals = {}
        for key in ["role_level", "department", "age_band", "job_title"]:
            vals = sorted(o["list"].get(key, set()))
            list_vals[key] = "\n".join(f"{i}. {v}" for i, v in enumerate(vals, 1)) if vals else ""

        for label, key, agg in selected_org_cols:
            if agg == "single":
                row_data[label] = o["single"].get(key, "")
            elif agg == "list":
                row_data[label] = list_vals.get(key, "")

        for sec, qid, qtext in qcols:
            answers = o["answers"].get((sec, qid), set())
            row_data[qid] = " | ".join(sorted(answers)) if answers else ""

        data_rows.append(row_data)

    df = pd.DataFrame(data_rows)

    # Search/filter
    search = st.text_input("Filter by organisation name", "")
    if search:
        df = df[df["Organisation Name"].str.contains(search, case=False, na=False)]

    st.dataframe(df, use_container_width=True, height=500)
    st.caption(f"Showing {len(df)} organisations")

with tab2:
    st.subheader("Scorecard — Performance by Section")

    score_data = compute_scores(orgs, questions, selected_sections, weights if use_weights else None)

    sc_headers = org_headers + selected_sections + ["OVERALL"]
    sc_rows = []
    for org_name in sorted(orgs.keys()):
        o = orgs[org_name]
        sd = score_data.get(org_name, {})
        row_data = {}

        list_vals = {}
        for key in ["role_level", "department", "age_band", "job_title"]:
            vals = sorted(o["list"].get(key, set()))
            list_vals[key] = "\n".join(f"{i}. {v}" for i, v in enumerate(vals, 1)) if vals else ""

        for label, key, agg in selected_org_cols:
            if agg == "single":
                row_data[label] = o["single"].get(key, "")
            elif agg == "list":
                row_data[label] = list_vals.get(key, "")

        for sec in selected_sections:
            row_data[sec] = sd.get("section_scores", {}).get(sec)

        row_data["OVERALL"] = sd.get("overall_score")
        sc_rows.append(row_data)

    sc_df = pd.DataFrame(sc_rows)

    if search:
        sc_df = sc_df[sc_df["Organisation Name"].str.contains(search, case=False, na=False)]

    # Color the score columns
    def color_scores(val):
        if pd.isna(val) or not isinstance(val, (int, float)):
            return ""
        if val >= 3:
            return "background-color: #63BE7B; color: white"
        elif val >= 2:
            return "background-color: #FFEB84"
        elif val >= 1:
            return "background-color: #F8696B; color: white"
        return "background-color: #C00000; color: white"

    score_cols = selected_sections + ["OVERALL"]
    styled = sc_df.style.applymap(color_scores, subset=score_cols).format(
        {c: "{:.2f}" for c in score_cols}, na_rep="-")

    st.dataframe(styled, use_container_width=True, height=500)
    st.caption(f"Showing {len(sc_df)} organisations | Scores: 0–4 scale | Green = high, Red = low")

with tab3:
    st.subheader("Question Reference")
    ref_data = [{"Section": s, "Question ID": q, "Question #": n, "Question Text": t}
                for s, q, n, t in questions]
    ref_df = pd.DataFrame(ref_data)
    st.dataframe(ref_df, use_container_width=True, height=500)

# ── Export ──
st.divider()
if st.button("Export to Excel", type="primary"):
    excel_bytes = export_to_excel(orgs, questions, selected_sections,
                                   selected_org_cols, weights if use_weights else None)
    st.download_button(
        label="Download Excel file",
        data=excel_bytes,
        file_name="SARI_Results_Processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.success("Ready! Click the download button above.")
