"""
SARI Organisation Statistics Generator
========================================
Generates a 10-sheet interactive workbook from a SARI raw Answers export.
Dashboard and Organisation Report use a modern layout engine with KPI cards,
clean typography, and professional styling.

Usage:
    python super_filter.py
    (edit INPUT_FILE / OUTPUT_FILE at top of this file)
"""

import math
import statistics
from pathlib import Path
from collections import Counter

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.workbook.defined_name import DefinedName

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════════════

INPUT_FILE = "SARI_Results_2026-08-04-00-36-28.xlsx"
OUTPUT_FILE = "SARI_Organisation.xlsx"

CFG = {
    "high_consensus": 0.80,
    "moderate_consensus": 0.60,
    "minimum_respondents_for_agreement": 2,
    "tier_boundaries": [0.20, 0.40, 0.60, 0.80],
    "section_order": [
        "Strategy & Leadership",
        "Governance, Policy & Ethics",
        "Talent & Organisational Culture",
        "Infrastructure & Technology",
        "Data Management & Readiness",
        "AI Implementation & Potential Impact",
        "Investment",
    ],
    "section_prefixes": {
        "strategy": "Strategy & Leadership",
        "governance": "Governance, Policy & Ethics",
        "talent": "Talent & Organisational Culture",
        "infrastructure": "Infrastructure & Technology",
        "data": "Data Management & Readiness",
        "aiapp": "AI Implementation & Potential Impact",
        "investment": "Investment",
        "background": "Background",
    },
    "multi_select_question_ids": ["background_2", "background_3", "background_4"],
}

REQ = [
    "Respondent ID", "Submitted at", "Question #", "Question ID", "Question",
    "Answer", "Answer value", "Answer score", "Max score", "Participant name",
    "Job title", "Organisation name", "Organisation type", "Organisation size",
    "Stakeholder category", "PCDS sector", "District", "Role level",
    "Department", "Age band", "Part of group", "Parent company",
]

RAW_OUT = [
    "Respondent ID", "Submitted at", "Standard section", "Question #",
    "Question ID", "Question", "Answer", "Answer value", "Answer score",
    "Max score", "Participant name", "Job title", "Organisation name",
    "Organisation type", "Organisation size", "Stakeholder category",
    "PCDS sector", "District", "Role level", "Department", "Age band",
    "Part of group", "Parent company",
]

# ═══════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM — Colours, Typography, Layout
# ═══════════════════════════════════════════════════════════════════════════

# ── Colours ──
NAVY = "FF163A63"
BLUE_LIGHT = "FFD9EAF7"
GREEN = "FF4CAF50"
ORANGE = "FFF4B400"
RED = "FFDB4437"
TEXT_DARK = "FF1F1F1F"
GREY = "FF6E6E6E"
WHITE = "FFFFFFFF"
BG_WHITE = "FFFFFFFF"
CARD_FILL = "FFF5F5F5"
ROW_ALT = "FFF8F9FA"
HEADER_BG = "FF2F5496"
CHART_COLORS = ["FF163A63", "FFDB4437", "FF4CAF50", "FF7B1FA2", "FF0097A7", "FFF4B400", "FF03A9F4"]

# ── Typography ──
FONT_FAMILY = "Calibri"
TITLE_FONT = Font(name=FONT_FAMILY, size=16, bold=True, color=WHITE)
SUBTITLE_FONT = Font(name=FONT_FAMILY, size=10, color=GREY)
CARD_LABEL_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color=TEXT_DARK)
CARD_VALUE_FONT = Font(name=FONT_FAMILY, size=12, bold=True, color=TEXT_DARK)
SECTION_TITLE_FONT = Font(name=FONT_FAMILY, size=11, bold=True, color=TEXT_DARK)
TABLE_HEADER_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color=WHITE)
TABLE_BODY_FONT = Font(name=FONT_FAMILY, size=9, color=TEXT_DARK)
CHART_FONT = Font(name=FONT_FAMILY, size=9, color=TEXT_DARK)

# ── Layout Engine ──
LAYOUT = {
    "margin_col": 2,       # left margin in columns
    "margin_row": 2,       # top margin in rows (after title)
    "card_width": 3,       # KPI card width in columns
    "card_gap": 1,         # gap between cards in columns
    "card_height": 3,      # KPI card height in rows (label + value)
    "section_gap": 1,      # gap between card rows
    "table_start_row": 17, # where the section table starts
    "chart_col": 9,        # chart anchor column
    "chart_width": 11,     # chart width in columns
    "chart_height": 10,    # chart height in rows
}

# ── Reusable Styles ──
THIN_BORDER = Border(
    left=Side(style="hair", color="FFD0D0D0"),
    right=Side(style="hair", color="FFD0D0D0"),
    top=Side(style="hair", color="FFD0D0D0"),
    bottom=Side(style="hair", color="FFD0D0D0"),
)
NO_BORDER = Border()

CARD_ALIGN = Alignment(horizontal="left", vertical="top", wrap_text=True)
CARD_VALUE_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")

# ═══════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def tier(x, b):
    if x < b[0]: return "AI Aware - 0"
    if x < b[1]: return "AI Explorer - 1"
    if x < b[2]: return "AI Follower - 2"
    if x < b[3]: return "AI Leader - 3"
    return "AI Pioneer - 4"


def agreement(consensus, n, min_n, high, moderate):
    if n < min_n: return "Not measurable"
    if consensus >= high: return "High"
    if consensus >= moderate: return "Moderate"
    return "Low"


def split_multi(answer):
    """Split a multi-select answer on commas, but NOT on commas inside brackets.

    Several options contain their own punctuation, e.g.
    "Shared infrastructure (e.g., cloud, computing power)". A naive
    str.split(",") shredded that single option into three phantom options
    across 39 organisations, inflating Answer Distribution by 77 rows and
    silently understating the real selection count for that option.
    """
    text = str(answer or "")
    parts, buf, depth = [], [], 0
    for ch in text:
        if ch in "([{":
            depth += 1
        elif ch in ")]}":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    parts.append("".join(buf))
    return [p.strip() for p in parts if p.strip()]


def mode_det(values):
    vals = [str(v).strip() for v in values if pd.notna(v) and str(v).strip()]
    if not vals: return "", 0
    c = Counter(vals)
    m = max(c.values())
    winners = sorted(k for k, v in c.items() if v == m)
    return winners[0], m


def apply_card(ws, row, col, label, value, width=None):
    """Draw a single KPI card at (row, col)."""
    w = width or LAYOUT["card_width"]
    # Card background
    for r in range(row, row + LAYOUT["card_height"]):
        for c in range(col, col + w):
            cell = ws.cell(r, c)
            cell.fill = PatternFill("solid", fgColor=CARD_FILL)
            cell.border = NO_BORDER
    # Label row
    label_cell = ws.cell(row, col)
    label_cell.value = label
    label_cell.font = CARD_LABEL_FONT
    label_cell.alignment = CARD_ALIGN
    # Value row
    val_cell = ws.cell(row + 1, col)
    val_cell.value = value
    val_cell.font = CARD_VALUE_FONT
    val_cell.alignment = CARD_VALUE_ALIGN
    # Merge value across card width for long text
    if w > 1:
        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + w - 1)


def apply_section_table(ws, start_row, start_col, sections, sdf_len):
    """Draw the section breakdown table."""
    headers = ["Section", "Average score", "Normalised score", "Respondents", "Agreement"]
    col_widths = [4, 2, 2, 2, 2]
    h = start_row

    # Header
    for i, (hdr, w) in enumerate(zip(headers, col_widths)):
        c = start_col + sum(col_widths[:i])
        cell = ws.cell(h, c)
        cell.value = hdr
        cell.font = TABLE_HEADER_FONT
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = HEADER_ALIGN
        cell.border = NO_BORDER
        if w > 1:
            ws.merge_cells(start_row=h, start_column=c, end_row=h, end_column=c + w - 1)

    # Data rows
    for ri, sec in enumerate(sections):
        r = h + 1 + ri
        bg = WHITE if ri % 2 == 0 else ROW_ALT
        for ci in range(start_col, start_col + sum(col_widths)):
            cell = ws.cell(r, ci)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = TABLE_BODY_FONT
            cell.border = NO_BORDER

        # Section name
        ws.cell(r, start_col).value = sec
        ws.cell(r, start_col).alignment = LEFT_ALIGN

        # Average score
        ac = start_col + col_widths[0]
        ws.cell(r, ac).value = f'=IFERROR(SUMIFS(\'Section Summary\'!$E:$E,\'Section Summary\'!$A:$A,$B$4,\'Section Summary\'!$B:$B,$A{r}),"")'
        ws.cell(r, ac).alignment = CENTER_ALIGN

        # Normalised score
        nc = ac + col_widths[1]
        ws.cell(r, nc).value = f'=IFERROR(SUMIFS(\'Section Summary\'!$J:$J,\'Section Summary\'!$A:$A,$B$4,\'Section Summary\'!$B:$B,$A{r}),"")'
        ws.cell(r, nc).number_format = "0.0%"
        ws.cell(r, nc).alignment = CENTER_ALIGN

        # Respondents
        rc = nc + col_widths[2]
        ws.cell(r, rc).value = f'=IFERROR(SUMIFS(\'Section Summary\'!$C:$C,\'Section Summary\'!$A:$A,$B$4,\'Section Summary\'!$B:$B,$A{r}),"")'
        ws.cell(r, rc).alignment = CENTER_ALIGN

        # Agreement
        gc = rc + col_widths[3]
        ws.cell(r, gc).value = f'=IF(D{r}<2,"Not measurable",IF(SUMIFS(\'Section Summary\'!$K:$K,\'Section Summary\'!$A:$A,$B$4,\'Section Summary\'!$B:$B,$A{r})>=0.8,"High",IF(SUMIFS(\'Section Summary\'!$K:$K,\'Section Summary\'!$A:$A,$B$4,\'Section Summary\'!$B:$B,$A{r})>=0.6,"Moderate","Low")))'
        ws.cell(r, gc).alignment = CENTER_ALIGN

    return h + 1 + len(sections)


def apply_banner(ws, title_text, subtitle_text, span):
    """Draw the navy banner at top of sheet."""
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    ws["A1"] = title_text
    ws["A1"].font = TITLE_FONT
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
    ws["A2"] = subtitle_text
    ws["A2"].font = SUBTITLE_FONT
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[2].height = 22


def apply_data_sheet(ws, name, subtitle, columns, rows, widths, freeze="A5", add_filter=True):
    """Write a data sheet with clean styling."""
    apply_banner(ws, name, subtitle, len(columns))

    # Header row
    for c, h in enumerate(columns, 1):
        cell = ws.cell(4, c)
        cell.value = h
        cell.font = TABLE_HEADER_FONT
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = HEADER_ALIGN
        cell.border = NO_BORDER
    ws.row_dimensions[4].height = 28

    # Data rows
    for r, row in enumerate(rows, 5):
        bg = WHITE if (r - 5) % 2 == 0 else ROW_ALT
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c)
            cell.value = v
            cell.font = TABLE_BODY_FONT
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = NO_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    ws.freeze_panes = freeze
    if add_filter:
        ws.auto_filter.ref = f"A4:{get_column_letter(len(columns))}{4 + len(rows)}"
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w
    ws.page_setup.orientation = "portrait"


# ═══════════════════════════════════════════════════════════════════════════
# DATA PROCESSING
# ═══════════════════════════════════════════════════════════════════════════

def load_data(path, cfg):
    df = pd.read_excel(path, sheet_name="Answers", engine="openpyxl", dtype=object)
    missing = [c for c in REQ if c not in df.columns]
    if missing:
        raise ValueError("Missing required columns: " + ", ".join(missing))
    for c in ["Answer score", "Max score", "Question #"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[df["Organisation name"].notna()].copy()
    df["Organisation name"] = df["Organisation name"].astype(str).str.strip()
    df = df[df["Organisation name"] != ""]
    df = df.drop_duplicates(["Respondent ID", "Question ID"], keep="last")
    pref = df["Question ID"].astype(str).str.split("_").str[0]
    df["Standard section"] = pref.map(cfg["section_prefixes"]).fillna(df.get("Section", "Other"))
    return df


def calculate(df, cfg):
    high = cfg["high_consensus"]
    mod = cfg["moderate_consensus"]
    min_n = cfg["minimum_respondents_for_agreement"]
    sections = cfg["section_order"]
    multi = set(cfg["multi_select_question_ids"])

    # ── Question Summary ──
    qrows = []
    for (org, sec, qid, qtext), g in df.groupby(
        ["Organisation name", "Standard section", "Question ID", "Question"],
        dropna=False, sort=True,
    ):
        n = g["Respondent ID"].nunique()
        mode, count = mode_det(g["Answer"])
        cons = count / n if n else 0
        scored = (g["Max score"].fillna(0) > 0).any()
        scores = g.loc[g["Max score"].fillna(0) > 0, "Answer score"].dropna().astype(float)
        maxs = g.loc[g["Max score"].fillna(0) > 0, "Max score"].dropna().astype(float)
        avg = scores.mean() if len(scores) else None
        med = scores.median() if len(scores) else None
        mn = scores.min() if len(scores) else None
        mx = scores.max() if len(scores) else None
        rng = (mx - mn) if len(scores) else None
        sd = scores.std(ddof=1) if len(scores) > 1 else (0 if len(scores) == 1 else None)
        norm = (avg / maxs.mean()) if len(scores) and maxs.mean() else None
        agr = agreement(cons, n, min_n, high, mod)
        flag = "Review" if n >= min_n and cons < mod else ""
        qrows.append([org, sec, qid, qtext, n, bool(scored), mode, count, cons, avg, med, mn, mx, rng, sd, norm, agr, flag])

    qcols = [
        "Organisation name", "Section", "Question ID", "Question", "Respondents",
        "Scored question", "Most common answer", "Most common count", "Consensus",
        "Average score", "Median score", "Minimum score", "Maximum score",
        "Score range", "Standard deviation", "Normalised score", "Agreement", "Review flag",
    ]
    qdf = pd.DataFrame(qrows, columns=qcols)

    # ── Section Summary ──
    srows = []
    for org in sorted(df["Organisation name"].unique()):
        for sec in sections:
            g = df[(df["Organisation name"] == org) & (df["Standard section"] == sec) & (df["Max score"].fillna(0) > 0)]
            if g.empty: continue
            scores = g["Answer score"].dropna().astype(float)
            maxs = g["Max score"].dropna().astype(float)
            n = g["Respondent ID"].nunique()
            nq = g["Question ID"].nunique()
            avg = scores.mean(); med = scores.median(); mn = scores.min(); mx = scores.max()
            maxp = maxs.mean(); norm = avg / maxp if maxp else None
            qs = qdf[(qdf["Organisation name"] == org) & (qdf["Section"] == sec)]
            cons = qs["Consensus"].mean() if len(qs) else None
            agr = agreement(cons or 0, n, min_n, high, mod)
            srows.append([org, sec, n, nq, avg, med, mn, mx, maxp, norm, cons, agr])

    scols = [
        "Organisation name", "Section", "Respondents", "Questions",
        "Average score", "Median score", "Minimum score", "Maximum score",
        "Max possible", "Normalised score", "Average consensus", "Agreement",
    ]
    sdf = pd.DataFrame(srows, columns=scols)

    # ── Organisation Summary ──
    orows = []
    for org, g in df.groupby("Organisation name", sort=True):
        n = g["Respondent ID"].nunique()
        scored = g[g["Max score"].fillna(0) > 0]
        avg = scored["Answer score"].dropna().astype(float).mean()
        denom = scored["Max score"].dropna().astype(float).mean()
        overall = avg / denom if denom else 0
        ss = sdf[sdf["Organisation name"] == org]
        strongest = ss.sort_values(["Normalised score", "Section"], ascending=[False, True]).iloc[0]["Section"] if len(ss) else ""
        weakest = ss.sort_values(["Normalised score", "Section"]).iloc[0]["Section"] if len(ss) else ""
        qs = qdf[qdf["Organisation name"] == org]
        cons = qs["Consensus"].mean() if len(qs) else 0
        reviews = int((qs["Review flag"] == "Review").sum())
        agr = agreement(cons, n, min_n, high, mod)
        first = g.iloc[0]
        typ = first.get("Organisation type", "")
        size = first.get("Organisation size", "")
        sector = first.get("PCDS sector", "")
        latest = g.iloc[-1].get("Submitted at", "")
        depts = g["Department"].dropna().astype(str).replace("", pd.NA).dropna().nunique()
        roles = g["Role level"].dropna().astype(str).replace("", pd.NA).dropna().nunique()
        interpretation = "Single respondent: perception only" if n < 2 else ("Directional: small sample" if n < 3 else "Multi-respondent view")
        b = cfg["tier_boundaries"]
        nxt = (1 - overall) if overall >= b[-1] else min(x for x in b if x > overall) - overall
        orows.append([org, typ, n, depts, roles, size, sector, latest, avg, overall, strongest, weakest, cons, reviews, agr, interpretation, tier(overall, b), nxt])

    ocols = [
        "Organisation name", "Organisation type", "Respondents", "Departments represented",
        "Role levels represented", "Organisation size", "Sector", "Latest submission",
        "Average score", "Overall score", "Strongest section", "Weakest section",
        "Average consensus", "Questions for review", "Agreement", "Interpretation",
        "Maturity tier", "Distance to next tier",
    ]
    odf = pd.DataFrame(orows, columns=ocols)

    # ── Answer Distribution ──
    drows = []
    for (org, sec, qid, qtext), g in df.groupby(
        ["Organisation name", "Standard section", "Question ID", "Question"], sort=True,
    ):
        n = g["Respondent ID"].nunique()
        if qid in multi:
            pairs = []
            for _, x in g.iterrows():
                for opt in split_multi(x["Answer"]):
                    pairs.append((x["Respondent ID"], opt, x["Answer score"]))
            for opt in sorted(set(p[1] for p in pairs)):
                selected = len(set(p[0] for p in pairs if p[1] == opt))
                scores = [p[2] for p in pairs if p[1] == opt and pd.notna(p[2])]
                drows.append([org, sec, qid, qtext, "Multi-select", opt, (sum(scores) / len(scores) if scores else None), selected, selected / n if n else 0, n])
        else:
            for opt, gg in g.groupby("Answer", dropna=False, sort=True):
                selected = gg["Respondent ID"].nunique()
                sc = gg["Answer score"].dropna().astype(float)
                drows.append([org, sec, qid, qtext, "Single-choice", opt, (sc.mean() if len(sc) else None), selected, selected / n if n else 0, n])

    dcols = [
        "Organisation name", "Standard section", "Question ID", "Question",
        "Question type", "Answer option", "Answer score", "Respondents selecting",
        "Percentage", "Question respondents",
    ]
    ddf = pd.DataFrame(drows, columns=dcols)

    return odf, sdf, qdf, ddf


# ═══════════════════════════════════════════════════════════════════════════
# WORKBOOK BUILDER
# ═══════════════════════════════════════════════════════════════════════════

def build(df, odf, sdf, qdf, ddf, cfg, out):
    wb = Workbook()
    wb.remove(wb.active)
    sections = cfg["section_order"]
    L = LAYOUT

    # ── 1. Read Me ──
    ws = wb.create_sheet("Read Me")
    apply_banner(ws, "SARI Organisation Statistics",
                 "Interactive organisation-level reporting workbook generated from the uploaded SARI results.", 8)
    notes = [
        ("Purpose", "Analyse questionnaire results by organisation, section, question and answer distribution."),
        ("Dashboard", "Choose an organisation in cell B4. The section score table and chart update using formulas."),
        ("Organisation Summary", "One row per organisation, including respondent count, maturity score, strongest/weakest sections and agreement."),
        ("Section Summary", "Scored question results aggregated by organisation and standardised section."),
        ("Question Summary", "Question-level mode, consensus, score statistics, agreement and review flags."),
        ("Answer Distribution", "Counts and percentages for each answer option. Multi-select percentages may exceed 100% in total."),
        ("Scoring rule", "Only rows where Max score is greater than zero are included in maturity scores. Background questions remain distribution-only."),
        ("Agreement rule", "High: at least 80% selected the modal answer; Moderate: 60% to below 80%; Low: below 60%; not measurable for one respondent."),
        ("Caution", "Results with one respondent represent an individual perception, not organisation-wide consensus."),
        ("Data standardisation", "Sections are standardised using the Question ID prefix, preventing English and Malay labels from splitting the same section."),
        ("Refresh", "This workbook is a snapshot. Re-run the analysis when new responses are added to the source export."),
        ("Enhancements in this version", ""),
        ("", "Maturity tiers, distance to next tier and a printable Organisation Report are included."),
    ]
    for r, (a, b) in enumerate(notes, 4):
        ws.cell(r, 1, a).font = Font(name=FONT_FAMILY, bold=True, color=NAVY)
        ws.cell(r, 2, b).font = Font(name=FONT_FAMILY, size=10, color=TEXT_DARK)
        # No per-row merge: the reference workbook merges only the two banner rows,
        # and column B is wide enough (100) to hold the text without it.
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 100
    ws.page_setup.orientation = "portrait"

    # ── 2. Lists (hidden) ──
    ws = wb.create_sheet("Lists")
    for r, o in enumerate(odf["Organisation name"], 1):
        ws.cell(r, 1, o)
    ws.sheet_state = "hidden"
    ws.page_setup.orientation = "portrait"

    # Dynamic named range so the dropdown grows with the organisation list rather
    # than being pinned to today's row count.
    wb.defined_names["OrganisationList"] = DefinedName(
        "OrganisationList",
        attr_text="Lists!$A$1:INDEX(Lists!$A:$A,COUNTA(Lists!$A:$A))")

    # ── 3. Dashboard ──
    # Layout matches chatgpt/SARI_Organisation_Fixed_Dropdowns.xlsx exactly:
    # A1:J24, selector at B4, label/value pairs at A/D/G, section table A17:E24.
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:J1")
    ws["A1"] = "Organisation Dashboard"
    ws["A1"].font = TITLE_FONT
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.merge_cells("A2:J2")
    ws["A2"] = ("Select an organisation. Green cells contain imported selections; "
                "black cells are formulas.")
    ws["A2"].font = SUBTITLE_FONT
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")

    # Selector, driven by the OrganisationList defined name so it grows with the data
    ws["A4"] = "Selected organisation"
    ws["A4"].font = CARD_LABEL_FONT
    first_org = odf.iloc[0]["Organisation name"] if len(odf) else ""
    ws["B4"] = first_org
    ws["B4"].font = Font(name=FONT_FAMILY, size=12, bold=True, color="FF008000")
    ws["B4"].fill = PatternFill("solid", fgColor="FFE2F0D9")
    ws["B4"].alignment = CARD_VALUE_ALIGN
    ws["B4"].border = THIN_BORDER
    dv = DataValidation(type="list", formula1="OrganisationList", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(ws["B4"])

    def _kpi(label_row, col_letter, label, org_col, pct=False, merge_to=None):
        """One label/value pair. Value reads Organisation Summary via the B4 selector."""
        lc = ws[f"{col_letter}{label_row}"]
        lc.value = label
        lc.font = CARD_LABEL_FONT
        lc.alignment = CARD_ALIGN
        vr = label_row + 1
        vc = ws[f"{col_letter}{vr}"]
        vc.value = (f"=IFERROR(INDEX('Organisation Summary'!${org_col}:${org_col},"
                    f"MATCH($B$4,'Organisation Summary'!$A:$A,0)),\"\")")
        vc.font = CARD_VALUE_FONT
        vc.alignment = CARD_VALUE_ALIGN
        vc.fill = PatternFill("solid", fgColor=BLUE_LIGHT)
        if pct:
            vc.number_format = "0.0%"
        if merge_to:
            ws.merge_cells(f"{col_letter}{vr}:{merge_to}{vr}")
            for cc in range(column_index_from_string(col_letter) + 1,
                            column_index_from_string(merge_to) + 1):
                ws.cell(vr, cc).fill = PatternFill("solid", fgColor=BLUE_LIGHT)

    _kpi(6, "A", "Respondents",          "C", merge_to="B")
    _kpi(6, "D", "Overall score",        "J", pct=True, merge_to="E")
    _kpi(6, "G", "Strongest section",    "K", merge_to="H")
    _kpi(9, "A", "Weakest section",      "L", merge_to="B")
    _kpi(9, "D", "Agreement",            "O", merge_to="E")
    _kpi(9, "G", "Interpretation",       "P", merge_to="H")
    _kpi(12, "A", "Maturity tier",       "Q")
    _kpi(12, "D", "Distance to next tier", "R", pct=True)
    _kpi(12, "G", "Questions for review", "N")

    # ── Section table, A17:E24 ──
    tbl_row = 17
    for i, hdr in enumerate(["Section", "Average score", "Normalised score",
                             "Respondents", "Agreement"]):
        c = ws.cell(tbl_row, 1 + i)
        c.value = hdr
        c.font = TABLE_HEADER_FONT
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = HEADER_ALIGN

    for ri, sec in enumerate(sections):
        r = tbl_row + 1 + ri
        bg = WHITE if ri % 2 == 0 else ROW_ALT
        for ci in range(1, 6):
            cell = ws.cell(r, ci)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = TABLE_BODY_FONT
        ws.cell(r, 1).value = sec
        ws.cell(r, 1).alignment = LEFT_ALIGN
        ws.cell(r, 2).value = (f"=IFERROR(SUMIFS('Section Summary'!$E:$E,"
                               f"'Section Summary'!$A:$A,$B$4,'Section Summary'!$B:$B,$A{r}),\"\")")
        ws.cell(r, 2).alignment = CENTER_ALIGN
        ws.cell(r, 3).value = (f"=IFERROR(SUMIFS('Section Summary'!$J:$J,"
                               f"'Section Summary'!$A:$A,$B$4,'Section Summary'!$B:$B,$A{r}),\"\")")
        ws.cell(r, 3).number_format = "0.0%"
        ws.cell(r, 3).alignment = CENTER_ALIGN
        ws.cell(r, 4).value = (f"=IFERROR(SUMIFS('Section Summary'!$C:$C,"
                               f"'Section Summary'!$A:$A,$B$4,'Section Summary'!$B:$B,$A{r}),\"\")")
        ws.cell(r, 4).alignment = CENTER_ALIGN
        ws.cell(r, 5).value = (
            f'=IF(D{r}<2,"Not measurable",'
            f"IF(SUMIFS('Section Summary'!$K:$K,'Section Summary'!$A:$A,$B$4,"
            f"'Section Summary'!$B:$B,$A{r})>=0.8,\"High\","
            f"IF(SUMIFS('Section Summary'!$K:$K,'Section Summary'!$A:$A,$B$4,"
            f"'Section Summary'!$B:$B,$A{r})>=0.6,\"Moderate\",\"Low\")))")
        ws.cell(r, 5).alignment = CENTER_ALIGN

    last_tbl_row = tbl_row + len(sections)
    ws.conditional_formatting.add(
        f"C{tbl_row+1}:C{last_tbl_row}",
        ColorScaleRule(start_type="min", start_color="FFF8696B",
                       mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
                       end_type="max", end_color="FF63BE7B"))

    ch = BarChart()
    ch.type = "bar"
    ch.style = 2
    ch.title = "Section maturity profile"
    ch.legend = None
    ch.add_data(Reference(ws, min_col=3, min_row=tbl_row + 1, max_row=last_tbl_row),
                titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=1, min_row=tbl_row + 1, max_row=last_tbl_row))
    ch.width = 15
    ch.height = 7.5
    ch.y_axis.title = "Normalised score"
    ch.x_axis.title = "Section"
    ch.y_axis.numFmt = "0.0%"
    # For a HORIZONTAL bar chart the category axis sits left and the value axis
    # sits bottom. openpyxl leaves both at "l", which is invalid OOXML: Excel
    # refuses the file and offers to "recover" it, silently dropping the chart
    # and some formatting on the way. LibreOffice accepts it, so this only shows
    # up in Excel. Set both explicitly.
    ch.x_axis.axPos = "l"
    ch.y_axis.axPos = "b"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    for idx, color in enumerate(CHART_COLORS[:len(sections)]):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = color
        ch.series[0].data_points.append(pt)
    ws.add_chart(ch, "G17")

    for col, w in {"A": 42.0, "B": 39.1, "C": 14.0, "D": 15.0, "E": 18.0, "G": 31.4}.items():
        ws.column_dimensions[col].width = w
    for r, h in {1: 27.75, 2: 30.0, 7: 15.0, 10: 15.0, 13: 34.5, 17: 27.75}.items():
        ws.row_dimensions[r].height = h
    ws.freeze_panes = "A17"
    ws.page_setup.orientation = "portrait"

    # ── 4. Organisation Report ──
    # A1:H26, selector at B3, pairs at A/D/G, section table A10:C17, priority A20:F26.
    ws = wb.create_sheet("Organisation Report")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:H1")
    ws["A1"] = "SARI Organisation Report"
    ws["A1"].font = TITLE_FONT
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(vertical="center")

    ws["A3"] = "Selected organisation"
    ws["A3"].font = CARD_LABEL_FONT
    ws["B3"] = first_org
    ws["B3"].font = Font(name=FONT_FAMILY, size=12, bold=True, color="FF008000")
    ws["B3"].fill = PatternFill("solid", fgColor="FFE2F0D9")
    ws["B3"].alignment = CARD_VALUE_ALIGN
    ws["B3"].border = THIN_BORDER
    dv2 = DataValidation(type="list", formula1="OrganisationList", allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add(ws["B3"])

    def _rkpi(row, lbl_col, val_col, label, org_col, pct=False):
        lc = ws[f"{lbl_col}{row}"]
        lc.value = label
        lc.font = CARD_LABEL_FONT
        lc.alignment = CARD_ALIGN
        vc = ws[f"{val_col}{row}"]
        vc.value = (f"=IFERROR(INDEX('Organisation Summary'!${org_col}:${org_col},"
                    f"MATCH($B$3,'Organisation Summary'!$A:$A,0)),\"\")")
        vc.font = CARD_VALUE_FONT
        vc.alignment = CARD_VALUE_ALIGN
        vc.fill = PatternFill("solid", fgColor=BLUE_LIGHT)
        if pct:
            vc.number_format = "0.0%"

    _rkpi(5, "A", "B", "Overall score",     "J", pct=True)
    _rkpi(5, "D", "E", "Maturity tier",     "Q")
    _rkpi(5, "G", "H", "Respondents",       "C")
    _rkpi(7, "A", "B", "Strongest section", "K")
    _rkpi(7, "D", "E", "Weakest section",   "L")
    _rkpi(7, "G", "H", "Agreement",         "O")

    st_row = 10
    for i, hdr in enumerate(["Section", "Score", "Agreement"]):
        c = ws.cell(st_row, 1 + i)
        c.value = hdr
        c.font = TABLE_HEADER_FONT
        c.fill = PatternFill("solid", fgColor=HEADER_BG)
        c.alignment = HEADER_ALIGN

    sec_last = 4 + len(sdf)
    for ri, sec in enumerate(sections):
        r = st_row + 1 + ri
        bg = WHITE if ri % 2 == 0 else ROW_ALT
        for ci in range(1, 4):
            ws.cell(r, ci).fill = PatternFill("solid", fgColor=bg)
            ws.cell(r, ci).font = TABLE_BODY_FONT
        ws.cell(r, 1).value = sec
        ws.cell(r, 1).alignment = LEFT_ALIGN
        ws.cell(r, 2).value = (f"=IFERROR(SUMIFS('Section Summary'!$J:$J,"
                               f"'Section Summary'!$A:$A,$B$3,'Section Summary'!$B:$B,$A{r}),\"\")")
        ws.cell(r, 2).number_format = "0.0%"
        ws.cell(r, 2).alignment = CENTER_ALIGN
        ws.cell(r, 3).value = (
            f"=IFERROR(LOOKUP(2,1/(('Section Summary'!$A$5:$A${sec_last}=$B$3)*"
            f"('Section Summary'!$B$5:$B${sec_last}=$A{r})),"
            f"'Section Summary'!$L$5:$L${sec_last}),\"\")")
        ws.cell(r, 3).alignment = CENTER_ALIGN

    pq_row = 20
    ws.cell(pq_row, 1).value = "Priority questions for review"
    ws.cell(pq_row, 1).font = SECTION_TITLE_FONT
    for i, hdr in enumerate(["Question ID", "Question", "Most common answer",
                             "Normalised score", "Agreement", "Review flag"]):
        c = ws.cell(pq_row + 1, 1 + i)
        c.value = hdr
        c.font = TABLE_HEADER_FONT
        c.fill = PatternFill("solid", fgColor=HEADER_BG)
        c.alignment = HEADER_ALIGN

    for rank in range(1, 6):
        r = pq_row + 1 + rank
        bg = WHITE if (rank - 1) % 2 == 0 else ROW_ALT
        for ci, src in enumerate(["B", "C", "D", "E", "F", "G"]):
            cell = ws.cell(r, 1 + ci)
            cell.value = (f"=IFERROR(INDEX('Priority Detail'!${src}:${src},"
                          f"MATCH($B$3&\"|{rank}\",'Priority Detail'!$I:$I,0)),\"\")")
            cell.font = TABLE_BODY_FONT
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.alignment = LEFT_ALIGN
        ws.cell(r, 4).number_format = "0.0%"

    for col, w in {"A": 35.0, "B": 28.0, "D": 20.0, "E": 18.0}.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 33.75
    ws.row_dimensions[3].height = 27.0
    ws.freeze_panes = "A10"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = "A1:H27"

    # ── 5-8. Data sheets ──
    data_sheets = [
        ("Organisation Summary", "One row per organisation. Sort or filter by respondent count, score or agreement.",
         odf, [47.71, 37, 14.14, 25, 22.71, 17.43, 25.71, 18.14, 17, 14.71, 27.29, 13, 19.43, 20.71, 12.57, 25.43, 20, 13], "A23", True),
        ("Section Summary", "Section-level maturity and internal agreement for each organisation.",
         sdf, [18.29, 15.71, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13], "A5", False),
        ("Question Summary", "Question-level statistics. Use Review flag to find low-consensus scored questions.",
         qdf, [45, 38, 13, 72, 12.29, 13, 60, 13, 10.86, 13, 13, 10.86, 9.71, 13, 13, 11.14, 11.57, 13], "D5", False),
        ("Answer Distribution", "Counts and percentages by answer option. Multi-select totals can exceed 100%.",
         ddf, [45, 20.57, 13.14, 72, 17.29, 65, 13.71, 16.29, 10.86, 16], "A5", False),
    ]

    for name, sub, frame, widths, freeze, add_filter in data_sheets:
        ws = wb.create_sheet(name)
        apply_data_sheet(ws, name, sub, list(frame.columns),
                         frame.where(pd.notna(frame), None).values.tolist(),
                         widths, freeze, add_filter)
        if name == "Organisation Summary":
            for r in range(5, 5 + len(frame)):
                ws.cell(r, 10).number_format = ws.cell(r, 13).number_format = ws.cell(r, 18).number_format = "0.0%"
        if name == "Section Summary":
            for r in range(5, 5 + len(frame)):
                ws.cell(r, 10).number_format = ws.cell(r, 11).number_format = "0.0%"
        if name == "Question Summary":
            for r in range(5, 5 + len(frame)):
                ws.cell(r, 9).number_format = ws.cell(r, 16).number_format = "0.0%"
        if name == "Answer Distribution":
            for r in range(5, 5 + len(frame)):
                ws.cell(r, 9).number_format = "0.0%"

    # ── 9. Raw Answers ──
    rframe = df.copy()
    for c in RAW_OUT:
        if c not in rframe.columns: rframe[c] = None
    rframe = rframe[RAW_OUT]
    ws = wb.create_sheet("Raw Answers")
    apply_data_sheet(ws, None, "Imported source rows with an added Standard section field. Personal email is intentionally omitted from this analysis copy.",
                     RAW_OUT, rframe.where(pd.notna(rframe), None).values.tolist(),
                     [22.86, 17.71, 38, 10.71, 13.29, 70, 65, 14.71, 12, 12.71, 28, 13, 45, 17.14, 15.43, 20.14, 10.71, 13, 13, 12.71, 13, 13, 14.57],
                     "A5", False)
    # Apply navy banner even without title text
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(vertical="center")

    # ── 10. Priority Detail (hidden) ──
    rows = []
    for org, g in qdf[qdf["Scored question"] == True].groupby("Organisation name"):
        gg = g.copy()
        gg["_review"] = (gg["Review flag"] == "Review").astype(int)
        gg = gg.sort_values(["_review", "Normalised score", "Question ID"], ascending=[False, True, True])
        for rank, (_, x) in enumerate(gg.iterrows(), 1):
            rows.append([org, x["Question ID"], x["Question"], x["Most common answer"],
                         x["Normalised score"], x["Agreement"], x["Review flag"], rank, f"{org}|{rank}"])

    ws = wb.create_sheet("Priority Detail")
    cols = ["Organisation", "Question ID", "Question", "Most common answer",
            "Normalised score", "Agreement", "Review flag", "Priority rank", "Lookup key"]
    for c, h in enumerate(cols, 1):
        cell = ws.cell(1, c)
        cell.value = h
        cell.font = TABLE_HEADER_FONT
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = HEADER_ALIGN
        cell.border = NO_BORDER
    for r, row in enumerate(rows, 2):
        bg = WHITE if (r - 2) % 2 == 0 else ROW_ALT
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c)
            cell.value = v
            cell.font = TABLE_BODY_FONT
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = NO_BORDER
    ws.sheet_state = "hidden"
    ws.page_setup.orientation = "portrait"

    # ── Finalise ──
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.active = wb.sheetnames.index("Dashboard")
    wb.save(out)


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    input_path = Path(INPUT_FILE)
    if not input_path.exists():
        print(f"ERROR: Input file not found: {INPUT_FILE}")
        return

    print(f"Reading: {INPUT_FILE}")
    df = load_data(str(input_path), CFG)
    print(f"  {len(df)} rows after dedup, {df['Respondent ID'].nunique()} respondents")

    print("Calculating...")
    odf, sdf, qdf, ddf = calculate(df, CFG)
    print(f"  {len(odf)} organisations, {len(sdf)} section rows, {len(qdf)} question rows")

    print("Building workbook...")
    build(df, odf, sdf, qdf, ddf, CFG, OUTPUT_FILE)
    print(f"Saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
