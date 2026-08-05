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
import shutil
from copy import copy
from openpyxl import Workbook, load_workbook
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
# The report design lives in this template, not in code. See build().
TEMPLATE_FILE = "report_template.xlsx"

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

def _clear_and_fill(ws, headers, rows, start_row):
    """Replace the data region of a template sheet, keeping its styling.

    Row `start_row` in the template carries the intended cell styles and number
    formats for a data row, so it is captured first and stamped onto every row
    written after it.
    """
    style_src = [(ws.cell(start_row, c)._style,
                  ws.cell(start_row, c).number_format,
                  copy(ws.cell(start_row, c).alignment))
                 for c in range(1, len(headers) + 1)]
    for r in range(start_row, max(ws.max_row, start_row) + 1):
        for c in range(1, max(ws.max_column, len(headers)) + 1):
            ws.cell(r, c).value = None
    for c, h in enumerate(headers, 1):
        ws.cell(start_row - 1, c).value = h
    for i, row in enumerate(rows, start_row):
        for c, v in enumerate(row, 1):
            cell = ws.cell(i, c)
            cell.value = v
            st, nf, al = style_src[c - 1]
            cell._style = copy(st)
            cell.number_format = nf
            cell.alignment = copy(al)
    last = ws.cell(start_row - 1, len(headers)).column_letter
    ws.auto_filter.ref = f"A{start_row-1}:{last}{start_row + max(len(rows), 1) - 1}"


def build(df, odf, sdf, qdf, ddf, cfg, out):
    """Fill the presentation template with freshly calculated data.

    The template owns everything visual: fonts, fills, merged cells, column
    widths, the chart, print settings, and the Dashboard / Organisation Report
    formulas. This function replaces only the DATA rows.

    Building the presentation in code never matched a workbook designed in
    Excel, and it also produced a chart with both axes at axPos="l", which
    Excel rejects outright. Separating calculation from presentation removes a
    whole class of that problem.
    """
    tpl = Path(TEMPLATE_FILE)
    if not tpl.exists():
        raise SystemExit(
            f"Template not found: {TEMPLATE_FILE}\n"
            "It carries the report design and must sit beside this script.")
    shutil.copy2(tpl, out)
    wb = load_workbook(out, data_only=False)

    def frame_rows(frame):
        return frame.where(pd.notna(frame), None).values.tolist()

    # ── Lists: one organisation per row, drives the dropdown ──
    ws = wb["Lists"]
    for r in range(1, max(ws.max_row, 1) + 1):
        ws.cell(r, 1).value = None
    orgs = list(odf["Organisation name"])
    for r, o in enumerate(orgs, 1):
        ws.cell(r, 1, o)
    ws.sheet_state = "hidden"

    # ── Data sheets: headers on row 4, data from row 5 ──
    for name, frame in (("Organisation Summary", odf),
                        ("Section Summary", sdf),
                        ("Question Summary", qdf),
                        ("Answer Distribution", ddf)):
        _clear_and_fill(wb[name], list(frame.columns), frame_rows(frame), 5)

    # ── Raw Answers ──
    rframe = df.copy()
    for c in RAW_OUT:
        if c not in rframe.columns:
            rframe[c] = None
    rframe = rframe[RAW_OUT]
    _clear_and_fill(wb["Raw Answers"], RAW_OUT, frame_rows(rframe), 5)

    # ── Priority Detail: hidden lookup table, header on row 1 ──
    prows = []
    for org, g in qdf[qdf["Scored question"] == True].groupby("Organisation name"):
        gg = g.copy()
        gg["_review"] = (gg["Review flag"] == "Review").astype(int)
        gg = gg.sort_values(["_review", "Normalised score", "Question ID"],
                            ascending=[False, True, True])
        for rank, (_, x) in enumerate(gg.iterrows(), 1):
            prows.append([org, x["Question ID"], x["Question"], x["Most common answer"],
                          x["Normalised score"], x["Agreement"], x["Review flag"],
                          rank, f"{org}|{rank}"])
    pcols = ["Organisation", "Question ID", "Question", "Most common answer",
             "Normalised score", "Agreement", "Review flag", "Priority rank", "Lookup key"]
    _clear_and_fill(wb["Priority Detail"], pcols, prows, 2)
    wb["Priority Detail"].auto_filter.ref = None
    wb["Priority Detail"].sheet_state = "hidden"

    # ── Dropdown: a dynamic named range, so it grows with the data ──
    if "OrganisationList" in wb.defined_names:
        del wb.defined_names["OrganisationList"]
    wb.defined_names["OrganisationList"] = DefinedName(
        "OrganisationList",
        attr_text="Lists!$A$1:INDEX(Lists!$A:$A,COUNTA(Lists!$A:$A))")

    first_org = orgs[0] if orgs else ""
    for sheet, cell_ref in (("Dashboard", "B4"), ("Organisation Report", "B3")):
        ws = wb[sheet]
        ws[cell_ref] = first_org
        dv = DataValidation(type="list", formula1="=OrganisationList", allow_blank=False,
                            showErrorMessage=True, showInputMessage=True,
                            errorTitle="Invalid organisation",
                            error="Select an organisation from the dropdown list.",
                            promptTitle="Organisation selector",
                            prompt="Choose an organisation from the list.")
        ws.add_data_validation(dv)
        dv.add(ws[cell_ref])

    wb.calculation.fullCalcOnLoad = True
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
