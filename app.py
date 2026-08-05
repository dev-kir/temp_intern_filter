"""
Super Filter — Desktop GUI
============================
Import a SARI survey Excel, view the Organisation Summary,
switch organisations via dropdown, and export the full workbook.

Usage:
    python app.py
"""

import io
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from collections import Counter

import pandas as pd

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG (mirrors super_filter.py)
# ═══════════════════════════════════════════════════════════════════════════

CFG = {
    "high_consensus": 0.80,
    "moderate_consensus": 0.60,
    "minimum_respondents_for_agreement": 2,
    "tier_boundaries": [0.20, 0.40, 0.60, 0.80],
    "section_order": [
        "Strategy & Leadership", "Governance, Policy & Ethics",
        "Talent & Organisational Culture", "Infrastructure & Technology",
        "Data Management & Readiness", "AI Implementation & Potential Impact",
        "Investment",
    ],
    "section_prefixes": {
        "strategy": "Strategy & Leadership", "governance": "Governance, Policy & Ethics",
        "talent": "Talent & Organisational Culture", "infrastructure": "Infrastructure & Technology",
        "data": "Data Management & Readiness", "aiapp": "AI Implementation & Potential Impact",
        "investment": "Investment", "background": "Background",
    },
    "multi_select_question_ids": ["background_2", "background_3", "background_4"],
}

# ═══════════════════════════════════════════════════════════════════════════
# DATA PROCESSING (same logic as super_filter.py)
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


def mode_det(values):
    vals = [str(v).strip() for v in values if pd.notna(v) and str(v).strip()]
    if not vals: return "", 0
    c = Counter(vals); m = max(c.values())
    winners = sorted(k for k, v in c.items() if v == m)
    return winners[0], m


def load_data(path, cfg):
    df = pd.read_excel(path, sheet_name="Answers", engine="openpyxl", dtype=object)
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
    high = cfg["high_consensus"]; mod = cfg["moderate_consensus"]
    min_n = cfg["minimum_respondents_for_agreement"]
    sections = cfg["section_order"]; multi = set(cfg["multi_select_question_ids"])

    qrows = []
    for (org, sec, qid, qtext), g in df.groupby(
        ["Organisation name", "Standard section", "Question ID", "Question"],
        dropna=False, sort=True,
    ):
        n = g["Respondent ID"].nunique(); mode, count = mode_det(g["Answer"])
        cons = count / n if n else 0; scored = (g["Max score"].fillna(0) > 0).any()
        scores = g.loc[g["Max score"].fillna(0) > 0, "Answer score"].dropna().astype(float)
        maxs = g.loc[g["Max score"].fillna(0) > 0, "Max score"].dropna().astype(float)
        avg = scores.mean() if len(scores) else None
        med = scores.median() if len(scores) else None
        mn = scores.min() if len(scores) else None; mx = scores.max() if len(scores) else None
        rng = (mx - mn) if len(scores) else None
        sd = scores.std(ddof=1) if len(scores) > 1 else (0 if len(scores) == 1 else None)
        norm = (avg / maxs.mean()) if len(scores) and maxs.mean() else None
        agr = agreement(cons, n, min_n, high, mod)
        flag = "Review" if n >= min_n and cons < mod else ""
        qrows.append([org, sec, qid, qtext, n, bool(scored), mode, count, cons, avg, med, mn, mx, rng, sd, norm, agr, flag])

    qcols = ["Organisation name", "Section", "Question ID", "Question", "Respondents",
             "Scored question", "Most common answer", "Most common count", "Consensus",
             "Average score", "Median score", "Minimum score", "Maximum score",
             "Score range", "Standard deviation", "Normalised score", "Agreement", "Review flag"]
    qdf = pd.DataFrame(qrows, columns=qcols)

    srows = []
    for org in sorted(df["Organisation name"].unique()):
        for sec in sections:
            g = df[(df["Organisation name"] == org) & (df["Standard section"] == sec) & (df["Max score"].fillna(0) > 0)]
            if g.empty: continue
            scores = g["Answer score"].dropna().astype(float); maxs = g["Max score"].dropna().astype(float)
            n = g["Respondent ID"].nunique(); nq = g["Question ID"].nunique()
            avg = scores.mean(); med = scores.median(); mn = scores.min(); mx = scores.max()
            maxp = maxs.mean(); norm = avg / maxp if maxp else None
            qs = qdf[(qdf["Organisation name"] == org) & (qdf["Section"] == sec)]
            cons = qs["Consensus"].mean() if len(qs) else None
            agr = agreement(cons or 0, n, min_n, high, mod)
            srows.append([org, sec, n, nq, avg, med, mn, mx, maxp, norm, cons, agr])

    scols = ["Organisation name", "Section", "Respondents", "Questions",
             "Average score", "Median score", "Minimum score", "Maximum score",
             "Max possible", "Normalised score", "Average consensus", "Agreement"]
    sdf = pd.DataFrame(srows, columns=scols)

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
        typ = first.get("Organisation type", ""); size = first.get("Organisation size", "")
        sector = first.get("PCDS sector", ""); latest = g.iloc[-1].get("Submitted at", "")
        depts = g["Department"].dropna().astype(str).replace("", pd.NA).dropna().nunique()
        roles = g["Role level"].dropna().astype(str).replace("", pd.NA).dropna().nunique()
        interpretation = "Single respondent: perception only" if n < 2 else ("Directional: small sample" if n < 3 else "Multi-respondent view")
        b = cfg["tier_boundaries"]
        nxt = (1 - overall) if overall >= b[-1] else min(x for x in b if x > overall) - overall
        orows.append([org, typ, n, depts, roles, size, sector, latest, avg, overall, strongest, weakest, cons, reviews, agr, interpretation, tier(overall, b), nxt])

    ocols = ["Organisation name", "Organisation type", "Respondents", "Departments represented",
             "Role levels represented", "Organisation size", "Sector", "Latest submission",
             "Average score", "Overall score", "Strongest section", "Weakest section",
             "Average consensus", "Questions for review", "Agreement", "Interpretation",
             "Maturity tier", "Distance to next tier"]
    odf = pd.DataFrame(orows, columns=ocols)
    return odf, sdf, qdf


# ═══════════════════════════════════════════════════════════════════════════
# GUI APP
# ═══════════════════════════════════════════════════════════════════════════

NAVY = "#163A63"
CARD_BG = "#F5F5F5"
WHITE = "#FFFFFF"
GREEN = "#4CAF50"
ORANGE = "#F4B400"
RED = "#DB4437"
TEXT = "#1F1F1F"
GREY = "#6E6E6E"
HEADER_BG = "#2F5496"


class SuperFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Super Filter — SARI Organisation Summary")
        self.root.geometry("1300x850")
        self.root.configure(bg=WHITE)

        self.df = None
        self.odf = None
        self.sdf = None
        self.qdf = None
        self.orgs = []
        self.selected_org = tk.StringVar()
        self.filepath = None

        self._build_ui()
        self._setup_drag_drop()

    def _setup_drag_drop(self):
        """Enable drag-and-drop of Excel files onto the window."""
        try:
            from tkinterdnd2 import DND_FILES
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
        except ImportError:
            pass

    def _on_drop(self, event):
        path = event.data.strip("{}")
        if path.endswith(".xlsx"):
            self._process_file(path)

    def _build_ui(self):
        # ── Top bar ──
        top = tk.Frame(self.root, bg=NAVY, height=50)
        top.pack(fill=tk.X)
        top.pack_propagate(False)

        title_lbl = tk.Label(top, text="SARI Organisation Summary", font=("Calibri", 16, "bold"),
                             fg=WHITE, bg=NAVY)
        title_lbl.pack(side=tk.LEFT, padx=20, pady=10)

        self.status_lbl = tk.Label(top, text="Open an Excel file or drag & drop it here",
                                   font=("Calibri", 10), fg="#AABBCC", bg=NAVY)
        self.status_lbl.pack(side=tk.RIGHT, padx=20, pady=10)

        # ── Toolbar ──
        toolbar = tk.Frame(self.root, bg="#F0F0F0", height=40)
        toolbar.pack(fill=tk.X)
        toolbar.pack_propagate(False)

        tk.Button(toolbar, text="Open Excel...", command=self._open_file,
                  font=("Calibri", 10), bg=WHITE, relief=tk.FLAT, padx=15, pady=4,
                  cursor="hand2").pack(side=tk.LEFT, padx=10, pady=5)

        tk.Label(toolbar, text="Organisation:", font=("Calibri", 10, "bold"),
                 bg="#F0F0F0", fg=TEXT).pack(side=tk.LEFT, padx=(20, 5), pady=5)

        self.org_dropdown = ttk.Combobox(toolbar, textvariable=self.selected_org,
                                         state="readonly", font=("Calibri", 10), width=50)
        self.org_dropdown.pack(side=tk.LEFT, padx=5, pady=5)
        self.org_dropdown.bind("<<ComboboxSelected>>", self._on_org_change)

        tk.Button(toolbar, text="Export Full Report", command=self._export,
                  font=("Calibri", 10, "bold"), bg=HEADER_BG, fg=WHITE,
                  relief=tk.FLAT, padx=15, pady=4, cursor="hand2").pack(side=tk.RIGHT, padx=10, pady=5)

        # ── Main content area ──
        self.main_frame = tk.Frame(self.root, bg=WHITE)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.placeholder = tk.Label(self.main_frame,
                                    text="Open a SARI survey Excel file to view the organisation summary.\n\n"
                                         "You can also drag & drop an .xlsx file onto this window.",
                                    font=("Calibri", 12), fg=GREY, bg=WHITE, justify=tk.CENTER)
        self.placeholder.pack(expand=True)

    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Select SARI Survey Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not path:
            return
        self._process_file(path)

    def _process_file(self, path):
        self.filepath = path
        self.status_lbl.config(text="Processing...")
        self.root.update()

        try:
            self.df = load_data(path, CFG)
            self.odf, self.sdf, self.qdf = calculate(self.df, CFG)
            self.orgs = sorted(self.odf["Organisation name"].tolist())
            self.org_dropdown["values"] = self.orgs
            if self.orgs:
                self.selected_org.set(self.orgs[0])
            self.status_lbl.config(
                text=f"{Path(path).name} — {len(self.orgs)} orgs, {self.df['Respondent ID'].nunique()} respondents")
            self._render_summary()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_lbl.config(text="Error loading file")

    def _on_org_change(self, event=None):
        self._render_summary()

    def _score_color(self, val):
        if val is None or pd.isna(val):
            return TEXT
        if val >= 0.75: return GREEN
        if val >= 0.50: return ORANGE
        return RED

    def _render_summary(self):
        for w in self.main_frame.winfo_children():
            w.destroy()

        org = self.selected_org.get()
        if not org or self.odf is None:
            return

        row_data = self.odf[self.odf["Organisation name"] == org]
        if row_data.empty:
            return
        d = row_data.iloc[0]

        # ── KPI Cards Row 1 ──
        cards_frame = tk.Frame(self.main_frame, bg=WHITE)
        cards_frame.pack(fill=tk.X, pady=(0, 10))

        overall_val = d["Overall score"]
        kpis = [
            ("Overall Score", f"{overall_val:.1%}" if pd.notna(overall_val) else "-"),
            ("Maturity Tier", str(d["Maturity tier"])),
            ("Respondents", str(int(d["Respondents"]))),
            ("Organisation Type", str(d["Organisation type"])),
            ("Organisation Size", str(d["Organisation size"])),
            ("Sector", str(d["Sector"])),
        ]

        for label, value in kpis:
            card = tk.Frame(cards_frame, bg=CARD_BG, padx=12, pady=8, relief=tk.FLAT, bd=0)
            card.pack(side=tk.LEFT, padx=5, fill=tk.BOTH, expand=True)

            tk.Label(card, text=label, font=("Calibri", 9, "bold"), fg=TEXT, bg=CARD_BG,
                     anchor=tk.W).pack(fill=tk.X)
            val_color = TEXT
            if label == "Overall Score":
                val_color = self._score_color(overall_val)
            tk.Label(card, text=value, font=("Calibri", 14, "bold"), fg=val_color, bg=CARD_BG,
                     anchor=tk.W).pack(fill=tk.X)

        # ── KPI Cards Row 2 ──
        cards_frame2 = tk.Frame(self.main_frame, bg=WHITE)
        cards_frame2.pack(fill=tk.X, pady=(0, 10))

        kpis2 = [
            ("Strongest Section", str(d["Strongest section"])),
            ("Weakest Section", str(d["Weakest section"])),
            ("Agreement", str(d["Agreement"])),
            ("Interpretation", str(d["Interpretation"])),
            ("Questions for Review", str(int(d["Questions for review"]))),
            ("Distance to Next Tier", f"{d['Distance to next tier']:.1%}" if pd.notna(d["Distance to next tier"]) else "-"),
        ]

        for label, value in kpis2:
            card = tk.Frame(cards_frame2, bg=CARD_BG, padx=12, pady=8, relief=tk.FLAT, bd=0)
            card.pack(side=tk.LEFT, padx=5, fill=tk.BOTH, expand=True)

            tk.Label(card, text=label, font=("Calibri", 9, "bold"), fg=TEXT, bg=CARD_BG,
                     anchor=tk.W).pack(fill=tk.X)
            tk.Label(card, text=value, font=("Calibri", 14, "bold"), fg=TEXT, bg=CARD_BG,
                     anchor=tk.W).pack(fill=tk.X)

        # ── Section Scores Table ──
        sec_frame = tk.Frame(self.main_frame, bg=WHITE)
        sec_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        tk.Label(sec_frame, text="Section Scores", font=("Calibri", 12, "bold"),
                 fg=TEXT, bg=WHITE).pack(anchor=tk.W, pady=(0, 5))

        tbl_header = tk.Frame(sec_frame, bg=HEADER_BG)
        tbl_header.pack(fill=tk.X)
        for h, w in [("Section", 35), ("Score", 12), ("Agreement", 12)]:
            tk.Label(tbl_header, text=h, font=("Calibri", 9, "bold"), fg=WHITE, bg=HEADER_BG,
                     width=w, anchor=tk.W, padx=8, pady=4).pack(side=tk.LEFT)

        sec_data = self.sdf[self.sdf["Organisation name"] == org]
        for i, (_, sr) in enumerate(sec_data.iterrows()):
            bg_color = WHITE if i % 2 == 0 else "#F8F9FA"
            row_frame = tk.Frame(sec_frame, bg=bg_color)
            row_frame.pack(fill=tk.X)

            score_val = sr["Normalised score"]
            score_text = f"{score_val:.1%}" if pd.notna(score_val) else "-"
            score_color = self._score_color(score_val)

            tk.Label(row_frame, text=sr["Section"], font=("Calibri", 9), fg=TEXT, bg=bg_color,
                     width=35, anchor=tk.W, padx=8, pady=3).pack(side=tk.LEFT)
            tk.Label(row_frame, text=score_text, font=("Calibri", 9, "bold"), fg=score_color,
                     bg=bg_color, width=12, anchor=tk.W, padx=8, pady=3).pack(side=tk.LEFT)
            tk.Label(row_frame, text=sr["Agreement"], font=("Calibri", 9), fg=TEXT, bg=bg_color,
                     width=12, anchor=tk.W, padx=8, pady=3).pack(side=tk.LEFT)

        # ── Priority Questions ──
        pq_frame = tk.Frame(self.main_frame, bg=WHITE)
        pq_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))

        tk.Label(pq_frame, text="Priority Questions for Review", font=("Calibri", 12, "bold"),
                 fg=TEXT, bg=WHITE).pack(anchor=tk.W, pady=(0, 5))

        pq_data = self.qdf[(self.qdf["Organisation name"] == org) & (self.qdf["Scored question"] == True)]
        pq_data = pq_data.copy()
        pq_data["_review"] = (pq_data["Review flag"] == "Review").astype(int)
        pq_data = pq_data.sort_values(["_review", "Normalised score", "Question ID"],
                                      ascending=[False, True, True]).head(5)

        pq_header = tk.Frame(pq_frame, bg=HEADER_BG)
        pq_header.pack(fill=tk.X)
        for h, w in [("Question ID", 18), ("Question", 55), ("Most Common Answer", 40),
                      ("Score", 8), ("Agreement", 10)]:
            tk.Label(pq_header, text=h, font=("Calibri", 9, "bold"), fg=WHITE, bg=HEADER_BG,
                     width=w, anchor=tk.W, padx=8, pady=4).pack(side=tk.LEFT)

        for i, (_, pr) in enumerate(pq_data.iterrows()):
            bg_color = WHITE if i % 2 == 0 else "#F8F9FA"
            row_frame = tk.Frame(pq_frame, bg=bg_color)
            row_frame.pack(fill=tk.X)

            score_val = pr["Normalised score"]
            score_text = f"{score_val:.1%}" if pd.notna(score_val) else "-"
            score_color = self._score_color(score_val)

            tk.Label(row_frame, text=pr["Question ID"], font=("Calibri", 9), fg=TEXT, bg=bg_color,
                     width=18, anchor=tk.W, padx=8, pady=3).pack(side=tk.LEFT)
            tk.Label(row_frame, text=str(pr["Question"])[:70], font=("Calibri", 9), fg=TEXT,
                     bg=bg_color, width=55, anchor=tk.W, padx=8, pady=3).pack(side=tk.LEFT)
            tk.Label(row_frame, text=str(pr["Most common answer"])[:50], font=("Calibri", 9),
                     fg=TEXT, bg=bg_color, width=40, anchor=tk.W, padx=8, pady=3).pack(side=tk.LEFT)
            tk.Label(row_frame, text=score_text, font=("Calibri", 9, "bold"), fg=score_color,
                     bg=bg_color, width=8, anchor=tk.W, padx=8, pady=3).pack(side=tk.LEFT)
            tk.Label(row_frame, text=pr["Agreement"], font=("Calibri", 9), fg=TEXT, bg=bg_color,
                     width=10, anchor=tk.W, padx=8, pady=3).pack(side=tk.LEFT)

    def _export(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Load a file first.")
            return

        path = filedialog.asksaveasfilename(
            title="Save Processed Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return

        try:
            self.status_lbl.config(text="Exporting...")
            self.root.update()

            import super_filter as sf
            odf, sdf, qdf, ddf = sf.calculate(self.df, sf.CFG)
            sf.build(self.df, odf, sdf, qdf, ddf, sf.CFG, path)

            messagebox.showinfo("Success", f"Saved to:\n{path}")
            self.status_lbl.config(text=f"Exported: {Path(path).name}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            self.status_lbl.config(text="Export failed")


def main():
    root = tk.Tk()
    try:
        root.lift()
        root.attributes("-topmost", True)
        root.after(100, lambda: root.attributes("-topmost", False))
    except: pass
    SuperFilterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
