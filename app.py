"""
Super Filter — Desktop GUI
============================
Import a SARI survey Excel, view the full Organisation Summary table,
filter by organisation, and export the complete workbook.

Usage:
    python app.py
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from collections import Counter

import pandas as pd

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG
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

NAVY = "#163A63"
CARD_BG = "#F5F5F5"
WHITE = "#FFFFFF"
GREEN = "#4CAF50"
ORANGE = "#F4B400"
RED = "#DB4437"
TEXT = "#1F1F1F"
GREY = "#6E6E6E"
HEADER_BG = "#2F5496"

# ═══════════════════════════════════════════════════════════════════════════
# DATA PROCESSING
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

class SuperFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Super Filter — SARI Organisation Summary")
        self.root.geometry("1400x850")
        self.root.configure(bg=WHITE)

        self.df = None
        self.odf = None
        self.sdf = None
        self.qdf = None
        self.orgs = []
        self.filepath = None
        self.sort_col = None
        self.sort_asc = True

        self._build_ui()

    def _build_ui(self):
        top = tk.Frame(self.root, bg=NAVY, height=50)
        top.pack(fill=tk.X)
        top.pack_propagate(False)

        tk.Label(top, text="SARI Organisation Summary", font=("Calibri", 16, "bold"),
                 fg=WHITE, bg=NAVY).pack(side=tk.LEFT, padx=20, pady=10)

        self.status_lbl = tk.Label(top, text="Open an Excel file to begin",
                                   font=("Calibri", 10), fg="#AABBCC", bg=NAVY)
        self.status_lbl.pack(side=tk.RIGHT, padx=20, pady=10)

        toolbar = tk.Frame(self.root, bg="#F0F0F0", height=40)
        toolbar.pack(fill=tk.X)
        toolbar.pack_propagate(False)

        tk.Button(toolbar, text="Open Excel...", command=self._open_file,
                  font=("Calibri", 10), bg=WHITE, relief=tk.FLAT, padx=15, pady=4,
                  cursor="hand2").pack(side=tk.LEFT, padx=10, pady=5)

        tk.Label(toolbar, text="Filter:", font=("Calibri", 10, "bold"),
                 bg="#F0F0F0", fg=TEXT).pack(side=tk.LEFT, padx=(20, 5), pady=5)

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._refresh_table())
        search_entry = tk.Entry(toolbar, textvariable=self.search_var, font=("Calibri", 10), width=30)
        search_entry.pack(side=tk.LEFT, padx=5, pady=5)

        tk.Button(toolbar, text="Export Full Report", command=self._export,
                  font=("Calibri", 10, "bold"), bg=HEADER_BG, fg=WHITE,
                  relief=tk.FLAT, padx=15, pady=4, cursor="hand2").pack(side=tk.RIGHT, padx=10, pady=5)

        # ── Main table area ──
        self.main_frame = tk.Frame(self.root, bg=WHITE)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.placeholder = tk.Label(self.main_frame,
                                    text="Open a SARI survey Excel file to view the organisation summary.",
                                    font=("Calibri", 12), fg=GREY, bg=WHITE)
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
            self.status_lbl.config(
                text=f"{Path(path).name} — {len(self.orgs)} orgs, {self.df['Respondent ID'].nunique()} respondents")
            self._build_table()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_lbl.config(text="Error loading file")

    def _build_table(self):
        for w in self.main_frame.winfo_children():
            w.destroy()

        # ── Treeview with scrollbars ──
        tree_frame = tk.Frame(self.main_frame, bg=WHITE)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = [
            "Organisation name", "Organisation type", "Respondents",
            "Departments", "Role levels", "Org size", "Sector",
            "Latest submission", "Avg score", "Overall score",
            "Strongest section", "Weakest section", "Consensus",
            "Review qs", "Agreement", "Interpretation", "Maturity tier", "Distance"
        ]
        display_cols = [
            "Organisation name", "Organisation type", "Resp",
            "Depts", "Roles", "Size", "Sector",
            "Latest", "Avg", "Overall",
            "Strongest", "Weakest", "Consensus",
            "Review", "Agreement", "Interpretation", "Tier", "Dist"
        ]

        self.tree = ttk.Treeview(tree_frame, columns=display_cols, show="headings", height=25)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        widths = [300, 200, 50, 50, 50, 60, 150, 120, 60, 70, 180, 180, 70, 50, 80, 180, 120, 60]
        for i, (col, display, w) in enumerate(zip(columns, display_cols, widths)):
            self.tree.heading(display, text=display, command=lambda c=col: self._sort_by(c))
            self.tree.column(display, width=w, minwidth=40)

        self._refresh_table()

    def _refresh_table(self):
        if self.odf is None:
            return
        self.tree.delete(*self.tree.get_children())

        search = self.search_var.get().lower()
        for _, row in self.odf.iterrows():
            if search and search not in str(row["Organisation name"]).lower():
                continue

            overall = row["Overall score"]
            overall_str = f"{overall:.1%}" if pd.notna(overall) else "-"
            avg_str = f"{row['Average score']:.2f}" if pd.notna(row["Average score"]) else "-"
            dist_str = f"{row['Distance to next tier']:.1%}" if pd.notna(row["Distance to next tier"]) else "-"

            values = (
                row["Organisation name"],
                row["Organisation type"],
                int(row["Respondents"]),
                int(row["Departments represented"]),
                int(row["Role levels represented"]),
                row["Organisation size"],
                row["Sector"],
                row["Latest submission"],
                avg_str,
                overall_str,
                row["Strongest section"],
                row["Weakest section"],
                f"{row['Average consensus']:.2f}" if pd.notna(row["Average consensus"]) else "-",
                int(row["Questions for review"]),
                row["Agreement"],
                row["Interpretation"],
                row["Maturity tier"],
                dist_str,
            )

            # Color tag based on overall score
            tag = "normal"
            if pd.notna(overall):
                if overall >= 0.75: tag = "high"
                elif overall >= 0.50: tag = "mid"
                elif overall > 0: tag = "low"

            self.tree.insert("", tk.END, values=values, tags=(tag,))

        self.tree.tag_configure("high", foreground=GREEN)
        self.tree.tag_configure("mid", foreground=ORANGE)
        self.tree.tag_configure("low", foreground=RED)

    def _sort_by(self, col):
        if self.sort_col == col:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = col
            self.sort_asc = False

        self.odf = self.odf.sort_values(col, ascending=self.sort_asc)
        self._refresh_table()

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
            sf.build(self.filepath, sf.TEMPLATE_FILE, path)

            messagebox.showinfo("Success", f"Saved to:\n{path}")
            self.status_lbl.config(text=f"Exported: {Path(path).name}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            self.status_lbl.config(text="Export failed")


def main():
    root = tk.Tk()

    # Set icon if available
    icon_path = Path(__file__).with_name("icon.png")
    if icon_path.exists():
        try:
            img = tk.PhotoImage(file=str(icon_path))
            root.iconphoto(True, img)
        except Exception:
            pass

    try:
        root.lift()
        root.attributes("-topmost", True)
        root.after(100, lambda: root.attributes("-topmost", False))
    except Exception:
        pass

    SuperFilterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
