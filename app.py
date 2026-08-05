"""
Super Filter — Desktop GUI
==========================
Drop in a raw SARI survey export, read the Organisation Summary on screen, then
export the complete formatted workbook.

The table shown here is not recalculated by this file. It is exactly the
`Organisation Summary` that super_filter.compute() produces and that
super_filter.build() writes into the workbook, so the screen and the file
cannot drift apart.

Run:
    python app.py            (or double-click run_app.command / run_app.bat)
"""

import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import super_filter as sf

# Drag-and-drop is optional. Without tkinterdnd2 the drop zone still works as a
# click target, so the app never fails to start over a missing extra.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False

# ── palette, matched to the workbook's own styling ──
NAVY = "#163A63"
HEADER_BG = "#2F5496"
WHITE = "#FFFFFF"
DROPZONE = "#F7F9FC"
TEXT = "#1F1F1F"
GREY = "#6E6E6E"
LINE = "#DDE3EA"      # the only divider in the window, under the column headings
STRIPE = "#F7F9FB"    # alternating row wash, quiet enough not to read as a band

ALL_ORGS = "All organisations"

# The on-screen table shows what is worth scanning and nothing more, so it fits the
# window without sideways scrolling. Every one of the 18 fields is still written to
# the exported workbook — this list only governs the display.
#
# label, index in sf.ORG_HEADERS, width, formatter, anchor, stretch
COLUMNS = [
    ("Organisation name",   0, 280, str,   tk.W,      True),
    ("Type",                1, 140, str,   tk.W,      True),
    ("Resp",                2,  52, "int", tk.CENTER, False),
    ("Size",                5,  72, str,   tk.CENTER, False),
    ("Sector",              6, 120, str,   tk.W,      True),
    ("Avg score",           8,  70, "f2",  tk.CENTER, False),
    ("Overall",             9,  76, "pct", tk.CENTER, False),
    ("Strongest section",  10, 140, str,   tk.W,      True),
    ("Weakest section",    11, 140, str,   tk.W,      True),
    ("Agreement",          14,  86, str,   tk.CENTER, False),
    ("Maturity tier",      16, 112, str,   tk.W,      False),
]

# Shown only in the exported workbook: Departments represented, Role levels
# represented, Latest submission, Average consensus, Questions for review,
# Interpretation, Distance to next tier.


def fmt(value, kind):
    if value is None or value == "":
        return "-"
    try:
        if kind == "int":
            return str(int(value))
        if kind == "f2":
            return f"{float(value):.2f}"
        if kind == "pct":
            return f"{float(value):.1%}"
    except (TypeError, ValueError):
        return str(value)
    return str(value)


class SuperFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Super Filter — SARI Organisation Summary")
        self.root.geometry("1400x850")
        self.root.minsize(900, 500)
        self.root.configure(bg=WHITE)

        self.data = None          # whatever compute() returned
        self.rows = []            # sf.ORG_HEADERS-shaped rows, current sort order
        self.filepath = None
        self.sort_idx = None
        self.sort_asc = True
        self.busy = False

        self._build_ui()

    # ── layout ────────────────────────────────────────────────────────────
    def _build_ui(self):
        bar = tk.Frame(self.root, bg=NAVY, height=52)
        bar.pack(fill=tk.X)
        bar.pack_propagate(False)
        tk.Label(bar, text="Organisation Summary", font=("Calibri", 16, "bold"),
                 fg=WHITE, bg=NAVY).pack(side=tk.LEFT, padx=20)
        self.status_lbl = tk.Label(bar, text="No file loaded", font=("Calibri", 10),
                                   fg="#AABBCC", bg=NAVY)
        self.status_lbl.pack(side=tk.RIGHT, padx=20)

        # The toolbar shares the table's background so the window reads as one
        # surface rather than a stack of separate grey strips.
        tb = tk.Frame(self.root, bg=WHITE, height=48)
        tb.pack(fill=tk.X)
        tb.pack_propagate(False)

        tk.Button(tb, text="Open Excel...", command=self._choose_file,
                  font=("Calibri", 10), bg=WHITE, relief=tk.FLAT,
                  highlightthickness=1, highlightbackground=LINE,
                  padx=14, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(14, 0), pady=9)

        tk.Label(tb, text="Organisation:", font=("Calibri", 10, "bold"),
                 bg=WHITE, fg=TEXT).pack(side=tk.LEFT, padx=(18, 6))
        self.org_var = tk.StringVar(value=ALL_ORGS)
        self.org_combo = ttk.Combobox(tb, textvariable=self.org_var, state="readonly",
                                      font=("Calibri", 10), width=38, values=[ALL_ORGS])
        self.org_combo.pack(side=tk.LEFT, pady=6)
        self.org_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh())

        tk.Label(tb, text="Search:", font=("Calibri", 10, "bold"),
                 bg=WHITE, fg=TEXT).pack(side=tk.LEFT, padx=(18, 6))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._refresh())
        tk.Entry(tb, textvariable=self.search_var, font=("Calibri", 10),
                 relief=tk.FLAT, highlightthickness=1, highlightbackground=LINE,
                 width=22).pack(side=tk.LEFT, pady=9)

        self.export_btn = tk.Button(tb, text="Export Full Report", command=self._export,
                                    font=("Calibri", 10, "bold"), bg=HEADER_BG, fg=WHITE,
                                    activebackground=NAVY, activeforeground=WHITE,
                                    relief=tk.FLAT, padx=14, pady=4, cursor="hand2")
        self.export_btn.pack(side=tk.RIGHT, padx=(0, 14), pady=9)

        self.main = tk.Frame(self.root, bg=WHITE)
        self.main.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 10))
        self._show_dropzone()

    def _show_dropzone(self):
        for w in self.main.winfo_children():
            w.destroy()

        wrap = tk.Frame(self.main, bg=WHITE)
        wrap.pack(expand=True)

        zone = tk.Frame(wrap, bg=DROPZONE, highlightbackground="#B8C4D4",
                        highlightthickness=2, width=560, height=200)
        zone.pack()
        zone.pack_propagate(False)

        headline = "Drop a SARI export here" if HAS_DND else "Choose a SARI export"
        tk.Label(zone, text=headline, font=("Calibri", 15, "bold"),
                 bg=DROPZONE, fg=NAVY).pack(pady=(52, 6))
        tk.Label(zone, text="or click to browse for an .xlsx file",
                 font=("Calibri", 11), bg=DROPZONE, fg=GREY).pack()
        tk.Label(zone, text="Needs the Answers and Scores sheets",
                 font=("Calibri", 9), bg=DROPZONE, fg=GREY).pack(pady=(14, 0))

        for w in (zone, *zone.winfo_children()):
            w.configure(cursor="hand2")
            w.bind("<Button-1>", lambda e: self._choose_file())

        if HAS_DND:
            zone.drop_target_register(DND_FILES)
            zone.dnd_bind("<<Drop>>", self._on_drop)
        self.dropzone = zone

    def _on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        if not paths:
            return
        path = paths[0]
        if not path.lower().endswith((".xlsx", ".xlsm")):
            messagebox.showwarning("Not a workbook",
                                   f"{Path(path).name} is not an .xlsx file.")
            return
        self._load(path)

    # ── loading ───────────────────────────────────────────────────────────
    def _choose_file(self):
        if self.busy:
            return
        path = filedialog.askopenfilename(
            title="Select SARI survey export",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self._load(path)

    def _load(self, path):
        self.filepath = path
        self._status(f"Reading {Path(path).name}...")
        self.root.update_idletasks()
        try:
            self.data = sf.compute(path)
        except Exception as e:
            messagebox.showerror("Could not read this file", str(e))
            self._status("No file loaded")
            return

        self.rows = list(self.data["organisation"])
        self.sort_idx, self.sort_asc = None, True
        self.org_combo.configure(values=[ALL_ORGS] + self.data["orgs"])
        self.org_var.set(ALL_ORGS)
        self._status(f"{Path(path).name}  —  {len(self.data['orgs'])} organisations, "
                     f"{self.data['respondents']} respondents")
        self._build_table()

    # ── table ─────────────────────────────────────────────────────────────
    def _build_table(self):
        for w in self.main.winfo_children():
            w.destroy()

        frame = tk.Frame(self.main, bg=WHITE)
        frame.pack(fill=tk.BOTH, expand=True)

        labels = [c[0] for c in COLUMNS]
        style = ttk.Style()
        style.theme_use("clam")          # the only theme that honours these colours on macOS
        style.configure("Summary.Treeview",
                        background=WHITE, fieldbackground=WHITE, foreground=TEXT,
                        rowheight=26, font=("Calibri", 11),
                        borderwidth=0, relief=tk.FLAT)
        # Headings sit on the same white as the rows, separated by one hairline only,
        # so the header does not read as a panel bolted on top of the body.
        style.configure("Summary.Treeview.Heading",
                        background=WHITE, foreground=GREY,
                        font=("Calibri", 10, "bold"),
                        relief=tk.FLAT, borderwidth=0, padding=(6, 8))
        style.map("Summary.Treeview.Heading", background=[("active", STRIPE)])
        style.map("Summary.Treeview",
                  background=[("selected", "#DCE7F5")],
                  foreground=[("selected", TEXT)])
        style.layout("Summary.Treeview", [
            ("Summary.Treeview.treearea", {"sticky": "nswe"})])

        # No horizontal scrollbar: the columns above are sized to fit the window.
        self.tree = ttk.Treeview(frame, columns=labels, show="headings",
                                 style="Summary.Treeview", selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        for label, idx, width, _, anchor, stretch in COLUMNS:
            self.tree.heading(label, text=label, anchor=anchor,
                              command=lambda i=idx: self._sort_by(i))
            self.tree.column(label, width=width, minwidth=44,
                             anchor=anchor, stretch=stretch)

        # Text stays the default colour in every column, matching the workbook, where
        # only Overall score carries a colour scale. Tkinter's table colours whole rows
        # and never single cells, so an alternating wash carries readability instead.
        self.tree.tag_configure("odd", background=STRIPE)

        tk.Frame(self.main, bg=LINE, height=1).pack(fill=tk.X)
        self.count_lbl = tk.Label(self.main, text="", font=("Calibri", 9),
                                  bg=WHITE, fg=GREY, anchor=tk.W)
        self.count_lbl.pack(fill=tk.X, pady=(6, 0))

        self._refresh()

    def _refresh(self):
        if self.data is None or not hasattr(self, "tree"):
            return
        self.tree.delete(*self.tree.get_children())

        picked = self.org_var.get()
        search = self.search_var.get().strip().lower()

        shown = 0
        for row in self.rows:
            name = str(row[0])
            if picked != ALL_ORGS and name != picked:
                continue
            if search and search not in name.lower():
                continue
            values = tuple(fmt(row[idx], kind) for _, idx, _, kind, _, _ in COLUMNS)
            self.tree.insert("", tk.END, values=values,
                             tags=("odd",) if shown % 2 else ())
            shown += 1

        self.count_lbl.config(text=f"Showing {shown} of {len(self.rows)} organisations")

    def _sort_by(self, idx):
        self.sort_asc = not self.sort_asc if self.sort_idx == idx else False
        self.sort_idx = idx

        def key(row):
            v = row[idx]
            if v is None or v == "":
                return (1, 0.0, "")
            if isinstance(v, (int, float)):
                return (0, float(v), "")
            return (0, 0.0, str(v).casefold())

        self.rows = sorted(self.rows, key=key, reverse=not self.sort_asc)
        self._refresh()

    # ── export ────────────────────────────────────────────────────────────
    def _export(self):
        if self.busy:
            return
        if self.data is None:
            messagebox.showwarning("Nothing to export", "Open a SARI export first.")
            return

        template = sf.TEMPLATE_FILE
        if not Path(template).exists():
            messagebox.showerror(
                "Template missing",
                f"report_template.xlsx was not found at:\n{template}\n\n"
                "It must sit beside the app. The workbook cannot be built without it.")
            return

        out = filedialog.asksaveasfilename(
            title="Save the full report",
            defaultextension=".xlsx",
            initialfile="SARI_Organisation.xlsx",
            filetypes=[("Excel workbooks", "*.xlsx")])
        if not out:
            return

        self.busy = True
        self.export_btn.config(state=tk.DISABLED, text="Exporting...")

        def work():
            try:
                sf.build(self.filepath, template, out,
                         data=self.data,
                         progress=lambda m: self.root.after(0, self._status, m))
                self.root.after(0, self._export_ok, out)
            except Exception as e:
                self.root.after(0, self._export_failed, e, traceback.format_exc())

        threading.Thread(target=work, daemon=True).start()

    def _export_ok(self, out):
        self.busy = False
        self.export_btn.config(state=tk.NORMAL, text="Export Full Report")
        self._status(f"Exported {Path(out).name}")
        messagebox.showinfo("Saved", f"Full report written to:\n{out}")

    def _export_failed(self, err, tb):
        self.busy = False
        self.export_btn.config(state=tk.NORMAL, text="Export Full Report")
        self._status("Export failed")
        print(tb, file=sys.stderr)
        messagebox.showerror("Export failed", f"{type(err).__name__}: {err}")

    def _status(self, text):
        self.status_lbl.config(text=text)


def main():
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()

    icon = Path(sf.resource_path("icon.png"))
    if icon.exists():
        try:
            root.iconphoto(True, tk.PhotoImage(file=str(icon)))
        except Exception:
            pass

    try:
        root.lift()
        root.attributes("-topmost", True)
        root.after(200, lambda: root.attributes("-topmost", False))
    except Exception:
        pass

    SuperFilterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
