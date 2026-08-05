"""
Super Filter — Desktop GUI
============================
Drop in a raw SARI survey export, read the Organisation Summary on screen,
then export the complete formatted workbook.

The table on screen is built with a plain Frame/Label grid instead of a ttk
Treeview so that the Overall score cell can be colour-coded independently.
Everything else stays default text colour, matching the workbook.

Run:
    python app.py
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
LINE = "#DDE3EA"
STRIPE = "#F7F9FB"
GREEN = "#4CAF50"
AMBER = "#F4B400"
RED = "#DB4437"

ALL_ORGS = "All organisations"

# Display columns only. Every field is still written to the exported workbook.
# (label, index in sf.ORG_HEADERS, width in px, anchor)
COLUMNS = [
    ("Organisation name",   0, 280, tk.W),
    ("Type",                1, 140, tk.W),
    ("Resp",                2,  52, tk.CENTER),
    ("Size",                5,  72, tk.CENTER),
    ("Sector",              6, 120, tk.W),
    ("Avg score",           8,  70, tk.CENTER),
    ("Overall",             9,  76, tk.CENTER),
    ("Strongest section",  10, 140, tk.W),
    ("Weakest section",    11, 140, tk.W),
    ("Agreement",          14,  86, tk.CENTER),
    ("Maturity tier",      16, 112, tk.W),
]
OVERALL_COL = 6       # index in COLUMNS for the Overall column
RAW_OVERALL_IDX = 9  # index in sf.ORG_HEADERS for the Overall score value


def fmt(value, idx):
    if value is None or value == "":
        return "-"
    if idx == 2:              # Respondents
        try:
            return str(int(value))
        except (TypeError, ValueError):
            return str(value)
    if idx == 8:              # Avg score
        try:
            return f"{float(value):.2f}"
        except (TypeError, ValueError):
            return str(value)
    if idx == 9:              # Overall
        try:
            return f"{float(value):.1%}"
        except (TypeError, ValueError):
            return str(value)
    return str(value)


def score_color(value):
    """Colour for the Overall score cell. Default text colour elsewhere."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return None
    if v >= 0.75: return GREEN
    if v >= 0.50: return AMBER
    return RED


# ───────────────────────────────────────────────────────────────────────
#  Custom grid table
# ───────────────────────────────────────────────────────────────────────

ROW_H = 30        # px, every row and the header
CELL_PAD = 10     # px of breathing room inside a cell


def _cell(parent, text, width, anchor, bg, fg, font, cursor=None):
    """One table cell of an exactly fixed pixel width.

    The label goes inside a Frame with pack_propagate off, so a long organisation
    name is clipped by its column instead of widening it. That is what keeps the
    header lined up with the rows: a column is `width` px everywhere, in the header
    and in all 124 rows, no matter what text lands in it.
    """
    holder = tk.Frame(parent, width=width, height=ROW_H, bg=bg)
    holder.pack_propagate(False)
    holder.pack(side=tk.LEFT)
    lbl = tk.Label(holder, text=text, bg=bg, fg=fg, font=font, anchor=anchor)
    lbl.pack(fill=tk.BOTH, expand=True, padx=CELL_PAD)
    if cursor:
        holder.configure(cursor=cursor)
        lbl.configure(cursor=cursor)
    return holder, lbl


class GridTable:
    """Header row plus a scrolling body, both laid out from the same width list.

    Previously the header was one grid and every row was its own grid, so each
    computed column widths independently from its own content and nothing lined up.
    Widths now come only from COLUMNS and are enforced per cell.
    """

    def __init__(self, parent, columns, on_click=None):
        self.columns = columns
        self.on_click = on_click
        self.total_width = sum(c[2] for c in columns)

        outer = tk.Frame(parent, bg=WHITE)
        outer.pack(fill=tk.BOTH, expand=True)

        # Header, pinned above the scrolling area
        self.header = tk.Frame(outer, bg=WHITE, height=ROW_H)
        self.header.pack(fill=tk.X, anchor="w")
        self.header.pack_propagate(False)
        self._draw_header()

        tk.Frame(outer, bg=LINE, height=1).pack(fill=tk.X)

        body_wrap = tk.Frame(outer, bg=WHITE)
        body_wrap.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(body_wrap, bg=WHITE, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(body_wrap, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.body = tk.Frame(self.canvas, bg=WHITE)
        self.body_id = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.bind("<Configure>", self._on_body_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self._bind_wheel(self.canvas)

    def _on_body_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        # A frame placed with create_window keeps its own requested width, which for
        # an empty container is 1px. Rows packed inside then have nothing to fill and
        # collapse. Give it at least the width of the columns, more if the window is
        # wider, so the row stripes run the full width.
        self.canvas.itemconfigure(self.body_id, width=max(event.width, self.total_width))

    def _bind_wheel(self, widget):
        # macOS delivers small integer deltas; Windows and X11 send multiples of 120.
        def on_wheel(e):
            if e.num == 4:      delta = -3
            elif e.num == 5:    delta = 3
            elif abs(e.delta) >= 120: delta = int(-e.delta / 120) * 3
            else:               delta = -e.delta
            self.canvas.yview_scroll(int(delta), "units")
            return "break"
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind_all(seq, on_wheel)

    def _draw_header(self):
        for label, idx, width, anchor in self.columns:
            holder, lbl = _cell(self.header, label, width, anchor, WHITE, GREY,
                                ("Calibri", 10, "bold"),
                                cursor="hand2" if self.on_click else None)
            if self.on_click:
                for w in (holder, lbl):
                    w.bind("<Button-1>", lambda e, c=idx, l=label: self.on_click(c, l))

    def set_rows(self, rows_data):
        for w in self.body.winfo_children():
            w.destroy()

        for r, row in enumerate(rows_data):
            bg = STRIPE if r % 2 else WHITE
            # No pack_propagate(False) here: the row takes its height from the cells
            # and stretches to the body width, so the stripe runs the full row.
            rowframe = tk.Frame(self.body, bg=bg)
            rowframe.pack(fill=tk.X, anchor="w")
            for i, (label, idx, width, anchor) in enumerate(self.columns):
                fg, font = TEXT, ("Calibri", 11)
                if i == OVERALL_COL:
                    col = score_color(row[RAW_OVERALL_IDX])
                    if col is not None:
                        fg, font = col, ("Calibri", 11, "bold")
                _cell(rowframe, fmt(row[idx], idx), width, anchor, bg, fg, font)

        self.body.update_idletasks()
        self._on_body_configure()
        self.canvas.yview_moveto(0)


# ───────────────────────────────────────────────────────────────────────
#  App
# ───────────────────────────────────────────────────────────────────────

class SuperFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Super Filter — SARI Organisation Summary")
        self.root.geometry("1400x850")
        self.root.minsize(900, 500)
        self.root.configure(bg=WHITE)

        self.data = None
        self.rows = []
        self.filepath = None
        self.sort_idx = None
        self.sort_asc = True
        self.busy = False
        self.table = None

        # Title, toolbar and status stay put; only the table body scrolls. Wrapping
        # the whole window in a Canvas made the table claim its natural height inside
        # an already-scrolling parent, which is why it filled only the top half and
        # its own scrollbar did nothing.
        self.body = tk.Frame(self.root, bg=WHITE)
        self.body.pack(fill=tk.BOTH, expand=True)

        tk.Label(self.body, text="Organisation Summary", font=("Calibri", 18, "bold"),
                 bg=WHITE, fg=NAVY, anchor=tk.W).pack(fill=tk.X, padx=14, pady=(14, 0))

        self._build_toolbar()

        self.main = tk.Frame(self.body, bg=WHITE)
        self.main.pack(fill=tk.BOTH, expand=True, padx=14, pady=(10, 14))
        self._show_dropzone()

    def _build_toolbar(self):
        tb = tk.Frame(self.body, bg=WHITE)
        tb.pack(fill=tk.X, padx=14, pady=(10, 0))

        tk.Button(tb, text="Open Excel...", command=self._choose_file,
                  font=("Calibri", 10), bg=WHITE, relief=tk.FLAT,
                  highlightthickness=1, highlightbackground=LINE,
                  padx=14, pady=4, cursor="hand2").pack(side=tk.LEFT, pady=4)

        tk.Label(tb, text="Organisation:", font=("Calibri", 10, "bold"),
                 bg=WHITE, fg=TEXT).pack(side=tk.LEFT, padx=(18, 6))
        self.org_var = tk.StringVar(value=ALL_ORGS)
        self.org_combo = ttk.Combobox(tb, textvariable=self.org_var, state="readonly",
                                       font=("Calibri", 10), width=38, values=[ALL_ORGS])
        self.org_combo.pack(side=tk.LEFT, pady=4)
        self.org_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh())

        tk.Label(tb, text="Search:", font=("Calibri", 10, "bold"),
                 bg=WHITE, fg=TEXT).pack(side=tk.LEFT, padx=(18, 6))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._refresh())
        tk.Entry(tb, textvariable=self.search_var, font=("Calibri", 10),
                 relief=tk.FLAT, highlightthickness=1, highlightbackground=LINE,
                 width=22).pack(side=tk.LEFT, pady=4)

        self.export_btn = tk.Button(tb, text="Export Full Report", command=self._export,
                                    font=("Calibri", 10, "bold"), bg=HEADER_BG, fg=WHITE,
                                    activebackground=NAVY, activeforeground=WHITE,
                                    relief=tk.FLAT, padx=14, pady=4, cursor="hand2")
        self.export_btn.pack(side=tk.RIGHT, pady=4)

        # Status line — moves with the toolbar
        self.status_lbl = tk.Label(self.body, text="No file loaded", font=("Calibri", 9),
                                   bg=WHITE, fg=GREY, anchor=tk.W)
        self.status_lbl.pack(fill=tk.X, padx=14, pady=(4, 0))

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

        # The count line claims its strip at the bottom first, so the table gets
        # every remaining pixel instead of being squeezed by a later sibling.
        self.count_lbl = tk.Label(self.main, text="", font=("Calibri", 9),
                                  bg=WHITE, fg=GREY, anchor=tk.W)
        self.count_lbl.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

        wrap = tk.Frame(self.main, bg=WHITE)
        wrap.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.table = GridTable(wrap, COLUMNS, on_click=self._sort_by)
        self._refresh()

    def _refresh(self):
        if self.data is None or self.table is None:
            return
        picked = self.org_var.get()
        search = self.search_var.get().strip().lower()

        shown = 0
        rows = []
        for row in self.rows:
            name = str(row[0])
            if picked != ALL_ORGS and name != picked:
                continue
            if search and search not in name.lower():
                continue
            rows.append(row)
            shown += 1

        self.table.set_rows(rows)
        self.count_lbl.config(text=f"Showing {shown} of {len(self.rows)} organisations")

    def _sort_by(self, idx, label):
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

    # Icon (optional)
    icon = Path(__file__).with_name("icon.png")
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
