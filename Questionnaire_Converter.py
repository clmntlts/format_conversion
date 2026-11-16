"""
Questionnaire Long -> Wide Converter  (GUI)

A desktop tool for converting optically-read questionnaire exports from
LONG format (one row per question/matrix item) to WIDE format
(one row per submission_id, one column per variable).

Workflow:
  1. Add one or more .xlsx data files (or a whole folder).
  2. Review the detected variables in the "Variables" tab and, if you like,
     rename the column that each variable will get in the wide file.
  3. Preview the wide result in the "Preview" tab.
  4. Export to .xlsx  (each output has a 'wide' sheet + a 'codebook' sheet).

No Python knowledge required once packaged as an .exe.
"""

import os
import sys
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

import converter_core as cc


APP_TITLE = "Questionnaire Long → Wide Converter"


class ConverterApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=8)
        self.master = master
        self.pack(fill="both", expand=True)

        # state
        self.files = []            # list of file paths
        self.long_frames = {}      # path -> long DataFrame
        self.codebook = None       # combined codebook across all files (DataFrame)
        self.rename_map = {}       # variable -> export name
        self._edit_widget = None

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # --- top: file selection ---------------------------------------
        top = ttk.LabelFrame(self, text="1. Data files", padding=6)
        top.pack(fill="x", pady=(0, 6))

        btns = ttk.Frame(top)
        btns.pack(fill="x")
        ttk.Button(btns, text="Add files…", command=self.add_files).pack(side="left")
        ttk.Button(btns, text="Add folder…", command=self.add_folder).pack(side="left", padx=4)
        ttk.Button(btns, text="Remove selected", command=self.remove_selected).pack(side="left")
        ttk.Button(btns, text="Clear all", command=self.clear_all).pack(side="left", padx=4)

        self.file_list = tk.Listbox(top, height=4)
        self.file_list.pack(fill="x", pady=(6, 0))

        # --- middle: notebook (variables / preview) --------------------
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        # Variables tab
        var_tab = ttk.Frame(self.nb, padding=4)
        self.nb.add(var_tab, text="2. Variables  (rename here)")

        hint = ttk.Label(
            var_tab,
            text="Double-click an 'Export name' cell to rename the column used in the wide file. "
            "Names are shared across all loaded files that use the same question code.",
            foreground="#555",
        )
        hint.pack(fill="x", pady=(0, 4))

        var_btns = ttk.Frame(var_tab)
        var_btns.pack(fill="x", pady=(0, 4))
        ttk.Button(var_btns, text="Reset names to codes", command=self.reset_names).pack(side="left")

        cols = ("variable", "export", "type", "matrix", "text")
        self.var_tree = ttk.Treeview(var_tab, columns=cols, show="headings", selectmode="browse")
        headings = {
            "variable": ("Variable (code)", 150),
            "export": ("Export name", 170),
            "type": ("Type", 90),
            "matrix": ("Matrix item", 220),
            "text": ("Question text", 460),
        }
        for c in cols:
            title, w = headings[c]
            self.var_tree.heading(c, text=title)
            self.var_tree.column(c, width=w, anchor="w", stretch=(c == "text"))
        vy = ttk.Scrollbar(var_tab, orient="vertical", command=self.var_tree.yview)
        self.var_tree.configure(yscrollcommand=vy.set)
        self.var_tree.pack(side="left", fill="both", expand=True)
        vy.pack(side="right", fill="y")
        self.var_tree.bind("<Double-1>", self._on_var_double_click)

        # Preview tab
        prev_tab = ttk.Frame(self.nb, padding=4)
        self.nb.add(prev_tab, text="3. Preview (wide)")
        pv_btns = ttk.Frame(prev_tab)
        pv_btns.pack(fill="x", pady=(0, 4))
        ttk.Button(pv_btns, text="Refresh preview", command=self.refresh_preview).pack(side="left")
        self.preview_info = ttk.Label(pv_btns, text="", foreground="#555")
        self.preview_info.pack(side="left", padx=8)

        grid_holder = ttk.Frame(prev_tab)
        grid_holder.pack(fill="both", expand=True)
        self.prev_tree = ttk.Treeview(grid_holder, show="headings")
        px = ttk.Scrollbar(grid_holder, orient="horizontal", command=self.prev_tree.xview)
        py = ttk.Scrollbar(grid_holder, orient="vertical", command=self.prev_tree.yview)
        self.prev_tree.configure(xscrollcommand=px.set, yscrollcommand=py.set)
        self.prev_tree.grid(row=0, column=0, sticky="nsew")
        py.grid(row=0, column=1, sticky="ns")
        px.grid(row=1, column=0, sticky="ew")
        grid_holder.rowconfigure(0, weight=1)
        grid_holder.columnconfigure(0, weight=1)

        # --- bottom: options + export ----------------------------------
        bottom = ttk.LabelFrame(self, text="4. Export", padding=6)
        bottom.pack(fill="x", pady=(6, 0))

        self.combine_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            bottom,
            text="Combine all files into ONE wide sheet (use when files are the same questionnaire)",
            variable=self.combine_var,
        ).pack(anchor="w")

        exp_row = ttk.Frame(bottom)
        exp_row.pack(fill="x", pady=(6, 0))
        ttk.Button(exp_row, text="Export wide .xlsx…", command=self.export).pack(side="left")
        self.status = ttk.Label(exp_row, text="Add one or more data files to begin.", foreground="#333")
        self.status.pack(side="left", padx=10)

    # -------------------------------------------------------------- files
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select questionnaire data files",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        self._load_paths(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Select a folder of data files")
        if not folder:
            return
        paths = [
            os.path.join(folder, f)
            for f in sorted(os.listdir(folder))
            if f.lower().endswith(".xlsx") and not f.lower().endswith("_wide.xlsx")
            and not f.startswith("~$")
        ]
        if not paths:
            messagebox.showinfo(APP_TITLE, "No .xlsx data files found in that folder.")
            return
        self._load_paths(paths)

    def _load_paths(self, paths):
        errors = []
        for p in paths:
            p = os.path.normpath(p)
            if p in self.files:
                continue
            try:
                df = cc.load_long(p)
            except Exception as e:  # noqa: BLE001
                errors.append(f"{os.path.basename(p)}: {e}")
                continue
            self.files.append(p)
            self.long_frames[p] = df
            self.file_list.insert("end", p)
        if errors:
            messagebox.showwarning(
                APP_TITLE, "Some files could not be loaded:\n\n" + "\n".join(errors)
            )
        self._rebuild_codebook()

    def remove_selected(self):
        sel = list(self.file_list.curselection())
        for i in reversed(sel):
            p = self.file_list.get(i)
            self.file_list.delete(i)
            self.files.remove(p)
            self.long_frames.pop(p, None)
        self._rebuild_codebook()

    def clear_all(self):
        self.file_list.delete(0, "end")
        self.files.clear()
        self.long_frames.clear()
        self._rebuild_codebook()

    # ---------------------------------------------------------- codebook
    def _rebuild_codebook(self):
        """Build a combined codebook across all loaded files (union of vars)."""
        if not self.files:
            self.codebook = None
            self._refresh_var_tree()
            self.status.config(text="Add one or more data files to begin.")
            return
        frames = []
        for p in self.files:
            frames.append(self.long_frames[p])
        combined = pd.concat(frames, ignore_index=True)
        _, cb = cc.to_wide(combined)  # cb ordered by questionnaire order
        self.codebook = cb
        # keep any existing renames, default the rest to the code
        for v in cb["variable"]:
            self.rename_map.setdefault(v, v)
        # drop renames for variables no longer present
        valid = set(cb["variable"])
        self.rename_map = {k: val for k, val in self.rename_map.items() if k in valid}
        self._refresh_var_tree()
        self.status.config(
            text=f"{len(self.files)} file(s), {len(cb)} variables detected."
        )

    def _refresh_var_tree(self):
        self.var_tree.delete(*self.var_tree.get_children())
        if self.codebook is None:
            return
        for _, r in self.codebook.iterrows():
            v = r["variable"]
            self.var_tree.insert(
                "",
                "end",
                iid=v,
                values=(
                    v,
                    self.rename_map.get(v, v),
                    r["question_type"],
                    str(r["matrix_row"])[:60],
                    str(r["question_text"])[:200],
                ),
            )

    def reset_names(self):
        if self.codebook is None:
            return
        for v in self.codebook["variable"]:
            self.rename_map[v] = v
        self._refresh_var_tree()

    # -------------------------------------------------- inline cell edit
    def _on_var_double_click(self, event):
        if self.var_tree.identify_region(event.x, event.y) != "cell":
            return
        col = self.var_tree.identify_column(event.x)
        if col != "#2":  # only the Export name column is editable
            return
        iid = self.var_tree.identify_row(event.y)
        if not iid:
            return
        x, y, w, h = self.var_tree.bbox(iid, col)
        self._destroy_editor()
        entry = tk.Entry(self.var_tree)
        entry.insert(0, self.rename_map.get(iid, iid))
        entry.select_range(0, "end")
        entry.focus_set()
        entry.place(x=x, y=y, width=w, height=h)

        def commit(_=None):
            new = entry.get().strip()
            if new:
                self.rename_map[iid] = new
                self.var_tree.set(iid, "export", new)
            self._destroy_editor()

        entry.bind("<Return>", commit)
        entry.bind("<Escape>", lambda e: self._destroy_editor())
        entry.bind("<FocusOut>", commit)
        self._edit_widget = entry

    def _destroy_editor(self):
        if self._edit_widget is not None:
            self._edit_widget.destroy()
            self._edit_widget = None

    def _apply_renames(self, wide):
        return wide.rename(columns={v: self.rename_map.get(v, v) for v in wide.columns})

    def _name_clashes(self):
        """Export names used by more than one column (incl. the meta columns)."""
        names = ["submission_id", *cc.META_COLS]
        names += [self.rename_map.get(v, v) for v in self.codebook["variable"]]
        seen, clashes = set(), set()
        for n in names:
            if n in seen:
                clashes.add(n)
            seen.add(n)
        return sorted(clashes)

    def _renamed_codebook(self, cb):
        cb = cb.copy()
        cb.insert(1, "export_name", cb["variable"].map(lambda v: self.rename_map.get(v, v)))
        return cb

    # ----------------------------------------------------------- preview
    def refresh_preview(self):
        if not self.files:
            messagebox.showinfo(APP_TITLE, "Add at least one data file first.")
            return
        try:
            if self.combine_var.get() or len(self.files) == 1:
                frames = [self.long_frames[p] for p in self.files]
                combined = pd.concat(frames, ignore_index=True)
                wide, _ = cc.to_wide(combined)
                label = "combined" if len(self.files) > 1 else os.path.basename(self.files[0])
            else:
                wide, _ = cc.to_wide(self.long_frames[self.files[0]])
                label = os.path.basename(self.files[0]) + "  (first file; tick 'Combine' or export to see all)"
            wide = self._apply_renames(wide)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Could not build preview:\n{e}")
            return

        self.prev_tree.delete(*self.prev_tree.get_children())
        cols = list(wide.columns)
        self.prev_tree["columns"] = cols
        for c in cols:
            self.prev_tree.heading(c, text=str(c))
            self.prev_tree.column(c, width=max(70, min(220, len(str(c)) * 9)), anchor="w", stretch=False)
        for _, row in wide.head(50).iterrows():
            self.prev_tree.insert("", "end", values=["" if pd.isna(x) else x for x in row.tolist()])
        self.preview_info.config(
            text=f"{label}: {wide.shape[0]} rows × {wide.shape[1]} columns (showing up to 50)."
        )
        self.nb.select(1)

    # ------------------------------------------------------------ export
    def export(self):
        if not self.files:
            messagebox.showinfo(APP_TITLE, "Add at least one data file first.")
            return
        self._destroy_editor()
        clashes = self._name_clashes()
        if clashes:
            messagebox.showerror(
                APP_TITLE,
                "Several columns would get the same export name:\n\n"
                + "\n".join(clashes)
                + "\n\nGive them distinct names in the Variables tab, then export again.",
            )
            return
        try:
            if self.combine_var.get() and len(self.files) > 1:
                out = filedialog.asksaveasfilename(
                    title="Save combined wide file",
                    defaultextension=".xlsx",
                    initialfile="combined_wide.xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                )
                if not out:
                    return
                frames = [self.long_frames[p] for p in self.files]
                combined = pd.concat(frames, ignore_index=True)
                wide, cb = cc.to_wide(combined)
                cc.export_wide(self._apply_renames(wide), self._renamed_codebook(cb), out)
                self.status.config(text=f"Exported: {out}")
                messagebox.showinfo(APP_TITLE, f"Done.\n\nCombined wide file written to:\n{out}")
            else:
                folder = filedialog.askdirectory(title="Choose output folder for wide files")
                if not folder:
                    return
                written = []
                for p in self.files:
                    wide, cb = cc.to_wide(self.long_frames[p])
                    base = os.path.splitext(os.path.basename(p))[0]
                    out = os.path.join(folder, base + "_wide.xlsx")
                    cc.export_wide(self._apply_renames(wide), self._renamed_codebook(cb), out)
                    written.append(out)
                self.status.config(text=f"Exported {len(written)} file(s) to {folder}")
                messagebox.showinfo(
                    APP_TITLE,
                    "Done. Wrote:\n\n" + "\n".join(os.path.basename(w) for w in written),
                )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Export failed:\n{e}\n\n{traceback.format_exc()}")


def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("1080x720")
    try:
        ConverterApp(root)
    except Exception:
        messagebox.showerror(APP_TITLE, traceback.format_exc())
        raise
    root.mainloop()


if __name__ == "__main__":
    main()
