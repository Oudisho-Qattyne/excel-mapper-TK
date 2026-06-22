import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import copy
import os

from state import AppState
from excel_io import read_excel, read_sheet, write_excel
from transformer import apply_mapping, get_supported_operations
from mapping_io import save_mapping, load_mapping
from tkinterdnd2 import DND_FILES

COLORS = {
    "bg": "#f0f0f5", "fg": "#1a1a2e", "select_bg": "#d0d0ff",
    "select_fg": "#1a1a2e", "border": "#c0c0d0", "button_bg": "#e0e0e8",
    "button_fg": "#1a1a2e", "accent": "#6c5ce7", "input_bg": "#ffffff",
    "input_fg": "#1a1a2e", "card_bg": "#ffffff", "card_fg": "#1a1a2e",
    "tree_bg": "#ffffff", "tree_fg": "#1a1a2e", "header_bg": "#e8e8f0",
    "header_fg": "#555570", "status_bg": "#4a6a8a", "status_fg": "white",
    "disabled_fg": "#a0a0b0",
}

OP_EXAMPLES = {
    "exact": "Copies value as-is e.g. Name -> Name",
    "constant": "Always outputs a fixed value e.g. \"Active\"",
    "split": "Splits by delimiter e.g. \"a,b\" -> \"a\"",
    "concat": "Joins columns e.g. First + Last -> Full",
    "condition": "If/elif rules e.g. Salary > 90k -> Senior",
    "lookup": "Maps old->new e.g. Y -> Yes, N -> No",
    "regex": "Extracts via regex e.g. name@domain -> name",
    "formula": "Python expression e.g. int(Salary) * 1.1",
    "date_format": "Reformats date e.g. 2023-01-15 -> 15/01/2023",
    "case": "Changes case: upper / lower / title",
    "strip": "Removes leading/trailing whitespace",
    "skip": "Leaves column empty (skip this field)",
    "count": "Counts non-empty cells across columns",
}


class ExcelMapperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Field Mapper & Transformer")
        self.root.geometry("1200x800")
        self.state = AppState()
        self._clipboard_mapping = None

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self._build_ui()
        self._apply_theme()
        self._bind_accelerators()

    # ── Theme ──────────────────────────────────────────────────
    def _apply_theme(self):
        c = COLORS
        self.style.configure(".", background=c["bg"], foreground=c["fg"],
                             fieldbackground=c["input_bg"], font=("Segoe UI", 10))
        self.style.configure("TFrame", background=c["bg"])
        self.style.configure("TLabel", background=c["bg"], foreground=c["fg"])
        self.style.configure("TButton", background=c["button_bg"], foreground=c["button_fg"],
                             bordercolor=c["border"], focuscolor=c["accent"],
                             padding=(12, 6))
        self.style.map("TButton", background=[("active", c["select_bg"])],
                       foreground=[("active", c["select_fg"])])
        self.style.configure("TEntry", fieldbackground=c["input_bg"], foreground=c["input_fg"],
                             bordercolor=c["border"])
        self.style.map("TEntry", bordercolor=[("focus", c["accent"])])
        self.style.configure("TCombobox", fieldbackground=c["input_bg"], foreground=c["input_fg"],
                             bordercolor=c["border"], arrowcolor=c["fg"])
        self.style.map("TCombobox", bordercolor=[("focus", c["accent"])])
        self.style.configure("TLabelframe", background=c["card_bg"], foreground=c["card_fg"],
                             bordercolor=c["border"])
        self.style.configure("TLabelframe.Label", background=c["bg"], foreground=c["fg"])
        self.style.configure("Treeview", background=c["tree_bg"], foreground=c["tree_fg"],
                             fieldbackground=c["tree_bg"], bordercolor=c["border"])
        self.style.map("Treeview", background=[("selected", c["select_bg"])],
                       foreground=[("selected", c["select_fg"])])
        self.style.configure("Treeview.Heading", background=c["header_bg"], foreground=c["header_fg"],
                             bordercolor=c["border"])
        self.style.configure("TNotebook", background=c["bg"], bordercolor=c["border"])
        self.style.configure("TNotebook.Tab", background=c["button_bg"], foreground=c["fg"],
                             bordercolor=c["border"], padding=(10, 4))
        self.style.map("TNotebook.Tab", background=[("selected", c["card_bg"])],
                       foreground=[("selected", c["fg"])])
        self.style.configure("Horizontal.TScrollbar", background=c["button_bg"],
                             bordercolor=c["border"], arrowcolor=c["fg"])
        self.style.configure("Vertical.TScrollbar", background=c["button_bg"],
                             bordercolor=c["border"], arrowcolor=c["fg"])

        self.root.config(bg=c["bg"])
        if hasattr(self, "status_bar"):
            self.status_bar.config(bg=c["status_bg"], fg=c["status_fg"])

        self._theme_tk_widgets(c)

    def _theme_tk_widgets(self, c):
        for w in self.root.winfo_children():
            self._theme_walk(w, c)

    def _theme_walk(self, w, c):
        cls = w.winfo_class()
        if cls == "Label":
            if hasattr(self, "status_bar") and w is self.status_bar:
                return
            current_fg = w.cget("fg")
            if not current_fg.startswith("#"):
                w.config(bg=c["bg"], fg=c["fg"])
            else:
                w.config(bg=c["bg"])
        elif cls == "Listbox":
            w.config(bg=c["input_bg"], fg=c["input_fg"],
                     selectbackground=c["select_bg"], selectforeground=c["select_fg"])
        elif cls == "Text":
            w.config(bg=c["input_bg"], fg=c["input_fg"],
                     insertbackground=c["fg"], selectbackground=c["select_bg"],
                     highlightbackground=c["border"])
        elif cls == "Canvas":
            w.config(bg=c["bg"], highlightbackground=c["border"])
        elif cls in ("Button", "Checkbutton", "Radiobutton"):
            try:
                w.config(bg=c["button_bg"], fg=c["button_fg"],
                         activebackground=c["select_bg"], activeforeground=c["select_fg"])
            except tk.TclError:
                pass
        elif cls == "Menu":
            w.config(bg=c["bg"], fg=c["fg"])
        for child in w.winfo_children():
            self._theme_walk(child, c)

    # ── Status Bar ──────────────────────────────────────────────
    def _set_status(self, msg, is_error=False):
        self.status_var.set(msg)
        self.status_bar.config(foreground="red" if is_error else "white")

    # ── Accelerators ────────────────────────────────────────────
    def _bind_accelerators(self):
        self.root.bind("<Control-o>", lambda e: self._browse_source())
        self.root.bind("<Control-s>", lambda e: self._save_mapping())

    # ── Build UI ────────────────────────────────────────────────
    def _build_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self._build_tab_load()
        self._build_tab_map()
        self._build_tab_export()
        self._build_tab_stats()
        self._setup_drag_drop()

        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = tk.Label(
            self.root, textvariable=self.status_var,
            relief="sunken", anchor="w", bg="#4a6a8a", fg="white",
            font=("Segoe UI", 9),
        )
        self.status_bar.pack(fill="x", side="bottom")

    # ═══════════════════════ TAB 1: LOAD ═══════════════════════
    def _build_tab_load(self):
        self.tab_load = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_load, text="Load Files")

        pw = ttk.PanedWindow(self.tab_load, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=5, pady=5)

        self._build_load_frame(pw, "Source", is_source=True)
        self._build_load_frame(pw, "Target", is_source=False)

    def _build_load_frame(self, parent, label, is_source):
        frame = ttk.LabelFrame(parent, text=label, padding=10)
        parent.add(frame, weight=1)

        btn = ttk.Button(frame, text="Browse Excel...",
                         command=lambda: self._browse_file(is_source))
        btn.pack(anchor="w", pady=(0, 5))

        fname_var = tk.StringVar(value="No file loaded")
        fname_lbl = tk.Label(frame, textvariable=fname_var,
                             anchor="w", wraplength=400, font=("Segoe UI", 9))
        fname_lbl.pack(fill="x", pady=(0, 5))

        sheet_frame = ttk.Frame(frame)
        sheet_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(sheet_frame, text="Sheet:").pack(side="left")
        sheet_cb = ttk.Combobox(sheet_frame, state="readonly", width=25)
        sheet_cb.pack(side="left", padx=(5, 0))

        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree = ttk.Treeview(tree_frame, show="headings",
                             yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)

        if is_source:
            self.src_fname_var = fname_var
            self.src_tree = tree
            self.src_sheet_cb = sheet_cb
            self.src_frame = frame
            self._src_is_source = True
        else:
            self.tgt_fname_var = fname_var
            self.tgt_tree = tree
            self.tgt_sheet_cb = sheet_cb
            self.tgt_frame = frame

        setattr(self, f'{"src" if is_source else "tgt"}_sheet_cb', sheet_cb)
        sheet_cb.bind("<<ComboboxSelected>>",
                      lambda e, s=is_source: self._on_sheet_change(s))

    def _load_file(self, path, is_source):
        try:
            rows, columns, sheet_names = read_excel(path)
            if is_source:
                self.state.source_rows = rows
                self.state.source_columns = columns
                self.state.source_file_name = path
                self.state.source_sheet_names = sheet_names
                self.src_fname_var.set(path)
                self._populate_tree(self.src_tree, rows, columns, max_rows=5)
                self._populate_sheet_cb(self.src_sheet_cb, sheet_names, is_source)
            else:
                self.state.target_rows = rows
                self.state.target_columns = columns
                self.state.target_file_name = path
                self.state.target_sheet_names = sheet_names
                self.tgt_fname_var.set(path)
                self._populate_tree(self.tgt_tree, rows, columns, max_rows=5)
                self._populate_sheet_cb(self.tgt_sheet_cb, sheet_names, is_source)
            self._set_status(f"Loaded {'source' if is_source else 'target'}: {path}")
            self._rebuild_mapping_cards()
            self._refresh_stats()
        except Exception as e:
            self._set_status(f"Error loading file: {e}", is_error=True)

    def _browse_file(self, is_source):
        path = filedialog.askopenfilename(
            title=f"Select {'Source' if is_source else 'Target'} Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self._load_file(path, is_source)

    def _setup_drag_drop(self):
        EXTENSIONS = (".xlsx", ".xls", ".csv")
        if hasattr(self.root, "drop_target_register"):
            self.root.drop_target_register(DND_FILES)

            def _on_drop(event):
                data = event.data.strip()
                for path in self.root.tk.splitlist(data):
                    p = path.strip("{}").strip()
                    if p.lower().endswith(EXTENSIONS):
                        self._load_file(p, is_source=True)
                        break

            self.root.dnd_bind("<<Drop>>", _on_drop)

    def _populate_sheet_cb(self, cb, sheet_names, is_source):
        if len(sheet_names) > 1:
            cb["values"] = sheet_names
            cb.set(sheet_names[0])
            cb.configure(state="readonly")
        else:
            cb["values"] = ()
            cb.set("")
            cb.configure(state="disabled")

    def _on_sheet_change(self, is_source):
        path = self.state.source_file_name if is_source else self.state.target_file_name
        if is_source:
            sheet = self.src_sheet_cb.get()
        else:
            sheet = self.tgt_sheet_cb.get()
        if not sheet or not path:
            return
        try:
            rows, columns = read_sheet(path, sheet)
            if is_source:
                self.state.source_rows = rows
                self.state.source_columns = columns
                self.state.source_selected_sheet = sheet
                self._populate_tree(self.src_tree, rows, columns, max_rows=5)
            else:
                self.state.target_rows = rows
                self.state.target_columns = columns
                self.state.target_selected_sheet = sheet
                self._populate_tree(self.tgt_tree, rows, columns, max_rows=5)
            self._rebuild_mapping_cards()
            self._set_status(f"Switched to sheet: {sheet}")
        except Exception as e:
            self._set_status(f"Error loading sheet: {e}", is_error=True)

    def _populate_tree(self, tree, rows, columns, max_rows=0):
        for item in tree.get_children():
            tree.delete(item)
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, minwidth=80)
        display = rows[:max_rows] if max_rows > 0 else rows
        for row in display:
            tree.insert("", "end", values=[row.get(c, "") for c in columns])

    # ═══════════════════════ TAB 2: MAP ════════════════════════
    def _build_tab_map(self):
        self.tab_map = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_map, text="Map & Transform")

        outer_frame = ttk.Frame(self.tab_map)
        outer_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer_frame, highlightthickness=0)
        vsb = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        self.map_inner = ttk.Frame(canvas)
        self.map_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.map_inner, anchor="nw", tags="inner")

        def _on_canvas_resize(e):
            canvas.itemconfig("inner", width=e.width)
        canvas.bind("<Configure>", _on_canvas_resize)

        def _on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Search bar
        search_frame = ttk.Frame(self.tab_map)
        search_frame.pack(fill="x", padx=10, pady=(5, 0))
        ttk.Label(search_frame, text="Search:", font=("Segoe UI", 9)).pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._filter_mapping_cards())
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side="left")
        ttk.Label(search_frame, text="Type to filter target columns", foreground="#888").pack(side="left", padx=(8, 0))

        btn_frame = ttk.Frame(self.tab_map)
        btn_frame.pack(fill="x", padx=5, pady=5)
        ttk.Button(btn_frame, text="Save Mapping", command=self._save_mapping).pack(side="left", padx=(0, 5))
        ttk.Button(btn_frame, text="Load Mapping", command=self._load_mapping).pack(side="left", padx=(0, 5))
        ttk.Button(btn_frame, text="Run Transformation", command=self._run_transform).pack(side="right")

    def _rebuild_mapping_cards(self):
        for w in self.map_inner.winfo_children():
            w.destroy()

        target_cols = self.state.target_columns
        if not target_cols:
            ttk.Label(self.map_inner, text="Load a target file first.").pack(pady=20)
            return

        source_cols = self.state.source_columns
        ops = sorted(get_supported_operations())

        if not self.state.mappings or len(self.state.mappings) != len(target_cols):
            self.state.mappings = []
            for col in target_cols:
                default_src = source_cols[0] if source_cols else ""
                self.state.mappings.append({
                    "target_col": col,
                    "op": "exact",
                    "params": {"source_col": default_src},
                })

        self._mapping_widgets = {}
        self._card_frames = []
        for i, m in enumerate(self.state.mappings):
            card = ttk.LabelFrame(self.map_inner, text=m["target_col"], padding=8)
            card.pack(fill="x", padx=10, pady=4)
            self._card_frames.append(card)
            self._build_mapping_card(card, i, m, source_cols, ops)
        self._filter_mapping_cards()

    def _filter_mapping_cards(self):
        if not hasattr(self, '_card_frames'):
            return
        query = self.search_var.get().strip().lower()
        for i, card in enumerate(self._card_frames):
            target = self.state.mappings[i]["target_col"] if i < len(self.state.mappings) else ""
            if not query or query in target.lower():
                card.pack(fill="x", padx=10, pady=4)
            else:
                card.pack_forget()

    def _build_mapping_card(self, parent, idx, mapping, source_cols, ops):
        top_frame = ttk.Frame(parent)
        top_frame.pack(fill="x")

        ttk.Label(top_frame, text="Operation:", font=("Segoe UI", 9, "bold")).pack(side="left")

        op_var = tk.StringVar(value=mapping["op"])
        op_cb = ttk.Combobox(top_frame, textvariable=op_var, values=ops, state="readonly", width=14)
        op_cb.pack(side="left", padx=(5, 10))

        example_lbl = tk.Label(top_frame, text=OP_EXAMPLES.get(mapping["op"], ""),
                               fg="#6a6a88", anchor="w", font=("Segoe UI", 8))
        example_lbl.pack(side="left", fill="x", padx=(0, 5))

        copy_btn = ttk.Button(top_frame, text="Copy", width=5,
                              command=lambda i=idx: self._copy_mapping(i))
        copy_btn.pack(side="right", padx=(2, 0))
        paste_btn = ttk.Button(top_frame, text="Paste", width=5,
                               command=lambda i=idx: self._paste_mapping(i))
        paste_btn.pack(side="right", padx=(2, 0))

        param_frame = ttk.Frame(parent)
        param_frame.pack(fill="x", pady=(5, 0))

        preview_frame = ttk.Frame(parent)
        preview_frame.pack(fill="x", pady=(5, 0))

        self._mapping_widgets[idx] = {
            "op_var": op_var, "param_frame": param_frame, "preview_frame": preview_frame,
            "example_lbl": example_lbl,
        }

        def on_op_change(*args):
            new_op = op_var.get()
            mapping["op"] = new_op
            mapping["params"] = self._default_params(new_op, source_cols)
            example_lbl.config(text=OP_EXAMPLES.get(new_op, ""))
            self._render_params(param_frame, idx, mapping, source_cols)
            self._update_preview(preview_frame, idx, mapping)

        op_var.trace_add("write", on_op_change)
        self._render_params(param_frame, idx, mapping, source_cols)
        self._update_preview(preview_frame, idx, mapping)

    def _default_params(self, op, source_cols):
        default_src = source_cols[0] if source_cols else ""
        defaults = {
            "exact": {"source_col": default_src},
            "constant": {"value": ""},
            "split": {"source_col": default_src, "delimiter": ",", "index": 0},
            "concat": {"source_cols": source_cols[:3] if len(source_cols) >= 3 else source_cols[:], "separator": " "},
            "condition": {"rules": [{"source_col": default_src, "operator": "==", "value": "", "output": ""}], "default": ""},
            "lookup": {"source_col": default_src, "mapping": {}},
            "regex": {"source_col": default_src, "pattern": "", "group": 0},
            "formula": {"expression": ""},
            "date_format": {"source_col": default_src, "input_format": "%Y-%m-%d", "output_format": "%d/%m/%Y"},
            "case": {"source_col": default_src, "case_type": "upper"},
            "strip": {"source_col": default_src},
            "skip": {},
            "count": {"source_cols": []},
        }
        return copy.deepcopy(defaults.get(op, {}))

    def _render_params(self, parent, idx, mapping, source_cols):
        for w in parent.winfo_children():
            w.destroy()
        op = mapping["op"]
        params = mapping["params"]

        def on_change():
            self._update_preview(self._mapping_widgets[idx]["preview_frame"], idx, mapping)

        renderers = {
            "exact": self._render_exact,
            "constant": self._render_constant,
            "split": self._render_split,
            "concat": self._render_concat,
            "condition": self._render_condition,
            "lookup": self._render_lookup,
            "regex": self._render_regex,
            "formula": self._render_formula,
            "date_format": self._render_date_format,
            "case": self._render_case,
            "strip": self._render_strip,
            "skip": self._render_skip,
            "count": self._render_count,
        }
        renderer = renderers.get(op, self._render_skip)
        renderer(parent, params, source_cols, on_change, idx)

    def _render_exact(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Source column:").pack(side="left")
        var = tk.StringVar(value=params.get("source_col", ""))
        cb = ttk.Combobox(parent, textvariable=var, values=source_cols, state="readonly", width=20)
        cb.pack(side="left", padx=(5, 0))

        def update(*args):
            params["source_col"] = var.get()
            on_change()
        var.trace_add("write", update)
        cb.bind("<<ComboboxSelected>>", lambda e: on_change())

    def _render_constant(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Value:").pack(side="left")
        var = tk.StringVar(value=params.get("value", ""))
        entry = ttk.Entry(parent, textvariable=var, width=30)
        entry.pack(side="left", padx=(5, 0))

        def update(*args):
            params["value"] = var.get()
            on_change()
        var.trace_add("write", update)

    def _render_split(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Source:").pack(side="left")
        var_src = tk.StringVar(value=params.get("source_col", ""))
        cb = ttk.Combobox(parent, textvariable=var_src, values=source_cols, state="readonly", width=15)
        cb.pack(side="left", padx=(5, 0))

        ttk.Label(parent, text="Delimiter:").pack(side="left", padx=(10, 0))
        var_delim = tk.StringVar(value=params.get("delimiter", ","))
        entry_delim = ttk.Entry(parent, textvariable=var_delim, width=5)
        entry_delim.pack(side="left", padx=(5, 0))

        ttk.Label(parent, text="Index:").pack(side="left", padx=(10, 0))
        var_idx = tk.IntVar(value=params.get("index", 0))
        spin = ttk.Spinbox(parent, from_=0, to=99, textvariable=var_idx, width=5)
        spin.pack(side="left", padx=(5, 0))

        def update(*args):
            params["source_col"] = var_src.get()
            params["delimiter"] = var_delim.get()
            params["index"] = var_idx.get()
            on_change()
        var_src.trace_add("write", update)
        var_delim.trace_add("write", update)
        var_idx.trace_add("write", update)

    def _render_concat(self, parent, params, source_cols, on_change, idx):
        inner = ttk.Frame(parent)
        inner.pack(fill="x")
        ttk.Label(inner, text="Columns:").pack(side="left")

        lb_frame = ttk.Frame(inner)
        lb_frame.pack(side="left", padx=(5, 0))
        lb = tk.Listbox(lb_frame, selectmode="multiple", height=4, width=25, exportselection=False)
        lb.pack()
        for c in source_cols:
            lb.insert("end", c)
        selected = params.get("source_cols", [])
        for i, c in enumerate(source_cols):
            if c in selected:
                lb.selection_set(i)

        sep_frame = ttk.Frame(parent)
        sep_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(sep_frame, text="Separator:").pack(side="left")
        var_sep = tk.StringVar(value=params.get("separator", " "))
        entry_sep = ttk.Entry(sep_frame, textvariable=var_sep, width=10)
        entry_sep.pack(side="left", padx=(5, 0))
        self._concat_listbox = lb

        def update(*args):
            sel = lb.curselection()
            params["source_cols"] = [source_cols[i] for i in sel]
            params["separator"] = var_sep.get()
            on_change()
        var_sep.trace_add("write", update)

        def on_lb_select(e):
            sel = lb.curselection()
            params["source_cols"] = [source_cols[i] for i in sel]
            on_change()
        lb.bind("<<ListboxSelect>>", on_lb_select)

    def _render_condition(self, parent, params, source_cols, on_change, idx):
        rules = params.get("rules", [])
        opers = ["==", "!=", "contains", "starts_with", "ends_with", ">", "<", ">=", "<=", "is_empty"]

        rules_frame = ttk.Frame(parent)
        rules_frame.pack(fill="x")

        def rebuild_rules():
            for w in rules_frame.winfo_children():
                w.destroy()
            row_vars = []
            for ri, rule in enumerate(rules):
                rf = ttk.Frame(rules_frame)
                rf.pack(fill="x", pady=2)

                var_sc = tk.StringVar(value=rule.get("source_col", ""))
                cb_sc = ttk.Combobox(rf, textvariable=var_sc, values=source_cols, state="readonly", width=12)
                cb_sc.pack(side="left")

                var_op = tk.StringVar(value=rule.get("operator", "=="))
                cb_op = ttk.Combobox(rf, textvariable=var_op, values=opers, state="readonly", width=10)
                cb_op.pack(side="left", padx=(3, 0))

                var_v = tk.StringVar(value=rule.get("value", ""))
                e_v = ttk.Entry(rf, textvariable=var_v, width=12)
                e_v.pack(side="left", padx=(3, 0))

                var_o = tk.StringVar(value=rule.get("output", ""))
                e_o = ttk.Entry(rf, textvariable=var_o, width=12)
                e_o.pack(side="left", padx=(3, 0))

                def make_updater(r, sv_sc, sv_op, sv_v, sv_o):
                    def upd(*args):
                        r["source_col"] = sv_sc.get()
                        r["operator"] = sv_op.get()
                        r["value"] = sv_v.get()
                        r["output"] = sv_o.get()
                        on_change()
                    return upd

                upd = make_updater(rule, var_sc, var_op, var_v, var_o)
                var_sc.trace_add("write", upd)
                var_op.trace_add("write", upd)
                var_v.trace_add("write", upd)
                var_o.trace_add("write", upd)

                def make_remover(r_idx):
                    def rem():
                        rules.pop(r_idx)
                        rebuild_rules()
                        on_change()
                    return rem
                ttk.Button(rf, text="X", width=2, command=make_remover(ri)).pack(side="left", padx=(5, 0))

                row_vars.append((var_sc, var_op, var_v, var_o))

            btn_add = ttk.Button(rules_frame, text="+ Add Rule",
                                 command=lambda: (rules.append({"source_col": source_cols[0] if source_cols else "", "operator": "==", "value": "", "output": ""}), rebuild_rules(), on_change()))
            btn_add.pack(pady=3)

            df_frame = ttk.Frame(parent)
            df_frame.pack(fill="x", pady=(5, 0))
            ttk.Label(df_frame, text="Default:").pack(side="left")
            var_def = tk.StringVar(value=params.get("default", ""))
            e_def = ttk.Entry(df_frame, textvariable=var_def, width=30)
            e_def.pack(side="left", padx=(5, 0))

            def upd_def(*args):
                params["default"] = var_def.get()
                on_change()
            var_def.trace_add("write", upd_def)

        rebuild_rules()

    def _render_lookup(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Source column:").pack(anchor="w")
        var_sc = tk.StringVar(value=params.get("source_col", ""))
        cb = ttk.Combobox(parent, textvariable=var_sc, values=source_cols, state="readonly", width=20)
        cb.pack(anchor="w", pady=(0, 5))

        def upd_sc(*args):
            params["source_col"] = var_sc.get()
            on_change()
        var_sc.trace_add("write", upd_sc)

        mapping = params.get("mapping", {})
        mapping_frame = ttk.Frame(parent)
        mapping_frame.pack(fill="x")

        var_pairs = []

        def rebuild_lookup():
            for w in mapping_frame.winfo_children():
                w.destroy()
            nonlocal var_pairs
            var_pairs = []
            items = list(mapping.items()) if mapping else [("", "")]
            for kv in items:
                rf = ttk.Frame(mapping_frame)
                rf.pack(fill="x", pady=2)
                v_old = tk.StringVar(value=kv[0])
                v_new = tk.StringVar(value=kv[1])
                e_old = ttk.Entry(rf, textvariable=v_old, width=15)
                e_old.pack(side="left")
                ttk.Label(rf, text="→").pack(side="left", padx=3)
                e_new = ttk.Entry(rf, textvariable=v_new, width=15)
                e_new.pack(side="left")

                def make_updater():
                    def upd(*args):
                        nonlocal mapping
                        new_map = {}
                        for vo, vn in var_pairs:
                            k = vo.get().strip()
                            if k:
                                new_map[k] = vn.get().strip()
                        params["mapping"] = new_map
                        on_change()
                    return upd
                upd = make_updater()
                v_old.trace_add("write", upd)
                v_new.trace_add("write", upd)
                var_pairs.append((v_old, v_new))

            ttk.Button(mapping_frame, text="+ Add Row",
                       command=lambda: (mapping.setdefault("", ""),
                                        rebuild_lookup(),
                                        on_change())).pack(pady=3)
        rebuild_lookup()

    def _render_regex(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Source:").pack(side="left")
        var_sc = tk.StringVar(value=params.get("source_col", ""))
        cb = ttk.Combobox(parent, textvariable=var_sc, values=source_cols, state="readonly", width=15)
        cb.pack(side="left", padx=(5, 0))

        ttk.Label(parent, text="Pattern:").pack(side="left", padx=(10, 0))
        var_pt = tk.StringVar(value=params.get("pattern", ""))
        e_pt = ttk.Entry(parent, textvariable=var_pt, width=20)
        e_pt.pack(side="left", padx=(5, 0))

        ttk.Label(parent, text="Group:").pack(side="left", padx=(10, 0))
        var_gr = tk.IntVar(value=params.get("group", 0))
        sp = ttk.Spinbox(parent, from_=0, to=9, textvariable=var_gr, width=5)
        sp.pack(side="left", padx=(5, 0))

        def update(*args):
            params["source_col"] = var_sc.get()
            params["pattern"] = var_pt.get()
            params["group"] = var_gr.get()
            on_change()
        var_sc.trace_add("write", update)
        var_pt.trace_add("write", update)
        var_gr.trace_add("write", update)

    def _render_formula(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Python expression (use row['col_name']):").pack(anchor="w")
        text = tk.Text(parent, height=4, width=60, font=("Consolas", 10))
        text.pack(anchor="w", pady=(5, 0))
        text.insert("1.0", params.get("expression", ""))

        def update():
            params["expression"] = text.get("1.0", "end-1c")
            on_change()

        def on_key(e):
            self.root.after(200, update)
        text.bind("<KeyRelease>", on_key)

    def _render_date_format(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Source:").pack(side="left")
        var_sc = tk.StringVar(value=params.get("source_col", ""))
        cb = ttk.Combobox(parent, textvariable=var_sc, values=source_cols, state="readonly", width=15)
        cb.pack(side="left", padx=(5, 0))

        ttk.Label(parent, text="Input fmt:").pack(side="left", padx=(10, 0))
        var_in = tk.StringVar(value=params.get("input_format", "%Y-%m-%d"))
        e_in = ttk.Entry(parent, textvariable=var_in, width=14)
        e_in.pack(side="left", padx=(5, 0))

        ttk.Label(parent, text="Output fmt:").pack(side="left", padx=(10, 0))
        var_out = tk.StringVar(value=params.get("output_format", "%d/%m/%Y"))
        e_out = ttk.Entry(parent, textvariable=var_out, width=14)
        e_out.pack(side="left", padx=(5, 0))

        def update(*args):
            params["source_col"] = var_sc.get()
            params["input_format"] = var_in.get()
            params["output_format"] = var_out.get()
            on_change()
        var_sc.trace_add("write", update)
        var_in.trace_add("write", update)
        var_out.trace_add("write", update)

    def _render_case(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Source:").pack(side="left")
        var_sc = tk.StringVar(value=params.get("source_col", ""))
        cb = ttk.Combobox(parent, textvariable=var_sc, values=source_cols, state="readonly", width=15)
        cb.pack(side="left", padx=(5, 0))

        ttk.Label(parent, text="Case:").pack(side="left", padx=(10, 0))
        var_cs = tk.StringVar(value=params.get("case_type", "upper"))
        cb_cs = ttk.Combobox(parent, textvariable=var_cs, values=["upper", "lower", "title"], state="readonly", width=8)
        cb_cs.pack(side="left", padx=(5, 0))

        def update(*args):
            params["source_col"] = var_sc.get()
            params["case_type"] = var_cs.get()
            on_change()
        var_sc.trace_add("write", update)
        var_cs.trace_add("write", update)

    def _render_strip(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Source column:").pack(side="left")
        var_sc = tk.StringVar(value=params.get("source_col", ""))
        cb = ttk.Combobox(parent, textvariable=var_sc, values=source_cols, state="readonly", width=20)
        cb.pack(side="left", padx=(5, 0))

        def update(*args):
            params["source_col"] = var_sc.get()
            on_change()
        var_sc.trace_add("write", update)

    def _render_skip(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Column will be empty", foreground="gray").pack(side="left")

    def _render_count(self, parent, params, source_cols, on_change, idx):
        ttk.Label(parent, text="Select columns to count (non-empty):").pack(anchor="w")
        lb_frame = ttk.Frame(parent)
        lb_frame.pack(fill="x", pady=(5, 0))
        lb = tk.Listbox(lb_frame, selectmode="multiple", height=5, width=35, exportselection=False)
        lb.pack()
        for c in source_cols:
            lb.insert("end", c)
        selected = params.get("source_cols", [])
        for i, c in enumerate(source_cols):
            if c in selected:
                lb.selection_set(i)
        self._count_listbox = lb

        def on_lb_select(e):
            sel = lb.curselection()
            params["source_cols"] = [source_cols[i] for i in sel]
            on_change()
        lb.bind("<<ListboxSelect>>", on_lb_select)

    def _update_preview(self, parent, idx, mapping):
        for w in parent.winfo_children():
            w.destroy()

        if not self.state.source_rows:
            return

        preview_rows = self.state.source_rows[:3]
        try:
            config = [{"target_col": "preview", "op": mapping["op"], "params": mapping["params"]}]
            results, errors = apply_mapping(preview_rows, config)
        except Exception:
            return

        ttk.Label(parent, text="Preview:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        for ri, (row, err_row) in enumerate(zip(results, errors)):
            val = row.get("preview", "")
            err = err_row.get("preview", "")
            color = "green" if not err else "red"
            text = f"Row {ri+1}: {val}"
            if err:
                text += f" (error: {err})"
            lbl = tk.Label(parent, text=text, fg=color, anchor="w",
                           font=("Segoe UI", 9))
            lbl.pack(fill="x")

    # ── Copy / Paste Mapping ──────────────────────────────────
    def _copy_mapping(self, idx):
        m = self.state.mappings[idx]
        self._clipboard_mapping = copy.deepcopy(m)

    def _paste_mapping(self, idx):
        if self._clipboard_mapping is None:
            self._set_status("Nothing copied.", is_error=True)
            return
        m = self.state.mappings[idx]
        m["op"] = self._clipboard_mapping["op"]
        m["params"] = copy.deepcopy(self._clipboard_mapping["params"])
        w = self._mapping_widgets.get(idx)
        if w:
            w["op_var"].set(m["op"])
            w["example_lbl"].config(text=OP_EXAMPLES.get(m["op"], ""))
            self._render_params(w["param_frame"], idx, m, self.state.source_columns)
            self._update_preview(w["preview_frame"], idx, m)
        self._set_status(f"Pasted mapping to column '{m['target_col']}'")

    # ── Save / Load Mapping ────────────────────────────────────
    def _save_mapping(self):
        if not self.state.mappings:
            self._set_status("No mappings to save.", is_error=True)
            return
        path = filedialog.asksaveasfilename(
            title="Save Mapping",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        if not path:
            return
        try:
            save_mapping(self.state.mappings, path)
            self._set_status(f"Mapping saved: {path}")
        except Exception as e:
            self._set_status(f"Error saving mapping: {e}", is_error=True)

    def _load_mapping(self):
        path = filedialog.askopenfilename(
            title="Load Mapping",
            filetypes=[("JSON files", "*.json")]
        )
        if not path:
            return
        try:
            loaded = load_mapping(path)
            target_cols = self.state.target_columns
            col_map = {m["target_col"]: m for m in loaded}
            new_mappings = []
            for tc in target_cols:
                if tc in col_map:
                    new_mappings.append(col_map[tc])
                else:
                    new_mappings.append({"target_col": tc, "op": "skip", "params": {}})
            self.state.mappings = new_mappings
            self._rebuild_mapping_cards()
            self._set_status(f"Mapping loaded: {path}")
        except Exception as e:
            self._set_status(f"Error loading mapping: {e}", is_error=True)

    # ── Run Transformation ─────────────────────────────────────
    def _run_transform(self):
        if not self.state.source_rows:
            self._set_status("No source data loaded.", is_error=True)
            return
        if not self.state.mappings:
            self._set_status("No mappings configured.", is_error=True)
            return
        try:
            result, errors = apply_mapping(self.state.source_rows, self.state.mappings)
            self.state.result_rows = result
            self.state.result_errors = errors
            self._populate_export_table()
            self.notebook.select(self.tab_export)
            self._set_status(f"Transformation complete: {len(result)} rows, {len(self.state.target_columns)} columns")
        except Exception as e:
            self._set_status(f"Transformation failed: {e}", is_error=True)

    # ═══════════════════════ TAB 3: EXPORT ═════════════════════
    def _build_tab_export(self):
        self.tab_export = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_export, text="Export")

        summary_frame = ttk.Frame(self.tab_export)
        summary_frame.pack(fill="x", pady=5)

        self.export_row_count = tk.StringVar(value="Rows: 0")
        self.export_col_count = tk.StringVar(value="Columns: 0")
        self.export_err_count = tk.StringVar(value="Errors: 0")

        ttk.Label(summary_frame, textvariable=self.export_row_count, font=("Segoe UI", 10, "bold")).pack(side="left", padx=10)
        ttk.Label(summary_frame, textvariable=self.export_col_count, font=("Segoe UI", 10, "bold")).pack(side="left", padx=10)
        ttk.Label(summary_frame, textvariable=self.export_err_count, font=("Segoe UI", 10, "bold")).pack(side="left", padx=10)

        tree_frame = ttk.Frame(self.tab_export)
        tree_frame.pack(fill="both", expand=True, pady=5)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        self.export_tree = ttk.Treeview(tree_frame, show="headings",
                                        yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=self.export_tree.yview)
        hsb.config(command=self.export_tree.xview)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.export_tree.pack(fill="both", expand=True)

        btn_frame = ttk.Frame(self.tab_export)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Download .xlsx", command=self._download_xlsx).pack(side="right")

    def _populate_export_table(self):
        result = self.state.result_rows
        errors = self.state.result_errors
        columns = self.state.target_columns

        tree = self.export_tree
        for item in tree.get_children():
            tree.delete(item)
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, minwidth=80)

        err_count = 0
        for ri, (row, err_row) in enumerate(zip(result, errors)):
            tags = ""
            vals = []
            has_err = False
            for c in columns:
                v = row.get(c, "")
                e = err_row.get(c, "") if err_row else ""
                if e:
                    has_err = True
                    err_count += 1
                vals.append(v)
            tag = "error" if has_err else ""
            tree.insert("", "end", values=vals, tags=tag)
        tree.tag_configure("error", background="#ffcccc")

        self.export_row_count.set(f"Rows: {len(result)}")
        self.export_col_count.set(f"Columns: {len(columns)}")
        self.export_err_count.set(f"Errors: {err_count}")

    def _download_xlsx(self):
        if not self.state.result_rows:
            self._set_status("No results to export.", is_error=True)
            return
        path = filedialog.asksaveasfilename(
            title="Save Excel Output",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
        try:
            write_excel(path, self.state.result_rows, self.state.target_columns)
            self._set_status(f"Exported: {path}")
        except Exception as e:
            self._set_status(f"Error exporting: {e}", is_error=True)

    # ═══════════════════════ TAB 4: STATISTICS ═════════════════
    def _build_tab_stats(self):
        self.tab_stats = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stats, text="Statistics")
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        paned = ttk.PanedWindow(self.tab_stats, orient="vertical")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        self._build_stats_panel(paned, "Source", "src")
        self._build_stats_panel(paned, "Target", "tgt")

    def _build_stats_panel(self, parent, label, prefix):
        frame = ttk.LabelFrame(parent, text=label, padding=5)
        parent.add(frame, weight=1)

        info_frame = ttk.Frame(frame)
        info_frame.pack(fill="x", pady=(0, 5))
        setattr(self, f"{prefix}_stats_info", tk.StringVar(value="Rows: 0  |  Columns: 0"))
        ttk.Label(info_frame, textvariable=getattr(self, f"{prefix}_stats_info"),
                  font=("Segoe UI", 10, "bold")).pack(side="left")

        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        cols = ("Column", "Non-Empty", "Fill %", "Unique", "Type")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                            yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100 if c != "Column" else 180)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        setattr(self, f"{prefix}_stats_tree", tree)

    def _on_tab_changed(self, event):
        if event.widget.index("current") == 3:
            self._refresh_stats()

    def _refresh_stats(self):
        for prefix in ("src", "tgt"):
            rows_key = f"{prefix}_stats_tree"
            tree = getattr(self, rows_key, None)
            if tree is None:
                continue
            for item in tree.get_children():
                tree.delete(item)

        for prefix, label, columns_key, rows_key in [
            ("src", "Source", "source_columns", "source_rows"),
            ("tgt", "Target", "target_columns", "target_rows"),
        ]:
            tree = getattr(self, f"{prefix}_stats_tree", None)
            if not tree:
                continue
            columns = getattr(self.state, columns_key, [])
            rows = getattr(self.state, rows_key, [])
            info = getattr(self, f"{prefix}_stats_info")
            info.set(f"Rows: {len(rows)}  |  Columns: {len(columns)}")

            for col in columns:
                non_empty = 0
                numeric = 0
                unique = set()
                for row in rows:
                    v = row.get(col, "")
                    if v != "" and v is not None:
                        non_empty += 1
                        unique.add(str(v))
                        try:
                            float(v)
                            numeric += 1
                        except (ValueError, TypeError):
                            pass
                total = len(rows) or 1
                fill_pct = f"{non_empty * 100 // total}%"

                if numeric == non_empty and non_empty > 0:
                    dtype = "Numeric"
                elif non_empty > 0:
                    dtype = "Text"
                else:
                    dtype = "Empty"

                tree.insert("", "end", values=(col, str(non_empty), fill_pct,
                                                str(len(unique)), dtype))
