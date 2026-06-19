# Build Prompt: Excel Field Mapper & Transformer — Python Tkinter GUI

Build a **desktop GUI application** using Python 3 + `tkinter` + `openpyxl` that replicates the functionality of the existing `index.html` (plain JS/HTML app in this directory) and adds a new **count** operation. Read `implementation_plan_js.md` first to understand the full workflow, then implement everything described below.

The app must run on Windows without needing admin install — single-file or small multi-file Python script launched with `python app.py`.

---

## Core Workflow (3 Tabs)

```
Tab 1: LOAD          Tab 2: MAP & TRANSFORM      Tab 3: EXPORT
┌──────────────┐     ┌────────────────────────┐  ┌──────────────┐
│ Browse Source│────▶│ For each target column  │─▶│ Preview table│
│ Browse Target│     │  • Pick source col(s)   │  │ Download.xlsx│
│ Preview rows │     │  • Pick operation       │  │ Summary stats│
│ (Treeview)   │     │  • Configure params     │  └──────────────┘
│ Sheet select │     │  • Live preview (3 rows)│
└──────────────┘     │  Save/Load mapping JSON │
                     │  Run Transformation btn │
                     └────────────────────────┘
```

---

## 13 Transformation Operations (12 existing + 1 new)

| # | Operation | Description | Parameters |
|---|-----------|-------------|------------|
| 1 | **exact** | Copy value as-is | `source_col: str` |
| 2 | **constant** | Always output a fixed value | `value: str` |
| 3 | **split** | Split by delimiter, take Nth part | `source_col, delimiter, index` |
| 4 | **concat** | Join multiple source columns | `source_cols: list[str], separator` |
| 5 | **condition** | If/elif/else chain | `source_col, rules: list[(op, val, out)], default` |
| 6 | **lookup** | Map old→new values (dictionary) | `source_col, mapping: dict` |
| 7 | **regex** | Extract using regex group | `source_col, pattern, group_index` |
| 8 | **formula** | Python expression with `row` dict | `expression: str` |
| 9 | **date_format** | Parse and reformat date | `source_col, input_format, output_format` |
| 10 | **case** | upper / lower / title | `source_col, case_type` |
| 11 | **strip** | Strip whitespace | `source_col` |
| 12 | **skip** | Leave column empty | — |
| **13** | **count** | **NEW** — count non-empty cells across selected columns per row | `source_cols: list[str]` |

### Count Operation — Detailed Spec

The `count` operation counts, for **each row**, how many of the selected source columns contain non-empty data. This is useful for:
- Counting number of children per family (select child1, child2, child3 columns → count of non-empty cells = number of children)
- Counting filled fields in a survey
- Counting items in repeating groups

**Parameters:**
- `source_cols: list[str]` — the columns to examine (multi-select via checklistbox or Listbox with `selectmode=multiple`)
- Returns: integer as string (e.g., `"3"`)

**Implementation:**
```python
def op_count(row: dict, params: dict) -> str:
    cols = params.get("source_cols", [])
    count = sum(1 for c in cols if row.get(c, "").strip() != "")
    return str(count)
```

---

## Technology Stack

| Layer | Choice | Reason |
|-------|--------|--------|
| **Language** | Python 3.10+ | Modern, widely available |
| **GUI** | `tkinter` + `ttk` (themed widgets) | Built-in, zero deps, native look |
| **Excel** | `openpyxl` | Read/write `.xlsx`, no Excel install needed |
| **JSON** | `json` (stdlib) | Mapping save/load |
| **Packaging** | Raw `.py` script(s) | No build tools, run with `python app.py` |

---

## File Structure (multi-file for clarity)

```
excel-mapper-tk/
├── app.py                # Entry point, launches main window
├── gui.py                # All tkinter UI: tabs, widgets, event handlers
├── transformer.py        # Pure functions: 13 operations, apply_mapping()
├── excel_io.py           # openpyxl read/write wrappers
├── mapping_io.py         # Save/load mapping JSON
├── state.py              # Global app state dataclass
└── requirements.txt      # openpyxl only
```

**Single-file option** also acceptable — put everything in `app.py` if preferred.

---

## State

Use a `dataclass` or plain dict. Must hold:

```python
@dataclass
class AppState:
    source_rows: list[dict]          # list of row dicts
    source_columns: list[str]
    source_file_name: str
    source_sheet_names: list[str]
    source_selected_sheet: str

    target_rows: list[dict]
    target_columns: list[str]
    target_file_name: str
    target_sheet_names: list[str]
    target_selected_sheet: str

    mappings: list[dict]             # [{target_col, op, params}, ...]
    result_rows: list[dict]
    result_errors: list[dict]        # [{row, col, msg}, ...]
```

---

## Tab 1 — Load Files

Layout: Two frames side-by-side (or top-bottom).

Each frame has:
- **Browse button** → `filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])`
- **File name label**
- **Sheet selector** (ttk.Combobox or OptionMenu) — shown only when >1 sheet
- **Preview Treeview** (ttk.Treeview) showing first 5 rows of the data

Use `openpyxl.load_workbook(filename, data_only=True)` then iterate sheet rows.

**Key detail:** `openpyxl` cells have `.value`. Convert all values to strings for display. Handle `None` as empty string.

When sheet selection changes, re-read the sheet and update preview.

---

## Tab 2 — Map & Transform

For each **target column**, render a **Labelframe** (or custom frame) containing:

1. **Target column name** as bold label
2. **Operation combobox** — all 13 operations
3. **Parameter frame** — dynamically rebuilt when operation changes
4. **Live preview** — first 3 transformed values with green (OK) / red (error) labels

### Parameter UI per operation (use ttk widgets):

| Operation | Widgets |
|-----------|---------|
| exact | Combobox of source columns |
| constant | Entry for value |
| split | Combobox (source col) + Entry (delimiter) + Spinbox (index) |
| concat | Listbox(multiple) of source cols + Entry (separator) |
| condition | Combobox (source col) + dynamic rows of (OpCombobox, Entry val, Entry out) + Add/Remove btns + Entry default |
| lookup | Combobox (source col) + dynamic rows of (Entry old, Entry new) + Add/Remove btns |
| regex | Combobox + Entry (pattern) + Spinbox (group) |
| formula | Text widget (expression, monospace font) |
| date_format | Combobox + Entry (in_fmt) + Entry (out_fmt) |
| case | Combobox + Combobox (upper/lower/title) |
| strip | Combobox |
| skip | Label saying "Column will be empty" |
| **count** | **Listbox(multiple) of source columns** |

### Live Preview Logic

After every parameter change:
1. Get first 3 rows from `state.source_rows`
2. Apply the single operation to each
3. Display results in the preview frame
4. Green foreground for success, red for error

### Bottom Action Buttons

- **Save Mapping** → `filedialog.asksaveasfilename` → dump `state.mappings` as JSON
- **Load Mapping** → `filedialog.askopenfilename` → load JSON → restore `state.mappings` → refresh UI
- **Run Transformation** → call `apply_mapping()` → switch to Export tab

---

## Tab 3 — Export

- **Summary labels**: Row count, Column count, Error count
- **ttk.Treeview** with full result table
  - Highlight error cells with red background tag
- Scrollbars (horizontal + vertical)
- **Download .xlsx button** → `filedialog.asksaveasfilename` → write with `openpyxl`

---

## Core Transformation Engine (`transformer.py`)

```python
OPERATIONS: dict[str, Callable] = {
    "exact": op_exact,
    "constant": op_constant,
    "split": op_split,
    "concat": op_concat,
    "condition": op_condition,
    "lookup": op_lookup,
    "regex": op_regex,
    "formula": op_formula,
    "date_format": op_date_format,
    "case": op_case,
    "strip": op_strip,
    "skip": op_skip,
    "count": op_count,        # NEW
}

def apply_mapping(source_rows: list[dict],
                  mappings: list[dict]) -> tuple[list[dict], list[dict]]:
    """Returns (result_rows, errors)."""
```

### Operation implementations:

**op_exact(row, params):** return str(row.get(params["source_col"], ""))

**op_constant(row, params):** return str(params.get("value", ""))

**op_split(row, params):** split by delimiter, return Nth piece or ""

**op_concat(row, params):** join values of listed columns with separator

**op_condition(row, params):**
```python
for rule in params["rules"]:
    val = str(row.get(rule["source_col"], ""))
    match rule["operator"]:
        case "==":         matched = val == rule["value"]
        case "!=":         matched = val != rule["value"]
        case "contains":   matched = rule["value"] in val
        case "starts_with": matched = val.startswith(rule["value"])
        case "ends_with":  matched = val.endswith(rule["value"])
        case ">":          matched = float(val) > float(rule["value"])
        case "<":          matched = float(val) < float(rule["value"])
        case ">=":         matched = float(val) >= float(rule["value"])
        case "<=":         matched = float(val) <= float(rule["value"])
        case "is_empty":   matched = val.strip() == ""
    if matched:
        return str(rule["output"])
return str(params.get("default", ""))
```

**op_lookup(row, params):** return `mapping.get(val, val)` where val = row[source_col]

**op_regex(row, params):** use `re.search(params["pattern"], val)`, return group or ""

**op_formula(row, params):** `eval(params["expression"], {"__builtins__": {}}, {"row": row})` — **use safe eval, restrict builtins**

**op_date_format(row, params):**
- Parse using `datetime.strptime(val, input_format)`
- Reformat using `datetime.strftime(output_format)`
- On failure, return original value

**op_case(row, params):** apply `str.upper()`, `.lower()`, or `.title()`

**op_strip(row, params):** return `str(row.get(params["source_col"], "")).strip()`

**op_skip(row, params):** return ""

**op_count(row, params):** **NEW** — count non-empty selected columns
```python
def op_count(row, params):
    cols = params.get("source_cols", [])
    count = sum(1 for c in cols if str(row.get(c, "")).strip() != "")
    return str(count)
```

### Condition operators (same as JS version):
`==`, `!=`, `contains`, `starts_with`, `ends_with`, `>`, `<`, `>=`, `<=`, `is_empty`

### Formula safety:
Use `eval()` with restricted globals. **DO NOT use unrestricted `eval`**. Acceptable:
```python
safe_globals = {"__builtins__": {"True": True, "False": False, "None": None,
                                 "abs": abs, "int": int, "float": float,
                                 "str": str, "len": len, "min": min, "max": max,
                                 "round": round, "sum": sum}}
try:
    result = eval(params["expression"], safe_globals, {"row": row})
    return str(result)
except Exception:
    return ""
```

---

## UI Design Guidelines

1. **Window size**: 1200x800, resizable
2. **Use ttk themed widgets** for modern look (ttk.Button, ttk.Combobox, ttk.Treeview, ttk.LabelFrame, ttk.Notebook for tabs)
3. **ttk.Notebook** for the 3 tabs (Load / Map / Export) — this is standard tkinter tab approach
4. **Font**: default system font, mono for code/expression areas
5. **Colors**: Use ttk theme "clam" or "vista" (Windows) for decent default styling
6. **Error handling**: Show error messages in status bar at bottom of window, not in popups for minor errors
7. **Mapping cards**: Use ttk.LabelFrame with padding inside a Canvas + Scrollbar (for many target columns)
8. **Drag & drop**: Optional (not required for v1)
9. **Accelerators**: Ctrl+O for open file, Ctrl+S for save mapping

---

## Development Steps (ordered)

1. **Scaffold**: Create files, install `openpyxl`, verify Python 3.10+
2. **State**: Define `AppState` dataclass in `state.py`
3. **Excel IO**: Implement `read_excel(path) -> (rows, columns, sheet_names)` and `write_excel(path, rows, columns)` in `excel_io.py`
4. **Transformer**: Implement all 13 operations + `apply_mapping()` in `transformer.py` — test with pytest or inline assertions
5. **Mapping IO**: Save/load JSON in `mapping_io.py`
6. **GUI Tab 1**: File browse, sheet selector, preview Treeview
7. **GUI Tab 2**: Dynamic mapping card rendering, parameter widgets, live preview
8. **GUI Tab 3**: Result table, summary, download button
9. **Integration**: Wire tabs together, connect Run button → transform → switch tab
10. **Test**: Open with test-source.xlsx / test-target.xlsx, set up all 13 operations, verify output

---

## Testing

- Use the provided `test-source.xlsx` (7 employees) and `test-target.xlsx` (6 columns)
- Test the **count** operation: select multiple columns where some cells are empty, verify count is correct
- Test save/load mapping round-trip
- Test multi-sheet workbook loading
- Test formula with restricted eval
- Test date_format with both valid and invalid dates

---

## Delivered Prompt Expectations

When the build is complete, deliver:
1. All source files (single-file or multi-file as decided)
2. Confirmation that the app opens and runs with `python app.py`
3. Brief summary of how the `count` operation works in the UI

---

**Build this now. Implement all 13 operations, all 3 tabs, and ensure the count operation works as described.**
