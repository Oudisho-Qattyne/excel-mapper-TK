# Excel Field Mapper & Transformer

A desktop GUI application for transforming and mapping data between Excel files. Built with Python + Tkinter + openpyxl.

---

## Features

- **Intuitive 4-tab interface** — Load, Map, Export, and Statistics
- **13 transformation operations** — exact copy, constant, split, concat, condition, lookup, regex, formula, date_format, case, strip, skip, count
- **Live preview** — see transformation results for the first 3 rows as you configure
- **Drag & drop** Excel files directly onto the window
- **Search/filter** target columns in the mapping tab
- **Column statistics** — view fill rate, unique count, and inferred type per column
- **Save/load mapping configurations** as JSON files
- **Keyboard shortcuts**: Ctrl+O (open source), Ctrl+S (save mapping)

---

## Requirements

- Python 3.10+
- `openpyxl` (installed automatically with the command below)
- `tkinterdnd2` (for drag-and-drop support)

```
pip install openpyxl tkinterdnd2
```

---

## How to Run

```
python app.py
```

Or use the pre-built executable: `dist/ExcelMapper.exe`

---

## Workflow

### 1. Load Files tab
- Select a **source** Excel file (the data you want to transform from)
- Select a **target** Excel file (defines the output column structure)
- Preview shows the first 5 rows
- Tabs with multiple sheets are supported

### 2. Map & Transform tab
- Each target column gets its own **configuration card**
- For each card:
  1. Choose a **transformation operation** (from 13 available)
  2. Adjust operation-specific parameters
  3. See a **live preview** of the first 3 rows
- Use the **search box** to quickly filter cards by target column name
- **Save/Load** your full mapping configuration as JSON
- Click **Run Transformation** to execute and switch to the Export tab

### 3. Export tab
- Full result table with all transformed rows
- Error cells highlighted in red
- Summary: row count, column count, error count
- **Download .xlsx** button to save the result

### 4. Statistics tab
- Per-column analysis for both source and target data:
  - Non-empty count and fill percentage
  - Unique value count
  - Inferred data type (Numeric / Text / Empty)
- Refreshes automatically when data is loaded

---

## Transformation Operations

| Operation | Description |
|-----------|-------------|
| **exact** | Copies source column value as-is |
| **constant** | Outputs the same fixed value for all rows |
| **split** | Splits by delimiter and picks a part by index |
| **concat** | Joins multiple source columns with a separator |
| **condition** | If/elif/else rules with 10 comparison operators |
| **lookup** | Maps old values to new values via a dictionary |
| **regex** | Extracts text using a regular expression |
| **formula** | Evaluates a Python expression (sandboxed) |
| **date_format** | Parses and reformats date strings |
| **case** | Converts to upper, lower, or title case |
| **strip** | Removes leading/trailing whitespace |
| **skip** | Leaves the column empty |
| **count** | Counts non-empty cells across selected columns |

---

## File Structure

```
excelMapperTk/
├── app.py                # Entry point
├── gui.py                # User interface (Tkinter)
├── transformer.py        # 13 transformation operations
├── excel_io.py           # Excel read/write (openpyxl)
├── mapping_io.py         # Save/load mapping JSON
├── state.py              # Application state
├── app_icon.ico          # Application icon
├── requirements.txt      # Dependencies
└── README.md             # This file
```

---

## Technical Notes

- No admin privileges required
- All processing runs locally — no internet connection needed
- Uses `openpyxl` for Excel I/O; Microsoft Excel is not required
- The `formula` operation uses a sandboxed `eval()` with a restricted set of allowed functions
- Drag-and-drop support provided by `tkinterdnd2`
