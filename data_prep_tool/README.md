# Train Type Data Preparation Tool

A Windows desktop application that generates a lean Excel file containing
**only the "Train Type" sheet**, pre-filled with values for a selected train type.
The output is intended to be copied manually into a large existing Excel workbook.

The application also includes a local workbook comparison feature that compares
a project workbook against a standard workbook and generates a deviation report.

> **The application never reads, opens, or modifies any user Excel files.**
> The Train Type generator still works on its own template copy, while the new
> comparison feature reads user-selected workbooks locally only.

---

## Directory Structure

```
data_prep_tool/
│
├── app/
│   ├── ui/
│   │   └── main_window.py          # Tkinter UI — no business logic
│   │
│   ├── services/
│   │   ├── train_type_service.py   # Resolves train type values from repository
│   │   ├── excel_generator.py      # Writes values into a copy of the template
│   │   └── data_update_service.py  # Reads/writes metadata.json (generation log)
│   │
│   ├── repositories/
│   │   ├── base_repository.py      # Abstract base — load_train_type_values / load_excel_mapping
│   │   ├── local_json_repository.py# Reads JSON from /data
│   │   └── sharepoint_repository.py# Downloads JSON from SharePoint, falls back to local
│   │
│   ├── models/
│   │   └── train_type.py           # TrainTypeValues dataclass
│   │
│   └── config.py                   # All paths and URLs (single source of truth)
│
├── data/
│   ├── train_type_values.json      # All train type parameter values
│   ├── excel_mapping.json          # Maps field names → sheet + cell addresses
│   └── metadata.json               # Generation history log (auto-updated)
│
├── templates/
│   └── train_type_template.xlsx    # Excel template (never modified by the app)
│
├── main.py                         # Entry point — wires dependencies, launches UI
├── create_template.py              # One-time helper to create the .xlsx template
├── requirements.txt
└── README.md
```

---

## Prerequisites

- Python 3.11 or newer
- Windows (required for the Tkinter UI and PyInstaller `.exe` target)

---

## Setup

### 1. Install dependencies

```bash
cd data_prep_tool
pip install -r requirements.txt
```

### 2. Create the Excel template (run once)

```bash
python create_template.py
```

This writes `templates/train_type_template.xlsx`. Only run it again if you need
to reset the template to its default layout.

### 3. Run the application

```bash
python main.py
```

---

## Configuration

### `app/config.py`

| Setting | Purpose |
|---|---|
| `SHAREPOINT_TRAIN_TYPE_VALUES_URL` | HTTPS URL to `train_type_values.json` on SharePoint |
| `SHAREPOINT_EXCEL_MAPPING_URL` | HTTPS URL to `excel_mapping.json` on SharePoint |
| `REQUEST_TIMEOUT_SECONDS` | HTTP timeout before falling back to local JSON |

Update the SharePoint URLs before deploying. If SharePoint is unreachable,
the app silently falls back to the local `/data` files and logs which source was used.

---

## Data files

### `data/train_type_values.json`

Defines all available train types and their parameter values:

```json
{
  "train_types": {
    "generic": {
      "max_speed_kmh": 160,
      "traction_type": "Electric",
      ...
    }
  }
}
```

Add a new train type by adding a new key under `"train_types"`.

### `data/excel_mapping.json`

Maps each field name to a specific sheet and cell in the template:

```json
{
  "mappings": [
    { "sheet": "Train Type", "cell": "B2", "field": "max_speed_kmh" },
    ...
  ]
}
```

To change where a value is written in the Excel output, edit this file —
no Python code changes are required.

## Compare Workbook Against Standard

The application now includes a second tab for deterministic workbook comparison.

### Inputs

- Standard Workbook (`.xlsm`)
- Project Workbook (`.xlsm`)

### Output

- `Workbook_Deviations.xlsx`
- Worksheet: `List of Project Deviations`
- Columns: `Change No`, `Change Type`, `Sheet Name`, `Cell Address`, `Standard Value`, `Project Value`

### Comparison rules

- Compares values only.
- Ignores formatting, macros, VBA, colors, borders, comments, row heights, column widths, and workbook metadata.
- Skips sheets `Cockpit`, `Revision history`, `Verify_Import_XML`, and `Passive Rules Check`.
- Treats whitespace-trimmed strings and empty values consistently.
- Supports password-protected workbooks using `IDDP_Protection_842` when needed.

The report is saved next to the selected Project Workbook.

### Preview workflow

After the comparison finishes, the app shows a preview of the first deviations
and the overall deviation counts. The user then confirms whether to save the
report and chooses the output path.

---

## Building a standalone `.exe`

```powershell
.uild.ps1
```

The resulting executable is at `dist\TrainTypeDataPrepTool.exe`.

> **Important:** The `data/` and `templates/` directories are bundled inside the
> `.exe` via `--add-data`. On first launch the app copies those defaults into a
> writable runtime folder next to the executable, so `metadata.json` and logs
> can be created safely.

If you want to run PyInstaller manually:

```powershell
pyinstaller --noconfirm --clean --onefile --windowed `
  --name TrainTypeDataPrepTool `
  --add-data "data;data" `
  --add-data "templates;templates" `
  main.py
```

The comparison feature requires the optional `msoffcrypto-tool` dependency so
password-protected workbooks can be opened locally when needed.

Application logging is written to `application.log` in the application folder.

---

## How it works

1. On startup the app attempts to fetch `train_type_values.json` and
   `excel_mapping.json` from the configured SharePoint URLs.
2. If SharePoint is unavailable (network error, timeout, wrong URL), it falls
   back to the local copies in `/data` and logs a warning.
3. The user selects a train type from the dropdown and clicks **Generate Excel**.
4. A Save dialog asks where to write the output file (default name: `Train type.xlsx`).
5. The app copies the template, writes the mapped values into the copy, and saves it.
6. The original template is never touched.
7. A log entry is appended to `data/metadata.json`.

### Mapping modes

The app supports two ways to map values into the workbook:

1. `mappings` list: explicit `sheet + cell + field` entries.
2. `parameter_column_mapping`: lookup by parameter name column (for large Train Type sheets).

Current configuration uses `parameter_column_mapping` and writes values by matching
JSON keys against parameter names in column `D` and writing to column `E`.

---

## Layering rules (do not violate)

| Layer | Allowed dependencies |
|---|---|
| `ui/` | `services/` only |
| `services/` | `models/`, `repositories/`, `config` |
| `repositories/` | `models/`, `config` |
| `models/` | stdlib only |
| `config.py` | stdlib only |
