# EQE (Excel Query Engine)

Refactored package for `ExcelQueryEngine`.

Quick start

- Install dependencies:

```bash
pip install -r requirements.txt
```

- Quick import test:

```bash
python -c "from eqe import ExcelQueryEngine; print('import ok')"
```

Files added

- `eqe/engine.py` — refactored `ExcelQueryEngine` implementation
- `eqe/utils.py` — helper functions for Excel references
- `eqe/__init__.py` — package exports
- `ExcelQueryEngine.py` — top-level re-export for backward compatibility
- `requirements.txt` — runtime dependency list

**About**

The Excel Query Engine is a small, dependency-light utility for programmatically
loading and querying Excel-style data. It provides a simple API to:

- Load spreadsheet data (from a filename or in-memory structures).
- Read single cells and rectangular ranges.
- Find cells by value and fetch adjacent cells.
- Extract header-based tables as lists of dictionaries.
- Select specific columns by Excel column letter.

It intentionally exposes both low-level primitives (cell/range access) and
higher-level helpers (table extraction, column selection) so you can use it for
quick automation tasks as well as in larger data pipelines.

**Why it's useful**

- Fast prototyping: quickly script small data extraction tasks without setting up
	a heavyweight Excel-processing pipeline.
- Reproducible extraction: encode extraction rules in code so you can re-run and
	version-control them alongside other project code.
- Lightweight integration: use it inside ETL jobs, test harnesses, or small CLI
	utilities where bringing in the full `openpyxl` workbook model is unnecessary.
- Safe defaults: helpers return plain Python lists and dictionaries making the
	output easy to inspect, test, and serialize.

**Where to use it**

- Data analysis and reporting: extract specific ranges or tables from shared
	spreadsheets for downstream analysis.
- Finance and operations: pull monthly figures, reconciliations, or small reports
	into scripts or scheduled tasks.
- QA and testing: verify spreadsheet-based test fixtures or validate exported
	reports.
- Prototyping and automation: build small utilities that transform spreadsheet
	data into CSV, JSON or feed into other systems.


Usage Examples

Below are a few quick examples showing how to use `ExcelQueryEngine`. These examples assume
you've installed the dependencies from `requirements.txt` and that `openpyxl` is available.

- Basic in-memory usage (zero-based row/column indexes):

```python
from ExcelQueryEngine import ExcelQueryEngine

# Initialize from an Excel filename (recommended). This will read all sheets
# into memory as list-of-lists (rows are lists of cell values).
engine = ExcelQueryEngine('path/to/workbook.xlsx')

# Get a single cell by zero-based row/column
print(engine.get_cell('Sheet1', 1, 0))  # e.g. 'Alice'

# Find coordinates of a value
print(engine.find_by_value('Sheet1', 'Bob'))

# Get adjacent value (to the right of 'Alice')
print(engine.get_adjacent_value('Sheet1', 'Alice', offset=(0, 1)))

# Extract rows using header row index (header at row 0)
print(engine.extract_table_from_header('Sheet1', 0))
```

- Using Excel-style references (requires `openpyxl`):

```python
from ExcelQueryEngine import ExcelQueryEngine

data = {
	'Sheet1': [
		['Name', 'Age', 'City'],
		['Alice', 30, 'NY'],
		['Bob', 25, 'LA'],
	]
}

engine = ExcelQueryEngine(data)

# Get a rectangular range by Excel reference
# A2:B3 corresponds to rows 2..3 and cols A..B -> zero-based indexes used internally
print(engine.get_range_by_ref('Sheet1', 'A2:B3'))
# [['Alice', 30], ['Bob', 25]]
```

- Selecting specific columns by letter from a start row:

```python
from ExcelQueryEngine import ExcelQueryEngine

engine = ExcelQueryEngine(data)

# Columns are specified by Excel letter(s) and start_row is 1-based
print(engine.get_columns_from_row('Sheet1', ['A', 'C'], start_row=2))
# [[ 'Alice', 'NY' ], [ 'Bob', 'LA' ]]
```

Note: methods that accept or return Excel-style references rely on the helper utilities in
`eqe.utils`, which use `openpyxl` to parse column letters and cell coordinates. If you want
to run the examples without installing `openpyxl`, avoid calling methods that parse Excel
references and use zero-based integer row/column indexes directly.

Examples

If you don't have an `.xlsx` file handy, the `examples/` folder contains scripts that
create a temporary workbook and demonstrate usage with a filename-based initialization.
Run them like:

```bash
python examples\basic_usage.py
python examples\columns_usage.py
python examples\excel_ref_usage.py
```

