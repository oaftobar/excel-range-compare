# Excel Range Comparison Tool

This is a small Python CLI tool that compares two worksheets in an Excel file over the **same cell range** and writes any differences to a new worksheet named **`Diffs`**.

This project was built as a learning exercise and intentionally kept simple and readable (good for internal tooling).
It works, it solves a problem, and I'm continuing to learn and improve.

## What it does

- Prompts for:
  - Worksheet 1 name
  - Worksheet 2 name
  - Range to compare (example: `A1:E21`)
- Compares cells **row-by-row, column-by-column**
- Records differences into `Diffs` with these columns:
  - `Row` (Excel row number)
  - `Cell` (Excel coordinate like `C12`)
  - `Column` (header text from the first row of the range)
  - `Value 1`, `Value 2`
  - `Delta %` (only when both values are numeric)

## Setup

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

Interactive:

```bash
python compare_ranges.py
```

Non-interactive (example):

```bash
python compare_ranges.py --input "Employee Sales.xlsx" --output "Employee Sales NEW.xlsx" --sheet1 "Dataset 1" --sheet2 "Dataset 2" --range "A1:E21"
```

## Notes / assumptions

- Both worksheets must contain the same range shape
- The first row of the chosen range is treated as a **header row** (used for the `Column` field in `Diffs`).
- The tool compares the **same range** on both sheets (simple by design).
- If a `Diffs` sheet already exists, it will be replaced on each run.

## Collaboration

I'm learning in public. If you have ideas or improvements and want to work together, cool. If not, it does what it says on the tin.
