# Product Detail Log Analysis

Console tool that turns a raw product log Excel file into a per-barcode summary.

## Setup

```bash
pip install pandas openpyxl python-dotenv
```

Edit `.env` and set your own directories:

```
INPUT_DIR=C:/Users/busra/Documents/logs
OUTPUT_DIR=C:/Users/busra/Documents/reports
```

## Run

```bash
python product_detail_log_analysis.py
```

The newest Excel file in `INPUT_DIR` is analysed. To target a specific file,
set `INPUT_FILE` in `.env` or pass it on the command line:

```bash
python product_detail_log_analysis.py --input raw_log.xlsx
python product_detail_log_analysis.py --input-dir D:/logs --output-dir D:/out
```

## Input (read by column position)

| Cell | Content |
|---|---|
| A | Barcode |
| B | Log status |
| C | Date |

## Output

| Cell | Header |
|---|---|
| A | Barcode |
| B | Total Log Count |
| C | First Log Date |
| D | First Log Status |
| E | Last Log Date |
| F | Last Log Status |

See `RFC-001-product-detail-log-analysis.md` for the full design.
