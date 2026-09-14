# RFC-001: Product Detail Log Analysis

| Field | Value |
|---|---|
| **RFC** | 001 |
| **Title** | Product Detail Log Analysis |
| **Status** | Implemented |
| **Author** | Busra Boyaci |
| **Created** | 2026-09-14 |
| **Component** | `product_detail_log_analysis.py` |

---

## 1. Summary

A console application that reads a raw product log Excel file, groups the records
by barcode, and produces a new Excel file reporting the first and last log entry
for each product, along with the total number of log records.

Input and output directories are supplied through a `.env` file, so paths can be
changed without editing source code.

---

## 2. Motivation

The raw log export contains one row per event, so a single product can appear
dozens of times across many dates. Answering "when did this product first enter
the system, what was its last status, and how many times was it touched?"
currently requires manual sorting and filtering in Excel — slow and error-prone
once the file grows past a few thousand rows.

This tool turns that raw event log into a one-row-per-product summary in a single
command.

---

## 3. Goals and non-goals

**Goals**

- Read the raw file from a configurable directory.
- Group records by barcode (one row of output per unique barcode).
- Report total log count, first log date + status, last log date + status.
- Write the header row exactly once, at the top of the output sheet.
- Keep every path in `.env` so nothing in the code changes between environments.
- Fail with a clear console message rather than a stack trace.

**Non-goals**

- No GUI; this is a console tool.
- No database, API, or scheduled-job integration.
- No modification of the raw input file — it is opened read-only.
- No analysis of statuses beyond first/last (no state-transition validation).

---

## 4. Definitions

| Term | Meaning |
|---|---|
| **Barcode** | The identifier of a single product; the grouping key. All rows sharing a barcode are one product. |
| **Log record** | One row of the raw file: barcode + status + date. |
| **First log** | The record with the earliest date within a barcode group. |
| **Last log** | The record with the latest date within a barcode group. |

---

## 5. Configuration (`.env`)

The `.env` file sits next to the script. Loading order: `python-dotenv` if
installed, otherwise a built-in parser; CLI arguments override everything.

| Key | Required | Default | Description |
|---|---|---|---|
| `INPUT_DIR` | Yes | `./input` | Directory holding the raw Excel file |
| `OUTPUT_DIR` | Yes | `./output` | Result directory; created if missing |
| `INPUT_FILE` | No | *(empty)* | Specific file name. If empty, the newest Excel file in `INPUT_DIR` is used |
| `SHEET_NAME` | No | *(empty)* | Sheet to read; empty means first sheet |
| `OUTPUT_SHEET_NAME` | No | `Analysis` | Sheet name of the output workbook |
| `OUTPUT_FILE_PREFIX` | No | `product_detail_log_analysis` | Output file name prefix |
| `INPUT_HAS_HEADER` | No | `true` | Whether row 1 of the raw file is a header |
| `SKIP_ROWS` | No | `0` | Rows to skip at the top (logo/title rows) |
| `DATE_FORMAT` | No | `%d.%m.%Y %H:%M:%S` | strftime format for dates in the output |
| `DATE_DAYFIRST` | No | `true` | Day-first parsing for text dates (`01.02.2025` → 1 February) |
| `TIMESTAMP_OUTPUT` | No | `true` | Append a timestamp to the output file name |

CLI overrides: `--input`, `--input-dir`, `--output-dir`.

---

## 6. Input specification

Columns are read **by position**, not by name, so the header text in the raw
file is irrelevant.

| Cell | Content | Expected type |
|---|---|---|
| A | Barcode | Text or number |
| B | Log status | Text |
| C | Date | Excel date, or text date |

Columns beyond C are ignored.

---

## 7. Processing

1. **Load configuration** — read `.env`, apply CLI overrides.
2. **Resolve the input file** — use `INPUT_FILE`, or fall back to the most
   recently modified Excel file in `INPUT_DIR`. Excel lock files (`~$...`) are
   ignored.
3. **Read** columns A–C into a dataframe; record the original row index as a
   tie-breaker.
4. **Normalise the barcode.** Excel stores numeric barcodes as floats, so
   `8690000000001` can arrive as `8690000000001.0`. The trailing `.0` is
   stripped, otherwise the same physical product would split into two groups.
   Status text is trimmed.
5. **Parse dates** with `to_datetime`. Unparseable values become `NaT`.
6. **Drop unusable rows** — empty barcode or invalid date. Each class of
   skipped row is counted and reported in the console; the run continues.
7. **Sort** by `(barcode, date, original row order)`.
8. **Group** by barcode, aggregating:
   - `size` → Total Log Count
   - `first` of date/status → first log
   - `last` of date/status → last log
9. **Format dates** into `DATE_FORMAT` as text so Excel cannot re-interpret
   them under a different locale.
10. **Write** the output workbook.

### Tie-breaking

When two records of the same barcode carry the identical timestamp, the row
that appeared earlier in the raw file is treated as the earlier log. The sort is
stable (`mergesort`), so the result is deterministic across runs.

---

## 8. Output specification

File: `{OUTPUT_FILE_PREFIX}_{input file name}_{YYYYMMDD_HHMMSS}.xlsx`
in `OUTPUT_DIR`. One row per unique barcode, sorted by barcode ascending.

| Cell | Header |
|---|---|
| A | Barcode |
| B | Total Log Count |
| C | First Log Date |
| D | First Log Status |
| E | Last Log Date |
| F | Last Log Status |

Formatting: header row written once (bold white on dark blue, frozen, with
autofilter), Arial body text, auto-fitted column widths.

---

## 9. Error handling

| Case | Behaviour |
|---|---|
| `.env` missing | Warn, continue with defaults / system environment |
| Input directory missing | Error message, exit code 1 |
| No Excel file in directory | Error message, exit code 1 |
| Fewer than 3 columns | Error message, exit code 1 |
| Empty barcode / invalid date row | Warn with a count, skip the row, continue |
| No valid rows remain | Error message, exit code 1 |
| Output file open in Excel | "File is in use" message, exit code 1 |

Exit code `0` on success, `1` on failure — suitable for chaining in a batch job.

---

## 10. Alternatives considered

- **Reading columns by header name.** Rejected: the raw export's header text is
  not stable and may be in either Turkish or English. Positional reading matches
  the spec (Cell A/B/C) exactly.
- **Writing dates as real Excel date values.** Rejected for now: the display
  format would then depend on the viewer's locale, and the report is meant to be
  read, not re-calculated. Text preserves the intended format everywhere.
- **Hardcoded paths.** Rejected — the whole point of the `.env` requirement.

---

## 11. Testing

Verified against a sample file covering: numeric and alphanumeric barcodes,
float-encoded barcodes, identical timestamps within one barcode, empty barcode
rows, unparseable dates, text dates in `dd.mm.yyyy` form, header-present and
header-absent files, and a missing `.env`.

---

## 12. Future work

- Optional `--status-filter` to analyse only a subset of statuses.
- Duration column (last − first) in days.
- CSV input support.
- Multi-file batch mode producing one consolidated report.
