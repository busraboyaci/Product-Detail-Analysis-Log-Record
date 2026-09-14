# How to Run — Daily Product Detail Log Analysis

This tool turns a raw product-log Excel export into a per-barcode summary
(one row per barcode: total log count, first log date + status, last log date + status).

There are two parts below:
- **Part A — One-time setup** (do this once on a computer).
- **Part B — Daily routine** (do this every day).

---

## Part A — One-time setup (do once)

You only need to do these steps the first time, or on a new computer.

### 1. Install Python
- Download Python from <https://www.python.org/downloads/> and install it.
- **Important:** on the first install screen, tick **“Add Python to PATH.”**
- To confirm, open a terminal and run:
  ```bash
  py --version
  ```
  You should see something like `Python 3.14.4`.

### 2. Install the required libraries
Open a terminal **in the project folder** and run:
```bash
py -m pip install -r requirements.txt
```
This installs `pandas`, `openpyxl`, and `python-dotenv`.

### 3. Confirm the `.env` file exists
The project already includes a `.env` file with these settings (no need to change
anything for normal use):
```
INPUT_DIR=./input
OUTPUT_DIR=./output
```
If `.env` is ever missing, copy `.env.example` to `.env`.

### 4. Test it once
Double-click **`run_daily.bat`**. Because the `input\` folder is empty on a fresh
setup, it will simply say *“No Excel files found”* — that means the setup works.

✅ Setup is done. You won't repeat Part A again.

---

## Part B — Daily routine (do every day)

### The short version
1. Put today's log export file(s) into the **`input\`** folder.
2. Double-click **`run_daily.bat`**.
3. Open the newest file(s) in the **`output\`** folder — those are your reports.

That's it. Details below.

### Step 1 — Get today's log export(s)
Download / export today's product-log Excel file(s) from your source
(e.g. Trendyol and Hepsiburada panels). Each file must have, by column position:

| Column A | Column B | Column C |
|---|---|---|
| Barcode | Log status | Date |

### Step 2 — Put the file(s) in the `input\` folder
Copy them into the project's **`input\`** folder. You can put in **one or several**
files — the runner processes every Excel file it finds.

### Step 3 — Run it
Double-click **`run_daily.bat`**. A window opens and shows progress, for example:
```
[RUN]  Analysing: Logkaydı-14.09.2026-trendyol.xlsx
[OK]   481 unique barcode(s) reported.
[DONE] Moved to archive\: Logkaydı-14.09.2026-trendyol.xlsx
```
When it finishes it prints *“Your reports are in the output\ folder”* and waits —
press any key to close the window.

### Step 4 — Read your reports
Open the **`output\`** folder. Each run creates one report per input file, named:
```
product_detail_log_analysis_<original file name>_<date>_<time>.xlsx
```
The report has these columns:

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| Barcode | Total Log Count | First Log Date | First Log Status | Last Log Date | Last Log Status |

---

## What happens to my files? (folders explained)

| Folder | Purpose |
|---|---|
| `input\`   | Where **you drop** today's raw log files before running. |
| `output\`  | Where **reports are written**. Files are timestamped, so nothing is ever overwritten — reports accumulate here as history. |
| `archive\` | After a file is analysed successfully, the runner **moves it here** so it isn't re-processed tomorrow. Your raw files are kept, just tidied away. |

You can safely delete old files from `output\` and `archive\` whenever you want to
free space; they are just history.

---

## Alternative: run one specific file manually

If you don't want to use the `.bat`, open a terminal in the project folder and run:
```bash
py product_detail_log_analysis.py --input "Logkaydı-14.09.2026-trendyol.xlsx"
```
(The file name is looked up inside the `input\` folder.) Running with **no**
`--input` analyses only the *newest* Excel file in `input\`.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| **“'py' is not recognized”** | Python isn't installed or wasn't added to PATH. Redo Part A, step 1, ticking “Add Python to PATH.” |
| **“No module named pandas”** | Redo Part A, step 2 (`py -m pip install -r requirements.txt`). |
| **`[FAIL]` and file stays in `input\`** | The report couldn't be written — usually because a previous report is open in Excel, or the input file has fewer than 3 columns. Close Excel and run again; the message in the window says which. |
| **“No Excel files found”** | You didn't put a file in `input\`, or it's still open/locked in Excel (files starting with `~$` are ignored). |
| **Turkish characters look wrong in the window** | Cosmetic only — the actual `.xlsx` report has the correct text. |

---

## Optional: make it even faster

- **Pin it:** right-click `run_daily.bat` → *Send to* → *Desktop (create shortcut)*
  so you can run it from your desktop.
- **Fully automatic:** it can be scheduled to run itself every day with Windows Task
  Scheduler. Ask and this can be set up, but note it only helps if today's export
  files are already in `input\` at that time.
