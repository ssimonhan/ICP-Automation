# ICP Automation

Python tools and a Streamlit app for processing fixed-template ICP export files into cleaned Excel workbooks and final concentration summaries.

Author: Shihua Han  
Version: 0.4.1  

---

## Overview

This project provides two ways to use the ICP processing pipeline:

### 1. Streamlit App (Recommended)
A user-friendly web interface that:
- Uploads raw ICP files
- Generates cleaned ICP sheets
- Allows editing dilution/sample directly in-browser
- Produces final concentration results
- Eliminates manual Excel editing steps

This app is compatible with [Streamlit Cloud](https://icp-automation.streamlit.app/). Some advanced Excel formatting features (e.g., Excel-native automatic data bars) are simplified in cloud environments due to platform limitations, but all core data processing and outputs remain unchanged.

### 2. Command-Line Scripts
Traditional workflow using:
- `process_icp_data.py`
- `generate_icp_concentrations.py`

---

## Installation

### Requirements
- Python 3.8+
- (Optional) Microsoft Excel for full formatting support

### Install dependencies
```bash
pip install -r requirements.txt
```

---

## Running the Streamlit App (Recommended)

### Start the app

#### Step 1 — Open a terminal

Navigate to the project folder:

```bash
cd path/to/ICP-Automation
```

For example (Windows):

```bash
cd C:\Users\YourName\ICP-Automation
```

#### Step 2 — (Optional) Activate your environment

If using virtual environment:

**Windows**

```bash
.\.venv\Scripts\activate
```
**Mac/Linux**

```bash
source .venv/bin/activate
```

#### Step 3 — Install dependencies (first time only)

```bash
pip install -r requirements.txt
```

#### Step 4 — Start the app

```bash
python -m streamlit run app.py
```

#### Step 5 — Open in browser

After running, you should see something like:

```
Local URL: http://localhost:8501
```

Open that URL in your browser:

```
http://localhost:8501
```

#### Step 6 — Stop the app

To stop the app, go back to the terminal and press:

```
Ctrl + C
```

---

### Streamlit Workflow (Fully Automated)

#### Step 1 — Upload raw ICP file
- Accepts .csv or .xlsx
- Generates a cleaned ICP workbook
- Optional: download intermediate cleaned file

#### Step 2 — Edit data in-app
- Editable table appears
- Modify:
  - Dilution (Column A)
  - Sample (Column B)
- Sample Name (Column C) is preserved

#### Step 3 — Generate final results
- Runs concentration calculations automatically
- Produces final Excel workbook

No manual Excel editing required

---

### Outputs

#### Intermediate output (optional)
- Cleaned ICP workbook
- Includes:
  - raw data
  - cleaned data
  - <initials> ICP sheet

#### Final output
- Concentration workbook:
  - grouped by sample
  - ppm calculations
  - selected final concentrations

---

### Excel Formatting Notes

- Full formatting requires Microsoft Excel
- Without Excel:
  - Calculations still work
  - Some formatting may be missing

---

## Command-Line Workflow (Optional)

### Step 1

```bash
python process_icp_data.py input.xlsx --initials SH --output SH_processed.xlsx
```

### Step 2

```bash
python generate_icp_concentrations.py --source SH_processed.xlsx --output SH_concentrations.xlsx
```

Manual editing required in CLI mode

---

## Key Features

- Automated ICP preprocessing
- Editable in-app workflow
- Excel output generation
- Concentration selection logic

---

## Optional Launcher (Windows)

```bat
@echo off
cd /d %~dp0
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -Command "python -m streamlit run app.py"
start http://localhost:8501
```

---

# Core ICP Processing Scripts

These scripts implement the underlying data processing pipeline used by the Streamlit app. They process fixed-template ICP export files into cleaned Excel workbooks and original-solution concentration summaries. They can also be run independently from the command line.

The workflow uses:

- `pandas` for reading and extracting raw ICP data
- `openpyxl` for workbook creation and formatting
- Microsoft Excel automation, when available, for native Excel conditional formatting and data bars

## Files

- `process_icp_data.py`  
  Validates a raw ICP export, extracts sample and element concentration columns, highlights standards/blanks, and creates an editable ICP workbook.

- `generate_icp_concentrations.py`  
  Uses the edited ICP workbook to calculate original-solution concentrations in ppm and select one concentration per element.

## Requirements

- Windows with Python installed
- Microsoft Excel installed, recommended for native Excel data bars and conditional formatting
- Python packages:

```bash
pip install pandas openpyxl
```

## Input Format

The raw ICP file should follow the expected template layout:

- Row 1 contains element names
- Row 2 contains repeated subheaders, including `Meas. Conc. [ ppb ]` or `Measured Conc. [ ppb ]`
- Each element uses a repeating 3-column group
- The sample name column is present in the raw data
- Sample names use a prefix followed by letters or numbers, with an optional underscore, such as `SH1`, `T20`, `H43`, `CS4`, `CS_10k`, or `CS_x10k`

## Step 1: Process Raw ICP Data

Run `process_icp_data.py` from the folder containing the scripts.

Example:

```bash
python process_icp_data.py "raw_icp_data.csv" --initials SH --output "SH_processed.xlsx"
```

For another sample prefix:

```bash
python process_icp_data.py "raw_icp_data.csv" --initials T --output "T_processed.xlsx"
python process_icp_data.py "raw_icp_data.csv" --initials H --output "H_processed.xlsx"
python process_icp_data.py "raw_icp_data.csv" --initials CS --output "CS_processed.xlsx"
```

If the input is an Excel workbook with multiple sheets, specify the source sheet:

```bash
python process_icp_data.py "raw_icp_data.xlsx" --sheet "Raw Data" --initials SH --output "SH_processed.xlsx"
```

Use quotes around file or sheet names that contain spaces.

To customize the calibration ranges from the command line, pass the same values to both scripts:

```bash
python process_icp_data.py "raw_icp_data.csv" --initials CS --orange-min 1 --green-min 10 --green-max 400 --internal-light-orange 20 --internal-orange 40 --blank-light-orange 0.2 --blank-orange 1 --output "CS_processed.xlsx"
python generate_icp_concentrations.py --source "CS_processed.xlsx" --orange-min 1 --green-min 10 --green-max 400 --output "CS_concentrations.xlsx"
```

## Output From Step 1

The processed workbook includes sheets such as:

- Original raw data
- Highlighted cleaned data
- `<initials> samples`
- `<initials> ICP`

For example, with `--initials SH`, the sample sheets are:

- `SH samples`
- `SH ICP`

The `<initials> ICP` sheet is the sheet to edit before running the second script.

## Editing the ICP Sheet

In the `<initials> ICP` sheet:

- Column A: `Dilution factor`
- Column B: `Sample`
- Column C: `Sample Name`
- Element concentrations begin in column D

Enter the dilution factor and sample name for each row.

Blank cells in the `Sample` column mean "same sample as above." If rows for the same sample are not adjacent, fill in the `Sample` column explicitly wherever the sample changes.

## Highlighting Rules in Step 1

### Internal Standards

Rows with `ppb` in the sample name are treated as internal standard rows. The script compares each measured value to the nominal ppb value in the sample name.

Cells are highlighted when the relative error is beyond these thresholds:

| Relative error | Fill color |
| --- | --- |
| <=20% | no fill |
| >20% to <=40% | orange |
| >40% | dark orange |

### Blank Samples

Rows with `blk` or `blank` in the sample name are treated as blank rows.

Cells are highlighted when values exceed these absolute concentration thresholds:

| Value | Fill color |
| --- | --- |
| <=0.2 ppb | no fill |
| >0.2 to <=1 ppb | orange |
| >1 ppb | dark orange |

### ICP Concentration Range

In the `<initials> ICP` sheet, element concentration cells in the accurate calibration range are highlighted using Excel conditional formatting. The default accurate range is `10` to `400` ppb.

Each row also receives Excel data bars across the element columns:

- Blue bars for positive values
- Red bars for negative values
- Zero-axis scaling
- Row-wise automatic minimum and maximum scaling

## Step 2: Generate Original-Solution Concentrations

After entering dilution factors and sample names in the `<initials> ICP` sheet, save and close the workbook. Then run:

```bash
python generate_icp_concentrations.py --source "SH_processed.xlsx" --output "SH_concentrations.xlsx"
```

For another prefix:

```bash
python generate_icp_concentrations.py --source "T_processed.xlsx" --output "T_concentrations.xlsx"
python generate_icp_concentrations.py --source "H_processed.xlsx" --output "H_concentrations.xlsx"
python generate_icp_concentrations.py --source "CS_processed.xlsx" --output "CS_concentrations.xlsx"
```

The script automatically uses the only sheet ending in ` ICP`, such as `SH ICP`, `T ICP`, or `CS ICP`.

If the workbook contains multiple ICP sheets, specify one:

```bash
python generate_icp_concentrations.py --source "processed.xlsx" --sheet "T ICP" --output "T_concentrations.xlsx"
```

## Concentration Calculation

For each sample group, the script copies the ICP concentrations in ppb, then calculates original-solution concentrations in ppm:

```text
c(ppm) = c(ppb) / 1000 * dilution factor
```

The calculated ppm rows are written below the copied ppb rows.

## Concentration Selection Rules

For each element in each sample group, the script selects one final concentration based on the ICP ppb values.

Selection priority:

1. If any ICP value is in the green accurate range, select the concentration corresponding to the lowest dilution factor in that range.
2. Otherwise, if any ICP value is in the orange lower-confidence range, select the concentration corresponding to the lowest dilution factor in that range.
3. Otherwise, if all ICP values are below the orange minimum, select the concentration corresponding to the lowest dilution factor.
4. Otherwise, if all ICP values are above the green maximum, select the concentration corresponding to the highest dilution factor.

The Streamlit app exposes these values in **Advanced settings**. Defaults are:

- Orange lower-confidence range: `1` ppb to below `10` ppb
- Green accurate range: `10` ppb through `400` ppb

## Highlighting Rules in Step 2

The concentration workbook uses Excel conditional formatting:

| Condition | Fill color |
| --- | --- |
| Selected from green accurate range | light green |
| Selected from orange lower-confidence range | light orange |
| All values below orange minimum | light gray |
| All values above green maximum | light red |
| Final selected concentration row | gold base fill |

Copied ICP ppb rows and calculated ppm rows also receive row-wise Excel data bars across the element columns.

## Notes

- Keep input filenames reasonably short if possible. Excel sheet names are limited to 31 characters, and very long filenames may require shortened sheet names.
- Close the workbook before rerunning a script that writes to it.
- Missing concentration values in selected sample rows are reported as warnings and left blank.
- If Excel automation is unavailable, the workbook is still created, but native Excel data bars or some conditional formatting may not be applied.

---

# License

This project is intended to be released under the GNU General Public License.

Copyright (C) 2026 Shihua Han

---

# Acknowledgments

* This script was developed with the assistance of OpenAI Codex and Microsoft Copilot to optimize the data parsing and automation workflows.
