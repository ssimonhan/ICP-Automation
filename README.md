# ICP Data Processing Scripts

Python scripts for processing fixed-template ICP export files into cleaned Excel workbooks and original-solution concentration summaries.

Author: Shihua Han  
Version: 0.1.0

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

## License

This project is intended to be released under the GNU General Public License.

Copyright (C) 2026 Shihua Han

## Input Format

The raw ICP file should follow the expected template layout:

- Row 1 contains element names
- Row 2 contains repeated subheaders, including `Conc.`
- Each element uses a repeating 3-column group
- The sample name column is present in the raw data
- Sample names use a prefix followed by a number, such as `SH1`, `T20`, `H43`, or `CS4`

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
| >10% | light orange |
| >20% | orange |
| >30% | darker orange |
| >40% | dark orange |
| >50% | darkest orange |

### Blank Samples

Rows with `blk` or `blank` in the sample name are treated as blank rows.

Cells are highlighted when values exceed these absolute concentration thresholds:

| Value | Fill color |
| --- | --- |
| >0.1 ppb | light orange |
| >0.2 ppb | orange |
| >0.3 ppb | darker orange |
| >0.4 ppb | dark orange |
| >0.5 ppb | darkest orange |

### ICP Concentration Range

In the `<initials> ICP` sheet, element concentration cells between `10` and `400` ppb are highlighted using Excel conditional formatting.

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

1. If any ICP value is between `10` and `400` ppb, select the concentration corresponding to the highest ICP value in that range.
2. Otherwise, if any ICP value is between `1` and `10` ppb, select the concentration corresponding to the highest ICP value in that range.
3. Otherwise, if all ICP values are `<1` ppb, select the concentration corresponding to the maximum ICP value.
4. Otherwise, if all ICP values are `>400` ppb, select the concentration corresponding to the minimum ICP value.

## Highlighting Rules in Step 2

The concentration workbook uses Excel conditional formatting:

| Condition | Fill color |
| --- | --- |
| Selected from 10-400 ppb range | light green |
| Selected from 1-10 ppb range | light orange |
| All values <1 ppb | light gray |
| All values >400 ppb | light gray |
| Final selected concentration row | gold base fill |

Copied ICP ppb rows and calculated ppm rows also receive row-wise Excel data bars across the element columns.

## Notes

- Keep input filenames reasonably short if possible. Excel sheet names are limited to 31 characters, and very long filenames may require shortened sheet names.
- Close the workbook before rerunning a script that writes to it.
- Missing concentration values in selected sample rows are reported as warnings and left blank.
- If Excel automation is unavailable, the workbook is still created, but native Excel data bars or some conditional formatting may not be applied.
