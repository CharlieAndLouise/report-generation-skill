---
name: excel-report
description: Generate a formatted Excel report from company data with the company logo, a title, and a summary row. Use this skill whenever the user asks to create a report, generate an Excel file, export data to Excel, build a spreadsheet from employee or company data, or wants to filter/summarize/reorganize data into a file — even if they don't say "Excel" or "report" explicitly. The user drives the content via natural language: they can specify filters (e.g., "top 2 by salary"), column selection, grouping, sorting, and formatting.
---

# Excel Report Generator

Create a polished Excel report from `excel-report/resources/data.json` using the company logo, based on what the user asks for.

## Report Layout (always use this structure)

| Row | Content |
|-----|---------|
| 1 | Logo (A1, top-left) + Report title (C1, bold, beside logo, make sure the the column is aligned to data's last column) |
| 2 | *(part of logo area — logo spans rows 1–2)* |
| 3 | Summary line: italic, light gray background, merged across all data columns |
| 4 | Column headers: bold, dark blue background, white text |
| 5+ | Data rows: alternating white / light blue shading |

## Steps

### 1. Understand the request
Parse the user's prompt for:
- **Filter**: what subset of data to show (e.g., "top 2 earners per department", "only Tech")
- **Columns**: which fields to include and in what order (default: all fields)
- **Grouping/sorting**: how to arrange rows (e.g., "grouped by department, sorted by salary desc")
- **Formatting**: any special treatment (e.g., "salary as currency", "department in uppercase")
- **Title**: derive a clear report title from the request (e.g., "Top Earners by Department")

### 2. Load and transform the data
Read `excel-report/resources/data.json`. Apply the filters, sorting, and transformations in Python. Keep it simple — a short script or inline logic is fine.

### 3. Generate the Excel file
Use the bundled Node.js script at `excel-report/scripts/create_report.js`. It handles all layout, logo placement, and formatting. You only need to:
- Prepare a list of objects (the transformed data rows)
- Write a temp JSON file with that data
- Call the script

```bash
# Install dependency if needed (once per project)
npm install exceljs --prefix c:/work/NewSkills

# Write your transformed data to a temp JSON file, then run:
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
node excel-report/scripts/create_report.js \
  --title "Your Report Title" \
  --summary "Your summary sentence" \
  --data /tmp/report_data.json \
  --logo excel-report/resources/company_logo.png \
  --output output/report_${TIMESTAMP}.xlsx
```

The script creates the `output/` folder automatically.

> On Windows, generate the timestamp inline with PowerShell if needed:
> `$ts = Get-Date -Format 'yyyyMMdd_HHmmss'`

### 4. Write the summary sentence
The summary should describe what's in the report in one sentence. Example:
> "Top 2 highest-paid employees per department — 10 employees across 5 departments. Generated 2025-01-15."

### 5. Tell the user
Report the output file path and a one-line description of what the report contains.

## Tips
- If the user asks for "top N per group", group with a plain JS object and sort with `Array.sort()` — no external libraries needed.
- Salary values in the source data are strings — parse with `parseInt()` before comparing.
- If the user doesn't specify columns, include all fields from the source data.
- The title in the Excel header should be human-readable (title case, spaces not underscores).
- `exceljs` is available — require it with `const ExcelJS = require('exceljs')`.
