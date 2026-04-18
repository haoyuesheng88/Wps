---
name: wps-excel-case
description: Open WPS Spreadsheets (ET) or Microsoft Excel, create a new workbook, and insert a formatted sample table or matrix. Use when the user asks to open WPS Excel, WPS表格, or ET and enter a sales case, inventory sheet, attendance table, expense sheet, demo workbook, permission matrix, or image-to-Excel table with polished formatting.
---

# WPS Excel Case

Use this skill to create clean, editable spreadsheets in WPS 表格 on Windows.

## Workflow

1. Resolve the table type before touching Excel.
If the user provides a screenshot or photo of a table, transcribe the structure first and choose the closest preset. Use `permission-matrix` for access-right tables with checkmarks and `-` placeholders.

2. Start or attach to WPS 表格.
Run [scripts/new-wps-excel-case.ps1](scripts/new-wps-excel-case.ps1). The script connects to an active spreadsheet app first and launches WPS ET automatically when needed.

3. Choose the preset.
Supported `CaseType` values are documented in [references/case-catalog.md](references/case-catalog.md).

4. Keep the result presentation-ready.
The script already applies title rows, merged headers, borders, alignment, widths, freeze panes, and type-specific highlighting. Use that instead of ad hoc cell styling when the preset already matches the request.

5. Confirm what was created.
Report the selected preset and that a new workbook was created.

## PowerShell Usage

```powershell
& 'C:\path\to\new-wps-excel-case.ps1'
& 'C:\path\to\new-wps-excel-case.ps1' -CaseType inventory
& 'C:\path\to\new-wps-excel-case.ps1' -CaseType permission-matrix
```

If the user only says “输入一个案例” or “做个示例表”, default to `sales`.

## Failure Handling

- If WPS 表格 is not installed and no spreadsheet COM object is available, stop and report that clearly.
- If the image contains a custom table that does not fit a preset, transcribe the data manually first instead of pretending a preset is exact.
- If workbook creation or formatting fails, surface the automation error instead of claiming success.
