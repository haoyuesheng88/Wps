---
name: wps-word-insert-text
description: Insert text into the active WPS Word or Microsoft Word document through desktop automation on Windows. Use when the user asks to type text into an already open WPS Word document, continue on the next line, add a translated line below, or open WPS Word and insert prepared content.
---

# WPS Word Insert Text

Use this skill to place prepared text into the current WPS 文字 or Microsoft Word document.

## Workflow

1. Draft the text first.
Assemble the final content before running the script. Keep relative dates resolved to absolute dates when the request says “today”, “tomorrow”, or similar.

2. Insert into the active document.
Run [scripts/insert-into-active-word.ps1](scripts/insert-into-active-word.ps1). Use `-NewParagraph` when the user asks for a new line or wants the new content below the current selection.

3. Start Word only when the task requires it.
If the user says WPS Word should be opened first, pass `-StartIfMissing`. Otherwise, prefer attaching to the already open document.

## PowerShell Usage

```powershell
& 'C:\path\to\insert-into-active-word.ps1' -Text $text
& 'C:\path\to\insert-into-active-word.ps1' -Text $text -NewParagraph
& 'C:\path\to\insert-into-active-word.ps1' -Text $text -StartIfMissing
```

## Failure Handling

- If no active Word application is available and the user did not ask to open one, say that clearly.
- Insert only at the current selection; do not overwrite unrelated content.
- If the task depends on live data, fetch that data before writing instead of inventing it.
