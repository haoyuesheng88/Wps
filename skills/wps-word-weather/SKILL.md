---
name: wps-word-weather
description: Look up live weather for a city or district and insert a concise weather line into an active WPS Word or Microsoft Word document. Use when the user asks to type today's weather into WPS Word, add a translated weather line below the current text, or continue weather notes on the next line in an open document.
---

# WPS Word Weather

Use this skill to prepare a short weather sentence and insert it into the active WPS 文字 or Microsoft Word document.

## Workflow

1. Resolve the place and the date.
Use absolute dates in both the response and the inserted text when the user says “today”, “tomorrow”, or similar relative wording.

2. Fetch live weather first.
Do not invent current conditions. Prefer a concise summary with current condition, approximate temperature, and a brief later-today trend if relevant.

3. Draft one compact weather line.
A good default pattern is:

`YYYY-MM-DD [location]天气：当前[condition]，约[temp]°C；稍后[trend]。`

4. Insert into the active document.
Run [scripts/insert-into-active-word.ps1](scripts/insert-into-active-word.ps1). Use `-NewParagraph` when the user wants the line on the next line.

## PowerShell Usage

```powershell
& 'C:\path\to\insert-into-active-word.ps1' -Text $text
& 'C:\path\to\insert-into-active-word.ps1' -Text $text -NewParagraph
```

## Translation Notes

If the user asks for Japanese, German, or another language, translate the already prepared weather content unless they explicitly request a fresh weather lookup.

## Failure Handling

- If Word automation fails because no active document is available, say so clearly.
- If live weather lookup is unavailable, stop and report that instead of fabricating a forecast.
- Insert only at the current selection.
