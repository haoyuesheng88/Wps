# Wps Skills Repository

A small cross-machine repository of Codex skills for WPS Office automation on Windows.

## Included Skills

- `wps-excel-case`: open WPS Spreadsheets (ET) and build formatted demo workbooks such as sales tables, inventory sheets, attendance summaries, expense logs, and permission matrices.
- `wps-word-insert-text`: type text into the active WPS Word or Microsoft Word document.
- `wps-word-weather`: look up live weather and insert a concise weather line into the active WPS Word or Microsoft Word document.

## Repository Layout

- `skills/`: installable Codex skills
- `scripts/install-skills.ps1`: copy the skills into the current machine's Codex skills directory
- `scripts/validate-skills.ps1`: run the local skill validator against every skill in this repo

## Install On Another Computer

```powershell
git clone https://github.com/haoyuesheng88/Wps.git
cd Wps
powershell -ExecutionPolicy Bypass -File .\\scripts\\install-skills.ps1 -Force
```

By default the installer copies the skills into `%USERPROFILE%\\.codex\\skills`. If `CODEX_HOME` is set, it uses `%CODEX_HOME%\\skills` instead.

## Validate

```powershell
powershell -ExecutionPolicy Bypass -File .\\scripts\\validate-skills.ps1
```

## Notes

- These skills target Windows desktop automation through COM, so WPS Office or Microsoft Office should already be installed.
- The Excel skill starts WPS Spreadsheets automatically when needed.
- The Word insertion skills expect an active document, unless the caller explicitly chooses to start Word and create a blank document first.
