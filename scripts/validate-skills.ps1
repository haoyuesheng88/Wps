$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$skillsRoot = Join-Path $repoRoot 'skills'
$codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $HOME '.codex' }
$validator = Join-Path $codexHome 'skills\.system\skill-creator\scripts\quick_validate.py'

if (-not (Test-Path $validator)) {
    throw "Validator not found: $validator"
}

$skills = Get-ChildItem -Path $skillsRoot -Directory
if (-not $skills) {
    throw "No skills found under $skillsRoot"
}

$env:PYTHONUTF8 = '1'
$failed = @()

foreach ($skill in $skills) {
    Write-Output "Validating $($skill.Name)..."
    & python $validator $skill.FullName
    if ($LASTEXITCODE -ne 0) {
        $failed += $skill.Name
    }
}

if ($failed.Count -gt 0) {
    throw ('Validation failed for: ' + ($failed -join ', '))
}

Write-Output 'All skills validated successfully.'
