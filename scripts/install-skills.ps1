param(
    [string[]]$SkillNames,
    [string]$Destination,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$skillsRoot = Join-Path $repoRoot 'skills'

if (-not (Test-Path $skillsRoot)) {
    throw "Skills directory not found: $skillsRoot"
}

if (-not $Destination) {
    $codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $HOME '.codex' }
    $Destination = Join-Path $codexHome 'skills'
}

New-Item -ItemType Directory -Force -Path $Destination | Out-Null

$availableSkills = Get-ChildItem -Path $skillsRoot -Directory
if ($SkillNames -and $SkillNames.Count -gt 0) {
    $selectedSkills = foreach ($name in $SkillNames) {
        $match = $availableSkills | Where-Object { $_.Name -eq $name }
        if (-not $match) {
            throw "Requested skill not found in repo: $name"
        }
        $match
    }
} else {
    $selectedSkills = $availableSkills
}

foreach ($skill in $selectedSkills) {
    $targetPath = Join-Path $Destination $skill.Name
    if (Test-Path $targetPath) {
        if (-not $Force) {
            Write-Output "Skipping existing skill (use -Force to overwrite): $($skill.Name)"
            continue
        }
        Remove-Item -LiteralPath $targetPath -Recurse -Force
    }

    Copy-Item -Path $skill.FullName -Destination $targetPath -Recurse -Force
    Write-Output "Installed: $($skill.Name) -> $targetPath"
}
