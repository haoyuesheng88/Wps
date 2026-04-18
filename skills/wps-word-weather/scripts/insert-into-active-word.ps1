param(
    [Parameter(Mandatory = $true)]
    [string]$Text,

    [switch]$NewParagraph,

    [switch]$StartIfMissing
)

$ErrorActionPreference = 'Stop'

function Get-ActiveWordApp {
    foreach ($progId in @('Word.Application', 'kwps.Application')) {
        try {
            return [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        } catch {
        }
    }

    return $null
}

function Find-WordExecutable {
    $patterns = @(
        (Join-Path $env:LOCALAPPDATA 'Kingsoft\WPS Office\*\office6\wps.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\WINWORD.EXE'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\WINWORD.EXE')
    )

    $matches = foreach ($pattern in $patterns) {
        Get-Item $pattern -ErrorAction SilentlyContinue
    }

    return $matches |
        Sort-Object FullName -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}

function Get-OrStartWordApp {
    $app = Get-ActiveWordApp
    if ($null -ne $app) {
        return $app
    }

    if (-not $StartIfMissing) {
        throw 'WPS Word 或 Microsoft Word 当前未打开。'
    }

    $exe = Find-WordExecutable
    if (-not $exe) {
        throw '未找到 WPS Word 或 Microsoft Word 的可执行文件。'
    }

    Start-Process -FilePath $exe | Out-Null

    for ($i = 0; $i -lt 12; $i++) {
        Start-Sleep -Seconds 1
        $app = Get-ActiveWordApp
        if ($null -ne $app) {
            return $app
        }
    }

    throw 'Word 应用已启动，但未能连接到自动化对象。'
}

$app = Get-OrStartWordApp
$app.Visible = $true

if ($app.Documents.Count -eq 0) {
    if (-not $StartIfMissing) {
        throw '已连接到 Word 应用，但当前没有打开的文档。'
    }

    $app.Documents.Add() | Out-Null
}

$selection = $app.Selection
if ($NewParagraph) {
    $selection.TypeParagraph()
}

$selection.TypeText($Text)
Write-Output '文本已插入到当前文档。'
