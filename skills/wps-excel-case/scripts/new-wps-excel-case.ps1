param(
    [ValidateSet('sales','inventory','attendance','expense','permission-matrix')]
    [string]$CaseType = 'sales'
)

$ErrorActionPreference = 'Stop'

function Get-ActiveSpreadsheetApp {
    foreach ($progId in @('ket.Application', 'Excel.Application')) {
        try {
            return [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        } catch {
        }
    }

    return $null
}

function Find-WpsSpreadsheetExe {
    $patterns = @(
        (Join-Path $env:LOCALAPPDATA 'Kingsoft\WPS Office\*\office6\et.exe'),
        (Join-Path $env:ProgramFiles 'WPS Office\office6\et.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'WPS Office\office6\et.exe')
    )

    $matches = foreach ($pattern in $patterns) {
        Get-Item $pattern -ErrorAction SilentlyContinue
    }

    return $matches |
        Sort-Object FullName -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}

function Get-OrStartSpreadsheetApp {
    $app = Get-ActiveSpreadsheetApp
    if ($null -ne $app) {
        return $app
    }

    $exe = Find-WpsSpreadsheetExe
    if (-not $exe) {
        throw 'WPS 表格未安装，且未找到 et.exe。'
    }

    Start-Process -FilePath $exe | Out-Null

    for ($i = 0; $i -lt 12; $i++) {
        Start-Sleep -Seconds 1
        $app = Get-ActiveSpreadsheetApp
        if ($null -ne $app) {
            return $app
        }
    }

    throw 'WPS 表格已启动，但未能连接到自动化对象。'
}

function Set-CellValue {
    param(
        $Worksheet,
        [int]$Row,
        [int]$Column,
        $Value
    )

    $Worksheet.Cells.Item($Row, $Column) = $Value
}

function Get-CaseDefinition {
    param([string]$Type)

    switch ($Type) {
        'sales' {
            return @{
                SheetName = '销售案例'
                Title = '销售统计案例'
                Subtitle = '月度销量、单价与销售额示例'
                Headers = @('月份','产品','销售数量','单价','销售额')
                Rows = @(
                    @('1月','A产品',120,88),
                    @('1月','B产品',95,76),
                    @('2月','A产品',132,88),
                    @('2月','B产品',110,76),
                    @('3月','A产品',145,88),
                    @('3月','B产品',118,76)
                )
                Formulas = @(
                    @{ Column = 5; Template = '=C{0}*D{0}' }
                )
                TotalLabel = '合计'
                TotalColumn = 5
                TotalFormula = '=SUM(E4:E9)'
                NumberFormats = @{ 'D4:E10' = '0.00' }
                ColumnWidths = @{ 'A' = 10; 'B' = 14; 'C' = 12; 'D' = 12; 'E' = 14 }
            }
        }
        'inventory' {
            return @{
                SheetName = '库存案例'
                Title = '库存管理案例'
                Subtitle = '库存、安全库存与补货建议示例'
                Headers = @('物料编码','物料名称','当前库存','安全库存','建议补货','状态')
                Rows = @(
                    @('RM-001','滤芯',180,200),
                    @('RM-002','包装箱',320,250),
                    @('RM-003','标签纸',90,120),
                    @('RM-004','密封圈',60,80),
                    @('RM-005','阀门',45,40)
                )
                Formulas = @(
                    @{ Column = 5; Template = '=IF(C{0}<D{0},D{0}-C{0},0)' },
                    @{ Column = 6; Template = '=IF(C{0}<D{0},"需补货","正常")' }
                )
                ColumnWidths = @{ 'A' = 14; 'B' = 14; 'C' = 12; 'D' = 12; 'E' = 12; 'F' = 12 }
            }
        }
        'attendance' {
            return @{
                SheetName = '考勤案例'
                Title = '员工考勤案例'
                Subtitle = '出勤、请假与出勤率示例'
                Headers = @('姓名','部门','应出勤','实出勤','迟到次数','请假天数','出勤率')
                Rows = @(
                    @('张三','生产部',22,21,1,0),
                    @('李四','质量部',22,20,2,1),
                    @('王五','仓储部',22,22,0,0),
                    @('赵六','采购部',22,19,1,2),
                    @('陈七','销售部',22,21,3,0)
                )
                Formulas = @(
                    @{ Column = 7; Template = '=D{0}/C{0}' }
                )
                NumberFormats = @{ 'G4:G8' = '0.00%' }
                ColumnWidths = @{ 'A' = 10; 'B' = 12; 'C' = 10; 'D' = 10; 'E' = 10; 'F' = 10; 'G' = 12 }
            }
        }
        'expense' {
            return @{
                SheetName = '报销案例'
                Title = '费用报销案例'
                Subtitle = '费用报销单据汇总示例'
                Headers = @('日期','申请人','部门','费用类型','金额','审批状态','备注')
                Rows = @(
                    @('2026-04-01','张三','销售部','交通费',126.5,'已审批','客户拜访'),
                    @('2026-04-03','李四','采购部','餐费',88,'待审批','供应商接待'),
                    @('2026-04-05','王五','质量部','住宿费',360,'已审批','外地验厂'),
                    @('2026-04-08','赵六','生产部','办公用品',245,'已审批','标签耗材'),
                    @('2026-04-10','陈七','仓储部','运输费',520,'待审批','紧急调货')
                )
                TotalLabel = '合计'
                TotalColumn = 5
                TotalFormula = '=SUM(E4:E8)'
                NumberFormats = @{ 'E4:E9' = '0.00' }
                ColumnWidths = @{ 'A' = 12; 'B' = 10; 'C' = 12; 'D' = 12; 'E' = 12; 'F' = 12; 'G' = 16 }
            }
        }
        'permission-matrix' {
            return @{
                SheetName = '权限矩阵'
                Title = '系统权限矩阵表'
                Subtitle = '适用于角色与权限交叉对照表'
                Headers = @('序号','权限名称','管理员','主管','工程师','操作员','访问者')
                Rows = @(
                    @('1','用户管理','√','-','-','-','-'),
                    @('2','系统退出','√','√','√','√','√'),
                    @('3','系统桌面','-','√','√','√','√'),
                    @('4','系统参数','-','√','√','-','-'),
                    @('5','报警配置','-','√','√','√','-'),
                    @('6','消警','-','√','√','-','-'),
                    @('7','手动操作','-','√','√','-','-'),
                    @('8','批操作','-','√','√','√','-'),
                    @('9','配方编辑','-','√','√','-','-'),
                    @('10','配方下载','-','√','√','-','-'),
                    @('11','报警确认','-','√','√','√','-'),
                    @('12','报表导出','-','√','√','√','-'),
                    @('13','电子记录查看','-','√','-','-','-'),
                    @('14','历史数据查看','-','√','√','√','-')
                )
                ColumnWidths = @{ 'A' = 8; 'B' = 18; 'C' = 11; 'D' = 11; 'E' = 11; 'F' = 11; 'G' = 11 }
                HighlightChecks = $true
                HighlightRangeStartColumn = 3
            }
        }
    }
}

function Write-CaseWorksheet {
    param(
        $Worksheet,
        [hashtable]$Case
    )

    $headerCount = $Case.Headers.Count
    $titleRow = 1
    $subtitleRow = 2
    $headerRow = 3
    $dataStartRow = 4

    $Worksheet.Name = $Case.SheetName

    $Worksheet.Range($Worksheet.Cells.Item($titleRow, 1), $Worksheet.Cells.Item($titleRow, $headerCount)).Merge() | Out-Null
    Set-CellValue -Worksheet $Worksheet -Row $titleRow -Column 1 -Value $Case.Title
    $titleRange = $Worksheet.Range($Worksheet.Cells.Item($titleRow, 1), $Worksheet.Cells.Item($titleRow, $headerCount))
    $titleRange.Font.Bold = $true
    $titleRange.Font.Size = 16
    $titleRange.Font.Name = 'Microsoft YaHei UI'
    $titleRange.HorizontalAlignment = -4108
    $titleRange.VerticalAlignment = -4108
    $titleRange.Interior.Color = 0xD9EAD3
    $titleRange.RowHeight = 30

    $Worksheet.Range($Worksheet.Cells.Item($subtitleRow, 1), $Worksheet.Cells.Item($subtitleRow, $headerCount)).Merge() | Out-Null
    Set-CellValue -Worksheet $Worksheet -Row $subtitleRow -Column 1 -Value $Case.Subtitle
    $subtitleRange = $Worksheet.Range($Worksheet.Cells.Item($subtitleRow, 1), $Worksheet.Cells.Item($subtitleRow, $headerCount))
    $subtitleRange.Font.Size = 10
    $subtitleRange.Font.Name = 'Microsoft YaHei UI'
    $subtitleRange.Font.Color = 0x666666
    $subtitleRange.HorizontalAlignment = -4108
    $subtitleRange.VerticalAlignment = -4108
    $subtitleRange.Interior.Color = 0xF3F6F4
    $subtitleRange.RowHeight = 20

    for ($column = 0; $column -lt $headerCount; $column++) {
        Set-CellValue -Worksheet $Worksheet -Row $headerRow -Column ($column + 1) -Value $Case.Headers[$column]
    }

    $headerRange = $Worksheet.Range($Worksheet.Cells.Item($headerRow, 1), $Worksheet.Cells.Item($headerRow, $headerCount))
    $headerRange.Font.Bold = $true
    $headerRange.Font.Name = 'Microsoft YaHei UI'
    $headerRange.Font.Size = 11
    $headerRange.HorizontalAlignment = -4108
    $headerRange.VerticalAlignment = -4108
    $headerRange.Interior.Color = 0xBDD7EE
    $headerRange.RowHeight = 24

    for ($rowIndex = 0; $rowIndex -lt $Case.Rows.Count; $rowIndex++) {
        $sheetRow = $dataStartRow + $rowIndex
        $rowValues = $Case.Rows[$rowIndex]

        for ($columnIndex = 0; $columnIndex -lt $rowValues.Count; $columnIndex++) {
            Set-CellValue -Worksheet $Worksheet -Row $sheetRow -Column ($columnIndex + 1) -Value $rowValues[$columnIndex]
            $cell = $Worksheet.Cells.Item($sheetRow, $columnIndex + 1)
            $cell.HorizontalAlignment = -4108
            $cell.VerticalAlignment = -4108
            $cell.Font.Name = 'Microsoft YaHei UI'
            $cell.Font.Size = 10
        }

        if ($Case.ContainsKey('Formulas')) {
            foreach ($formula in $Case.Formulas) {
                $Worksheet.Cells.Item($sheetRow, $formula.Column).Formula = [string]::Format($formula.Template, $sheetRow)
            }
        }

        if ($rowIndex % 2 -eq 0) {
            $Worksheet.Range($Worksheet.Cells.Item($sheetRow, 1), $Worksheet.Cells.Item($sheetRow, $headerCount)).Interior.Color = 0xF9FBFD
        }

        $Worksheet.Rows.Item($sheetRow).RowHeight = 22
    }

    $lastDataRow = $dataStartRow + $Case.Rows.Count - 1
    $lastRow = $lastDataRow
    if ($Case.ContainsKey('TotalLabel')) {
        $lastRow = $lastDataRow + 1
        Set-CellValue -Worksheet $Worksheet -Row $lastRow -Column 1 -Value $Case.TotalLabel
        $Worksheet.Cells.Item($lastRow, $Case.TotalColumn).Formula = $Case.TotalFormula
        $totalRange = $Worksheet.Range($Worksheet.Cells.Item($lastRow, 1), $Worksheet.Cells.Item($lastRow, $headerCount))
        $totalRange.Font.Bold = $true
        $totalRange.Font.Name = 'Microsoft YaHei UI'
        $totalRange.Interior.Color = 0xE2F0D9
    }

    if ($Case.ContainsKey('NumberFormats')) {
        foreach ($entry in $Case.NumberFormats.GetEnumerator()) {
            $Worksheet.Range($entry.Key).NumberFormatLocal = $entry.Value
        }
    }

    if ($Case.ContainsKey('HighlightChecks') -and $Case.HighlightChecks) {
        for ($row = $dataStartRow; $row -le $lastDataRow; $row++) {
            for ($column = $Case.HighlightRangeStartColumn; $column -le $headerCount; $column++) {
                $cell = $Worksheet.Cells.Item($row, $column)
                if ([string]$cell.Value2 -eq '√') {
                    $cell.Font.Bold = $true
                    $cell.Font.Color = 0x2E7D32
                    $cell.Interior.Color = 0xE2F0D9
                } else {
                    $cell.Font.Color = 0x7F7F7F
                }
            }
        }
    }

    $usedRange = $Worksheet.Range($Worksheet.Cells.Item($headerRow, 1), $Worksheet.Cells.Item($lastRow, $headerCount))
    $usedRange.Borders.LineStyle = 1
    $usedRange.Borders.Weight = 2
    $usedRange.WrapText = $true

    foreach ($entry in $Case.ColumnWidths.GetEnumerator()) {
        $Worksheet.Columns.Item($entry.Key).ColumnWidth = $entry.Value
    }

    $Worksheet.Range($Worksheet.Cells.Item($headerRow, 1), $Worksheet.Cells.Item($lastRow, $headerCount)).HorizontalAlignment = -4108
    $Worksheet.Range($Worksheet.Cells.Item($headerRow, 1), $Worksheet.Cells.Item($lastRow, $headerCount)).VerticalAlignment = -4108

    $Worksheet.Activate() | Out-Null
    $Worksheet.Application.ActiveWindow.SplitRow = $headerRow
    $Worksheet.Application.ActiveWindow.FreezePanes = $true
    $Worksheet.Range('A3').Select() | Out-Null

    return $lastRow
}

$case = Get-CaseDefinition -Type $CaseType
if (-not $case) {
    throw "Unsupported case type: $CaseType"
}

$app = Get-OrStartSpreadsheetApp
$app.Visible = $true
$workbook = $app.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
[void](Write-CaseWorksheet -Worksheet $worksheet -Case $case)

Write-Output ("已在 WPS 表格中创建案例：{0}" -f $case.Title)
