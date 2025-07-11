# 休祭日データ収集と加工スクリプト
    $holidays = @()
    $lastRow = $Worksheet.Cells.Item($Worksheet.Rows# 休祭日データ統合処理 - PowerShell COMデモ
param(
    [string]$ExcelFile = "holiday_data.xlsx",
    [int]$Year = 0
)

# Excel COM オブジェクト初期化
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false

# 年の入力処理
if ($Year -eq 0) {
    $Year = Read-Host "処理する年を入力してください (例: 2025)"
    $Year = [int]$Year
}

Write-Host "処理対象年: $Year" -ForegroundColor Green

try {
    # ワークブック・シート取得
    $Workbook = $Excel.Workbooks.Open((Resolve-Path $ExcelFile).Path)
    $Worksheet = $Workbook.Worksheets.Item(1)
    
    # 指定年の列を検索
    $TargetColumn = 0
    $maxCol = 10  # 最大10列まで検索
    
    for ($col = 1; $col -le $maxCol; $col++) {
        $yearValue = $Worksheet.Cells.Item(1, $col).Value2
        if ($yearValue -eq $Year) {
            $TargetColumn = $col
            break
        }
    }
    
    if ($TargetColumn -eq 0) {
        Write-Host "年 $Year のデータが見つかりません" -ForegroundColor Red
        return
    }
    
    Write-Host "列 $TargetColumn で年 $Year のデータを処理中..." -ForegroundColor Yellow
    
    # 休祭日データ収集（指定年の列から）
    $holidays = @()
    $lastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, $TargetColumn).End(-4162).Row  # xlUp = -4162
    
    # 2行目以降がデータ（1行目は年）
    for ($i = 2; $i -le $lastRow; $i++) {
        $cellValue = $Worksheet.Cells.Item($i, $TargetColumn).Value2
        if ($cellValue -and $cellValue -match "^\d+$") {  # 数値（日付シリアル値）のみ処理
            # Excelシリアル値をDateTime変換
            $date = [DateTime]::FromOADate($cellValue)
            if ($date.Year -eq $Year) {  # 指定年のデータのみ
                $formatted = $date.ToString("yyyy/MM/dd")
                $holidays += $formatted
            }
        }
    }
    
    # 指定年の土日データ生成（全月）
    function Get-WeekendDates {
        param([int]$Year)
        
        $weekends = @()
        for ($month = 1; $month -le 12; $month++) {
            $daysInMonth = [DateTime]::DaysInMonth($Year, $month)
            
            for ($day = 1; $day -le $daysInMonth; $day++) {
                $date = Get-Date -Year $Year -Month $month -Day $day
                if ($date.DayOfWeek -in @('Saturday', 'Sunday')) {
                    $weekends += $date.ToString("yyyy/MM/dd")
                }
            }
        }
        return $weekends
    }
    
    # 休日データ統合（指定年の全年間）
    $weekends = Get-WeekendDates $Year
    $allHolidays = $holidays + $weekends | Sort-Object | Get-Unique
    
    # 結果出力
    Write-Host "`n=== $Year 年 統合休日データ ===" -ForegroundColor Green
    Write-Host "祝祭日: $($holidays.Count) 日, 土日: $($weekends.Count) 日, 合計: $($allHolidays.Count) 日" -ForegroundColor Cyan
    
    $allHolidays | ForEach-Object { 
        $date = [DateTime]::Parse($_)
        $dayType = if ($holidays -contains $_) { "祝祭日" } else { "土日" }
        Write-Host "$_ ($($date.ToString('ddd'))) - $dayType"
    }
    
} finally {
    # COMオブジェクト解放
    $Workbook.Close($false)
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

Write-Host "`n処理完了" -ForegroundColor Cyan
