<#
.SYNOPSIS
    Excel装飾自動化ツール - PowerShell COMオブジェクト デモンストレーション

.DESCRIPTION
    既存のExcelファイルに対して、セルの装飾（背景色、網掛け、斜線等）を
    会話形式で適用するデモンストレーションツールです。

.NOTES
    Author: PowerShell COMオブジェクト デモ
    Version: 1.0
    PowerShell: 5.1以上対応
#>

# エラー時にスクリプトを停止
$ErrorActionPreference = "Stop"

# メイン処理開始
Write-Host "=== Excel装飾自動化ツール ===" -ForegroundColor Green
Write-Host "既存のExcelファイルにセル装飾を適用します。`n"

try {
    # 1. Excelファイルパスの入力
    Write-Host "1. Excelファイルを指定してください"
    $filePath = Read-Host "Excelファイルのパス"
    
    # ファイル存在チェック
    if (-not (Test-Path $filePath)) {
        throw "指定されたファイルが存在しません: $filePath"
    }
    
    # 拡張子チェック
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
    if ($extension -notmatch "\.(xlsx|xls)$") {
        throw "Excelファイル(.xlsx または .xls)を指定してください"
    }
    
    Write-Host "✓ ファイルを確認しました" -ForegroundColor Green

    # 2. Excelアプリケーションを起動
    Write-Host "`n2. Excelを起動しています..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    
    # 3. ワークブックを開く
    $workbook = $excel.Workbooks.Open($filePath)
    Write-Host "✓ ファイルを開きました" -ForegroundColor Green
    
    # 4. シート名の入力
    Write-Host "`n3. 利用可能なシート一覧:"
    for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
        Write-Host "  - $($workbook.Worksheets.Item($i).Name)"
    }
    
    $sheetName = Read-Host "`n操作するシート名"
    
    # シート存在チェック
    $worksheet = $null
    try {
        $worksheet = $workbook.Worksheets.Item($sheetName)
    }
    catch {
        throw "指定されたシート '$sheetName' が見つかりません"
    }
    
    Write-Host "✓ シート '$sheetName' を選択しました" -ForegroundColor Green

    # 5. 装飾機能の選択
    Write-Host "`n4. 適用する装飾を選択してください"
    Write-Host "1. 背景色変更"
    Write-Host "2. 網掛けパターン"
    Write-Host "3. 斜線"
    Write-Host "4. 罫線"
    Write-Host "5. フォント書式"
    
    $functionChoice = Read-Host "選択 (1-5)"
    
    # 入力値チェック
    if ($functionChoice -notmatch "^[1-5]$") {
        throw "1から5の数字を入力してください"
    }

    # 6. セル範囲の入力
    Write-Host "`n5. セル範囲を指定してください"
    Write-Host "例: A1:C5, A1, B2:D10"
    $cellRange = Read-Host "セル範囲"
    
    # 簡易的な範囲チェック
    if ($cellRange -notmatch "^[A-Z]+\d+(:[A-Z]+\d+)?$") {
        throw "正しいセル範囲形式で入力してください (例: A1:C5)"
    }

    # 7. 設定内容の確認
    Write-Host "`n=== 設定内容確認 ===" -ForegroundColor Yellow
    Write-Host "ファイル: $filePath"
    Write-Host "シート: $sheetName"
    Write-Host "範囲: $cellRange"
    
    $functionName = switch ($functionChoice) {
        "1" { "背景色変更" }
        "2" { "網掛けパターン" }
        "3" { "斜線" }
        "4" { "罫線" }
        "5" { "フォント書式" }
    }
    Write-Host "装飾: $functionName"
    
    # 8. 実行確認
    $confirmation = Read-Host "`n実行しますか？ (y/n)"
    if ($confirmation.ToLower() -ne "y") {
        Write-Host "処理をキャンセルしました。" -ForegroundColor Yellow
        return
    }

    # 9. 装飾の適用
    Write-Host "`n装飾を適用しています..."
    $range = $worksheet.Range($cellRange)
    
    switch ($functionChoice) {
        "1" {
            # 背景色変更（薄い黄色）
            $range.Interior.Color = 0xFFFFCC
            Write-Host "✓ 背景色を薄い黄色に変更しました" -ForegroundColor Green
        }
        "2" {
            # 網掛けパターン（斜線パターン）
            $range.Interior.Pattern = 17  # 斜線パターン
            $range.Interior.PatternColor = 0x0000FF  # 青色パターン
            $range.Interior.Color = 0xFFFFFF  # 白背景
            Write-Host "✓ 網掛けパターンを適用しました" -ForegroundColor Green
        }
        "3" {
            # 斜線（両方向）
            $range.Borders.Item(5).LineStyle = 1  # 右上がり斜線
            $range.Borders.Item(5).Color = 0x000000
            $range.Borders.Item(6).LineStyle = 1  # 右下がり斜線
            $range.Borders.Item(6).Color = 0x000000
            Write-Host "✓ 斜線を追加しました" -ForegroundColor Green
        }
        "4" {
            # 罫線（外枠）
            $range.Borders.Item(7).LineStyle = 1   # 左
            $range.Borders.Item(8).LineStyle = 1   # 上
            $range.Borders.Item(9).LineStyle = 1   # 下
            $range.Borders.Item(10).LineStyle = 1  # 右
            $range.Borders.Item(7).Weight = 3      # 太線
            $range.Borders.Item(8).Weight = 3
            $range.Borders.Item(9).Weight = 3
            $range.Borders.Item(10).Weight = 3
            Write-Host "✓ 罫線を追加しました" -ForegroundColor Green
        }
        "5" {
            # フォント書式（太字、赤色）
            $range.Font.Bold = $true
            $range.Font.Color = 0x0000FF  # 青色
            $range.Font.Size = 12
            Write-Host "✓ フォント書式を変更しました（太字・青色）" -ForegroundColor Green
        }
    }

    # 10. 完了メッセージ
    Write-Host "`n=== 処理完了 ===" -ForegroundColor Green
    Write-Host "装飾の適用が完了しました。"
    Write-Host "Excelファイルを確認してください。"
    
    # 保存確認
    $saveConfirmation = Read-Host "`nファイルを保存しますか？ (y/n)"
    if ($saveConfirmation.ToLower() -eq "y") {
        $workbook.Save()
        Write-Host "✓ ファイルを保存しました" -ForegroundColor Green
    }

}
catch {
    # エラーハンドリング
    Write-Host "`nエラーが発生しました:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "`nプログラムを終了します。" -ForegroundColor Red
}
finally {
    # リソースの解放
    if ($workbook) {
        try { $workbook.Close() } catch { }
    }
    if ($excel) {
        try { 
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch { }
    }
    
    # ガベージコレクション実行
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
