# Edge IE互換モードのクッキー削除スクリプト
# SharePoint Online WebDAV接続用

Write-Host "Edge IE互換モードのクッキー削除を開始します..." -ForegroundColor Yellow

try {
    # RunDll32を使用してIE互換モードのクッキーとキャッシュを削除
    # パラメータ説明:
    # 1 = 履歴
    # 2 = クッキー
    # 4 = 一時ファイル
    # 8 = フォームデータ
    # 16 = パスワード
    # 32 = すべて
    
    Write-Host "クッキーを削除中..." -ForegroundColor Cyan
    Start-Process "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess 2" -Wait -WindowStyle Hidden
    
    Write-Host "一時ファイルを削除中..." -ForegroundColor Cyan
    Start-Process "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess 4" -Wait -WindowStyle Hidden
    
    Write-Host "" # 空行
    Write-Host "✓ Edge IE互換モードのクッキーとキャッシュを正常に削除しました" -ForegroundColor Green
    Write-Host "✓ SharePoint OnlineのWebDAV接続を再試行できます" -ForegroundColor Green
    
    # 削除完了の時刻を表示
    $currentTime = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    Write-Host "削除完了時刻: $currentTime" -ForegroundColor Gray
}
catch {
    Write-Host "" # 空行
    Write-Host "✗ エラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "手動でEdgeの設定からクッキーを削除してください" -ForegroundColor Yellow
}

Write-Host "" # 空行
Write-Host "スクリプト実行完了。何かキーを押して終了してください..." -ForegroundColor White
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
