# powershell_tool_report
PowerShellScriptの、ちょっと便利なツールやレポート集  
  
## Excel-COM-HolidayProcessor.ps1
外部のエクセルファイルに登録してい休祭日データと土日データを合わせて出力するコード  
休祭日のデータ形式
```
A列    B列    C列    D列    E列
1行 2023   2024   2025   2026   2027  ← 各列の年
2行 祝日1  祝日1  祝日1  祝日1  祝日1  ← 各年の祝祭日データ
3行 祝日2  祝日2  祝日2  祝日2  祝日2
```

##clear_sharepoint_webdav_cache.ps1
edgeのIE互換モードのクッキー一時ファイル削除ツール
SharePointOnlineの「エクスプローラーで開く」機能AD認証のトークンの期限があり、期間を超えた時に使用するIEの互換モードのクッキーと一時ファイルを削除ツール
経緯は以下に
https://qiita.com/ya-man-kys/items/9ba1e5d039cbc431fb41

## make_view_markdown.ps1
HTTPListenerを使ったmarkdownエディタ/ビューワー
コンプライアンス上厳しい制限が掛けられていてもmaekdownは便利に使いたいため作成

## powershell-excel-com-demo.ps1
Excel装飾自動化ツール - PowerShell COMオブジェクト デモンストレーション  
既存のExcelファイルに対して、セルの装飾（背景色、網掛け、斜線等）を会話形式で適用するデモンストレーションツールです。

## set_file_timestamps.ps1
よくあるファイルのタイムスタンプ変更ツールです。PowerShellしか使えない環境でどうぞ。
