
# powershell_tool_report
PowerShellScriptの、ちょっと便利なツールやレポート集  
<hr></hr>

#### File-Clipboard-Manager

複数のファイルをグループ管理し、クリップボードに一括登録できるPowerShellツールです。
主にClaudeなどの生成AIのチャットでファイルをアップロードする際のサポートになります

<hr></hr>

### Get-PathTree
Windows/WSL2両対応のフォルダツリー生成ツール - GUIで簡単操作、コメント記号付きツリー

以下のようなフォルダー（ディレクトリ）のツリーテキストを作成するツールです。  
AIプログラミングなどで、チャットで受け答えする際に必要なので作りました。
```
MyProject/                          //
├── docs/                           //
│   ├── manual.pdf                  //
│   └── readme.txt                  //
├── src/                            //
│   ├── main.js                     //
│   └── style.css                   //
├── tests/                          //
│   └── test_main.js                //
├── .gitignore                      //
├── package.json                    //
└── README.md                       //
```
<hr></hr>

### PowerShell_Clipboard_Manager
外部ツールを導入できない環境でのcliborライクなクリップボード管理ツール
<hr></hr>

### Excel-COM-HolidayProcessor.ps1
外部のエクセルファイルに登録してい休祭日データと土日データを合わせて出力するコード  
休祭日のデータ形式
```
A列    B列    C列    D列    E列
1行 2023   2024   2025   2026   2027  ← 各列の年
2行 祝日1  祝日1  祝日1  祝日1  祝日1  ← 各年の祝祭日データ
3行 祝日2  祝日2  祝日2  祝日2  祝日2
```
<hr></hr>

### clear_sharepoint_webdav_cache.ps1
edgeのIE互換モードのクッキー一時ファイル削除ツール
SharePointOnlineの「エクスプローラーで開く」機能AD認証のトークンの期限があり、期間を超えた時に使用するIEの互換モードのクッキーと一時ファイルを削除ツール
経緯は以下に
https://qiita.com/ya-man-kys/items/9ba1e5d039cbc431fb41
<hr></hr>

### make_view_markdown.ps1
HTTPListenerを使ったmarkdownエディタ/ビューワー
コンプライアンス上厳しい制限が掛けられていてもmaekdownは便利に使いたいため作成
<hr></hr>

### powershell-excel-com-demo.ps1
Excel装飾自動化ツール - PowerShell COMオブジェクト デモンストレーション  
既存のExcelファイルに対して、セルの装飾（背景色、網掛け、斜線等）を会話形式で適用するデモンストレーションツールです。
<hr></hr>

### set_file_timestamps.ps1
よくあるファイルのタイムスタンプ変更ツールです。PowerShellしか使えない環境でどうぞ。
