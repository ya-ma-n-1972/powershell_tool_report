# PowerShell Markdown Editor + Viewer
# 
# 使用方法：
# 1. PowerShellでこのスクリプトを実行: .\markdown_editor.ps1
# 2. ブラウザで http://localhost:8080 にアクセス
# 3. Ctrl+Cでサーバーを停止
#
# 機能：
# - Markdownエディタ（左右分割、リアルタイムプレビュー、保存機能）
# - Markdownビューア（ファイル名指定で表示）
# - PowerShellScriptファイルと同じディレクトリにMarkdownファイルを置いてください。保存先も同じディレクトリになります。

# HTMLコンテンツ
$html_content = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Markdown Editor + Viewer</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/github-markdown-css/5.8.1/github-markdown.min.css">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background-color: #f8f9fa;
            height: 100vh;
            overflow: hidden;
        }
        
        .header {
            background: #ffffff;
            border-bottom: 1px solid #e1e4e8;
            padding: 10px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .tabs {
            display: flex;
            gap: 10px;
        }
        
        .tab {
            padding: 8px 16px;
            border: none;
            background: #f1f3f4;
            cursor: pointer;
            border-radius: 4px;
            font-size: 14px;
        }
        
        .tab.active {
            background: #007bff;
            color: white;
        }
        
        .save-btn {
            padding: 8px 16px;
            background: #28a745;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }
        
        .save-btn:hover {
            background: #218838;
        }
        
        .content {
            height: calc(100vh - 60px);
            display: flex;
        }
        
        .editor-mode {
            display: flex;
            width: 100%;
        }
        
        .editor-pane {
            width: 50%;
            border-right: 1px solid #e1e4e8;
        }
        
        .preview-pane {
            width: 50%;
            background: white;
            overflow-y: auto;
        }
        
        .editor-textarea {
            width: 100%;
            height: 100%;
            border: none;
            outline: none;
            padding: 20px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            line-height: 1.6;
            resize: none;
            background: white;
        }
        
        .markdown-body {
            padding: 20px;
            min-height: 100%;
        }
        
        .viewer-mode {
            display: none;
            flex-direction: column;
            width: 100%;
        }
        
        .viewer-controls {
            padding: 20px;
            background: white;
            border-bottom: 1px solid #e1e4e8;
        }
        
        .viewer-controls input {
            padding: 8px 12px;
            border: 1px solid #d1d5da;
            border-radius: 4px;
            margin-right: 10px;
            font-size: 14px;
        }
        
        .viewer-controls button {
            padding: 8px 16px;
            background: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }
        
        .viewer-content {
            flex: 1;
            overflow-y: auto;
            background: white;
        }
        
        .status {
            padding: 10px 20px;
            background: #d4edda;
            color: #155724;
            border-bottom: 1px solid #c3e6cb;
            display: none;
        }
        
        .status.error {
            background: #f8d7da;
            color: #721c24;
            border-color: #f5c6cb;
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="tabs">
            <button class="tab active" onclick="switchMode('editor')">エディタ</button>
            <button class="tab" onclick="switchMode('viewer')">ビューア</button>
        </div>
        <button class="save-btn" onclick="saveFile()" id="saveBtn">保存</button>
    </div>
    
    <div class="status" id="status"></div>
    
    <div class="content">
        <div class="editor-mode" id="editorMode">
            <div class="editor-pane">
                <textarea class="editor-textarea" id="editor" placeholder="Markdownを入力してください..."></textarea>
            </div>
            <div class="preview-pane">
                <div class="markdown-body" id="preview">プレビューがここに表示されます</div>
            </div>
        </div>
        
        <div class="viewer-mode" id="viewerMode">
            <div class="viewer-controls">
                <input type="text" id="filename" placeholder="ファイル名を入力 (例: test.md)" />
                <button onclick="loadFile()">読み込み</button>
            </div>
            <div class="viewer-content">
                <div class="markdown-body" id="viewerContent">ファイルを選択してください</div>
            </div>
        </div>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/marked/4.3.0/marked.min.js"></script>
    <script>
        let currentMode = 'editor';
        let currentFilename = 'new.md';
        
        // モード切り替え
        function switchMode(mode) {
            currentMode = mode;
            
            // タブの切り替え
            document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
            event.target.classList.add('active');
            
            // コンテンツの切り替え
            if (mode === 'editor') {
                document.getElementById('editorMode').style.display = 'flex';
                document.getElementById('viewerMode').style.display = 'none';
                document.getElementById('saveBtn').style.display = 'block';
            } else {
                document.getElementById('editorMode').style.display = 'none';
                document.getElementById('viewerMode').style.display = 'flex';
                document.getElementById('saveBtn').style.display = 'none';
            }
        }
        
        // リアルタイムプレビュー
        function updatePreview() {
            const markdown = document.getElementById('editor').value;
            const html = marked.parse(markdown);
            document.getElementById('preview').innerHTML = html;
        }
        
        // ファイル保存
        async function saveFile() {
            const content = document.getElementById('editor').value;
            const filename = prompt('ファイル名を入力してください:', currentFilename);
            
            if (!filename) return;
            
            try {
                const response = await fetch('/save', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        filename: filename,
                        content: content
                    })
                });
                
                if (response.ok) {
                    currentFilename = filename;
                    showStatus('ファイルを保存しました: ' + filename, 'success');
                } else {
                    showStatus('保存に失敗しました', 'error');
                }
            } catch (error) {
                showStatus('保存エラー: ' + error.message, 'error');
            }
        }
        
        // ファイル読み込み
        async function loadFile() {
            const filename = document.getElementById('filename').value;
            
            if (!filename) {
                showStatus('ファイル名を入力してください', 'error');
                return;
            }
            
            try {
                const response = await fetch('/markdown/' + encodeURIComponent(filename));
                
                if (response.ok) {
                    const markdown = await response.text();
                    const html = marked.parse(markdown);
                    document.getElementById('viewerContent').innerHTML = html;
                    showStatus('ファイルを読み込みました: ' + filename, 'success');
                } else {
                    showStatus('ファイルが見つかりません: ' + filename, 'error');
                }
            } catch (error) {
                showStatus('読み込みエラー: ' + error.message, 'error');
            }
        }
        
        // エディタモードでファイル読み込み
        async function loadToEditor(filename) {
            try {
                const response = await fetch('/markdown/' + encodeURIComponent(filename));
                
                if (response.ok) {
                    const markdown = await response.text();
                    document.getElementById('editor').value = markdown;
                    currentFilename = filename;
                    updatePreview();
                    showStatus('ファイルを読み込みました: ' + filename, 'success');
                } else {
                    showStatus('ファイルが見つかりません: ' + filename, 'error');
                }
            } catch (error) {
                showStatus('読み込みエラー: ' + error.message, 'error');
            }
        }
        
        // ステータス表示
        function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = 'status' + (type === 'error' ? ' error' : '');
            status.style.display = 'block';
            
            setTimeout(() => {
                status.style.display = 'none';
            }, 3000);
        }
        
        // イベントリスナー
        document.getElementById('editor').addEventListener('input', updatePreview);
        document.getElementById('filename').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                loadFile();
            }
        });
        
        // 初期化
        document.addEventListener('DOMContentLoaded', function() {
            updatePreview();
            
            // Ctrl+Sで保存
            document.addEventListener('keydown', function(e) {
                if (e.ctrlKey && e.key === 's') {
                    e.preventDefault();
                    if (currentMode === 'editor') {
                        saveFile();
                    }
                }
            });
        });
    </script>
</body>
</html>
"@

try {
    # HTTPListenerの作成
    $http = [System.Net.HttpListener]::new()
    $http.Prefixes.Add("http://localhost:8080/")
    
    # リスナーの開始
    Write-Host "Markdown Editor + Viewer を起動中..."
    $http.Start()
    Write-Host "サーバー起動完了: http://localhost:8080/"
    Write-Host "ブラウザでアクセスしてください"
    Write-Host "停止するにはCtrl+Cを押してください"
    Write-Host ""
    
    while ($http.IsListening) {
        # リクエストの待機
        $context = $http.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        # レスポンスの作成
        try {
            if ($request.Url.AbsolutePath -eq "/") {
                # メインページ（HTML）
                $content = [System.Text.Encoding]::UTF8.GetBytes($html_content)
                $response.ContentType = "text/html; charset=utf-8"
                $response.OutputStream.Write($content, 0, $content.Length)
                Write-Host "メインページを配信しました"
                
            } elseif ($request.Url.AbsolutePath -match "^/markdown/(.+)$") {
                # 指定されたMarkdownファイルの内容を配信
                $filename = [System.Web.HttpUtility]::UrlDecode($matches[1])
                $script_dir = Split-Path -Parent $MyInvocation.MyCommand.Path
                $markdown_path = Join-Path $script_dir $filename
                
                if (Test-Path $markdown_path) {
                    $markdown_content = Get-Content -Path $markdown_path -Encoding UTF8 -Raw
                    $content = [System.Text.Encoding]::UTF8.GetBytes($markdown_content)
                    $response.ContentType = "text/plain; charset=utf-8"
                    $response.OutputStream.Write($content, 0, $content.Length)
                    Write-Host "Markdownファイルを配信しました: $filename"
                } else {
                    $error_message = "ファイルが見つかりません: $filename"
                    $content = [System.Text.Encoding]::UTF8.GetBytes($error_message)
                    $response.StatusCode = 404
                    $response.ContentType = "text/plain; charset=utf-8"
                    $response.OutputStream.Write($content, 0, $content.Length)
                    Write-Host "エラー: ファイルが見つかりません - $markdown_path"
                }
                
            } elseif ($request.Url.AbsolutePath -eq "/save" -and $request.HttpMethod -eq "POST") {
                # ファイル保存
                $reader = [System.IO.StreamReader]::new($request.InputStream)
                $json_data = $reader.ReadToEnd()
                $reader.Close()
                
                $data = ConvertFrom-Json $json_data
                $filename = $data.filename
                $file_content = $data.content
                
                $script_dir = Split-Path -Parent $MyInvocation.MyCommand.Path
                $file_path = Join-Path $script_dir $filename
                
                # UTF-8で保存
                [System.IO.File]::WriteAllText($file_path, $file_content, [System.Text.Encoding]::UTF8)
                
                $success_message = "ファイルを保存しました: $filename"
                $content = [System.Text.Encoding]::UTF8.GetBytes($success_message)
                $response.ContentType = "text/plain; charset=utf-8"
                $response.OutputStream.Write($content, 0, $content.Length)
                Write-Host "ファイルを保存しました: $filename"
                
            } else {
                # 存在しないパス
                $error_message = "404 Not Found"
                $content = [System.Text.Encoding]::UTF8.GetBytes($error_message)
                $response.StatusCode = 404
                $response.ContentType = "text/plain; charset=utf-8"
                $response.OutputStream.Write($content, 0, $content.Length)
                Write-Host "404エラー: $($request.Url.AbsolutePath)"
            }
        }
        catch {
            # リクエスト処理中のエラー
            Write-Host "リクエスト処理エラー: $_"
            $error_message = "サーバーエラーが発生しました"
            $content = [System.Text.Encoding]::UTF8.GetBytes($error_message)
            $response.StatusCode = 500
            $response.ContentType = "text/plain; charset=utf-8"
            $response.OutputStream.Write($content, 0, $content.Length)
        }
        
        # レスポンスの終了
        $response.Close()
    }
}
catch {
    Write-Host "サーバーエラー: $_"
    Write-Host "ポート8080が既に使用されている可能性があります"
}
finally {
    # クリーンアップ
    if ($null -ne $http) {
        Write-Host "`nサーバーを停止しています..."
        $http.Stop()
        $http.Close()
        Write-Host "サーバーを停止しました"
    }
}
