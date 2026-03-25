<#
.SYNOPSIS
    File Clipboard Manager - プロジェクトフォルダのファイルをクリップボードに登録するツール

.NOTES
    Author: ya-man
    Version: 2.0
    Requires: PowerShell 5.1 or later, Windows Forms, Python + pathspec
#>

# ================================================================================
# アセンブリの読み込み
# ================================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ================================================================================
# グローバル変数
# ================================================================================
$script:currentFolder      = ""
$script:currentFiles       = @()
$script:folderHistory      = @()
$script:gitignoreEnabled   = $true
$script:settingsPath       = Join-Path $PSScriptRoot "settings.json"
$script:filterScriptPath   = Join-Path $PSScriptRoot "filter.py"
$script:messageLabel       = $null
$script:notifyIcon         = $null

# ================================================================================
# 設定ファイル操作
# ================================================================================

function Load-Settings {
    # settings.json が存在しない場合はデフォルト値で自動生成
    if (-not (Test-Path $script:settingsPath)) {
        $defaultSettings = @{
            gitignore_enabled  = $true
            folder_history     = @()
            exclude_extensions = @(
                ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".ico", ".webp",
                ".tiff", ".tif", ".psd", ".ai", ".eps", ".raw", ".heic", ".avif", ".icns",
                ".mp4", ".mp3", ".wav", ".flac", ".aac", ".ogg", ".avi", ".mov",
                ".mkv", ".wmv", ".m4a", ".m4v", ".webm", ".wma", ".m4p", ".3gp",
                ".exe", ".dll", ".so", ".dylib", ".bin", ".obj", ".class",
                ".apk", ".aab", ".ipa", ".jar", ".war", ".ear", ".jmod",
                ".zip", ".tar", ".gz", ".7z", ".rar", ".bz2", ".xz",
                ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
                ".db", ".sqlite", ".sqlite3",
                ".o", ".a", ".lib", ".pyc", ".pyo"
            )
            exclude_filenames  = @(
                "Cargo.lock", "package-lock.json", "yarn.lock", "pnpm-lock.yaml",
                "poetry.lock", "Pipfile.lock", "Gemfile.lock", "composer.lock",
                "go.sum", "packages.lock.json", "Package.resolved", "pubspec.lock"
            )
        }
        $defaultSettings | ConvertTo-Json -Depth 10 | Set-Content $script:settingsPath -Encoding UTF8
    }

    if (Test-Path $script:settingsPath) {
        try {
            $json = Get-Content $script:settingsPath -Raw -Encoding UTF8
            $settings = $json | ConvertFrom-Json
            $script:gitignoreEnabled = $settings.gitignore_enabled

            $needsSave = $false

            # マイグレーション: last_folder → folder_history
            if ($null -eq $settings.folder_history -and $null -ne $settings.last_folder) {
                $settings | Add-Member -NotePropertyName "folder_history" -NotePropertyValue @($settings.last_folder)
                $settings.PSObject.Properties.Remove('last_folder')
                $needsSave = $true
            }

            # folder_history の読み込み
            if ($null -ne $settings.folder_history) {
                $script:folderHistory = @($settings.folder_history)
                if ($script:folderHistory.Count -gt 0) {
                    $script:currentFolder = $script:folderHistory[0]
                }
            }

            # 除外ルールのキーが必須
            if ($null -eq $settings.exclude_extensions -or $null -eq $settings.exclude_filenames) {
                [System.Windows.Forms.MessageBox]::Show(
                    "settings.json に exclude_extensions または exclude_filenames がありません。`nsettings.json を確認してください。",
                    "設定エラー",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                exit
            }

            if ($needsSave) {
                $settings | ConvertTo-Json -Depth 10 | Set-Content $script:settingsPath -Encoding UTF8
            }
        }
        catch {
            $script:currentFolder    = ""
            $script:gitignoreEnabled = $true
        }
    }
}

function Save-Settings {
    try {
        if (Test-Path $script:settingsPath) {
            $json = Get-Content $script:settingsPath -Raw -Encoding UTF8
            $settings = $json | ConvertFrom-Json
            $settings.folder_history = $script:folderHistory
            $settings.gitignore_enabled = $script:gitignoreEnabled

            # マイグレーション済み確認: last_folder が残っていたら削除
            if ($null -ne $settings.last_folder) {
                $settings.PSObject.Properties.Remove('last_folder')
            }
        }
        else {
            $settings = @{
                folder_history     = $script:folderHistory
                gitignore_enabled  = $script:gitignoreEnabled
                exclude_extensions = @(
                    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".ico", ".webp",
                    ".tiff", ".tif", ".psd", ".ai", ".eps", ".raw", ".heic", ".avif", ".icns",
                    ".mp4", ".mp3", ".wav", ".flac", ".aac", ".ogg", ".avi", ".mov",
                    ".mkv", ".wmv", ".m4a", ".m4v", ".webm", ".wma", ".m4p", ".3gp",
                    ".exe", ".dll", ".so", ".dylib", ".bin", ".obj", ".class",
                    ".apk", ".aab", ".ipa", ".jar", ".war", ".ear", ".jmod",
                    ".zip", ".tar", ".gz", ".7z", ".rar", ".bz2", ".xz",
                    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
                    ".db", ".sqlite", ".sqlite3",
                    ".o", ".a", ".lib", ".pyc", ".pyo"
                )
                exclude_filenames  = @(
                    "Cargo.lock", "package-lock.json", "yarn.lock", "pnpm-lock.yaml",
                    "poetry.lock", "Pipfile.lock", "Gemfile.lock", "composer.lock",
                    "go.sum", "packages.lock.json", "Package.resolved", "pubspec.lock"
                )
            }
        }
        $settings | ConvertTo-Json -Depth 10 | Set-Content $script:settingsPath -Encoding UTF8
    }
    catch {
        Show-Message "設定の保存に失敗しました: $_"
    }
}

function Add-FolderHistory($folderPath) {
    # 既に履歴にある場合は一度削除して先頭に移動
    $script:folderHistory = @($folderPath) + @($script:folderHistory | Where-Object { $_ -ne $folderPath })

    # 上限10件に切り詰め
    if ($script:folderHistory.Count -gt 10) {
        $script:folderHistory = $script:folderHistory[0..9]
    }
}

# ================================================================================
# 起動時チェック
# ================================================================================

function Test-Requirements {
    # Pythonチェック
    try {
        $null = & python --version 2>&1
        if ($LASTEXITCODE -ne 0) { throw }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Pythonが見つかりません。`nPythonをインストールしてください。",
            "起動エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }

    # pathspecチェック
    try {
        $null = & python -c "import pathspec" 2>&1
        if ($LASTEXITCODE -ne 0) { throw }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "pathspecが見つかりません。`n以下のコマンドを実行してください:`n`npip install pathspec",
            "起動エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }

    return $true
}

# ================================================================================
# フィルタリング
# ================================================================================

function Invoke-Filter {
    if ([string]::IsNullOrEmpty($script:currentFolder)) {
        return @()
    }

    if (-not (Test-Path $script:currentFolder)) {
        Show-Message "フォルダが見つかりません: $script:currentFolder"
        return @()
    }

    # .gitignore チェック
    if ($script:gitignoreEnabled) {
        $gitignorePath = Join-Path $script:currentFolder ".gitignore"
        if (-not (Test-Path $gitignorePath)) {
            Show-Message ".gitignoreが見つかりません。チェックをOFFにするか、.gitignoreを作成してください。"
            return @()
        }
    }

    try {
        $filterArgs = @($script:filterScriptPath, $script:currentFolder)
        if (-not $script:gitignoreEnabled) {
            $filterArgs += "--no-gitignore"
        }

        $output = & python @filterArgs 2>&1
        if ($LASTEXITCODE -ne 0) {
            $errMsg = $output | Where-Object { $_ -match "^ERROR:" }
            Show-Message "フィルタリングエラー: $errMsg"
            return @()
        }

        return @($output | Where-Object { $_ -ne "" })
    }
    catch {
        Show-Message "filter.py の実行に失敗しました: $_"
        return @()
    }
}

# ================================================================================
# ファイルサイズ計算
# ================================================================================

function Get-TotalSize($files) {
    $totalBytes = 0
    foreach ($f in $files) {
        if (Test-Path $f) {
            $totalBytes += (Get-Item $f).Length
        }
    }
    return "{0:N1} MB" -f ($totalBytes / 1MB)
}

# ================================================================================
# クリップボード操作
# ================================================================================

function Set-ClipboardFiles($files) {
    $col = New-Object System.Collections.Specialized.StringCollection
    foreach ($f in $files) {
        if (Test-Path $f) { [void]$col.Add($f) }
    }
    if ($col.Count -eq 0) {
        Show-Message "対象ファイルがありません"
        return
    }
    [System.Windows.Forms.Clipboard]::SetFileDropList($col)
    Show-Message "クリップボードに登録しました（$($col.Count)件）"
}

function Set-ClipboardWithFence($files) {
    if ($files.Count -eq 0) {
        Show-Message "対象ファイルがありません"
        return
    }

    $fence = '```'
    $allText = ""
    $count = 0
    $skipCount = 0

    foreach ($f in $files) {
        if (-not (Test-Path $f)) { continue }
        try {
            $content = Get-Content $f -Raw -Encoding UTF8
            $allText += "ファイルパス: $f`n"
            $allText += $fence + "`n"
            $allText += $content + "`n"
            $allText += $fence + "`n`n"
            $count++
        }
        catch { $skipCount++ }
    }

    if ($allText.Length -eq 0) {
        Show-Message "読み込めるファイルがありません"
        return
    }

    [System.Windows.Forms.Clipboard]::SetText($allText)
    if ($skipCount -gt 0) {
        Show-Message "フェンス付きでクリップボードに登録しました（${count}件、${skipCount}件スキップ）"
    }
    else {
        Show-Message "フェンス付きでクリップボードに登録しました（${count}件）"
    }
}

function Set-ClipboardAddFence {
    $textExtensions = @(
        '.txt','.ps1','.psm1','.psd1','.cs','.vb','.js','.ts','.tsx','.jsx',
        '.html','.htm','.xml','.json','.css','.md','.log','.csv','.ini',
        '.config','.bat','.cmd','.py','.rb','.java','.c','.cpp','.h','.hpp',
        '.php','.sh','.rs','.toml','.yaml','.yml'
    )

    $fileList = [System.Windows.Forms.Clipboard]::GetFileDropList()

    if ($null -ne $fileList -and $fileList.Count -gt 0) {
        if ($fileList.Count -gt 1) {
            [System.Windows.Forms.MessageBox]::Show(
                "複数ファイルには対応していません。1つのファイルをコピーしてください。",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $filePath = $fileList[0]

        if (-not (Test-Path $filePath)) {
            Show-Message "ファイルが見つかりません"
            return
        }

        $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
        if ($textExtensions -notcontains $ext) {
            Show-Message "テキストファイルではありません"
            return
        }

        $fileSize = (Get-Item $filePath).Length
        if ($fileSize -gt 30MB) {
            Show-Message "ファイルサイズが30MBを超えています"
            return
        }

        try {
            $content = Get-Content $filePath -Raw -Encoding UTF8
            $fence = '```'
            $text = "ファイルパス: $filePath`n" + $fence + "`n" + $content + "`n" + $fence
            [System.Windows.Forms.Clipboard]::SetText($text)
            $fileName = [System.IO.Path]::GetFileName($filePath)
            Show-Message "$fileName にフェンスを追加してクリップボードに登録しました"
        }
        catch {
            Show-Message "ファイルの読み込みに失敗しました: $_"
        }
        return
    }

    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $content = [System.Windows.Forms.Clipboard]::GetText()
        if ([string]::IsNullOrWhiteSpace($content)) {
            Show-Message "クリップボードのテキストが空です"
            return
        }
        $fence = '```'
        $text = $fence + "`n" + $content + "`n" + $fence
        [System.Windows.Forms.Clipboard]::SetText($text)
        Show-Message "テキストにフェンスを追加してクリップボードに登録しました"
        return
    }

    Show-Message "クリップボードにファイルまたはテキストがありません"
}

# ================================================================================
# BOM変換
# ================================================================================

function Convert-ToBomUtf8($filePath) {
    $textExtensions = @(
        '.txt','.ps1','.psm1','.psd1','.cs','.vb','.js','.ts','.tsx','.jsx',
        '.html','.htm','.xml','.json','.css','.md','.log','.csv','.ini',
        '.config','.bat','.cmd','.py','.rb','.java','.c','.cpp','.h','.hpp',
        '.php','.sh','.rs','.toml','.yaml','.yml'
    )

    $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
    if ($textExtensions -notcontains $ext) {
        Show-Message "テキストファイルではありません"
        return
    }

    $fileName = [System.IO.Path]::GetFileName($filePath)
    $result = [System.Windows.Forms.MessageBox]::Show(
        "$fileName をBOM付きUTF-8に変換しますか？`n`n※バックアップは作成されません",
        "確認",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $content = Get-Content $filePath -Raw
            $utf8Bom = New-Object System.Text.UTF8Encoding($true)
            [System.IO.File]::WriteAllText($filePath, $content, $utf8Bom)
            Show-Message "$fileName : BOM変換完了"
        }
        catch {
            Show-Message "変換に失敗しました: $_"
        }
    }
}

# ================================================================================
# メッセージ表示
# ================================================================================

function Show-Message($message) {
    if ($null -ne $script:messageLabel) {
        $script:messageLabel.Text = $message
    }
}

# ================================================================================
# ファイル詳細ウィンドウ
# ================================================================================

function Show-FileDetailWindow {
    if ($script:currentFiles.Count -eq 0) {
        Show-Message "ファイルがありません"
        return
    }

    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = "ファイル詳細"
    $detailForm.Size = New-Object System.Drawing.Size(700, 500)
    $detailForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $detailForm.MaximizeBox = $false
    $detailForm.MinimizeBox = $false
    $detailForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # ファイル一覧
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 10)
    $listBox.Size = New-Object System.Drawing.Size(665, 370)
    $listBox.HorizontalScrollbar = $true
    $listBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended

    foreach ($f in $script:currentFiles) {
        [void]$listBox.Items.Add($f)
    }
    $detailForm.Controls.Add($listBox)

    # BOM変換ボタン
    $btnBom = New-Object System.Windows.Forms.Button
    $btnBom.Text = "BOM変換"
    $btnBom.Location = New-Object System.Drawing.Point(10, 395)
    $btnBom.Size = New-Object System.Drawing.Size(100, 30)
    $btnBom.Add_Click({
        if ($listBox.SelectedItems.Count -eq 0) {
            Show-Message "ファイルを選択してください"
            return
        }
        foreach ($f in $listBox.SelectedItems) {
            Convert-ToBomUtf8 $f
        }
    })
    $detailForm.Controls.Add($btnBom)

    # 選択項目をクリップボードへ
    $btnClipSelected = New-Object System.Windows.Forms.Button
    $btnClipSelected.Text = "クリップボードへ"
    $btnClipSelected.Location = New-Object System.Drawing.Point(120, 395)
    $btnClipSelected.Size = New-Object System.Drawing.Size(140, 30)
    $btnClipSelected.Add_Click({
        Set-ClipboardFiles @($listBox.SelectedItems)
    })
    $detailForm.Controls.Add($btnClipSelected)

    # 選択項目をフェンス付きでクリップボードへ
    $btnFenceSelected = New-Object System.Windows.Forms.Button
    $btnFenceSelected.Text = "フェンス付き"
    $btnFenceSelected.Location = New-Object System.Drawing.Point(270, 395)
    $btnFenceSelected.Size = New-Object System.Drawing.Size(140, 30)
    $btnFenceSelected.Add_Click({
        Set-ClipboardWithFence @($listBox.SelectedItems)
    })
    $detailForm.Controls.Add($btnFenceSelected)

    # 閉じるボタン
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "閉じる"
    $btnClose.Location = New-Object System.Drawing.Point(575, 395)
    $btnClose.Size = New-Object System.Drawing.Size(100, 30)
    $btnClose.Add_Click({ $detailForm.Close() })
    $detailForm.Controls.Add($btnClose)

    # メッセージエリア
    $msgLabel = New-Object System.Windows.Forms.Label
    $msgLabel.Location = New-Object System.Drawing.Point(420, 400)
    $msgLabel.Size = New-Object System.Drawing.Size(245, 20)
    $msgLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $detailForm.Controls.Add($msgLabel)
    $script:messageLabel = $msgLabel

    $detailForm.Add_FormClosed({
        $script:messageLabel = $null
    })

    [void]$detailForm.ShowDialog()
}

# ================================================================================
# メインウィンドウ
# ================================================================================

function Show-MainWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "File Clipboard Manager"
    $form.Size = New-Object System.Drawing.Size(500, 280)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # フォルダ選択ボタン
    $btnFolder = New-Object System.Windows.Forms.Button
    $btnFolder.Text = "フォルダ選択"
    $btnFolder.Location = New-Object System.Drawing.Point(10, 15)
    $btnFolder.Size = New-Object System.Drawing.Size(100, 30)
    $form.Controls.Add($btnFolder)

    # フォルダ履歴（ComboBox）
    $cmbFolder = New-Object System.Windows.Forms.ComboBox
    $cmbFolder.Location = New-Object System.Drawing.Point(120, 17)
    $cmbFolder.Size = New-Object System.Drawing.Size(355, 25)
    $cmbFolder.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbFolder.DropDownWidth = 500

    foreach ($folder in $script:folderHistory) {
        [void]$cmbFolder.Items.Add($folder)
    }
    if ($cmbFolder.Items.Count -gt 0) {
        $cmbFolder.SelectedIndex = 0
    }

    $form.Controls.Add($cmbFolder)

    # .gitignore 有効チェックボックス
    $chkGitignore = New-Object System.Windows.Forms.CheckBox
    $chkGitignore.Text = ".gitignore を有効にする"
    $chkGitignore.Location = New-Object System.Drawing.Point(10, 55)
    $chkGitignore.Size = New-Object System.Drawing.Size(200, 25)
    $chkGitignore.Checked = $script:gitignoreEnabled
    $form.Controls.Add($chkGitignore)

    # ---- ボタン群 ----

    # クリップボードへ
    $btnClip = New-Object System.Windows.Forms.Button
    $btnClip.Text = "クリップボードへ"
    $btnClip.Location = New-Object System.Drawing.Point(10, 95)
    $btnClip.Size = New-Object System.Drawing.Size(140, 30)
    $btnClip.Add_Click({
        Set-ClipboardFiles $script:currentFiles
    })
    $form.Controls.Add($btnClip)

    # フェンス付きクリップボード
    $btnFence = New-Object System.Windows.Forms.Button
    $btnFence.Text = "フェンス付きクリップボード"
    $btnFence.Location = New-Object System.Drawing.Point(160, 95)
    $btnFence.Size = New-Object System.Drawing.Size(200, 30)
    $btnFence.Add_Click({
        Set-ClipboardWithFence $script:currentFiles
    })
    $form.Controls.Add($btnFence)

    # クリップボードにフェンス追加
    $btnAddFence = New-Object System.Windows.Forms.Button
    $btnAddFence.Text = "クリップボードにフェンス追加"
    $btnAddFence.Location = New-Object System.Drawing.Point(10, 135)
    $btnAddFence.Size = New-Object System.Drawing.Size(200, 30)
    $btnAddFence.Add_Click({
        Set-ClipboardAddFence
    })
    $form.Controls.Add($btnAddFence)

    # ファイル詳細
    $btnDetail = New-Object System.Windows.Forms.Button
    $btnDetail.Text = "ファイル詳細"
    $btnDetail.Location = New-Object System.Drawing.Point(220, 135)
    $btnDetail.Size = New-Object System.Drawing.Size(140, 30)
    $btnDetail.Add_Click({
        Show-FileDetailWindow
    })
    $form.Controls.Add($btnDetail)

    # コード更新
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "コード更新"
    $btnRefresh.Location = New-Object System.Drawing.Point(370, 135)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 30)
    $btnRefresh.Add_Click({
        if ([string]::IsNullOrEmpty($script:currentFolder)) {
            Show-Message "フォルダが選択されていません"
            return
        }

        $script:currentFiles = Invoke-Filter
        Show-Message "更新しました（$($script:currentFiles.Count)件）"
    })
    $form.Controls.Add($btnRefresh)

    # メッセージエリア
    $script:messageLabel = New-Object System.Windows.Forms.Label
    $script:messageLabel.Text = ""
    $script:messageLabel.Location = New-Object System.Drawing.Point(10, 185)
    $script:messageLabel.Size = New-Object System.Drawing.Size(465, 25)
    $script:messageLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $script:messageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($script:messageLabel)

    # タスクトレイアイコン
    $script:notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $script:notifyIcon.Icon = [System.Drawing.SystemIcons]::Application
    $script:notifyIcon.Text = "File Clipboard Manager"
    $script:notifyIcon.Visible = $true
    $script:notifyIcon.Add_Click({
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $form.Activate()
    })

    $form.Add_FormClosed({
        $script:notifyIcon.Dispose()
    })

    # ---- イベント ----

    # フォルダ選択
    $btnFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "プロジェクトフォルダを選択してください"
        if (-not [string]::IsNullOrEmpty($script:currentFolder)) {
            $dialog.SelectedPath = $script:currentFolder
        }

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:currentFolder = $dialog.SelectedPath

            # 履歴に追加
            Add-FolderHistory $script:currentFolder

            # ComboBox を更新
            $cmbFolder.Items.Clear()
            foreach ($folder in $script:folderHistory) {
                [void]$cmbFolder.Items.Add($folder)
            }
            $cmbFolder.SelectedIndex = 0

            $script:currentFiles = Invoke-Filter
            Save-Settings
            Show-Message "フォルダを読み込みました（$($script:currentFiles.Count)件）"
        }
    })

    # 履歴からフォルダ切替
    $cmbFolder.Add_SelectedIndexChanged({
        $selected = $cmbFolder.SelectedItem
        if ($null -eq $selected) { return }
        if ($selected -eq $script:currentFolder) { return }

        $script:currentFolder = $selected

        # 選択したフォルダを履歴の先頭に移動
        Add-FolderHistory $script:currentFolder

        # ComboBox を更新（先頭に移動を反映）
        $cmbFolder.Items.Clear()
        foreach ($folder in $script:folderHistory) {
            [void]$cmbFolder.Items.Add($folder)
        }
        $cmbFolder.SelectedIndex = 0

        $script:currentFiles = Invoke-Filter
        Save-Settings
        Show-Message "フォルダを切り替えました（$($script:currentFiles.Count)件）"
    })

    # .gitignore チェック切替
    $chkGitignore.Add_CheckedChanged({
        $script:gitignoreEnabled = $chkGitignore.Checked
        if (-not [string]::IsNullOrEmpty($script:currentFolder)) {
            $script:currentFiles = Invoke-Filter
            Show-Message "フィルタリングを更新しました（$($script:currentFiles.Count)件）"
        }
        Save-Settings
    })

    # 起動時に前回フォルダを復元
    if (-not [string]::IsNullOrEmpty($script:currentFolder)) {
        $script:currentFiles = Invoke-Filter
        if ($script:currentFiles.Count -gt 0) {
            Show-Message "前回のフォルダを読み込みました（$($script:currentFiles.Count)件）"
        }
    }

    [void]$form.ShowDialog()
}

# ================================================================================
# メイン処理
# ================================================================================

if (-not (Test-Requirements)) { exit }

Load-Settings
Show-MainWindow