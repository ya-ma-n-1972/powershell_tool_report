#requires -version 5.1

<#
.SYNOPSIS
    PowerShell Clipboard Manager - Main Process
    PowerShell版クリップボード管理ツール（メインプロセス）
.DESCRIPTION
    A clipboard history management tool with persistent storage and hotkey support
    Fixed: Multi-monitor taskbar blinking issue + Ctrl+M implementation with type casting
    クリップボード履歴管理ツール（永続化・ホットキー対応）
    修正: マルチモニター環境でのタスクバー点滅問題 + Ctrl+M実装（型キャスト対応）
.NOTES
    PowerShell: 5.1+
    Author: Professional implementation for public release
#>

# Set UTF-8 encoding to prevent character corruption
# UTF-8エンコーディングを設定して文字化けを防ぐ
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {
    Write-Verbose "Console encoding setting skipped / エンコーディング設定スキップ: $($_.Exception.Message)"
}

# Load required assemblies / 必要なアセンブリの読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# System.IO.Pipes is built-in, no need to load / System.IO.Pipesは組み込みのため読み込み不要

#region Constants Definition / 定数定義

# Application constants / アプリケーション定数
[int]$script:MAX_HISTORY_COUNT = 20        # Maximum clipboard history items / 履歴最大件数
[int]$script:MAX_TEMPLATE_COUNT = 10       # Maximum template items / 定型文最大件数
[int]$script:DISPLAY_TEXT_LENGTH = 30      # Display text truncation length / 表示テキスト最大長
[int]$script:MONITOR_INTERVAL_MS = 500     # Clipboard monitoring interval (ms) / 監視間隔
[string]$script:PIPE_NAME = "PSClipboardPipe"  # Named pipe identifier / 名前付きパイプ識別子

# Clipboard access retry settings / クリップボードアクセスリトライ設定
[int]$script:CLIPBOARD_RETRY_COUNT = 3
[int]$script:CLIPBOARD_RETRY_DELAY_MS = 100

# Data persistence settings / データ永続化設定
[int]$script:SAVE_FREQUENCY = 5  # Save every N clipboard updates / N回の更新ごとに保存

#endregion

#region Global Variables / グローバル変数

# Data storage / データストレージ
[System.Collections.ArrayList]$script:clipboard_history = [System.Collections.ArrayList]::new()
[System.Collections.ArrayList]$script:template_data = [System.Collections.ArrayList]::new()
[string]$script:current_clipboard = ""

# File paths / ファイルパス
[string]$script:data_file_path = Join-Path $PSScriptRoot "clipboard_data.json"

# Application state / アプリケーション状態
[bool]$script:persist_history = $true      # Persist history across sessions / 履歴の永続化
[bool]$script:auto_hide_to_tray = $true    # Auto-hide after clipboard set / 設定後の自動格納
[bool]$script:is_editing = $false          # Edit dialog display flag / 編集ダイアログ表示中フラグ

# UI components / UIコンポーネント
$script:main_form = $null
$script:listbox_history = $null
$script:listbox_template = $null
$script:notify_icon = $null
$script:timer = $null

# Process management / プロセス管理
$script:hotkey_process = $null
$script:pipe_runspace = $null

#endregion

#region Utility Functions / ユーティリティ関数

# --- 安全なハッシュ値取得関数 ---
<#
.SYNOPSIS
    Safely retrieve value from hashtable or object
    ハッシュテーブルまたはオブジェクトから安全に値を取得
#>
function Get-SafeHashValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $HashTable,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$KeyName,
        
        [string]$DefaultValue = ""
    )
    
    try {
        if ($HashTable -is [hashtable] -and $HashTable.ContainsKey($KeyName)) {
            return $HashTable[$KeyName]
        }
        elseif ($HashTable.PSObject.Properties[$KeyName]) {
            return $HashTable.PSObject.Properties[$KeyName].Value
        }
        elseif ($null -ne $HashTable.$KeyName) {
            return $HashTable.$KeyName
        }
        else {
            Write-Verbose "Key '$KeyName' not found, using default value / キー '$KeyName' が見つかりません。デフォルト値を使用"
            return $DefaultValue
        }
    }
    catch {
        Write-Warning "[HashAccess] Failed to access key / キーアクセス失敗: $($_.Exception.Message)"
        return $DefaultValue
    }
}

# --- テキスト入力検証関数 ---
<#
.SYNOPSIS
    Validate text input for clipboard operations
    クリップボード操作用のテキスト入力を検証
#>
function Test-ValidTextInput {
    [CmdletBinding()]
    param(
        [string]$Text,
        [int]$MaxLength = 10000
    )
    
    if ([string]::IsNullOrWhiteSpace($Text)) {
        Write-Verbose "Invalid input: empty text / 無効な入力: 空のテキスト"
        return $false
    }
    
    if ($Text.Length -gt $MaxLength) {
        Write-Verbose "Invalid input: text too long / 無効な入力: テキストが長すぎます ($($Text.Length) > $MaxLength)"
        return $false
    }
    
    return $true
}

# --- ListBox表示用テキスト整形関数 ---
<#
.SYNOPSIS
    Format text for display in ListBox
    ListBox表示用にテキストをフォーマット
#>
function Format-DisplayText {
    [CmdletBinding()]
    param(
        [string]$Text,
        [int]$MaxLength = $script:DISPLAY_TEXT_LENGTH,
        [string]$Prefix = ""
    )
    
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return "${Prefix}[Empty text / 空のテキスト]"
    }
    
    if ($Text.Length -gt $MaxLength) {
        return "${Prefix}$($Text.Substring(0, $MaxLength))..."
    }
    
    return "${Prefix}${Text}"
}

#endregion

#region Data Management Functions / データ管理関数

# --- アプリケーションデータ読込関数 ---
<#
.SYNOPSIS
    Load application data from JSON file
    JSONファイルからアプリケーションデータを読み込み
#>
function Import-ApplicationData {
    [CmdletBinding()]
    param()
    
    try {
        Write-Information "Loading application data / アプリケーションデータ読み込み中..."
        
        if (-not (Test-Path $script:data_file_path)) {
            Write-Information "Data file not found, initializing / データファイルが存在しません。初期化します"
            return
        }
        
        $jsonContent = Get-Content -Path $script:data_file_path -Encoding UTF8 -Raw
        
        if ([string]::IsNullOrWhiteSpace($jsonContent)) {
            Write-Warning "Empty data file / 空のデータファイル"
            return
        }
        
        $data = $jsonContent | ConvertFrom-Json
        
        # Load template data / 定型文データの読み込み
        if ($null -ne $data.template_data) {
            $script:template_data.Clear()
            
            foreach ($item in $data.template_data) {
                $hashItem = @{
                    id = [int](Get-SafeHashValue $item "id" "1")
                    text = Get-SafeHashValue $item "text" ""
                }
                
                if (Test-ValidTextInput $hashItem.text) {
                    [void]$script:template_data.Add($hashItem)
                }
            }
            
            Write-Information "Templates loaded / 定型文読み込み完了: $($script:template_data.Count) items"
        }
        
        # Load clipboard history / クリップボード履歴の読み込み
        if ($script:persist_history -and $null -ne $data.clipboard_history) {
            $script:clipboard_history.Clear()
            
            $historyCount = 0
            foreach ($historyItem in $data.clipboard_history) {
                if ((Test-ValidTextInput $historyItem) -and $historyCount -lt $script:MAX_HISTORY_COUNT) {
                    [void]$script:clipboard_history.Add($historyItem)
                    $historyCount++
                }
            }
            
            Write-Information "History loaded / 履歴読み込み完了: $($script:clipboard_history.Count) items"
        }
        
        # Load settings / 設定の読み込み
        if ($null -ne $data.settings) {
            if ($null -ne $data.settings.persist_history) {
                $script:persist_history = [bool]$data.settings.persist_history
            }
            if ($null -ne $data.settings.auto_hide_to_tray) {
                $script:auto_hide_to_tray = [bool]$data.settings.auto_hide_to_tray
            }
        }
    }
    catch {
        Write-Error "[DataLoad] Failed to load data / データ読み込み失敗: $($_.Exception.Message)"
        $script:template_data.Clear()
        $script:clipboard_history.Clear()
    }
}

# --- アプリケーションデータ保存関数 ---
<#
.SYNOPSIS
    Save application data to JSON file
    アプリケーションデータをJSONファイルに保存
#>
function Export-ApplicationData {
    [CmdletBinding()]
    param()
    
    try {
        Write-Verbose "Saving application data / アプリケーションデータ保存中..."
        
        $saveData = @{
            template_data = @($script:template_data)
            clipboard_history = if ($script:persist_history) { 
                @($script:clipboard_history | Select-Object -First $script:MAX_HISTORY_COUNT) 
            } else { 
                @() 
            }
            settings = @{
                persist_history = $script:persist_history
                auto_hide_to_tray = $script:auto_hide_to_tray
                max_history_count = $script:MAX_HISTORY_COUNT
                max_template_count = $script:MAX_TEMPLATE_COUNT
            }
            metadata = @{
                version = "2.5"
                last_saved = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            }
        }
        
        $jsonContent = $saveData | ConvertTo-Json -Depth 4 -Compress
        [System.IO.File]::WriteAllText($script:data_file_path, $jsonContent, [System.Text.Encoding]::UTF8)
        
        Write-Verbose "Data saved successfully / データ保存完了"
    }
    catch {
        Write-Error "[DataSave] Failed to save data / データ保存失敗: $($_.Exception.Message)"
    }
}

#endregion

#region Clipboard Management Functions / クリップボード管理関数

# --- クリップボード内容取得（リトライ付き） ---
<#
.SYNOPSIS
    Safely retrieve clipboard content with retry logic
    リトライロジック付きでクリップボード内容を安全に取得
#>
function Get-ClipboardContent {
    [CmdletBinding()]
    param()
    
    for ($i = 0; $i -lt $script:CLIPBOARD_RETRY_COUNT; $i++) {
        try {
            if ([System.Windows.Forms.Clipboard]::ContainsText()) {
                return [System.Windows.Forms.Clipboard]::GetText()
            }
            return ""
        }
        catch {
            if ($i -eq ($script:CLIPBOARD_RETRY_COUNT - 1)) {
                Write-Warning "[Clipboard] Failed to get content / 取得失敗: $($_.Exception.Message)"
                return ""
            }
            Start-Sleep -Milliseconds $script:CLIPBOARD_RETRY_DELAY_MS
        }
    }
    return ""
}

# --- クリップボード内容設定（リトライ付き） ---
<#
.SYNOPSIS
    Safely set clipboard content with retry logic
    リトライロジック付きでクリップボード内容を安全に設定
#>
function Set-ClipboardContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Text,
        
        [bool]$AutoHide = $true
    )
    
    if (-not (Test-ValidTextInput $Text)) {
        Write-Warning "[Clipboard] Invalid text input / 無効なテキスト入力"
        return $false
    }
    
    for ($i = 0; $i -lt $script:CLIPBOARD_RETRY_COUNT; $i++) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($Text)
            
            $displayText = Format-DisplayText $Text 50
            Write-Information "Clipboard set / クリップボード設定: $displayText"
            
            if ($AutoHide -and $script:auto_hide_to_tray -and $null -ne $script:main_form) {
                Start-Sleep -Milliseconds 200
                Hide-ToTray
                Write-Verbose "Auto-hidden to tray / トレイに自動格納"
            }
            
            return $true
        }
        catch {
            if ($i -eq ($script:CLIPBOARD_RETRY_COUNT - 1)) {
                Write-Error "[Clipboard] Failed to set content / 設定失敗: $($_.Exception.Message)"
                return $false
            }
            Start-Sleep -Milliseconds $script:CLIPBOARD_RETRY_DELAY_MS
        }
    }
    return $false
}

# --- 履歴更新（重複除去付き） ---
<#
.SYNOPSIS
    Update clipboard history with deduplication
    重複除去付きでクリップボード履歴を更新
#>
function Update-ClipboardHistory {
    [CmdletBinding()]
    param()
    
    $currentContent = Get-ClipboardContent
    
    if ([string]::IsNullOrWhiteSpace($currentContent) -or 
        $currentContent -eq $script:current_clipboard) {
        return
    }
    
    try {
        $script:clipboard_history = [System.Collections.ArrayList]@(
            $script:clipboard_history | Where-Object { $_ -ne $currentContent }
        )
        
        $script:clipboard_history.Insert(0, $currentContent)
        
        while ($script:clipboard_history.Count -gt $script:MAX_HISTORY_COUNT) {
            $script:clipboard_history.RemoveAt($script:clipboard_history.Count - 1)
        }
        
        $script:current_clipboard = $currentContent
        
        Update-HistoryDisplay
        
        if ($script:persist_history -and ($script:clipboard_history.Count % $script:SAVE_FREQUENCY -eq 0)) {
            Export-ApplicationData
        }
        
        Write-Verbose "History updated / 履歴更新: $($script:clipboard_history.Count) items"
    }
    catch {
        Write-Error "[History] Update failed / 更新失敗: $($_.Exception.Message)"
    }
}

#endregion

#region GUI Update Functions / GUI更新関数

# --- 履歴ListBox表示更新 ---
<#
.SYNOPSIS
    Update history ListBox display
    履歴ListBoxの表示を更新
#>
function Update-HistoryDisplay {
    [CmdletBinding()]
    param()
    
    if ($null -eq $script:listbox_history) { 
        return 
    }
    
    try {
        $script:listbox_history.BeginUpdate()
        $script:listbox_history.Items.Clear()
        
        for ($i = 0; $i -lt $script:clipboard_history.Count; $i++) {
            $itemText = $script:clipboard_history[$i]
            $displayText = Format-DisplayText $itemText -Prefix "$($i + 1): "
            [void]$script:listbox_history.Items.Add($displayText)
        }
    }
    catch {
        Write-Error "[Display] History update failed / 履歴表示更新失敗: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $script:listbox_history) {
            $script:listbox_history.EndUpdate()
        }
    }
}

# --- 定型文ListBox表示更新 ---
<#
.SYNOPSIS
    Update template ListBox display
    定型文ListBoxの表示を更新
#>
function Update-TemplateDisplay {
    [CmdletBinding()]
    param()
    
    if ($null -eq $script:listbox_template) { 
        return 
    }
    
    try {
        $script:listbox_template.BeginUpdate()
        $script:listbox_template.Items.Clear()
        
        for ($i = 0; $i -lt $script:template_data.Count; $i++) {
            $item = $script:template_data[$i]
            $itemText = Get-SafeHashValue $item "text" "[Error / エラー]"
            $displayText = Format-DisplayText $itemText -Prefix "T$($i + 1): "
            [void]$script:listbox_template.Items.Add($displayText)
        }
    }
    catch {
        Write-Error "[Display] Template update failed / 定型文表示更新失敗: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $script:listbox_template) {
            $script:listbox_template.EndUpdate()
        }
    }
}

#endregion

#region Edit Dialog and Template Management / 編集ダイアログと定型文管理

# --- テキスト編集ダイアログ表示 ---
<#
.SYNOPSIS
    Show text edit dialog with multiline support
    複数行対応のテキスト編集ダイアログを表示
#>
function Show-EditDialog {
    [CmdletBinding()]
    param(
        [string]$InitialText = "",
        [string]$Title = "Edit Text / テキスト編集",
        [int]$MaxLength = 10000
    )
    
    try {
        $script:is_editing = $true
        
        $editForm = New-Object System.Windows.Forms.Form
        $editForm.Text = $Title
        $editForm.Size = New-Object System.Drawing.Size(520, 350)
        $editForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $editForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $editForm.MaximizeBox = $false
        $editForm.MinimizeBox = $false
        
        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Multiline = $true
        $textbox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $textbox.Location = New-Object System.Drawing.Point(10, 10)
        $textbox.Size = New-Object System.Drawing.Size(480, 220)
        $textbox.Text = $InitialText
        $textbox.MaxLength = $MaxLength
        $textbox.AcceptsReturn = $true
        
        $charCountLabel = New-Object System.Windows.Forms.Label
        $charCountLabel.Location = New-Object System.Drawing.Point(10, 240)
        $charCountLabel.Size = New-Object System.Drawing.Size(300, 20)
        $charCountLabel.Text = "Characters / 文字数: $($InitialText.Length)/$MaxLength"
        
        $textbox.add_TextChanged({
            $charCountLabel.Text = "Characters / 文字数: $($textbox.Text.Length)/$MaxLength"
        })
        
        $editForm.Controls.AddRange(@($textbox, $charCountLabel))
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(335, 270)
        $okButton.Size = New-Object System.Drawing.Size(75, 30)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(415, 270)
        $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        
        $editForm.Controls.AddRange(@($okButton, $cancelButton))
        $editForm.AcceptButton = $okButton
        $editForm.CancelButton = $cancelButton
        
        $textbox.Select()
        
        $result = $editForm.ShowDialog($script:main_form)
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $textbox.Text
        }
        
        return $null
    }
    catch {
        Write-Error "[Dialog] Edit dialog failed / 編集ダイアログ失敗: $($_.Exception.Message)"
        return $null
    }
    finally {
        if ($null -ne $editForm) {
            $editForm.Dispose()
        }
        $script:is_editing = $false
    }
}

# --- 定型文追加（検証付き） ---
function Add-TemplateItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TextContent
    )
    
    if (-not (Test-ValidTextInput $TextContent)) {
        Write-Warning "[Template] Invalid text / 無効なテキスト"
        return $false
    }
    
    if ($script:template_data.Count -ge $script:MAX_TEMPLATE_COUNT) {
        $message = "Maximum template limit reached / 定型文の最大件数に達しました ($script:MAX_TEMPLATE_COUNT)"
        Write-Warning $message
        [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Limit Reached / 制限到達",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return $false
    }
    
    try {
        $newId = if ($script:template_data.Count -eq 0) { 
            1 
        } else { 
            ($script:template_data | ForEach-Object { $_.id } | Measure-Object -Maximum).Maximum + 1
        }
        
        $newItem = @{
            id = $newId
            text = $TextContent
        }
        
        [void]$script:template_data.Add($newItem)
        
        Update-TemplateDisplay
        Export-ApplicationData
        
        Write-Information "Template added / 定型文追加 ($($script:template_data.Count)/$script:MAX_TEMPLATE_COUNT)"
        return $true
    }
    catch {
        Write-Error "[Template] Failed to add / 追加失敗: $($_.Exception.Message)"
        return $false
    }
}

# --- 定型文上移動 ---
function Move-TemplateUp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$Index
    )
    
    if ($Index -le 0 -or $Index -ge $script:template_data.Count) {
        return
    }
    
    try {
        $temp = $script:template_data[$Index]
        $script:template_data[$Index] = $script:template_data[$Index - 1]
        $script:template_data[$Index - 1] = $temp
        
        Update-TemplateDisplay
        Export-ApplicationData
        
        $script:listbox_template.SelectedIndex = $Index - 1
    }
    catch {
        Write-Error "[Template] Failed to move up / 上移動失敗: $($_.Exception.Message)"
    }
}

# --- 定型文下移動 ---
function Move-TemplateDown {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$Index
    )
    
    if ($Index -lt 0 -or $Index -ge ($script:template_data.Count - 1)) {
        return
    }
    
    try {
        $temp = $script:template_data[$Index]
        $script:template_data[$Index] = $script:template_data[$Index + 1]
        $script:template_data[$Index + 1] = $temp
        
        Update-TemplateDisplay
        Export-ApplicationData
        
        $script:listbox_template.SelectedIndex = $Index + 1
    }
    catch {
        Write-Error "[Template] Failed to move down / 下移動失敗: $($_.Exception.Message)"
    }
}

#endregion

#region System Tray Functions / システムトレイ関数

# --- システムトレイアイコン初期化 ---
<#
.SYNOPSIS
    Initialize system tray icon
    システムトレイアイコンを初期化
#>
function Initialize-NotifyIcon {
    [CmdletBinding()]
    param()
    
    try {
        $script:notify_icon = New-Object System.Windows.Forms.NotifyIcon
        $script:notify_icon.Text = "PS Clipboard Manager`nCtrl+Shift+Space"
        $script:notify_icon.Visible = $true
        
        $iconPath = [System.IO.Path]::Combine($PSHOME, "powershell.exe")
        if (Test-Path $iconPath) {
            $script:notify_icon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
        } else {
            $script:notify_icon.Icon = [System.Drawing.SystemIcons]::Application
        }
        
        $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
        
        $showItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $showItem.Text = "Show / 表示"
        $showItem.add_Click({
            Show-MainWindowAtCursor
        })
        [void]$contextMenu.Items.Add($showItem)
        
        [void]$contextMenu.Items.Add("-")
        
        $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $exitItem.Text = "Exit / 終了"
        $exitItem.add_Click({
            $script:main_form.Close()
        })
        [void]$contextMenu.Items.Add($exitItem)
        
        $script:notify_icon.ContextMenuStrip = $contextMenu
        
        $script:notify_icon.add_MouseClick({
            param($sender, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                if ($script:main_form.Visible) {
                    Hide-ToTray
                } else {
                    Show-MainWindowAtCursor
                }
            }
        })
        
        Write-Information "System tray icon initialized / システムトレイアイコン初期化完了"
    }
    catch {
        Write-Error "[SystemTray] Initialization failed / 初期化失敗: $($_.Exception.Message)"
    }
}

# --- カーソル位置にメインウィンドウ表示（マルチモニター対応） ---
<#
.SYNOPSIS
    Show main window at cursor position (FIXED)
    カーソル位置にメインウィンドウを表示（修正版）
.DESCRIPTION
    Fixed ShowInTaskbar issue for multi-monitor environments
    マルチモニター環境のShowInTaskbar問題を修正
#>
function Show-MainWindowAtCursor {
    [CmdletBinding()]
    param()
    
    if ($null -eq $script:main_form) { 
        return 
    }
    
    try {
        # Get cursor position / カーソル位置を取得
        $cursorPosition = [System.Windows.Forms.Cursor]::Position
        
        # Get screen information / スクリーン情報を取得
        $screen = [System.Windows.Forms.Screen]::FromPoint($cursorPosition)
        $workingArea = $screen.WorkingArea
        
        # Calculate window position / ウィンドウ位置を計算
        $formWidth = $script:main_form.Width
        $formHeight = $script:main_form.Height
        
        $x = $cursorPosition.X - 20
        $y = $cursorPosition.Y - $formHeight - 10
        
        # Adjust to fit within screen / 画面内に収まるように調整
        $x = [Math]::Max($workingArea.X, [Math]::Min($x, $workingArea.Right - $formWidth))
        $y = [Math]::Max($workingArea.Y, [Math]::Min($y, $workingArea.Bottom - $formHeight))
        
        # Show window / ウィンドウを表示
        $script:main_form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $script:main_form.Location = New-Object System.Drawing.Point($x, $y)
        $script:main_form.Show()
        $script:main_form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        
        # FIXED: Keep ShowInTaskbar as false to prevent taskbar blinking
        # 修正: タスクバーの点滅を防ぐためShowInTaskbarはfalseのまま
        
        $script:main_form.TopMost = $true
        $script:main_form.Activate()
        $script:main_form.BringToFront()
        
        # Set focus to history listbox / 履歴リストボックスにフォーカス
        if ($null -ne $script:listbox_history -and $script:listbox_history.Items.Count -gt 0) {
            $script:listbox_history.Focus()
            if ($script:listbox_history.SelectedIndex -lt 0) {
                $script:listbox_history.SelectedIndex = 0
            }
        }
        
        # Release TopMost with longer delay / より長い遅延でTopMostを解除
        Start-Sleep -Milliseconds 300
        $script:main_form.TopMost = $false
        
        Write-Verbose "Window shown at cursor / カーソル位置に表示: X=$x, Y=$y"
    }
    catch {
        Write-Error "[Window] Failed to show at cursor / カーソル位置表示失敗: $($_.Exception.Message)"
    }
}

# --- ウィンドウをトレイに格納 ---
function Hide-ToTray {
    [CmdletBinding()]
    param()
    
    if ($null -eq $script:main_form) { 
        return 
    }
    
    $script:main_form.Hide()
    # FIXED: ShowInTaskbar remains false / 修正: ShowInTaskbarはfalseのまま
    # $script:main_form.ShowInTaskbar = $false  # Already false / すでにfalse
}

#endregion

#region Named Pipe Server / 名前付きパイプサーバー

# --- 名前付きパイプサーバー起動（Runspace利用） ---
<#
.SYNOPSIS
    Start named pipe server in separate runspace (FIXED)
    別のRunspaceで名前付きパイプサーバーを起動（修正版）
#>
function Start-PipeServer {
    [CmdletBinding()]
    param()
    
    try {
        # Create runspace with STA apartment state / STAアパートメント状態でRunspaceを作成
        $script:pipe_runspace = [runspacefactory]::CreateRunspace()
        $script:pipe_runspace.ApartmentState = "STA"
        $script:pipe_runspace.ThreadOptions = "ReuseThread"
        $script:pipe_runspace.Open()
        
        # Share variables with runspace / Runspaceと変数を共有
        $script:pipe_runspace.SessionStateProxy.SetVariable("main_form", $script:main_form)
        $script:pipe_runspace.SessionStateProxy.SetVariable("pipe_name", $script:PIPE_NAME)
        $script:pipe_runspace.SessionStateProxy.SetVariable("listbox_history", $script:listbox_history)
        $script:pipe_runspace.SessionStateProxy.SetVariable("listbox_template", $script:listbox_template)
        
        $powershell = [powershell]::Create()
        $powershell.Runspace = $script:pipe_runspace
        
        [void]$powershell.AddScript({
            Write-Host "[PipeServer] Starting / 起動中..." -ForegroundColor Green
            
            $retryDelay = 1000
            $maxRetryDelay = 30000
            
            while ($true) {
                try {
                    $pipe = New-Object System.IO.Pipes.NamedPipeServerStream(
                        $pipe_name,
                        [System.IO.Pipes.PipeDirection]::InOut
                    )
                    
                    Write-Host "[PipeServer] Waiting for connection / 接続待機中..." -ForegroundColor Cyan
                    $pipe.WaitForConnection()
                    Write-Host "[PipeServer] Client connected / クライアント接続" -ForegroundColor Green
                    
                    $reader = New-Object System.IO.StreamReader($pipe)
                    $command = $reader.ReadLine()
                    Write-Host "[PipeServer] Command received / コマンド受信: $command" -ForegroundColor Yellow
                    
                    if ($command -eq "SHOW") {
                        # Show window in main thread / メインスレッドでウィンドウを表示
                        if ($null -ne $main_form) {
                            $main_form.Invoke([Action]{
                                # Get cursor position / カーソル位置を取得
                                $cursorPosition = [System.Windows.Forms.Cursor]::Position
                                $screen = [System.Windows.Forms.Screen]::FromPoint($cursorPosition)
                                $workingArea = $screen.WorkingArea
                                
                                # Calculate position / 位置を計算
                                $formWidth = $main_form.Width
                                $formHeight = $main_form.Height
                                
                                $x = $cursorPosition.X - 20
                                $y = $cursorPosition.Y - $formHeight - 10
                                
                                # Adjust within screen / 画面内に調整
                                $x = [Math]::Max($workingArea.X, [Math]::Min($x, $workingArea.Right - $formWidth))
                                $y = [Math]::Max($workingArea.Y, [Math]::Min($y, $workingArea.Bottom - $formHeight))
                                
                                # Show window / ウィンドウを表示
                                $main_form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
                                $main_form.Location = New-Object System.Drawing.Point($x, $y)
                                $main_form.Show()
                                $main_form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
                                
                                # FIXED: Keep ShowInTaskbar as false / 修正: ShowInTaskbarはfalseのまま
                                
                                $main_form.TopMost = $true
                                $main_form.Activate()
                                $main_form.BringToFront()
                                
                                # Set focus / フォーカスを設定
                                if ($null -ne $listbox_history -and $listbox_history.Items.Count -gt 0) {
                                    $listbox_history.Focus()
                                    if ($listbox_history.SelectedIndex -lt 0) {
                                        $listbox_history.SelectedIndex = 0
                                    }
                                }
                                
                                # Release TopMost with longer delay / より長い遅延でTopMostを解除
                                Start-Sleep -Milliseconds 300
                                $main_form.TopMost = $false
                            })
                        }
                        
                        # Send response / 応答を送信
                        $writer = New-Object System.IO.StreamWriter($pipe)
                        $writer.WriteLine("OK")
                        $writer.Flush()
                        $writer.Close()
                    }
                    
                    # Cleanup / クリーンアップ
                    $reader.Close()
                    $pipe.Close()
                    $pipe.Dispose()
                    
                    $retryDelay = 1000
                    
                } catch {
                    Write-Host "[PipeServer] Error / エラー: $_ - Retrying in $($retryDelay/1000) seconds..." -ForegroundColor Red
                    Start-Sleep -Milliseconds $retryDelay
                    $retryDelay = [Math]::Min($retryDelay * 2, $maxRetryDelay)
                }
            }
        })
        
        # Start asynchronous execution / 非同期実行を開始
        [void]$powershell.BeginInvoke()
        
        # Wait for pipe server initialization / パイプサーバーの初期化を待つ
        Write-Host "Waiting for pipe server initialization / パイプサーバー初期化待機中..." -ForegroundColor Yellow
        Start-Sleep -Milliseconds 500
        
        Write-Information "Pipe server started / パイプサーバー起動完了"
    }
    catch {
        Write-Error "[PipeServer] Failed to start / 起動失敗: $($_.Exception.Message)"
    }
}

#endregion

#region Event Handlers / イベントハンドラー

# --- 履歴項目コンテキストメニュー表示 ---
<#
.SYNOPSIS
    Show context menu for history items
    履歴項目のコンテキストメニューを表示
#>
function Show-HistoryContextMenu {
    [CmdletBinding()]
    param($Sender, $EventArgs)
    
    try {
        $selectedIndex = $script:listbox_history.SelectedIndex
        $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
        
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $script:clipboard_history.Count) {
            $editItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $editItem.Text = "Edit / 編集"
            $editItem.add_Click({
                $index = $script:listbox_history.SelectedIndex
                if ($index -ge 0) {
                    $currentText = $script:clipboard_history[$index]
                    $editedText = Show-EditDialog $currentText "Edit History / 履歴の編集"
                    
                    if ($null -ne $editedText -and $editedText -ne $currentText) {
                        $script:clipboard_history[$index] = $editedText
                        Update-HistoryDisplay
                        if ($script:persist_history) {
                            Export-ApplicationData
                        }
                    }
                }
            })
            [void]$contextMenu.Items.Add($editItem)
            
            $deleteItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $deleteItem.Text = "Delete / 削除"
            $deleteItem.add_Click({
                $index = $script:listbox_history.SelectedIndex
                if ($index -ge 0) {
                    $targetText = $script:clipboard_history[$index]
                    $displayTarget = Format-DisplayText $targetText 100
                    
                    $message = "Delete this item? / この項目を削除しますか？`n`n$displayTarget"
                    $result = [System.Windows.Forms.MessageBox]::Show(
                        $message,
                        "Confirm Delete / 削除確認",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )
                    
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        $script:clipboard_history.RemoveAt($index)
                        Update-HistoryDisplay
                        if ($script:persist_history) {
                            Export-ApplicationData
                        }
                    }
                }
            })
            [void]$contextMenu.Items.Add($deleteItem)
            
            [void]$contextMenu.Items.Add("-")
            
            $addTemplateItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $addTemplateItem.Text = "Add to Template / 定型文に追加"
            $addTemplateItem.add_Click({
                $index = $script:listbox_history.SelectedIndex
                if ($index -ge 0) {
                    $textToAdd = $script:clipboard_history[$index]
                    Add-TemplateItem $textToAdd
                }
            })
            [void]$contextMenu.Items.Add($addTemplateItem)
        }
        
        $contextMenu.Show($script:listbox_history, $EventArgs.Location)
    }
    catch {
        Write-Error "[ContextMenu] History menu failed / 履歴メニュー失敗: $($_.Exception.Message)"
    }
}

# --- 定型文項目コンテキストメニュー表示 ---
<#
.SYNOPSIS
    Show context menu for template items
    定型文項目のコンテキストメニューを表示
#>
function Show-TemplateContextMenu {
    [CmdletBinding()]
    param($Sender, $EventArgs)
    
    try {
        $selectedIndex = $script:listbox_template.SelectedIndex
        $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
        
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $script:template_data.Count) {
            $editItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $editItem.Text = "Edit / 編集"
            $editItem.add_Click({
                $index = $script:listbox_template.SelectedIndex
                if ($index -ge 0) {
                    $currentText = Get-SafeHashValue $script:template_data[$index] "text"
                    $editedText = Show-EditDialog $currentText "Edit Template / 定型文の編集"
                    
                    if ($null -ne $editedText -and $editedText -ne $currentText) {
                        $script:template_data[$index]["text"] = $editedText
                        Update-TemplateDisplay
                        Export-ApplicationData
                    }
                }
            })
            [void]$contextMenu.Items.Add($editItem)
            
            $deleteItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $deleteItem.Text = "Delete / 削除"
            $deleteItem.add_Click({
                $index = $script:listbox_template.SelectedIndex
                if ($index -ge 0) {
                    $script:template_data.RemoveAt($index)
                    Update-TemplateDisplay
                    Export-ApplicationData
                }
            })
            [void]$contextMenu.Items.Add($deleteItem)
            
            [void]$contextMenu.Items.Add("-")
            
            if ($selectedIndex -gt 0) {
                $moveUpItem = New-Object System.Windows.Forms.ToolStripMenuItem
                $moveUpItem.Text = "Move Up / 上へ移動"
                $moveUpItem.add_Click({
                    Move-TemplateUp $script:listbox_template.SelectedIndex
                })
                [void]$contextMenu.Items.Add($moveUpItem)
            }
            
            if ($selectedIndex -lt ($script:template_data.Count - 1)) {
                $moveDownItem = New-Object System.Windows.Forms.ToolStripMenuItem
                $moveDownItem.Text = "Move Down / 下へ移動"
                $moveDownItem.add_Click({
                    Move-TemplateDown $script:listbox_template.SelectedIndex
                })
                [void]$contextMenu.Items.Add($moveDownItem)
            }
        } else {
            $newItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $newItem.Text = "New Template / 新規登録"
            $newItem.add_Click({
                $newText = Show-EditDialog "" "New Template / 新規定型文"
                if (-not [string]::IsNullOrWhiteSpace($newText)) {
                    Add-TemplateItem $newText
                }
            })
            [void]$contextMenu.Items.Add($newItem)
            
            $clipboardItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $clipboardItem.Text = "Add from Clipboard / クリップボードから登録"
            $clipboardItem.add_Click({
                $currentText = Get-ClipboardContent
                if (-not [string]::IsNullOrWhiteSpace($currentText)) {
                    Add-TemplateItem $currentText
                }
            })
            [void]$contextMenu.Items.Add($clipboardItem)
        }
        
        $contextMenu.Show($script:listbox_template, $EventArgs.Location)
    }
    catch {
        Write-Error "[ContextMenu] Template menu failed / 定型文メニュー失敗: $($_.Exception.Message)"
    }
}

# --- 履歴項目クリック処理 ---
<#
.SYNOPSIS
    Handle history item click
    履歴項目クリックを処理
#>
function Invoke-HistoryClick {
    [CmdletBinding()]
    param()
    
    try {
        $selectedIndex = $script:listbox_history.SelectedIndex
        
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $script:clipboard_history.Count) {
            $selectedText = $script:clipboard_history[$selectedIndex]
            Set-ClipboardContent $selectedText $true
        }
    }
    catch {
        Write-Error "[Click] History click failed / 履歴クリック失敗: $($_.Exception.Message)"
    }
}

# --- 定型文項目クリック処理 ---
<#
.SYNOPSIS
    Handle template item click
    定型文項目クリックを処理
#>
function Invoke-TemplateClick {
    [CmdletBinding()]
    param()
    
    try {
        $selectedIndex = $script:listbox_template.SelectedIndex
        
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $script:template_data.Count) {
            $selectedText = Get-SafeHashValue $script:template_data[$selectedIndex] "text"
            
            if (-not [string]::IsNullOrWhiteSpace($selectedText)) {
                Set-ClipboardContent $selectedText $true
            }
        }
    }
    catch {
        Write-Error "[Click] Template click failed / 定型文クリック失敗: $($_.Exception.Message)"
    }
}

# --- ListBoxキーダウン処理（Enter/Tab/Esc/Ctrl+M/Appsキー対応） ---
<#
.SYNOPSIS
    Handle keyboard operations (FIXED WITH CTRL+M)
    キーボード操作を処理（Ctrl+M対応修正版）
#>
function Invoke-ListBoxKeyDown {
    [CmdletBinding()]
    param($Sender, $EventArgs)
    
    try {
        switch ($EventArgs.KeyCode) {
            ([System.Windows.Forms.Keys]::Enter) {
                if ($Sender -eq $script:listbox_history) {
                    Invoke-HistoryClick
                } elseif ($Sender -eq $script:listbox_template) {
                    Invoke-TemplateClick
                }
                $EventArgs.Handled = $true
            }
            
            ([System.Windows.Forms.Keys]::Escape) {
                Hide-ToTray
                $EventArgs.Handled = $true
            }
            
            ([System.Windows.Forms.Keys]::Tab) {
                if ($Sender -eq $script:listbox_history -and $script:listbox_template.Items.Count -gt 0) {
                    $script:listbox_template.Focus()
                    $script:listbox_template.SelectedIndex = 0
                } elseif ($Sender -eq $script:listbox_template -and $script:listbox_history.Items.Count -gt 0) {
                    $script:listbox_history.Focus()
                    $script:listbox_history.SelectedIndex = 0
                }
                $EventArgs.Handled = $true
                $EventArgs.SuppressKeyPress = $true
            }
            
            # M key with Ctrl: Alternative context menu / Ctrl+M: 代替コンテキストメニュー
            ([System.Windows.Forms.Keys]::M) {
                if ($EventArgs.Control -and -not $EventArgs.Shift -and -not $EventArgs.Alt) {
                    $selectedIndex = $Sender.SelectedIndex
                    
                    if ($Sender -eq $script:listbox_history) {
                        $location = if ($selectedIndex -ge 0) {
                            $bounds = $Sender.GetItemRectangle($selectedIndex)
                            # Explicit type casting to avoid operator error / 演算エラー回避のため明示的に型キャスト
                            $x = [int]$bounds.Left + 20
                            $y = [int]$bounds.Bottom
                            New-Object System.Drawing.Point($x, $y)
                        } else {
                            New-Object System.Drawing.Point(10, 10)
                        }
                        Show-HistoryContextMenu $Sender ([PSCustomObject]@{Location = $location})
                    } elseif ($Sender -eq $script:listbox_template) {
                        $location = if ($selectedIndex -ge 0) {
                            $bounds = $Sender.GetItemRectangle($selectedIndex)
                            # Explicit type casting to avoid operator error / 演算エラー回避のため明示的に型キャスト
                            $x = [int]$bounds.Left + 20
                            $y = [int]$bounds.Bottom
                            New-Object System.Drawing.Point($x, $y)
                        } else {
                            New-Object System.Drawing.Point(10, 10)
                        }
                        Show-TemplateContextMenu $Sender ([PSCustomObject]@{Location = $location})
                    }
                    
                    $EventArgs.Handled = $true
                    $EventArgs.SuppressKeyPress = $true
                }
            }
            
            ([System.Windows.Forms.Keys]::Apps) {
                $selectedIndex = $Sender.SelectedIndex
                
                if ($Sender -eq $script:listbox_history) {
                    $location = if ($selectedIndex -ge 0) {
                        $bounds = $Sender.GetItemRectangle($selectedIndex)
                        # Explicit type casting to avoid operator error / 演算エラー回避のため明示的に型キャスト
                        $x = [int]$bounds.Left + 20
                        $y = [int]$bounds.Bottom
                        New-Object System.Drawing.Point($x, $y)
                    } else {
                        New-Object System.Drawing.Point(10, 10)
                    }
                    Show-HistoryContextMenu $Sender ([PSCustomObject]@{Location = $location})
                } elseif ($Sender -eq $script:listbox_template) {
                    $location = if ($selectedIndex -ge 0) {
                        $bounds = $Sender.GetItemRectangle($selectedIndex)
                        # Explicit type casting to avoid operator error / 演算エラー回避のため明示的に型キャスト
                        $x = [int]$bounds.Left + 20
                        $y = [int]$bounds.Bottom
                        New-Object System.Drawing.Point($x, $y)
                    } else {
                        New-Object System.Drawing.Point(10, 10)
                    }
                    Show-TemplateContextMenu $Sender ([PSCustomObject]@{Location = $location})
                }
                
                $EventArgs.Handled = $true
                $EventArgs.SuppressKeyPress = $true
            }
        }
    }
    catch {
        Write-Error "[Keyboard] Operation failed / 操作失敗: $($_.Exception.Message)"
    }
}

# --- タイマーTick（クリップボード監視） ---
<#
.SYNOPSIS
    Handle timer tick for clipboard monitoring
    クリップボード監視用のタイマーティックを処理
#>
function Invoke-TimerTick {
    Update-ClipboardHistory
}

# --- フォームクローズ処理（各種リソース解放） ---
<#
.SYNOPSIS
    Handle form closing event
    フォームクローズイベントを処理
#>
function Invoke-FormClosing {
    [CmdletBinding()]
    param($Sender, $EventArgs)
    
    try {
        Write-Information "Shutting down / 終了処理中..."
        
        if ($null -ne $script:hotkey_process -and -not $script:hotkey_process.HasExited) {
            Write-Information "Stopping Hotkey_Monitor.ps1 / Hotkey_Monitor.ps1を停止中..."
            Stop-Process -Id $script:hotkey_process.Id -Force
        }
        
        if ($null -ne $script:timer) {
            $script:timer.Stop()
            $script:timer.Dispose()
            Write-Information "Clipboard monitoring stopped / クリップボード監視停止"
        }
        
        if ($null -ne $script:pipe_runspace) {
            $script:pipe_runspace.Close()
            $script:pipe_runspace.Dispose()
            Write-Information "Pipe server stopped / パイプサーバー停止"
        }
        
        if ($null -ne $script:notify_icon) {
            $script:notify_icon.Visible = $false
            $script:notify_icon.Dispose()
            Write-Information "Tray icon removed / トレイアイコン削除"
        }
        
        Export-ApplicationData
        Write-Information "Shutdown complete / 終了処理完了"
    }
    catch {
        Write-Error "[Shutdown] Failed / 終了処理失敗: $($_.Exception.Message)"
    }
}

# --- フォームリサイズ処理（最小化時トレイ格納） ---
<#
.SYNOPSIS
    Handle form resize event
    フォームリサイズイベントを処理
#>
function Invoke-FormResize {
    [CmdletBinding()]
    param($Sender, $EventArgs)
    
    if ($script:main_form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
        Hide-ToTray
    }
}

#endregion

#region Main Processing / メイン処理

# --- メインフォーム初期化（ShowInTaskbar=false） ---
<#
.SYNOPSIS
    Initialize main form (FIXED)
    メインフォームを初期化（修正版）
#>
function Initialize-MainForm {
    [CmdletBinding()]
    param()
    
    try {
        # Create main form / メインフォームを作成
        $script:main_form = New-Object System.Windows.Forms.Form
        $script:main_form.Text = "PS Clipboard Manager"
        $script:main_form.Size = New-Object System.Drawing.Size(400, 550)
        $script:main_form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $script:main_form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
        $script:main_form.MaximizeBox = $false
        
        # FIXED: Set ShowInTaskbar to false initially
        # 修正: 初期状態でShowInTaskbarをfalseに設定
        $script:main_form.ShowInTaskbar = $false
        
        # Auto-hide when focus lost / フォーカスを失ったときに自動格納
        $script:main_form.add_Deactivate({
            if ($script:auto_hide_to_tray -and -not $script:is_editing) {
                Start-Sleep -Milliseconds 100
                if (-not $script:main_form.Focused) {
                    Hide-ToTray
                    Write-Verbose "Auto-hidden due to focus loss / フォーカス喪失により自動格納"
                }
            }
        })
        
        # History section / 履歴セクション
        $historyLabel = New-Object System.Windows.Forms.Label
        $historyLabel.Text = "Clipboard History / 履歴 (Max $script:MAX_HISTORY_COUNT)"
        $historyLabel.Location = New-Object System.Drawing.Point(10, 10)
        $historyLabel.Size = New-Object System.Drawing.Size(360, 20)
        $script:main_form.Controls.Add($historyLabel)
        
        $script:listbox_history = New-Object System.Windows.Forms.ListBox
        $script:listbox_history.Location = New-Object System.Drawing.Point(10, 35)
        $script:listbox_history.Size = New-Object System.Drawing.Size(360, 250)
        $script:listbox_history.IntegralHeight = $false
        $script:listbox_history.add_Click({ Invoke-HistoryClick })
        $script:listbox_history.add_KeyDown({ param($s, $e) Invoke-ListBoxKeyDown $s $e })
        
        $script:listbox_history.add_MouseDown({ param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $point = New-Object System.Drawing.Point($e.X, $e.Y)
                $index = $s.IndexFromPoint($point)
                if ($index -eq -1 -or $index -ge $script:clipboard_history.Count) {
                    $s.SelectedIndex = -1
                }
            }
        })
        
        $script:listbox_history.add_MouseUp({ param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                $point = New-Object System.Drawing.Point($e.X, $e.Y)
                $index = $s.IndexFromPoint($point)
                if ($index -ge 0 -and $index -lt $script:clipboard_history.Count) {
                    $s.SelectedIndex = $index
                } else {
                    $s.SelectedIndex = -1
                }
                Show-HistoryContextMenu $s $e
            }
        })
        
        $script:main_form.Controls.Add($script:listbox_history)
        
        # Template section / 定型文セクション
        $templateLabel = New-Object System.Windows.Forms.Label
        $templateLabel.Text = "Templates / 定型文 (Max $script:MAX_TEMPLATE_COUNT)"
        $templateLabel.Location = New-Object System.Drawing.Point(10, 300)
        $templateLabel.Size = New-Object System.Drawing.Size(360, 20)
        $script:main_form.Controls.Add($templateLabel)
        
        $script:listbox_template = New-Object System.Windows.Forms.ListBox
        $script:listbox_template.Location = New-Object System.Drawing.Point(10, 325)
        $script:listbox_template.Size = New-Object System.Drawing.Size(360, 130)
        $script:listbox_template.IntegralHeight = $false
        $script:listbox_template.add_Click({ Invoke-TemplateClick })
        $script:listbox_template.add_KeyDown({ param($s, $e) Invoke-ListBoxKeyDown $s $e })
        
        $script:listbox_template.add_MouseDown({ param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $point = New-Object System.Drawing.Point($e.X, $e.Y)
                $index = $s.IndexFromPoint($point)
                if ($index -eq -1 -or $index -ge $script:template_data.Count) {
                    $s.SelectedIndex = -1
                }
            }
        })
        
        $script:listbox_template.add_MouseUp({ param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                $point = New-Object System.Drawing.Point($e.X, $e.Y)
                $index = $s.IndexFromPoint($point)
                if ($index -ge 0 -and $index -lt $script:template_data.Count) {
                    $s.SelectedIndex = $index
                } else {
                    $s.SelectedIndex = -1
                }
                Show-TemplateContextMenu $s $e
            }
        })
        
        $script:main_form.Controls.Add($script:listbox_template)
        
        # Help label / ヘルプラベル
        $helpLabel = New-Object System.Windows.Forms.Label
        $helpLabel.Text = "Ctrl+Shift+Space | Enter=Set | Tab=Switch | Esc=Hide | Ctrl+M=Context"
        $helpLabel.Location = New-Object System.Drawing.Point(10, 465)
        $helpLabel.Size = New-Object System.Drawing.Size(360, 35)
        $helpLabel.ForeColor = [System.Drawing.Color]::DarkGray
        $script:main_form.Controls.Add($helpLabel)
        
        # Event handlers / イベントハンドラー
        $script:main_form.add_FormClosing({ param($s, $e) Invoke-FormClosing $s $e })
        $script:main_form.add_Resize({ param($s, $e) Invoke-FormResize $s $e })
        
        Write-Information "Main form initialized / メインフォーム初期化完了"
    }
    catch {
        Write-Error "[MainForm] Initialization failed / 初期化失敗: $($_.Exception.Message)"
        throw
    }
}

# --- クリップボード監視開始 ---
<#
.SYNOPSIS
    Start clipboard monitoring
    クリップボード監視を開始
#>
function Start-ClipboardMonitor {
    [CmdletBinding()]
    param()
    
    try {
        $script:timer = New-Object System.Windows.Forms.Timer
        $script:timer.Interval = $script:MONITOR_INTERVAL_MS
        $script:timer.add_Tick({ Invoke-TimerTick })
        $script:timer.Start()
        Write-Information "Clipboard monitoring started / クリップボード監視開始: ${script:MONITOR_INTERVAL_MS}ms"
    }
    catch {
        Write-Error "[Monitor] Failed to start / 開始失敗: $($_.Exception.Message)"
        throw
    }
}

# --- Hotkey_Monitor.ps1プロセス起動 ---
<#
.SYNOPSIS
    Start Hotkey_Monitor.ps1 process
    Hotkey_Monitor.ps1プロセスを起動
#>
function Start-HotkeyProcess {
    [CmdletBinding()]
    param()
    
    try {
        $hotkeyScriptPath = Join-Path $PSScriptRoot "Hotkey_Monitor.ps1"
        
        if (Test-Path $hotkeyScriptPath) {
            Write-Information "Starting Hotkey_Monitor.ps1 / Hotkey_Monitor.ps1を起動中..."
            
            $script:hotkey_process = Start-Process powershell.exe `
                -ArgumentList "-ExecutionPolicy Bypass -File `"$hotkeyScriptPath`"" `
                -WindowStyle Hidden `
                -PassThru
                
            Write-Information "Hotkey_Monitor.ps1 started / 起動完了 (PID: $($script:hotkey_process.Id))"
        } else {
            Write-Warning "[Hotkey] Script not found / スクリプトが見つかりません: $hotkeyScriptPath"
            Write-Warning "Hotkey functionality unavailable / ホットキー機能は利用できません"
        }
    }
    catch {
        Write-Warning "[Hotkey] Failed to start / 起動失敗: $($_.Exception.Message)"
        Write-Warning "Hotkey functionality unavailable / ホットキー機能は利用できません"
    }
}

# --- アプリケーション起動処理（各初期化・監視開始） ---
<#
.SYNOPSIS
    Start application (FIXED)
    アプリケーションを開始（修正版）
#>
function Start-Application {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host " PS Clipboard Manager Starting..." -ForegroundColor Cyan
        Write-Host " PS Clipboard Manager 起動中..." -ForegroundColor Cyan
        Write-Host " Version 2.5 - Type Casting Fixed" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Cyan
        
        # Load data / データを読み込み
        Import-ApplicationData
        
        # Process initial clipboard / 起動時のクリップボードを処理
        $startupClipboard = Get-ClipboardContent
        if (-not [string]::IsNullOrWhiteSpace($startupClipboard)) {
            if ($script:clipboard_history -notcontains $startupClipboard) {
                [void]$script:clipboard_history.Insert(0, $startupClipboard)
                if ($script:clipboard_history.Count -gt $script:MAX_HISTORY_COUNT) {
                    $script:clipboard_history.RemoveRange(
                        $script:MAX_HISTORY_COUNT, 
                        $script:clipboard_history.Count - $script:MAX_HISTORY_COUNT
                    )
                }
                Write-Information "Added startup clipboard to history / 起動時のクリップボードを履歴に追加"
            }
            $script:current_clipboard = $startupClipboard
        }
        
        # Initialize GUI / GUIを初期化
        Initialize-MainForm
        
        # Initialize system tray / システムトレイを初期化
        Initialize-NotifyIcon
        
        # Update displays / 表示を更新
        Update-HistoryDisplay
        Update-TemplateDisplay
        
        # Start pipe server (with initialization delay) / パイプサーバーを起動（初期化遅延付き）
        Start-PipeServer
        
        # Start Hotkey_Monitor.ps1 / Hotkey_Monitor.ps1を起動
        Start-HotkeyProcess
        
        # Start monitoring / 監視を開始
        Start-ClipboardMonitor
        
        Write-Host ""
        Write-Host "Startup Complete! / 起動完了！" -ForegroundColor Green
        Write-Host "History / 履歴: $($script:clipboard_history.Count) | Templates / 定型文: $($script:template_data.Count)"
        Write-Host ""
        Write-Host "【Operations / 操作方法】" -ForegroundColor Yellow
        Write-Host "- Ctrl+Shift+Space: Show window / ウィンドウ表示"
        Write-Host "- Tray icon click: Show window / トレイアイコンクリック"
        Write-Host "- Enter: Set clipboard / クリップボード設定"
        Write-Host "- Tab: Switch lists / リスト切替"
        Write-Host "- Esc: Hide window / ウィンドウを隠す"
        Write-Host "- Ctrl+M: Context menu / コンテキストメニュー"
        Write-Host "========================================" -ForegroundColor Cyan
        
        # Hide to tray initially / 初期状態でトレイに格納
        Hide-ToTray
        
        # Start message loop / メッセージループを開始
        [System.Windows.Forms.Application]::Run($script:main_form)
    }
    catch {
        Write-Error "[Application] Failed to start / 起動失敗: $($_.Exception.Message)"
        throw
    }
}

#endregion

#region Main Execution / メイン実行

# --- アプリケーション開始・例外処理・クリーンアップ ---
try {
    # Start application / アプリケーションを開始
    Start-Application
}
catch {
    $errorMessage = "Critical error occurred / 重大なエラーが発生しました:`n`n$($_.Exception.Message)"
    Write-Error $errorMessage
    
    try {
        [System.Windows.Forms.MessageBox]::Show(
            $errorMessage,
            "PS Clipboard Manager - Error / エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    catch {
        Write-Host "Failed to display GUI / GUI表示失敗"
    }
    
    Read-Host "Press Enter to exit / Enterキーで終了"
}
finally {
    # Cleanup / クリーンアップ
    try {
        if ($null -ne $script:timer) {
            $script:timer.Stop()
            $script:timer.Dispose()
        }
        if ($null -ne $script:notify_icon) {
            $script:notify_icon.Visible = $false
            $script:notify_icon.Dispose()
        }
        if ($null -ne $script:hotkey_process -and -not $script:hotkey_process.HasExited) {
            Stop-Process -Id $script:hotkey_process.Id -Force
        }
        if ($null -ne $script:pipe_runspace) {
            $script:pipe_runspace.Close()
            $script:pipe_runspace.Dispose()
        }
        Write-Information "Cleanup complete / クリーンアップ完了"
    }
    catch {
        Write-Warning "Cleanup error / クリーンアップエラー: $($_.Exception.Message)"
    }
}

#endregion
