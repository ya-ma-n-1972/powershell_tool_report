# FileTimeUpdater.ps1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# グローバル変数
$global:selected_file_path = ""
$global:original_creation_time = ""
$global:original_last_write_time = ""
$global:original_last_access_time = ""

# ファイル日時情報取得関数
function Get-FileTimeInfo {
   param([string]$file_path)
   
   try {
       if (-not (Test-Path $file_path)) {
           throw "ファイルが見つかりません"
       }
       
       $file_info = Get-Item $file_path
       
       return @{
           CreationTime = $file_info.CreationTime.ToString("yyyy/MM/dd HH:mm:ss")
           LastWriteTime = $file_info.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss")
           LastAccessTime = $file_info.LastAccessTime.ToString("yyyy/MM/dd HH:mm:ss")
       }
   }
   catch {
       throw "ファイル情報の取得に失敗しました: $($_.Exception.Message)"
   }
}

# ファイル日時情報更新関数
function Set-FileTimeInfo {
   param(
       [string]$file_path,
       [string]$creation_time,
       [string]$last_write_time,
       [string]$last_access_time
   )
   
   try {
       # 読み取り専用チェック
       $file_info = Get-Item $file_path
       if ($file_info.IsReadOnly) {
           throw "読み取り専用ファイルのため更新できません"
       }
       
       # 日時変換
       $creation_datetime = [DateTime]::ParseExact($creation_time, "yyyy/MM/dd HH:mm:ss", $null)
       $write_datetime = [DateTime]::ParseExact($last_write_time, "yyyy/MM/dd HH:mm:ss", $null)
       $access_datetime = [DateTime]::ParseExact($last_access_time, "yyyy/MM/dd HH:mm:ss", $null)
       
       # ファイル日時更新
       $file_info.CreationTime = $creation_datetime
       $file_info.LastWriteTime = $write_datetime
       $file_info.LastAccessTime = $access_datetime
       
       return $true
   }
   catch [System.UnauthorizedAccessException] {
       throw "ファイルへのアクセス権限がありません"
   }
   catch [System.IO.IOException] {
       throw "ファイルが使用中です（SharePoint同期中の可能性があります）"
   }
   catch [System.FormatException] {
       throw "日時フォーマットが正しくありません"
   }
   catch {
       throw "ファイル更新に失敗しました: $($_.Exception.Message)"
   }
}

# 日時フォーマット検証関数
function Validate-DateTime {
   param([string]$datetime_string)
   
   if ([string]::IsNullOrWhiteSpace($datetime_string)) {
       return $false, "日時が入力されていません"
   }
   
   try {
       [DateTime]::ParseExact($datetime_string, "yyyy/MM/dd HH:mm:ss", $null) | Out-Null
       return $true, ""
   }
   catch {
       return $false, "日時フォーマットが正しくありません（yyyy/MM/dd HH:mm:ss）"
   }
}

# エラーメッセージ表示関数
function Show-ErrorMessage {
   param([string]$message)
   
   [System.Windows.Forms.MessageBox]::Show(
       $message, 
       "エラー", 
       [System.Windows.Forms.MessageBoxButtons]::OK, 
       [System.Windows.Forms.MessageBoxIcon]::Error
   )
}

# 成功メッセージ表示関数
function Show-SuccessMessage {
   param([string]$message)
   
   [System.Windows.Forms.MessageBox]::Show(
       $message, 
       "完了", 
       [System.Windows.Forms.MessageBoxButtons]::OK, 
       [System.Windows.Forms.MessageBoxIcon]::Information
   )
}

# フォーム初期化関数
function Clear-Form {
   $global:selected_file_path = ""
   $pathTextBox.Text = "ファイルを選択してください..."
   $creationTimeTextBox.Text = ""
   $lastWriteTimeTextBox.Text = ""
   $lastAccessTimeTextBox.Text = ""
   $creationTimeTextBox.Enabled = $false
   $lastWriteTimeTextBox.Enabled = $false
   $lastAccessTimeTextBox.Enabled = $false
   $applyButton.Enabled = $false
}

# 日時表示更新関数
function Update-TimeDisplay {
   param([hashtable]$time_info)
   
   $global:original_creation_time = $time_info.CreationTime
   $global:original_last_write_time = $time_info.LastWriteTime
   $global:original_last_access_time = $time_info.LastAccessTime
   
   $creationTimeTextBox.Text = $time_info.CreationTime
   $lastWriteTimeTextBox.Text = $time_info.LastWriteTime
   $lastAccessTimeTextBox.Text = $time_info.LastAccessTime
   
   $creationTimeTextBox.Enabled = $true
   $lastWriteTimeTextBox.Enabled = $true
   $lastAccessTimeTextBox.Enabled = $true
   $applyButton.Enabled = $true
}

# ファイル選択時処理関数
function Handle-FileSelection {
   $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
   $openFileDialog.Filter = "すべてのファイル(*.*)|*.*"
   $openFileDialog.Title = "ファイルを選択してください"
   
   if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
       try {
           $global:selected_file_path = $openFileDialog.FileName
           $pathTextBox.Text = $global:selected_file_path
           $pathTextBox.SelectAll()
           $pathTextBox.Focus()
           
           # ファイル情報を取得して表示
           $time_info = Get-FileTimeInfo -file_path $global:selected_file_path
           Update-TimeDisplay -time_info $time_info
           
       }
       catch {
           Show-ErrorMessage -message $_.Exception.Message
           Clear-Form
       }
   }
}

# 適用ボタン処理関数
function Handle-ApplyButton {
   try {
       # 入力値検証
       $creation_valid, $creation_message = Validate-DateTime -datetime_string $creationTimeTextBox.Text
       if (-not $creation_valid) {
           Show-ErrorMessage -message "作成日時: $creation_message"
           $creationTimeTextBox.Text = $global:original_creation_time
           return
       }
       
       $write_valid, $write_message = Validate-DateTime -datetime_string $lastWriteTimeTextBox.Text
       if (-not $write_valid) {
           Show-ErrorMessage -message "更新日時: $write_message"
           $lastWriteTimeTextBox.Text = $global:original_last_write_time
           return
       }
       
       $access_valid, $access_message = Validate-DateTime -datetime_string $lastAccessTimeTextBox.Text
       if (-not $access_valid) {
           Show-ErrorMessage -message "アクセス日時: $access_message"
           $lastAccessTimeTextBox.Text = $global:original_last_access_time
           return
       }
       
       # ファイル更新実行
       $result = Set-FileTimeInfo -file_path $global:selected_file_path `
                                -creation_time $creationTimeTextBox.Text `
                                -last_write_time $lastWriteTimeTextBox.Text `
                                -last_access_time $lastAccessTimeTextBox.Text
       
       if ($result) {
           Show-SuccessMessage -message "ファイル情報を更新しました"
       }
       
   }
   catch {
       Show-ErrorMessage -message $_.Exception.Message
       
       # エラー種別による処理分岐
       if ($_.Exception.Message -match "読み取り専用|アクセス権限") {
           Clear-Form
       }
       # SharePoint同期中などの場合は元の状態を維持（何もしない）
   }
}

# メインフォーム作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "ファイル日時更新ツール"
$form.Size = New-Object System.Drawing.Size(800,400)
$form.StartPosition = "CenterScreen"

# ファイル選択ボタン
$selectButton = New-Object System.Windows.Forms.Button
$selectButton.Location = New-Object System.Drawing.Point(20,20)
$selectButton.Size = New-Object System.Drawing.Size(120,30)
$selectButton.Text = "ファイル選択"

# パス表示用テキストボックス
$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(20,60)
$pathTextBox.Size = New-Object System.Drawing.Size(740,80)
$pathTextBox.Multiline = $true
$pathTextBox.ScrollBars = "Vertical"
$pathTextBox.ReadOnly = $true
$pathTextBox.Text = "ファイルを選択してください..."

# 作成日時ラベル
$creationTimeLabel = New-Object System.Windows.Forms.Label
$creationTimeLabel.Location = New-Object System.Drawing.Point(20,160)
$creationTimeLabel.Size = New-Object System.Drawing.Size(100,20)
$creationTimeLabel.Text = "作成日時:"

# 作成日時テキストボックス
$creationTimeTextBox = New-Object System.Windows.Forms.TextBox
$creationTimeTextBox.Location = New-Object System.Drawing.Point(130,160)
$creationTimeTextBox.Size = New-Object System.Drawing.Size(200,20)
$creationTimeTextBox.Enabled = $false

# 更新日時ラベル
$lastWriteTimeLabel = New-Object System.Windows.Forms.Label
$lastWriteTimeLabel.Location = New-Object System.Drawing.Point(20,190)
$lastWriteTimeLabel.Size = New-Object System.Drawing.Size(100,20)
$lastWriteTimeLabel.Text = "更新日時:"

# 更新日時テキストボックス
$lastWriteTimeTextBox = New-Object System.Windows.Forms.TextBox
$lastWriteTimeTextBox.Location = New-Object System.Drawing.Point(130,190)
$lastWriteTimeTextBox.Size = New-Object System.Drawing.Size(200,20)
$lastWriteTimeTextBox.Enabled = $false

# アクセス日時ラベル
$lastAccessTimeLabel = New-Object System.Windows.Forms.Label
$lastAccessTimeLabel.Location = New-Object System.Drawing.Point(20,220)
$lastAccessTimeLabel.Size = New-Object System.Drawing.Size(100,20)
$lastAccessTimeLabel.Text = "アクセス日時:"

# アクセス日時テキストボックス
$lastAccessTimeTextBox = New-Object System.Windows.Forms.TextBox
$lastAccessTimeTextBox.Location = New-Object System.Drawing.Point(130,220)
$lastAccessTimeTextBox.Size = New-Object System.Drawing.Size(200,20)
$lastAccessTimeTextBox.Enabled = $false

# 適用ボタン
$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Location = New-Object System.Drawing.Point(20,260)
$applyButton.Size = New-Object System.Drawing.Size(100,30)
$applyButton.Text = "適用"
$applyButton.Enabled = $false

# イベントハンドラー設定
$selectButton.Add_Click({ Handle-FileSelection })
$applyButton.Add_Click({ Handle-ApplyButton })

# コントロールの追加
$form.Controls.AddRange(@(
   $selectButton,
   $pathTextBox,
   $creationTimeLabel,
   $creationTimeTextBox,
   $lastWriteTimeLabel,
   $lastWriteTimeTextBox,
   $lastAccessTimeLabel,
   $lastAccessTimeTextBox,
   $applyButton
))

# フォームの表示
[void]$form.ShowDialog()
