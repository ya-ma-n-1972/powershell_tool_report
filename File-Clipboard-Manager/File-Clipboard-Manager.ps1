<#
.SYNOPSIS
    File Clipboard Manager - ファイルをクリップボードに登録するツール

.DESCRIPTION
    複数のファイルをグループ管理し、クリップボードに一括登録できるツールです。
    ファイルパス方式とフェンス付きテキスト方式の2つの登録方法をサポートします。

.NOTES
    Author: ya-man
    Version: 1.1
    Date: 2025-11-03
    Requires: PowerShell 5.1 or later, Windows Forms
#>

# ================================================================================
# アセンブリの読み込み
# ================================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ================================================================================
# グローバル変数
# ================================================================================
$script:config = $null
$script:configPath = Join-Path $PSScriptRoot "config.json"
$script:messageLabel = $null

# ================================================================================
# ヘルパー関数: 表示形式変換
# ================================================================================

<#
.SYNOPSIS
    ファイルパスを表示用フォーマットに変換
.DESCRIPTION
    フルパスを "ファイル名":"フルパス" 形式に変換します。
    同名ファイルの識別を容易にするための表示用フォーマットです。
#>
function ConvertTo-DisplayFormat($filePath) {
    $fileName = [System.IO.Path]::GetFileName($filePath)
    return "`"$fileName`":`"$filePath`""
}

<#
.SYNOPSIS
    表示用フォーマットからファイルパスを抽出
.DESCRIPTION
    "ファイル名":"フルパス" 形式からフルパスを抽出します。
    正規表現で2つ目のダブルクォート内の文字列を取得します。
#>
function ConvertFrom-DisplayFormat($displayText) {
    if ($displayText -match '^"[^"]+":"(.+)"') {
        return $matches[1]
    }
    return $displayText
}

# ================================================================================
# 設定ファイル操作
# ================================================================================

<#
.SYNOPSIS
    設定ファイルの読み込み
.DESCRIPTION
    config.jsonから設定を読み込みます。
    ファイルが存在しない場合はデフォルト設定を作成します。
#>
function Load-Config {
    if (Test-Path $script:configPath) {
        try {
            $json = Get-Content $script:configPath -Raw -Encoding UTF8
            $script:config = $json | ConvertFrom-Json
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "設定ファイルの読み込みに失敗しました。デフォルト設定を使用します。`n`nエラー: $_",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $script:config = @{
                groups = @(
                    @{ name = "グループ1"; files = @() },
                    @{ name = "グループ2"; files = @() },
                    @{ name = "グループ3"; files = @() },
                    @{ name = "グループ4"; files = @() }
                )
            }
        }
    }
    else {
        # デフォルト設定の作成
        $script:config = @{
            groups = @(
                @{ name = "グループ1"; files = @() },
                @{ name = "グループ2"; files = @() },
                @{ name = "グループ3"; files = @() },
                @{ name = "グループ4"; files = @() }
            )
        }
        Save-Config
    }
}

<#
.SYNOPSIS
    設定ファイルの保存
.DESCRIPTION
    現在の$script:configをJSON形式でconfig.jsonに保存します。
#>
function Save-Config {
    try {
        $json = $script:config | ConvertTo-Json -Depth 10
        $json | Set-Content $script:configPath -Encoding UTF8
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "設定ファイルの保存に失敗しました。`n`nエラー: $_",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# ================================================================================
# ファイルサイズ計算
# ================================================================================

<#
.SYNOPSIS
    グループ内のファイル合計サイズを計算
.DESCRIPTION
    指定グループの全ファイルサイズを合計し、MB単位で返します。
    存在しないファイルはスキップします。
#>
function Get-GroupSize($groupIndex) {
    $totalBytes = 0
    foreach ($filePath in $script:config.groups[$groupIndex].files) {
        if (Test-Path $filePath) {
            $totalBytes += (Get-Item $filePath).Length
        }
    }
    $totalMB = $totalBytes / 1MB
    return "{0:N1} MB" -f $totalMB
}

# ================================================================================
# BOM付きUTF-8変換
# ================================================================================

<#
.SYNOPSIS
    選択ファイルをBOM付きUTF-8に変換
.DESCRIPTION
    指定されたファイルをBOM付きUTF-8エンコーディングで再保存します。
    バックアップは作成されないため注意が必要です。
#>
function Convert-ToBomUtf8($groupIndex) {
    # このバージョンでは使用されていません（設定ウィンドウから直接処理）
}

# ================================================================================
# クリップボード登録（ファイルパス方式）
# ================================================================================

<#
.SYNOPSIS
    ファイルパスをクリップボードに登録
.DESCRIPTION
    指定グループのファイルパスをクリップボードに登録します。
    エクスプローラーに貼り付け可能な形式（FileDropList）で登録されます。
#>
function Set-ClipboardFiles($groupIndex) {
    $files = $script:config.groups[$groupIndex].files
    $groupName = $script:config.groups[$groupIndex].name

    $fileCollection = New-Object System.Collections.Specialized.StringCollection

    foreach ($filePath in $files) {
        if (Test-Path $filePath) {
            [void]$fileCollection.Add($filePath)
        }
    }

    [System.Windows.Forms.Clipboard]::SetFileDropList($fileCollection)
    Show-Message "クリップボードに登録しました（$groupName、$($fileCollection.Count)件）"
}

# ================================================================================
# クリップボード登録（個別ファイルのフェンス付きテキスト方式）
# ================================================================================

<#
.SYNOPSIS
    個別ファイルをフェンス付きテキストとしてクリップボードに登録
.DESCRIPTION
    ファイル内容をMarkdownコードフェンス（```）で囲んでクリップボードに登録します。
    AIチャットツールへのコード送信に便利です。
#>
function Set-ClipboardTextWithFence($filePath) {
    if (-not (Test-Path $filePath)) {
        Show-Message "ファイルが見つかりません: $filePath"
        return
    }

    try {
        $content = Get-Content $filePath -Raw -Encoding UTF8
        $fence = '```'
        $text = $fence + "`n" + $content + "`n" + $fence

        [System.Windows.Forms.Clipboard]::SetText($text)

        $fileName = [System.IO.Path]::GetFileName($filePath)
        Show-Message "$fileName にフェンスを追加してクリップボードに登録しました"
    }
    catch {
        Show-Message "ファイルの読み込みに失敗しました: $_"
    }
}

<#
.SYNOPSIS
    クリップボードの内容をフェンス付きテキストに変換
.DESCRIPTION
    クリップボード内のファイルまたはテキストをフェンス付きテキストに変換します。
    ファイルとテキストの両方に対応し、設定ウィンドウへの登録なしで使用できます。
#>
function Set-ClipboardFileWithFence {
    # クリップボードからファイルリストを取得
    $fileList = [System.Windows.Forms.Clipboard]::GetFileDropList()

    # ファイル形式が入っている場合
    if ($null -ne $fileList -and $fileList.Count -gt 0) {
        # エラーチェック1: 複数ファイルでないか
        if ($fileList.Count -gt 1) {
            [System.Windows.Forms.MessageBox]::Show(
                "複数ファイルには対応していません。1つのファイルを選択してください",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $filePath = $fileList[0]

        # エラーチェック2: ファイルが存在するか
        if (-not (Test-Path $filePath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "ファイルが見つかりません",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        # エラーチェック3: テキストファイルか（拡張子チェック）
        $textExtensions = @('.txt', '.ps1', '.psm1', '.psd1', '.cs', '.vb', '.js', '.ts',
                            '.html', '.htm', '.xml', '.json', '.css', '.md', '.log',
                            '.csv', '.ini', '.config', '.bat', '.cmd', '.py', '.rb',
                            '.java', '.c', '.cpp', '.h', '.hpp', '.php', '.sh')

        $ext = [System.IO.Path]::GetExtension($filePath).ToLower()

        if ($textExtensions -notcontains $ext) {
            [System.Windows.Forms.MessageBox]::Show(
                "テキストファイルではありません",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # エラーチェック4: ファイルサイズが30MB以下か
        $fileSize = (Get-Item $filePath).Length
        $maxSize = 30MB

        if ($fileSize -gt $maxSize) {
            [System.Windows.Forms.MessageBox]::Show(
                "ファイルサイズが30MBを超えています",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # ファイル内容を読み込み＋フェンスで囲む
        try {
            $content = Get-Content $filePath -Raw -Encoding UTF8
            $fence = '```'
            $text = $fence + "`n" + $content + "`n" + $fence

            # クリップボードに再登録
            [System.Windows.Forms.Clipboard]::SetText($text)

            # 成功メッセージ
            $fileName = [System.IO.Path]::GetFileName($filePath)
            Show-Message "$fileName にフェンスを追加してクリップボードに登録しました"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "ファイルの読み込みに失敗しました: $_",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        return
    }

    # テキスト形式が入っている場合
    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        try {
            $content = [System.Windows.Forms.Clipboard]::GetText()

            # 空のテキストチェック
            if ([string]::IsNullOrWhiteSpace($content)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "クリップボードのテキストが空です",
                    "エラー",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }

            # フェンスで囲む
            $fence = '```'
            $text = $fence + "`n" + $content + "`n" + $fence

            # クリップボードに再登録
            [System.Windows.Forms.Clipboard]::SetText($text)

            # 成功メッセージ
            Show-Message "テキストにフェンスを追加してクリップボードに登録しました"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "テキストの処理に失敗しました: $_",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        return
    }

    # ファイルもテキストも入っていない場合
    [System.Windows.Forms.MessageBox]::Show(
        "クリップボードにファイルまたはテキストがコピーされていません",
        "エラー",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
}

# ================================================================================
# クリップボード登録（グループ全体のフェンス付きテキスト方式）
# ================================================================================

<#
.SYNOPSIS
    グループ全体をフェンス付きテキストとしてクリップボードに登録
.DESCRIPTION
    グループ内の全ファイルをフェンス付きテキストとして連結し、
    クリップボードに登録します。各ファイルの先頭にファイル名を記載します。
#>
function Set-ClipboardGroupTextWithFence($groupIndex) {
    $files = $script:config.groups[$groupIndex].files
    $groupName = $script:config.groups[$groupIndex].name
    $allText = ""
    $fence = '```'
    $processedCount = 0

    foreach ($filePath in $files) {
        if (-not (Test-Path $filePath)) {
            continue
        }

        try {
            $content = Get-Content $filePath -Raw -Encoding UTF8
            $fileName = [System.IO.Path]::GetFileName($filePath)

            # ファイル名 + コードフェンス + 内容 + コードフェンス + 空行
            $allText += $fileName + "`n"
            $allText += $fence + "`n"
            $allText += $content + "`n"
            $allText += $fence + "`n`n"

            $processedCount++
        }
        catch {
            # エラーは無視して次のファイルへ
        }
    }

    if ($allText.Length -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText($allText)
        Show-Message "$groupName の全ファイル（${processedCount}件）をフェンス付きでクリップボードに登録しました"
    }
    else {
        Show-Message "登録可能なファイルがありません"
    }
}

# ================================================================================
# メッセージ表示
# ================================================================================

<#
.SYNOPSIS
    メッセージ表示エリアにメッセージを表示
.DESCRIPTION
    メインウィンドウ下部のラベルにメッセージを表示します。
#>
function Show-Message($message) {
    if ($script:messageLabel -ne $null) {
        $script:messageLabel.Text = $message
    }
}

# ================================================================================
# ファイル移動
# ================================================================================

<#
.SYNOPSIS
    ファイルを別グループへ移動
.DESCRIPTION
    指定されたファイルをソースグループからターゲットグループへ移動します。
    ターゲットグループが20ファイル上限に達している場合はエラーを表示します。
#>
function Move-FileToGroup($sourceGroupIndex, $fileIndex, $targetGroupIndex) {
    if ($script:config.groups[$targetGroupIndex].files.Count -ge 20) {
        Show-Message "移動先グループは既に20ファイル登録されています"
        return $false
    }

    $filePath = $script:config.groups[$sourceGroupIndex].files[$fileIndex]

    # ソースから削除
    $script:config.groups[$sourceGroupIndex].files = @(
        $script:config.groups[$sourceGroupIndex].files | Where-Object { $_ -ne $filePath }
    )

    # ターゲットに追加
    $script:config.groups[$targetGroupIndex].files += $filePath

    $targetGroupName = $script:config.groups[$targetGroupIndex].name
    Show-Message "ファイルを $targetGroupName へ移動しました"

    return $true
}

# ================================================================================
# メインウィンドウ構築
# ================================================================================

<#
.SYNOPSIS
    メインウィンドウの表示
.DESCRIPTION
    4つのグループを管理するメインUIを構築・表示します。
#>
function Show-MainWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "File Clipboard Manager"
    $form.Size = New-Object System.Drawing.Size(450, 430)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # グループごとのUI要素を格納する配列
    $groupLabels = @()
    $sizeLabels = @()

    # 4グループ分のUIを作成
    for ($i = 0; $i -lt 4; $i++) {
        $yPos = 10 + ($i * 70)

        # グループ名ラベル
        $lblGroup = New-Object System.Windows.Forms.Label
        $lblGroup.Text = $script:config.groups[$i].name
        $lblGroup.Location = New-Object System.Drawing.Point(10, $yPos)
        $lblGroup.Size = New-Object System.Drawing.Size(200, 20)
        $form.Controls.Add($lblGroup)
        $groupLabels += $lblGroup

        # ファイルサイズラベル
        $lblSize = New-Object System.Windows.Forms.Label
        $lblSize.Text = Get-GroupSize $i
        $lblSize.Location = New-Object System.Drawing.Point(320, $yPos)
        $lblSize.Size = New-Object System.Drawing.Size(100, 20)
        $lblSize.TextAlign = [System.Drawing.ContentAlignment]::TopRight
        $form.Controls.Add($lblSize)
        $sizeLabels += $lblSize

        # 設定ボタン
        $btnConfig = New-Object System.Windows.Forms.Button
        $btnConfig.Text = "設定"
        $btnConfig.Location = New-Object System.Drawing.Point(10, ($yPos + 25))
        $btnConfig.Size = New-Object System.Drawing.Size(100, 25)
        $btnConfig.Tag = $i
        $btnConfig.Add_Click({
            Show-ConfigWindow $this.Tag $groupLabels $sizeLabels
        })
        $form.Controls.Add($btnConfig)

        # クリップボードへボタン（ファイルパス方式）
        $btnClipboard = New-Object System.Windows.Forms.Button
        $btnClipboard.Text = "クリップボードへ"
        $btnClipboard.Location = New-Object System.Drawing.Point(120, ($yPos + 25))
        $btnClipboard.Size = New-Object System.Drawing.Size(130, 25)
        $btnClipboard.Tag = $i
        $btnClipboard.Add_Click({
            Set-ClipboardFiles $this.Tag
        })
        $form.Controls.Add($btnClipboard)

        # フェンス付きでクリップボードボタン（テキスト方式）
        $btnFence = New-Object System.Windows.Forms.Button
        $btnFence.Text = "フェンス付きでクリップボード"
        $btnFence.Location = New-Object System.Drawing.Point(260, ($yPos + 25))
        $btnFence.Size = New-Object System.Drawing.Size(170, 25)
        $btnFence.Tag = $i
        $btnFence.Add_Click({
            Set-ClipboardGroupTextWithFence $this.Tag
        })
        $form.Controls.Add($btnFence)
    }

    # クリップボードにフェンス追加ボタン
    $btnClipboardFence = New-Object System.Windows.Forms.Button
    $btnClipboardFence.Text = "クリップボードにフェンス追加"
    $btnClipboardFence.Location = New-Object System.Drawing.Point(10, 290)
    $btnClipboardFence.Size = New-Object System.Drawing.Size(250, 30)
    $btnClipboardFence.Add_Click({
        Set-ClipboardFileWithFence
    })
    $form.Controls.Add($btnClipboardFence)

    # メッセージ表示エリア
    $script:messageLabel = New-Object System.Windows.Forms.Label
    $script:messageLabel.Text = ""
    $script:messageLabel.Location = New-Object System.Drawing.Point(10, 340)
    $script:messageLabel.Size = New-Object System.Drawing.Size(420, 30)
    $script:messageLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $script:messageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($script:messageLabel)

    [void]$form.ShowDialog()
}

# ================================================================================
# 設定ウィンドウ構築
# ================================================================================

<#
.SYNOPSIS
    設定ウィンドウの表示
.DESCRIPTION
    指定グループの設定ウィンドウを表示します。
    ファイルの追加・削除・移動・BOM変換などの操作が可能です。
#>
function Show-ConfigWindow($groupIndex, $groupLabels, $sizeLabels) {
    $configForm = New-Object System.Windows.Forms.Form
    $configForm.Text = "$($script:config.groups[$groupIndex].name) - 設定"
    $configForm.Size = New-Object System.Drawing.Size(650, 550)
    $configForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $configForm.MaximizeBox = $false
    $configForm.MinimizeBox = $false

    # メインウィンドウの右側に配置
    $configForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $configForm.Location = New-Object System.Drawing.Point(($form.Location.X + $form.Width + 10), $form.Location.Y)

    # グループ名ラベル
    $lblGroupName = New-Object System.Windows.Forms.Label
    $lblGroupName.Text = "グループ名:"
    $lblGroupName.Location = New-Object System.Drawing.Point(10, 15)
    $lblGroupName.Size = New-Object System.Drawing.Size(100, 20)
    $configForm.Controls.Add($lblGroupName)

    # グループ名入力
    $txtGroupName = New-Object System.Windows.Forms.TextBox
    $txtGroupName.Text = $script:config.groups[$groupIndex].name
    $txtGroupName.Location = New-Object System.Drawing.Point(120, 13)
    $txtGroupName.Size = New-Object System.Drawing.Size(500, 20)
    $configForm.Controls.Add($txtGroupName)

    # ファイルリストボックス
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 45)
    $listBox.Size = New-Object System.Drawing.Size(610, 300)
    $listBox.HorizontalScrollbar = $true

    # ファイルリストを表示用フォーマットで追加
    foreach ($filePath in $script:config.groups[$groupIndex].files) {
        $displayText = ConvertTo-DisplayFormat $filePath
        [void]$listBox.Items.Add($displayText)
    }
    $configForm.Controls.Add($listBox)

    # 登録数ラベル
    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "$($listBox.Items.Count)/20"
    $lblCount.Location = New-Object System.Drawing.Point(10, 355)
    $lblCount.Size = New-Object System.Drawing.Size(100, 20)
    $configForm.Controls.Add($lblCount)

    # リストボックス更新用関数
    $updateListBox = {
        $listBox.Items.Clear()
        foreach ($filePath in $script:config.groups[$groupIndex].files) {
            $displayText = ConvertTo-DisplayFormat $filePath
            [void]$listBox.Items.Add($displayText)
        }
        $lblCount.Text = "$($listBox.Items.Count)/20"
    }

    # ファイル追加ボタン
    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "ファイル追加"
    $btnAdd.Location = New-Object System.Drawing.Point(10, 385)
    $btnAdd.Size = New-Object System.Drawing.Size(100, 30)
    $btnAdd.Add_Click({
        if ($script:config.groups[$groupIndex].files.Count -ge 20) {
            Show-Message "最大20ファイルまでです"
            return
        }

        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "ファイルを選択"
        $openFileDialog.Filter = "すべてのファイル (*.*)|*.*"

        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFile = $openFileDialog.FileName

            if ($script:config.groups[$groupIndex].files -contains $selectedFile) {
                Show-Message "既に登録されています"
                return
            }

            $script:config.groups[$groupIndex].files += $selectedFile
            & $updateListBox
            Show-Message "ファイルを追加しました"
        }
    })
    $configForm.Controls.Add($btnAdd)

    # 選択ファイルを削除ボタン
    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "選択ファイルを削除"
    $btnRemove.Location = New-Object System.Drawing.Point(120, 385)
    $btnRemove.Size = New-Object System.Drawing.Size(120, 30)
    $btnRemove.Add_Click({
        if ($listBox.SelectedIndex -lt 0) {
            Show-Message "ファイルを選択してください"
            return
        }

        $selectedDisplayText = $listBox.SelectedItem
        $selectedFilePath = ConvertFrom-DisplayFormat $selectedDisplayText

        $script:config.groups[$groupIndex].files = @(
            $script:config.groups[$groupIndex].files | Where-Object { $_ -ne $selectedFilePath }
        )

        & $updateListBox
        Show-Message "ファイルを削除しました"
    })
    $configForm.Controls.Add($btnRemove)

    # 全て削除ボタン
    $btnRemoveAll = New-Object System.Windows.Forms.Button
    $btnRemoveAll.Text = "全て削除"
    $btnRemoveAll.Location = New-Object System.Drawing.Point(250, 385)
    $btnRemoveAll.Size = New-Object System.Drawing.Size(100, 30)
    $btnRemoveAll.Add_Click({
        if ($script:config.groups[$groupIndex].files.Count -eq 0) {
            Show-Message "削除するファイルがありません"
            return
        }

        $result = [System.Windows.Forms.MessageBox]::Show(
            "グループ内の全てのファイル（$($script:config.groups[$groupIndex].files.Count)件）を削除しますか？",
            "確認",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:config.groups[$groupIndex].files = @()
            & $updateListBox
            Show-Message "全てのファイルを削除しました"
        }
    })
    $configForm.Controls.Add($btnRemoveAll)

    # BOM変換ボタン
    $btnBom = New-Object System.Windows.Forms.Button
    $btnBom.Text = "BOM変換"
    $btnBom.Location = New-Object System.Drawing.Point(360, 385)
    $btnBom.Size = New-Object System.Drawing.Size(100, 30)
    $btnBom.Add_Click({
        if ($listBox.SelectedIndex -lt 0) {
            Show-Message "ファイルを選択してください"
            return
        }

        $selectedDisplayText = $listBox.SelectedItem
        $selectedFilePath = ConvertFrom-DisplayFormat $selectedDisplayText

        if (-not (Test-Path $selectedFilePath)) {
            Show-Message "ファイルが見つかりません"
            return
        }

        # テキストファイルかチェック
        $textExtensions = @('.txt', '.ps1', '.psm1', '.psd1', '.cs', '.vb', '.js', '.ts',
                            '.html', '.htm', '.xml', '.json', '.css', '.md', '.log',
                            '.csv', '.ini', '.config', '.bat', '.cmd', '.py', '.rb',
                            '.java', '.c', '.cpp', '.h', '.hpp', '.php', '.sh')

        $ext = [System.IO.Path]::GetExtension($selectedFilePath).ToLower()

        if ($textExtensions -notcontains $ext) {
            Show-Message "テキストファイルではありません"
            return
        }

        $fileName = [System.IO.Path]::GetFileName($selectedFilePath)
        $result = [System.Windows.Forms.MessageBox]::Show(
            "$fileName をBOM付きUTF-8に変換しますか？`n`n※バックアップは作成されません",
            "確認",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $content = Get-Content $selectedFilePath -Raw
                $utf8WithBom = New-Object System.Text.UTF8Encoding($true)
                [System.IO.File]::WriteAllText($selectedFilePath, $content, $utf8WithBom)
                Show-Message "$fileName : BOM変換終了"
            }
            catch {
                Show-Message "変換に失敗しました: $_"
            }
        }
    })
    $configForm.Controls.Add($btnBom)

    # 別グループへ移動ボタン
    $btnMove = New-Object System.Windows.Forms.Button
    $btnMove.Text = "別グループへ移動"
    $btnMove.Location = New-Object System.Drawing.Point(470, 385)
    $btnMove.Size = New-Object System.Drawing.Size(130, 30)
    $btnMove.Add_Click({
        if ($listBox.SelectedIndex -lt 0) {
            Show-Message "ファイルを選択してください"
            return
        }

        # 移動先選択ダイアログ
        $moveForm = New-Object System.Windows.Forms.Form
        $moveForm.Text = "移動先グループを選択"
        $moveForm.Size = New-Object System.Drawing.Size(300, 200)
        $moveForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $moveForm.MaximizeBox = $false
        $moveForm.MinimizeBox = $false
        $moveForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent

        $lblMove = New-Object System.Windows.Forms.Label
        $lblMove.Text = "移動先グループ:"
        $lblMove.Location = New-Object System.Drawing.Point(10, 20)
        $lblMove.Size = New-Object System.Drawing.Size(100, 20)
        $moveForm.Controls.Add($lblMove)

        $comboBox = New-Object System.Windows.Forms.ComboBox
        $comboBox.Location = New-Object System.Drawing.Point(10, 50)
        $comboBox.Size = New-Object System.Drawing.Size(260, 20)
        $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

        for ($i = 0; $i -lt 4; $i++) {
            if ($i -ne $groupIndex) {
                [void]$comboBox.Items.Add("グループ$($i+1): $($script:config.groups[$i].name)")
            }
        }
        $comboBox.SelectedIndex = 0
        $moveForm.Controls.Add($comboBox)

        $btnMoveOk = New-Object System.Windows.Forms.Button
        $btnMoveOk.Text = "移動"
        $btnMoveOk.Location = New-Object System.Drawing.Point(60, 100)
        $btnMoveOk.Size = New-Object System.Drawing.Size(80, 30)
        $btnMoveOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $moveForm.Controls.Add($btnMoveOk)

        $btnMoveCancel = New-Object System.Windows.Forms.Button
        $btnMoveCancel.Text = "キャンセル"
        $btnMoveCancel.Location = New-Object System.Drawing.Point(150, 100)
        $btnMoveCancel.Size = New-Object System.Drawing.Size(80, 30)
        $btnMoveCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $moveForm.Controls.Add($btnMoveCancel)

        $moveForm.AcceptButton = $btnMoveOk
        $moveForm.CancelButton = $btnMoveCancel

        if ($moveForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # 選択されたグループインデックスを計算
            $targetGroupIndex = $comboBox.SelectedIndex
            if ($targetGroupIndex -ge $groupIndex) {
                $targetGroupIndex++
            }

            if (Move-FileToGroup $groupIndex $listBox.SelectedIndex $targetGroupIndex) {
                & $updateListBox
            }
        }
    })
    $configForm.Controls.Add($btnMove)

    # フェンス追加ボタン
    $btnAddFence = New-Object System.Windows.Forms.Button
    $btnAddFence.Text = "フェンス追加"
    $btnAddFence.Location = New-Object System.Drawing.Point(10, 425)
    $btnAddFence.Size = New-Object System.Drawing.Size(100, 30)
    $btnAddFence.Add_Click({
        if ($listBox.SelectedIndex -lt 0) {
            Show-Message "ファイルを選択してください"
            return
        }

        $selectedDisplayText = $listBox.SelectedItem
        $selectedFilePath = ConvertFrom-DisplayFormat $selectedDisplayText

        Set-ClipboardTextWithFence $selectedFilePath
    })
    $configForm.Controls.Add($btnAddFence)

    # OKボタン
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(400, 465)
    $btnOk.Size = New-Object System.Drawing.Size(100, 30)
    $btnOk.Add_Click({
        $script:config.groups[$groupIndex].name = $txtGroupName.Text
        Save-Config

        # メインウィンドウのラベルを更新
        $groupLabels[$groupIndex].Text = $script:config.groups[$groupIndex].name
        $sizeLabels[$groupIndex].Text = Get-GroupSize $groupIndex

        $configForm.Close()
    })
    $configForm.Controls.Add($btnOk)

    # キャンセルボタン
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "キャンセル"
    $btnCancel.Location = New-Object System.Drawing.Point(510, 465)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Add_Click({
        $configForm.Close()
    })
    $configForm.Controls.Add($btnCancel)

    $configForm.AcceptButton = $btnOk
    $configForm.CancelButton = $btnCancel

    [void]$configForm.ShowDialog()
}

# ================================================================================
# メイン処理
# ================================================================================

# 設定ファイルの読み込み
Load-Config

# メインウィンドウの表示
Show-MainWindow
