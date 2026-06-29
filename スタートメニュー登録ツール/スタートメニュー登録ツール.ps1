Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- メインフォーム ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "スタートメニューに追加"
$form.Size = New-Object System.Drawing.Size(520, 490)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)

$y = 20

# --- ショートカット名 ---
$lblName = New-Object System.Windows.Forms.Label
$lblName.Text = "ショートカット名："
$lblName.Location = New-Object System.Drawing.Point(20, $y)
$lblName.AutoSize = $true
$form.Controls.Add($lblName)

$y += 22
$txtName = New-Object System.Windows.Forms.TextBox
$txtName.Location = New-Object System.Drawing.Point(20, $y)
$txtName.Size = New-Object System.Drawing.Size(460, 25)
$form.Controls.Add($txtName)

$y += 40

# --- ターゲットパス ---
$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "ターゲット（.exe / .pyw / .ps1 / .bat など）："
$lblTarget.Location = New-Object System.Drawing.Point(20, $y)
$lblTarget.AutoSize = $true
$form.Controls.Add($lblTarget)

$y += 22
$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(20, $y)
$txtTarget.Size = New-Object System.Drawing.Size(370, 25)
$form.Controls.Add($txtTarget)

$btnTarget = New-Object System.Windows.Forms.Button
$btnTarget.Text = "参照..."
$btnTarget.Location = New-Object System.Drawing.Point(400, ($y - 1))
$btnTarget.Size = New-Object System.Drawing.Size(80, 25)
$btnTarget.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "すべてのファイル (*.*)|*.*|実行ファイル (*.exe)|*.exe|Python (*.py;*.pyw)|*.py;*.pyw|PowerShell (*.ps1)|*.ps1|バッチ (*.bat;*.cmd)|*.bat;*.cmd"
    if ($dlg.ShowDialog() -eq "OK") {
        $txtTarget.Text = $dlg.FileName
        # ショートカット名が空なら、ファイル名で自動入力
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
            $txtName.Text = [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName)
        }
    }
})
$form.Controls.Add($btnTarget)

# ターゲットを手入力／貼り付けした場合も、名前が空なら自動入力する
$txtTarget.Add_TextChanged({
    if ([string]::IsNullOrWhiteSpace($txtName.Text) -and
        (Test-Path -LiteralPath $txtTarget.Text -PathType Leaf)) {
        $txtName.Text = [System.IO.Path]::GetFileNameWithoutExtension($txtTarget.Text)
    }
})

$y += 40

# --- 引数 ---
$lblArgs = New-Object System.Windows.Forms.Label
$lblArgs.Text = "引数（オプション）："
$lblArgs.Location = New-Object System.Drawing.Point(20, $y)
$lblArgs.AutoSize = $true
$form.Controls.Add($lblArgs)

$y += 22
$txtArgs = New-Object System.Windows.Forms.TextBox
$txtArgs.Location = New-Object System.Drawing.Point(20, $y)
$txtArgs.Size = New-Object System.Drawing.Size(460, 25)
$form.Controls.Add($txtArgs)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($txtArgs, "スペースを含む引数は `"...`" で囲んでください")

$y += 40

# --- アイコン ---
$lblIcon = New-Object System.Windows.Forms.Label
$lblIcon.Text = "アイコン（オプション / .ico ファイル）："
$lblIcon.Location = New-Object System.Drawing.Point(20, $y)
$lblIcon.AutoSize = $true
$form.Controls.Add($lblIcon)

$y += 22
$txtIcon = New-Object System.Windows.Forms.TextBox
$txtIcon.Location = New-Object System.Drawing.Point(20, $y)
$txtIcon.Size = New-Object System.Drawing.Size(370, 25)
$form.Controls.Add($txtIcon)

$btnIcon = New-Object System.Windows.Forms.Button
$btnIcon.Text = "参照..."
$btnIcon.Location = New-Object System.Drawing.Point(400, ($y - 1))
$btnIcon.Size = New-Object System.Drawing.Size(80, 25)
$btnIcon.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "アイコン (*.ico)|*.ico|すべてのファイル (*.*)|*.*"
    if ($dlg.ShowDialog() -eq "OK") {
        $txtIcon.Text = $dlg.FileName
    }
})
$form.Controls.Add($btnIcon)

$y += 40

# --- .ps1 自動ラップのチェックボックス ---
$chkWrap = New-Object System.Windows.Forms.CheckBox
$chkWrap.Text = ".ps1 は PowerShell でラップ起動する（-ExecutionPolicy Bypass -File）"
$chkWrap.Location = New-Object System.Drawing.Point(20, $y)
$chkWrap.Size = New-Object System.Drawing.Size(460, 22)
$chkWrap.Checked = $true
$form.Controls.Add($chkWrap)

$y += 26

# --- ウィンドウ非表示のチェックボックス（ラップの下位オプション。右にインデント） ---
$chkHidden = New-Object System.Windows.Forms.CheckBox
$chkHidden.Text = "ウィンドウを隠す（GUI/常駐向け。Read-Host 等の対話型は OFF）"
$chkHidden.Location = New-Object System.Drawing.Point(40, $y)
$chkHidden.Size = New-Object System.Drawing.Size(440, 22)
$chkHidden.Checked = $false
$form.Controls.Add($chkHidden)

$y += 26

# --- PowerShell 7 (pwsh) で起動するチェックボックス（ラップの下位オプション） ---
$chkPwsh = New-Object System.Windows.Forms.CheckBox
$chkPwsh.Text = "PowerShell 7 (pwsh) で起動する（無ければ Windows PowerShell）"
$chkPwsh.Location = New-Object System.Drawing.Point(40, $y)
$chkPwsh.Size = New-Object System.Drawing.Size(440, 22)
$chkPwsh.Checked = $true
$form.Controls.Add($chkPwsh)

$y += 26

# --- .py を pythonw.exe で起動するチェックボックス（コンソール窓を出さない） ---
$chkPy = New-Object System.Windows.Forms.CheckBox
$chkPy.Text = ".py は pythonw.exe で起動する（コンソール非表示。GUI アプリ向け）"
$chkPy.Location = New-Object System.Drawing.Point(20, $y)
$chkPy.Size = New-Object System.Drawing.Size(460, 22)
$chkPy.Checked = $false
$form.Controls.Add($chkPy)

$y += 36

# --- 追加ボタン ---
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "スタートメニューに追加"
$btnAdd.Location = New-Object System.Drawing.Point(150, $y)
$btnAdd.Size = New-Object System.Drawing.Size(200, 35)
$btnAdd.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10, [System.Drawing.FontStyle]::Bold)
$btnAdd.Add_Click({

    # --- バリデーション ---
    if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
        [System.Windows.Forms.MessageBox]::Show("ショートカット名を入力してください。", "エラー", "OK", "Warning")
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtTarget.Text)) {
        [System.Windows.Forms.MessageBox]::Show("ターゲットのパスを入力してください。", "エラー", "OK", "Warning")
        return
    }
    if (-not (Test-Path -LiteralPath $txtTarget.Text -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("指定されたターゲット（ファイル）が見つかりません：`n$($txtTarget.Text)", "エラー", "OK", "Warning")
        return
    }
    # ショートカット名にファイル名禁止文字が含まれていないか
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    if ($txtName.Text.IndexOfAny($invalidChars) -ge 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "ショートカット名に次の文字は使用できません：`n\ / : * ? `" < > |",
            "エラー", "OK", "Warning")
        return
    }
    # アイコンが指定されている場合は存在チェック
    if (-not [string]::IsNullOrWhiteSpace($txtIcon.Text) -and -not (Test-Path -LiteralPath $txtIcon.Text -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("指定されたアイコンが見つかりません：`n$($txtIcon.Text)", "エラー", "OK", "Warning")
        return
    }

    # --- ショートカット作成 ---
    try {
        $startMenu = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
        $lnkPath = Join-Path $startMenu "$($txtName.Text).lnk"

        # 同名ショートカットが存在する場合は上書き確認
        if (Test-Path -LiteralPath $lnkPath) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                "同名のショートカットが既に存在します。上書きしますか？`n`n$lnkPath",
                "確認", "YesNo", "Question")
            if ($answer -ne "Yes") { return }
        }

        $targetPath = $txtTarget.Text
        $arguments  = $txtArgs.Text
        $ext = [System.IO.Path]::GetExtension($targetPath).ToLower()

        # .ps1 ラップ時に使用するエンジンを先に決定する
        #   既定で PowerShell 7 (pwsh) を使い、無ければ Windows PowerShell にフォールバック
        $engine = "powershell.exe"
        $usePwsh = $false
        if ($chkWrap.Checked -and $ext -eq ".ps1" -and $chkPwsh.Checked) {
            $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($pwshCmd) {
                $engine = $pwshCmd.Source
                $usePwsh = $true
            } else {
                # pwsh を指定されたが見つからない → フォールバックする旨を通知
                [System.Windows.Forms.MessageBox]::Show(
                    "PowerShell 7 (pwsh.exe) が見つからないため、Windows PowerShell (powershell.exe) で起動するショートカットを作成します。",
                    "情報", "OK", "Information")
            }
        }

        # BOM なし UTF-8 の .ps1 は Windows PowerShell 5.1 が文字化けして
        # 構文エラーで即終了する。pwsh は UTF-8 を既定で読むため、BOM 付与が
        # 必要なのは Windows PowerShell で実行されるケースのみに限定する。
        if ($ext -eq ".ps1" -and -not $usePwsh) {
            $bytes = [System.IO.File]::ReadAllBytes($targetPath)
            $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
            $hasNonAscii = $false
            foreach ($bb in $bytes) { if ($bb -gt 0x7F) { $hasNonAscii = $true; break } }
            if (-not $hasBom -and $hasNonAscii) {
                # 別エンコーディングを壊さないよう UTF-8 として妥当か検証してから付与
                $isValidUtf8 = $true
                $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
                try { [void]$strictUtf8.GetString($bytes) } catch { $isValidUtf8 = $false }
                if ($isValidUtf8) {
                    $ans = [System.Windows.Forms.MessageBox]::Show(
                        "対象の .ps1 は BOM なし UTF-8 です。`nこのままだと Windows PowerShell で日本語が文字化けし、起動に失敗する場合があります。`n`nBOM を付与して修正しますか？（推奨）`n※元ファイルは .bak としてバックアップします。",
                        "確認", "YesNo", "Question")
                    if ($ans -eq "Yes") {
                        # 書き換え前にバックアップを作成
                        Copy-Item -LiteralPath $targetPath -Destination "$targetPath.bak" -Force
                        $text = $strictUtf8.GetString($bytes)
                        [System.IO.File]::WriteAllText($targetPath, $text, (New-Object System.Text.UTF8Encoding($true)))
                    }
                }
            }
        }

        # .ps1 のラップ処理
        if ($chkWrap.Checked -and $ext -eq ".ps1") {
            $hideOpt = if ($chkHidden.Checked) { "-WindowStyle Hidden " } else { "" }
            $arguments = "-ExecutionPolicy Bypass ${hideOpt}-File `"$targetPath`" $arguments".Trim()
            $targetPath = $engine
        }

        # .pyw のラップ処理（pythonw.exe が関連付けされていない場合の保険）
        # および .py を pythonw で起動（コンソール非表示）するオプション
        if (($ext -eq ".pyw") -or ($ext -eq ".py" -and $chkPy.Checked)) {
            # pythonw.exe を探す
            $pythonw = Get-Command pythonw.exe -ErrorAction SilentlyContinue
            if ($pythonw) {
                $arguments = "`"$targetPath`" $arguments".Trim()
                $targetPath = $pythonw.Source
            }
            elseif ($ext -eq ".py") {
                # .py で明示的に pythonw を選んだのに見つからない場合だけ通知
                # （.pyw は従来どおり関連付け任せのフォールバック）
                [System.Windows.Forms.MessageBox]::Show(
                    "pythonw.exe が見つからないため、関連付け（通常 python.exe）で起動します。コンソール窓が表示される場合があります。",
                    "情報", "OK", "Information")
            }
            # .pyw で見つからなければ .pyw 直接指定（関連付けに依存）
        }

        $shell = New-Object -ComObject WScript.Shell
        try {
            $sc = $shell.CreateShortcut($lnkPath)
            $sc.TargetPath = $targetPath
            if (-not [string]::IsNullOrWhiteSpace($arguments)) {
                $sc.Arguments = $arguments
            }
            $sc.WorkingDirectory = [System.IO.Path]::GetDirectoryName($txtTarget.Text)
            if (-not [string]::IsNullOrWhiteSpace($txtIcon.Text)) {
                $sc.IconLocation = $txtIcon.Text
            }
            $sc.Save()
        }
        finally {
            # COM オブジェクトを解放（複数回登録時の RCW 蓄積を防ぐ）
            if ($sc)    { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sc) }
            if ($shell) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
        }

        [System.Windows.Forms.MessageBox]::Show(
            "スタートメニューに追加しました！`n`n場所: $lnkPath",
            "完了", "OK", "Information"
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "エラーが発生しました：`n$($_.Exception.Message)",
            "エラー", "OK", "Error"
        )
    }
})
$form.Controls.Add($btnAdd)

# --- 対象拡張子に応じて .ps1 関連オプションの有効/無効を切り替える ---
$updatePs1Options = {
    $ext = ""
    if (-not [string]::IsNullOrWhiteSpace($txtTarget.Text)) {
        $ext = [System.IO.Path]::GetExtension($txtTarget.Text).ToLower()
    }
    $isPs1 = ($ext -eq ".ps1")
    $chkWrap.Enabled = $isPs1
    # 「隠す」「pwsh」はラップが有効なときだけ意味を持つ（入れ子の下位オプション）
    $sub = $isPs1 -and $chkWrap.Checked
    $chkHidden.Enabled = $sub
    $chkPwsh.Enabled   = $sub
    # 「.py を pythonw で起動」は対象が .py のときだけ意味を持つ
    $chkPy.Enabled     = ($ext -eq ".py")
}
$txtTarget.Add_TextChanged($updatePs1Options)
$chkWrap.Add_CheckedChanged($updatePs1Options)
& $updatePs1Options

# --- ターゲット欄はパスが長いと頭が切れるため、フルパスを ToolTip 表示 ---
$txtTarget.Add_TextChanged({ $toolTip.SetToolTip($txtTarget, $txtTarget.Text) })

# --- フォームへのドラッグ＆ドロップでターゲットを設定 ---
$form.AllowDrop = $true
$form.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
})
$form.Add_DragDrop({
    param($s, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files -and $files.Count -gt 0) {
        # Text 変更で TextChanged が発火し、名前自動入力・オプション切替・ToolTip も更新される
        $txtTarget.Text = $files[0]
    }
})

# --- 表示 ---
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
$form.Dispose()
