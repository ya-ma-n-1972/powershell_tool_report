###############################################################
# Windows/WSL2両対応 フォルダツリー生成器
# GUIでフォルダツリーを取得し、ツリー形式で出力するツール
# - Windowsパス/WSLパス両対応
# - ツリー表示・チェックボックス選択・ツリー編集・コピー機能
# - .NET WinFormsを利用したGUI
###############################################################

###############################################################
# .NET Framework のアセンブリ（ライブラリ）を読み込み
# GUIアプリケーション用（WinForms）とグラフィック描画用（Drawing）
###############################################################
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

###############################################################
# データ構造の定義
# フォルダ/ファイルのツリー構造を保持するクラス
# - Name: 名前
# - Path: フルパス
# - IsDirectory: ディレクトリであるかどうか
# - Children: 子ノード（ArrayList）
# - IsChecked: チェック状態（ツリー出力対象か）
###############################################################
class FileSystemNode {
    [string]$Name
    [string]$Path
    [bool]$IsDirectory
    [System.Collections.ArrayList]$Children
    [bool]$IsChecked
    
    FileSystemNode([string]$name, [string]$path, [bool]$isDirectory) {
        $this.Name = $name
        $this.Path = $path
        $this.IsDirectory = $isDirectory
        $this.Children = [System.Collections.ArrayList]::new() # 子ノード格納用
        $this.IsChecked = $false # 初期状態は未選択
    }
    # 子ノード追加メソッド
    [void]AddChild([FileSystemNode]$child) {
        [void]$this.Children.Add($child)
    }
}

###############################################################
# パス判定機能
# Windowsパス/WSLパスの判定（UNC形式やLinux形式に対応）
###############################################################
function Get-PathType {
    param([string]$Path)
    
    # WSLのUNCパス（\\wsl.localhost\...）またはLinuxスタイルのパス（/home/...）を判定
    if ($Path -match '^\\\\wsl[\.\$]' -or $Path -match '^/') {
        return "WSL"
    }
    # Windowsのドライブレター形式（C:\等）または実在するWindowsパスを判定
    elseif ($Path -match '^[A-Z]:\\' -or (Test-Path $Path -ErrorAction SilentlyContinue)) {
        return "Windows"
    }
    else {
        return "Unknown"
    }
}

###############################################################
# WindowsパスをWSLパスに変換
# Windows形式のパス（UNC形式やドライブレター）をLinux形式に変換
# 戻り値: @{ Path = "変換後のパス", Distro = "ディストリビューション名" }
###############################################################
function Convert-ToWSLPath {
    param([string]$WindowsPath)

    if ($WindowsPath -match '^\\\\wsl[\.\$](?:localhost\\)?([^\\]+)(.*)') {
        $distro = $Matches[1]
        $path = $Matches[2] -replace '\\', '/'

        # ディストリビューションルートの場合
        if ([string]::IsNullOrEmpty($path)) {
            $path = "/"
        }

        # ディストロとパスの両方を返す
        return @{
            Path = $path
            Distro = $distro
        }
    }

    if ($WindowsPath -match '^([A-Z]):(.*)') {
        $drive = $Matches[1].ToLower()
        $path = $Matches[2] -replace '\\', '/'
        return @{
            Path = "/mnt/$drive$path"
            Distro = $null
        }
    }

    return @{
        Path = $WindowsPath
        Distro = $null
    }
}

###############################################################
# Windowsパス用のディレクトリツリー取得
# 再帰的にディレクトリ・ファイルを探索し、FileSystemNodeツリーを構築
###############################################################
function Get-WindowsDirectoryTree {
    param(
        [string]$Path, # 探索するルートパス
        [int]$MaxDepth = 3, # 探索する最大深度
        [bool]$IncludeFiles = $true # ファイルも含めるかどうか
    )
    
    if (-not (Test-Path $Path)) {
        [System.Windows.Forms.MessageBox]::Show(
            "指定されたパスが存在しません: $Path",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return $null
    }
    
    $rootName = Split-Path -Leaf $Path # パスの最後の要素を取得
    if (-not $rootName) { $rootName = $Path } # ルートドライブの場合
    $rootNode = [FileSystemNode]::new($rootName, $Path, $true) # ルートノード作成

    # 再帰的にディレクトリをスキャンする内部ScriptBlock
    # - 指定深度までディレクトリ/ファイルを探索
    # - FileSystemNodeツリーを構築
    # ScriptBlockとして定義してスコープ汚染を防止
    $scanDirectory = {
        param(
            [string]$DirectoryPath, # 現在探索中のディレクトリ
            [FileSystemNode]$ParentNode, # 親ノード
            [int]$CurrentDepth, # 現在の深度
            [int]$MaxDepth, # 最大深度
            [bool]$IncludeFiles # ファイル含有フラグ
        )

        # 指定された深度に達したら、それ以上深く探索しない
        if ($CurrentDepth -ge $MaxDepth) { return }

        try {
            # Get-ChildItemでディレクトリ内のサブディレクトリを取得
            # -Attributes !ReparsePoint: ジャンクション/シンボリックリンクを除外
            #   （循環参照や予期しない巨大ツリーの展開を防ぐため）
            # Sort-Object Name: アルファベット順にソート（WSL側と統一、実行ごとの安定性向上）
            $directories = Get-ChildItem -Path $DirectoryPath -Directory -Attributes !ReparsePoint -ErrorAction SilentlyContinue | Sort-Object Name
            # サブディレクトリごとにノードを作成し、再帰的にスキャン
            foreach ($dir in $directories) {
                $dirNode = [FileSystemNode]::new($dir.Name, $dir.FullName, $true)
                $ParentNode.AddChild($dirNode)

                # 再帰呼び出し（& 演算子でScriptBlockを実行）
                & $scanDirectory -DirectoryPath $dir.FullName `
                              -ParentNode $dirNode `
                              -CurrentDepth ($CurrentDepth + 1) `
                              -MaxDepth $MaxDepth `
                              -IncludeFiles $IncludeFiles
            }

            if ($IncludeFiles) {
                # Get-ChildItemでディレクトリ内のファイルを取得
                # Sort-Object Name: アルファベット順にソート（WSL側と統一、実行ごとの安定性向上）
                $files = Get-ChildItem -Path $DirectoryPath -File -ErrorAction SilentlyContinue | Sort-Object Name
                foreach ($file in $files) {
                    $fileNode = [FileSystemNode]::new($file.Name, $file.FullName, $false)
                    $ParentNode.AddChild($fileNode)
                }
            }
        }
        catch {
            Write-Host "警告: $DirectoryPath のスキャン中にエラー: $_" -ForegroundColor Yellow
        }
    }

    # ScriptBlock のエントリーポイント
    & $scanDirectory -DirectoryPath $Path -ParentNode $rootNode -CurrentDepth 0 -MaxDepth $MaxDepth -IncludeFiles $IncludeFiles
    
    return $rootNode
}

###############################################################
# WSL2パス用のディレクトリツリー取得
# WSL環境でLinuxコマンド(find)を使い、ディレクトリ/ファイル一覧を取得
# 文字列リストからFileSystemNodeツリーを再構築
###############################################################
function Get-WSLDirectoryTree {
    param(
        [string]$Path,              # 探索するルートパス（WSL形式 or Windows UNC形式）
        [int]$MaxDepth = 3,         # 探索する最大深度
        [bool]$IncludeFiles = $true # ファイルも含めるかどうか
    )

    # Windows UNC形式（\\wsl.localhost\Ubuntu\path）をLinux形式（/path）に変換
    # 異なる環境間でのパス形式の統一が必要
    $wslInfo = if ($Path -match '^\\\\wsl') {
        Convert-ToWSLPath -WindowsPath $Path
    } else {
        @{ Path = $Path; Distro = $null }
    }

    $wslPath = $wslInfo.Path
    $distro = $wslInfo.Distro
    
    # ルートディレクトリの場合は警告して中止
    if ($wslPath -eq "/") {
        [System.Windows.Forms.MessageBox]::Show(
            "WSLディストリビューションのルートディレクトリ全体のツリー取得はサポートされていません。`n`n" +
            "特定のディレクトリ（例: /home）を選択するか、`n" +
            "コンソールで直接コマンドを実行してください。",
            "警告",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return $null
    }
    
    # パスの最後の要素を取得（ルートの場合はパス全体）
    $rootName = Split-Path -Leaf $wslPath
    if (-not $rootName) { $rootName = $wslPath }
    $rootNode = [FileSystemNode]::new($rootName, $wslPath, $true) # ルートノード作成
    
    try {
        # WSL環境でパスの存在確認（Linuxコマンドを実行）
        # Windows側からは直接アクセスできないため、WSL経由で確認が必要
        $checkCommand = "test -d '$wslPath' && echo 'EXISTS' || echo 'NOT_EXISTS'"
        $exists = if ($distro) {
            wsl -d $distro bash -c $checkCommand 2>$null
        } else {
            wsl bash -c $checkCommand 2>$null
        }
        
        if ($exists -ne 'EXISTS') {
            [System.Windows.Forms.MessageBox]::Show(
                "指定されたWSLパスが存在しません: $wslPath",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return $rootNode
        }
        
        # findコマンドで全ディレクトリを一括取得
        # -maxdepthで効率的に深度を制限（Windows版のようなPowerShellでの再帰処理は不要）
        # WSL側で効率的に処理し、結果を文字列リストとして受け取る
        $findCommand = "cd '$wslPath' 2>/dev/null && find . -maxdepth $MaxDepth -type d 2>/dev/null | sort"
        $directories = if ($distro) {
            @(wsl -d $distro bash -c $findCommand)
        } else {
            @(wsl bash -c $findCommand)
        }
        
    # パスマップで親子関係を管理（文字列から階層構造を再構築するため）
    # キー: パス文字列、値: 対応するノードオブジェクト
        $pathMap = @{ "." = $rootNode }
        
    # findの結果（フラットな文字列リスト）からツリー構造を再構築
        foreach ($dir in $directories) {
            # カレントディレクトリや空行はスキップ
            if ($dir -eq "." -or [string]::IsNullOrWhiteSpace($dir)) { continue }
            
            # "./folder1/subfolder" → "folder1/subfolder" に正規化
            $relativePath = $dir.TrimStart('./')
            $parts = $relativePath -split '/'  # パスを階層ごとに分割
            
            # 親ディレクトリのパスと現在のディレクトリ名を特定
            if ($parts.Count -eq 1) {
                # 第1階層のディレクトリ（例: "folder1"）
                $parentKey = "."                # 親はルート
                $dirName = $parts[0]            # ディレクトリ名
            } else {
                # 第2階層以降（例: "folder1/subfolder"）
                $parentParts = $parts[0..($parts.Count - 2)]  # 親パスの部分を抽出
                $parentKey = "./" + ($parentParts -join '/')  # 親のキーを構築
                $dirName = $parts[-1]                          # 最後の要素がディレクトリ名
            }
            
            # 親ノードが存在する場合のみ、子ノードとして追加
            # （親が先に処理されていることを前提とする）
            if ($pathMap.ContainsKey($parentKey)) {
                # フルパスを構築（WSL内部での絶対パス）
                $fullPath = if ($wslPath -eq "/") { "/$relativePath" } else { "$wslPath/$relativePath" }
                $newNode = [FileSystemNode]::new($dirName, $fullPath, $true)
                $pathMap[$parentKey].AddChild($newNode)    # 親ノードに子として追加
                $pathMap[$dir] = $newNode                  # マップに登録（後続の子の親になる可能性）
            }
        }
        
    # ファイルも含める場合の処理（ディレクトリと同様の流れ）
        if ($IncludeFiles) {
            # findコマンドで全ファイルを一括取得
            $filesCommand = "cd '$wslPath' 2>/dev/null && find . -maxdepth $MaxDepth -type f 2>/dev/null | sort"
            $files = if ($distro) {
                @(wsl -d $distro bash -c $filesCommand)
            } else {
                @(wsl bash -c $filesCommand)
            }
            
            # ファイルごとに親ディレクトリを特定して追加
            foreach ($file in $files) {
                if ([string]::IsNullOrWhiteSpace($file)) { continue }
                
                # ディレクトリと同じロジックでパスを解析
                $relativePath = $file.TrimStart('./')
                $parts = $relativePath -split '/'
                
                # 親ディレクトリとファイル名を特定
                if ($parts.Count -eq 1) {
                    # ルート直下のファイル
                    $parentKey = "."
                    $fileName = $parts[0]
                } else {
                    # サブディレクトリ内のファイル
                    $parentParts = $parts[0..($parts.Count - 2)]
                    $parentKey = "./" + ($parentParts -join '/')
                    $fileName = $parts[-1]
                }
                
                # 親ディレクトリが存在する場合、ファイルノードを追加
                if ($pathMap.ContainsKey($parentKey)) {
                    $fullPath = if ($wslPath -eq "/") { "/$relativePath" } else { "$wslPath/$relativePath" }
                    $fileNode = [FileSystemNode]::new($fileName, $fullPath, $false)
                    $pathMap[$parentKey].AddChild($fileNode)
                }
            }
        }
    }
    catch {
        # WSL関連のエラー（WSL未インストール、起動失敗など）をキャッチ
        [System.Windows.Forms.MessageBox]::Show(
            "WSLアクセスエラー: $_`n`nWSLがインストールされていることを確認してください",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    
    return $rootNode
}

###############################################################
# WindowsとWSL2の両環境に対応した統合ディレクトリツリー取得関数
# パス形式を判定し、適切な取得関数を呼び出す
###############################################################
function Get-DirectoryTree {
    param(
        [string]$Path,           # 探索するパス（Windows/WSL両対応）
        [int]$MaxDepth = 3,      # 探索する最大深度（デフォルト3階層）
        [bool]$IncludeFiles = $true  # ファイルを含めるか（デフォルトは含める）
    )
    
    # 入力されたパスが「Windowsパス」か「WSLパス」かを判定
    $pathType = Get-PathType -Path $Path

    # パスの形式に応じて処理を振り分け
    switch ($pathType) {
        "Windows" {
            Write-Host "Windowsパスとして処理: $Path" -ForegroundColor Green
            return Get-WindowsDirectoryTree -Path $Path -MaxDepth $MaxDepth -IncludeFiles $IncludeFiles
        }
        "WSL" {
            Write-Host "WSLパスとして処理: $Path" -ForegroundColor Cyan
            return Get-WSLDirectoryTree -Path $Path -MaxDepth $MaxDepth -IncludeFiles $IncludeFiles
        }
        default {
            [System.Windows.Forms.MessageBox]::Show(
                "パスの形式を認識できません: $Path`n`n" +
                "サポートされている形式:`n" +
                "- Windows: C:\folder または D:\path\to\folder`n" +
                "- WSL: /home/user または \\wsl.localhost\Ubuntu\path",
                "エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $null
        }
    }
}

###############################################################
# ツリー文字列生成
# FileSystemNodeツリーからツリー形式の文字列を生成
# - チェックされたノードのみ出力
# - 階層/接続記号（├──, └──）で見やすく表示
# - コメント記号（//、#、なし）を選択可能
###############################################################
function ConvertTo-TreeOutput {
    param(
        [FileSystemNode]$Node,    # 処理対象のノード
        [string]$Prefix = "",      # 現在の階層のプレフィックス（インデント用）
        [bool]$IsLast = $true,     # 兄弟ノード中で最後かどうか
        [bool]$IsRoot = $true,     # ルートノードかどうか

        # コメント記号スタイル
        [ValidateSet("//", "#", "None")]
        [string]$CommentStyle = "//"
    )
    
    # 文字列の表示幅を計算する内部関数（全角文字対応）
    # 全角文字は2文字分、半角文字は1文字分としてカウントし、見た目の幅を揃えるために使用
    function Get-DisplayWidth {
        param([string]$Text)
        $width = 0
        foreach ($char in $Text.ToCharArray()) {
            # 全角文字判定（Unicode範囲に基づく簡易判定）
            if ([int]$char -ge 0x3000 -and [int]$char -le 0x9FFF) {
                $width += 2
            }
            elseif ([int]$char -ge 0xFF00 -and [int]$char -le 0xFFEF) {
                $width += 2
            }
            else {
                $width += 1
            }
        }
        return $width
    }
    
    # まず通常のツリーを生成（現在の処理をそのまま実行）
    # ローカルスコープで配列を保持（関数内でのみ有効）
    # ArrayList を使用することで、内部関数からも直接変更可能
    $lines = [System.Collections.ArrayList]::new()

    # ツリー生成の内部再帰ScriptBlock
    # 注: ArrayList は参照型なので、ScriptBlockから直接 $lines を操作できます
    # ScriptBlockとして定義してスコープ汚染を防止
    $buildTreeLines = {
        param(
            [FileSystemNode]$Node,
            [string]$Prefix = "",
            [bool]$IsLast = $true,
            [bool]$IsRoot = $true
        )

        if ($IsRoot) {
            $nodeName = if ($Node.IsDirectory) { "$($Node.Name)/" } else { $Node.Name }
            [void]$lines.Add($nodeName)

            $checkedChildren = $Node.Children | Where-Object { $_.IsChecked }

            for ($i = 0; $i -lt $checkedChildren.Count; $i++) {
                $child = $checkedChildren[$i]
                $isLastChild = ($i -eq $checkedChildren.Count - 1)
                # 再帰呼び出し（& 演算子でScriptBlockを実行）
                & $buildTreeLines -Node $child -Prefix "" -IsLast $isLastChild -IsRoot $false
            }
        }
        else {
            if (-not $Node.IsChecked) { return }

            $connector = if ($IsLast) { "└── " } else { "├── " }
            $nodeName = if ($Node.IsDirectory) { "$($Node.Name)/" } else { $Node.Name }
            [void]$lines.Add("$Prefix$connector$nodeName")

            $childPrefix = $Prefix + $(if ($IsLast) { "    " } else { "│   " })

            $checkedChildren = $Node.Children | Where-Object { $_.IsChecked }
            for ($i = 0; $i -lt $checkedChildren.Count; $i++) {
                $child = $checkedChildren[$i]
                $isLastChild = ($i -eq $checkedChildren.Count - 1)
                # 再帰呼び出し（& 演算子でScriptBlockを実行）
                & $buildTreeLines -Node $child -Prefix $childPrefix -IsLast $isLastChild -IsRoot $false
            }
        }
    }

    # ツリー構造を生成
    & $buildTreeLines -Node $Node -Prefix $Prefix -IsLast $IsLast -IsRoot $IsRoot

    # 最長行の表示幅を計算
    $maxWidth = 0
    foreach ($line in $lines) {
        $width = Get-DisplayWidth -Text $line
        if ($width -gt $maxWidth) {
            $maxWidth = $width
        }
    }

    # コメント位置を決定（最長行 + 5文字、ただし最小40文字）
    $commentPosition = [Math]::Max($maxWidth + 5, 40)

    # 各行にコメント記号を付与
    $outputLines = @()
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            $outputLines += $line
        }
        else {
            if ($CommentStyle -eq "None") {
                # コメントなしモード
                $outputLines += $line
            }
            else {
                # 標準モード: スペースで埋めてコメント記号を配置
                $currentWidth = Get-DisplayWidth -Text $line
                $paddingCount = $commentPosition - $currentWidth

                # パディングが負の値にならないよう保護
                if ($paddingCount -lt 0) {
                    $paddingCount = 2  # 最小でも2スペース確保
                }

                $padding = " " * $paddingCount
                $outputLines += "${line}${padding}${CommentStyle}"
            }
        }
    }

    # 改行で結合して返す
    return ($outputLines -join "`r`n") + "`r`n"
}

###############################################################
# GUI定義
# WinFormsによるフォーム・各種コントロールの作成
# - パス入力/参照/深度/オプション/ツリービュー/編集エリア/各種ボタン
###############################################################
$script:rootNode = $null

###############################################################
# メインフォームの作成
# - タイトル/サイズ/中央表示
# - 最小サイズ設定（ウィンドウ最大化対応）
###############################################################
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows/WSL2 フォルダツリー生成"
$form.Size = New-Object System.Drawing.Size(950, 750)
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)  # 最小サイズ設定
$form.StartPosition = "CenterScreen"

###############################################################
# パス入力部分（ラベル・テキストボックス・参照ボタン）
###############################################################
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "パスを入力 (Windows: C:\folder、WSL: /home/user または \\wsl.localhost\Ubuntu\path):"
$pathLabel.Location = New-Object System.Drawing.Point(10, 15)
$pathLabel.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(10, 35)
$pathTextBox.Size = New-Object System.Drawing.Size(650, 20)
$pathTextBox.Text = [Environment]::GetFolderPath("MyDocuments")
$pathTextBox.Anchor = "Top,Left,Right"  # 横幅を伸縮
$form.Controls.Add($pathTextBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "参照"
$browseButton.Location = New-Object System.Drawing.Point(670, 33)
$browseButton.Size = New-Object System.Drawing.Size(75, 23)
$browseButton.Anchor = "Top,Right"  # 右端に固定
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "フォルダを選択"
    $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $pathTextBox.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($browseButton)

###############################################################
# 深度設定（ラベル・数値入力）
###############################################################
$depthLabel = New-Object System.Windows.Forms.Label
$depthLabel.Text = "深度:"
$depthLabel.Location = New-Object System.Drawing.Point(680, 15)
$depthLabel.Size = New-Object System.Drawing.Size(40, 20)
$depthLabel.Anchor = "Top,Right"  # 右端に固定
$form.Controls.Add($depthLabel)

$depthNumeric = New-Object System.Windows.Forms.NumericUpDown
$depthNumeric.Location = New-Object System.Drawing.Point(760, 35)
$depthNumeric.Size = New-Object System.Drawing.Size(50, 20)
$depthNumeric.Minimum = 1
$depthNumeric.Maximum = 10
$depthNumeric.Value = 3
$depthNumeric.Anchor = "Top,Right"  # 右端に固定
$form.Controls.Add($depthNumeric)

###############################################################
# オプション（ファイル含むチェックボックス）
###############################################################
$includeFilesCheckBox = New-Object System.Windows.Forms.CheckBox
$includeFilesCheckBox.Text = "ファイルを含む"
$includeFilesCheckBox.Location = New-Object System.Drawing.Point(820, 35)
$includeFilesCheckBox.Size = New-Object System.Drawing.Size(120, 20)
$includeFilesCheckBox.Checked = $true
$includeFilesCheckBox.Anchor = "Top,Right"  # 右端に固定
$form.Controls.Add($includeFilesCheckBox)

###############################################################
# 読み込みボタン
# - 入力パス/深度/オプションでツリー取得
# - ツリービューにノード追加
###############################################################
$loadButton = New-Object System.Windows.Forms.Button
$loadButton.Text = "読み込み"
$loadButton.Location = New-Object System.Drawing.Point(10, 65)
$loadButton.Size = New-Object System.Drawing.Size(100, 30)
$loadButton.BackColor = [System.Drawing.Color]::LightGreen
$loadButton.Anchor = "Top,Left"  # 上左に固定
$loadButton.Add_Click({
    # 1. 初期化処理
    $treeView.Nodes.Clear()        # ツリービューの既存ノードをクリア
    $previewTextBox.Text = ""      # 編集エリアをクリア
    $statusLabel.Text = "読み込み中..." # ステータスバーに進捗表示
    
    $script:rootNode = Get-DirectoryTree -Path $pathTextBox.Text `
                                        -MaxDepth $depthNumeric.Value `
                                        -IncludeFiles $includeFilesCheckBox.Checked
    
    # ツリービューにノードを追加
    if ($script:rootNode) {
        $rootTreeNode = New-Object System.Windows.Forms.TreeNode
        $rootTreeNode.Text = if ($script:rootNode.IsDirectory) { 
            "$($script:rootNode.Name)/"    # ディレクトリなら末尾に/を付ける
        } else { 
            $script:rootNode.Name 
        }
        $rootTreeNode.Tag = $script:rootNode  # FileSystemNodeオブジェクトを紐付け
        $rootTreeNode.Checked = $false        # 初期状態は未選択
        
    # 再帰的な子ノード追加（ScriptBlockとして定義してスコープ汚染を防止）
    $addTreeNodes = {
            param($ParentTreeNode, $ParentFileNode)

            foreach ($child in $ParentFileNode.Children) {
                # 子ノードを作成
                $childTreeNode = New-Object System.Windows.Forms.TreeNode

                # ディレクトリなら末尾に/を付ける
                $childTreeNode.Text = if ($child.IsDirectory) {
                    "$($child.Name)/"
                } else {
                    $child.Name
                }

                # FileSystemNodeオブジェクトを紐付け
                $childTreeNode.Tag = $child
                $childTreeNode.Checked = $false

                # ディレクトリを青色で表示
                if ($child.IsDirectory) {
                    $childTreeNode.ForeColor = [System.Drawing.Color]::Blue
                }

                # 親ノードに追加
                $ParentTreeNode.Nodes.Add($childTreeNode)

                # 子がさらに子を持つ場合は再帰呼び出し（& 演算子でScriptBlockを実行）
                if ($child.Children.Count -gt 0) {
                    & $addTreeNodes -ParentTreeNode $childTreeNode -ParentFileNode $child
                }
            }
        }

        # ScriptBlockを呼び出して全階層を構築
        & $addTreeNodes -ParentTreeNode $rootTreeNode -ParentFileNode $script:rootNode
        
        # ルートノードをツリービューに追加
        $treeView.Nodes.Add($rootTreeNode)
        
        # ルートノードを展開（第1階層を表示）
        $rootTreeNode.Expand()
        
    # ステータスバーに完了メッセージ
    $statusLabel.Text = "読み込み完了 - $($script:rootNode.Children.Count) 個の項目"
    } else {
        $statusLabel.Text = "読み込み失敗"
    }
})
$form.Controls.Add($loadButton)

###############################################################
# 全選択/全解除ボタン
# - ツリービューの全ノードのチェック状態を一括変更
###############################################################
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "全選択"
$selectAllButton.Location = New-Object System.Drawing.Point(120, 65)
$selectAllButton.Size = New-Object System.Drawing.Size(75, 30)
$selectAllButton.Anchor = "Top,Left"  # 上左に固定
# 再帰的なチェック処理（ScriptBlockとして定義してスコープ汚染を防止）
$selectAllButton.Add_Click({
    $setAllNodes = {
        param($Node, $Checked)
        $Node.Checked = $Checked        # 現在のノードをチェック
        foreach ($child in $Node.Nodes) {
            # 再帰呼び出し（& 演算子でScriptBlockを実行）
            & $setAllNodes -Node $child -Checked $Checked
        }
    }
    foreach ($node in $treeView.Nodes) {
        & $setAllNodes -Node $node -Checked $true  # trueで全選択
    }
})
$form.Controls.Add($selectAllButton)

$deselectAllButton = New-Object System.Windows.Forms.Button
$deselectAllButton.Text = "全解除"
$deselectAllButton.Location = New-Object System.Drawing.Point(200, 65)
$deselectAllButton.Size = New-Object System.Drawing.Size(75, 30)
$deselectAllButton.Anchor = "Top,Left"  # 上左に固定
# 再帰的なチェック解除処理（ScriptBlockとして定義してスコープ汚染を防止）
$deselectAllButton.Add_Click({
    $setAllNodes = {
        param($Node, $Checked)
        $Node.Checked = $Checked       # 現在のノードをチェック解除
        foreach ($child in $Node.Nodes) {
            # 再帰呼び出し（& 演算子でScriptBlockを実行）
            & $setAllNodes -Node $child -Checked $Checked
        }
    }

    foreach ($node in $treeView.Nodes) {
        & $setAllNodes -Node $node -Checked $false # falseで全解除
    }
})
$form.Controls.Add($deselectAllButton)

###############################################################
# ツリービュー
# - フォルダ/ファイル構造を階層表示
# - チェックボックスで選択可能
# - チェック時は子ノードも自動選択
###############################################################
$treeViewLabel = New-Object System.Windows.Forms.Label
$treeViewLabel.Text = "ファイルとフォルダ（チェックして選択）:"
$treeViewLabel.Location = New-Object System.Drawing.Point(10, 100)
$treeViewLabel.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($treeViewLabel)

$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Location = New-Object System.Drawing.Point(10, 120)
$treeView.Size = New-Object System.Drawing.Size(450, 400)
$treeView.CheckBoxes = $true
$treeView.Anchor = "Top,Left,Bottom"  # 高さを伸縮
$treeView.Add_AfterCheck({
    param($sender, $e)

    # ScriptBlockとして定義してスコープ汚染を防止
    $setChildNodes = {
        param($Node, $Checked)
        foreach ($child in $Node.Nodes) {
            $child.Checked = $Checked
            # 再帰呼び出し（& 演算子でScriptBlockを実行）
            & $setChildNodes -Node $child -Checked $Checked
        }
    }

    if ($e.Node) {
        & $setChildNodes -Node $e.Node -Checked $e.Node.Checked
    }
})
$form.Controls.Add($treeView)

###############################################################
# ツリー編集エリア
# - RichTextBoxでツリーを表示・編集可能
# - 大きいツリーへの対応（サイズ拡張、無制限テキスト）
###############################################################
$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Text = "ツリー編集エリア:"
$previewLabel.Location = New-Object System.Drawing.Point(470, 100)
$previewLabel.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($previewLabel)

# RichTextBoxを使用してより良い表示
$previewTextBox = New-Object System.Windows.Forms.RichTextBox
$previewTextBox.Location = New-Object System.Drawing.Point(470, 120)
$previewTextBox.Size = New-Object System.Drawing.Size(460, 400) 
$previewTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$previewTextBox.ReadOnly = $false  # 編集可能に変更
$previewTextBox.WordWrap = $false  # 横スクロール可能
$previewTextBox.ScrollBars = "Both"
$previewTextBox.MaxLength = 0  # テキスト長制限を解除（無制限）
$previewTextBox.Anchor = "Top,Left,Bottom,Right"  # 縦横両方向に伸縮
$form.Controls.Add($previewTextBox)

###############################################################
# コメント設定用ComboBox
# - コメント記号スタイル選択
###############################################################

# ComboBox: コメント記号スタイル選択
$comboBoxCommentStyle = New-Object System.Windows.Forms.ComboBox
$comboBoxCommentStyle.Location = New-Object System.Drawing.Point(10, 530)
$comboBoxCommentStyle.Size = New-Object System.Drawing.Size(140, 25)
$comboBoxCommentStyle.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBoxCommentStyle.Anchor = "Bottom,Left"
$comboBoxCommentStyle.Items.AddRange(@(
    "//",
    "#",
    "コメントなし"
))
$comboBoxCommentStyle.SelectedIndex = 0  # デフォルト: //
$form.Controls.Add($comboBoxCommentStyle)

###############################################################
# ツリー生成ボタン
# - チェック状態をFileSystemNodeに反映
# - 折りたたまれた子ノードは出力に含めない
# - 生成したツリーを編集エリアに表示
###############################################################

$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "ツリー生成"
$generateButton.Location = New-Object System.Drawing.Point(160, 530)
$generateButton.Size = New-Object System.Drawing.Size(120, 30)
$generateButton.BackColor = [System.Drawing.Color]::LightBlue
$generateButton.Anchor = "Bottom,Left"  # 下に固定
$generateButton.Add_Click({
    if (-not $script:rootNode) {
        [System.Windows.Forms.MessageBox]::Show(
            "先にフォルダ構造を読み込んでください",
            "警告",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # 修正版：ツリービューの展開状態を考慮してチェック状態を更新する
    # 親ノードが折りたたまれている場合、その配下の子ノードはすべて非選択として扱う
    $updateNodeCheckedState = {
        param(
            $TreeNode,
            [bool]$ParentIsExpanded = $true  # 親が展開されているかどうか
        )

        # 現在のノードのチェック状態を反映
        if ($TreeNode.Tag) {
            # 親が折りたたまれている場合、子要素は無効化
            if ($ParentIsExpanded) {
                $TreeNode.Tag.IsChecked = $TreeNode.Checked
            } else {
                $TreeNode.Tag.IsChecked = $false
            }
        }

        # 子ノードの処理
        foreach ($child in $TreeNode.Nodes) {
            # 現在のノードが展開されているかどうかを子に伝える
            # - TreeNode.IsExpandedがtrueかつ親も展開されている場合のみ、子を有効とする
            # - 折りたたまれている場合（IsExpanded = false）、子要素は全て無効
            $childShouldBeActive = $ParentIsExpanded -and $TreeNode.IsExpanded

            # 再帰呼び出し（& 演算子でScriptBlockを実行）
            & $updateNodeCheckedState -TreeNode $child -ParentIsExpanded $childShouldBeActive
        }
    }

    # ルートノードから処理開始
    foreach ($node in $treeView.Nodes) {
        # ルートノードは常に有効（ParentIsExpanded = true）
        & $updateNodeCheckedState -TreeNode $node -ParentIsExpanded $true
    }

    # ComboBoxから選択値を取得
    $selectedCommentStyle = switch ($comboBoxCommentStyle.SelectedItem.ToString()) {
        "//" { "//" }
        "#" { "#" }
        "コメントなし" { "None" }
        default { "//" }
    }

    # ツリー生成
    $treeOutput = ConvertTo-TreeOutput `
        -Node $script:rootNode `
        -CommentStyle $selectedCommentStyle
    $previewTextBox.Text = $treeOutput
    $statusLabel.Text = "ツリー生成完了"
})
$form.Controls.Add($generateButton)

###############################################################
# クリップボードにコピーボタン
# - ツリー編集エリアの内容をクリップボードへコピー
###############################################################
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = "クリップボードにコピー"
$copyButton.Location = New-Object System.Drawing.Point(290, 530)
$copyButton.Size = New-Object System.Drawing.Size(150, 30)
$copyButton.Anchor = "Bottom,Left"  # 下に固定
$copyButton.Add_Click({
    # RichTextBoxのテキストを取得
    $textToCopy = ""
    if ($previewTextBox -and $previewTextBox.Text) {
        $textToCopy = $previewTextBox.Text
    }
    
    if (-not [string]::IsNullOrWhiteSpace($textToCopy)) {
        [System.Windows.Forms.Clipboard]::SetText($textToCopy)
        [System.Windows.Forms.MessageBox]::Show(
            "ツリーをクリップボードにコピーしました",
            "成功",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    $statusLabel.Text = "クリップボードにコピー完了"
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "コピーするツリーがありません",
            "警告",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})
$form.Controls.Add($copyButton)

###############################################################
# 使用方法ラベル
# - ユーザー向けの簡易説明
###############################################################
$helpLabel = New-Object System.Windows.Forms.Label
$helpLabel.Text = @"
使用方法:
1. パスを入力（例: D:\folder, /home/user, \\wsl.localhost\Ubuntu\home, \\wsl$\Ubuntu\home）
2. [読み込み]をクリックしてフォルダ構造を取得
3. 出力したい項目にチェック（親をチェックすると子も自動選択）
4. [ツリー生成]をクリックして出力を作成
5. 必要に応じてクリップボードにコピー

サポート: Windows/WSL2両対応
"@
$helpLabel.Location = New-Object System.Drawing.Point(10, 570) 
$helpLabel.Size = New-Object System.Drawing.Size(920, 100) 
$helpLabel.Font = New-Object System.Drawing.Font("MS UI Gothic", 9)
$helpLabel.Anchor = "Bottom,Left,Right"  # 下に固定、横幅を伸縮
$form.Controls.Add($helpLabel)


###############################################################
# ステータスバー
# - 現在の状態を表示
###############################################################
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "準備完了"
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

###############################################################
# フォーム表示（メイン処理の開始）
###############################################################
$form.ShowDialog()
