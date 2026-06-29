###############################################################
# Claude Code プロジェクト永続化データ クリーナー
# Windows/WSL2両対応 GUI
#
# - ~/.claude/ 配下のプロジェクト単位永続化データを完全削除
# - JSONL内の cwd フィールドで正規パス識別（日本語パス対応）
# - claude project purge 相当の射程をカバー
#   削除対象: projects/<encoded>/、file-history/<sid>/、
#             tasks/<sid>*（旧 todos/ も対象）、session-env/<sid>/、
#             debug/<sid>.txt、history.jsonl 該当行（project キーで照合）、
#             ~/.claude.json 該当エントリ
#   温存: shell-snapshots/、backups/、settings.json、plugins/
###############################################################

###############################################################
# .NET Framework のアセンブリ（ライブラリ）を読み込み
###############################################################
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

###############################################################
# 他プロセスの作業ディレクトリ(cwd)を PEB から読み取る型
# - claude.exe セッションがどのプロジェクトを開いているか識別するため
# - 同一ユーザー・x64 プロセスなら管理者権限なしで読める
###############################################################
try {
    Add-Type -Language CSharp -TypeDefinition @'
using System;using System.Runtime.InteropServices;using System.Text;
public static class ClaudeProcPeb {
 [StructLayout(LayoutKind.Sequential)] struct PBI { public IntPtr R1; public IntPtr Peb; public IntPtr A; public IntPtr B; public IntPtr Pid; public IntPtr R3; }
 [DllImport("ntdll.dll")] static extern int NtQueryInformationProcess(IntPtr h,int c,ref PBI p,int l,out int r);
 [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr OpenProcess(int a,bool i,int pid);
 [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
 [DllImport("kernel32.dll",SetLastError=true)] static extern bool ReadProcessMemory(IntPtr h,IntPtr b,byte[] buf,int s,out int r);
 static IntPtr RPtr(IntPtr h,IntPtr a){ byte[] b=new byte[8]; int r; if(!ReadProcessMemory(h,a,b,8,out r))return IntPtr.Zero; return (IntPtr)BitConverter.ToInt64(b,0); }
 public static string GetCwd(int pid){
  IntPtr h=OpenProcess(0x410,false,pid); if(h==IntPtr.Zero)return null;
  try{ var pbi=new PBI(); int ret; if(NtQueryInformationProcess(h,0,ref pbi,Marshal.SizeOf(pbi),out ret)!=0)return null;
   IntPtr pp=RPtr(h,(IntPtr)((long)pbi.Peb+0x20)); if(pp==IntPtr.Zero)return null;       // PEB.ProcessParameters
   byte[] us=new byte[16]; int r; if(!ReadProcessMemory(h,(IntPtr)((long)pp+0x38),us,16,out r))return null; // CurrentDirectory.DosPath
   ushort len=BitConverter.ToUInt16(us,0); IntPtr buf=(IntPtr)BitConverter.ToInt64(us,8); if(len==0||buf==IntPtr.Zero)return null;
   byte[] sb=new byte[len]; if(!ReadProcessMemory(h,buf,sb,len,out r))return null; return Encoding.Unicode.GetString(sb,0,r);
  } finally{ CloseHandle(h); } }
}
'@
} catch {
    # 既に定義済み等は無視
}

###############################################################
# データクラス定義
# ClaudeProject: 検出された1プロジェクトの状態を保持
#   - Cwd: JSONLから取得した正規の絶対パス（識別子）
#   - EncodedFolderName: ~/.claude/projects/ 配下のフォルダ名
#   - Environment: "Windows" または "WSL2: <distro>"
#   - Distro: WSL2の場合のディストロ名（Windowsならnull）
#   - SessionIds: このプロジェクト配下のセッションUUID群
#   - LastModified: フォルダの最終更新日時
#   - SizeBytes: フォルダのサイズ（バイト）
###############################################################
class ClaudeProject {
    [string]$Cwd
    [string]$EncodedFolderName
    [string]$Environment
    [string]$Distro
    [string[]]$SessionIds
    [datetime]$LastModified
    [long]$SizeBytes

    # 対象種別: "Claude"（既定）または "Codex"
    [string]$Kind = "Claude"
    # Codex用: このプロジェクトに属する rollout ファイルのフルパス群
    [string[]]$Files
    # Codex用: CODEX_HOME（削除時の history.jsonl / session_index.jsonl 位置解決に使用）
    [string]$BasePath

    ClaudeProject() {
        $this.SessionIds = @()
        $this.Files = @()
    }
}

###############################################################
# スクリプト全体で共有する状態
###############################################################
$script:Environments = @()        # 検出された環境一覧
$script:CurrentProjects = @()      # 現在表示中のプロジェクト一覧（ClaudeProject[]）
$script:LastScanWarning = $null   # スキャン中の警告メッセージ（ステータスバー表示用）
$script:SortColumn = -1            # 一覧ソート中の列インデックス（-1=未ソート）
$script:SortAscending = $true      # ソート方向（true=昇順）

###############################################################
# エンコーディング初期設定
# WSLとのやり取りはUTF-8に統一
# WSL_UTF8=1 はWSL 0.64.0以降で有効
###############################################################
$env:WSL_UTF8 = "1"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

###############################################################
# 利用可能な環境を列挙
# - 常にWindowsを含める
# - wsl --list --quiet で取得したディストロを追加
###############################################################
function Get-AvailableEnvironments {
    $envs = @()

    # Windowsは常に利用可能
    $envs += [pscustomobject]@{
        Name   = "Windows"
        Distro = $null
        IsWsl  = $false
    }

    # WSL2ディストロを列挙
    try {
        $rawList = @(wsl --list --quiet 2>$null)
        foreach ($line in $rawList) {
            $trimmed = $line.Trim()
            # 空行・NULL文字混入をスキップ（UTF-16残骸対策）
            $trimmed = $trimmed -replace "`0", ""
            if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }

            $envs += [pscustomobject]@{
                Name   = "WSL2: $trimmed"
                Distro = $trimmed
                IsWsl  = $true
            }
        }
    }
    catch {
        Write-Host "WSL列挙でエラー: $_" -ForegroundColor Yellow
    }

    return $envs
}

###############################################################
# Windows側のClaude Codeプロジェクトを列挙
# - CLAUDE_CONFIG_DIR が設定されていれば優先
# - 各プロジェクトフォルダから先頭JSONLを読み、cwdを取得
###############################################################
###############################################################
# .claude.json の projects キーから「エンコード名 -> 実cwd」対応表を作る
# - Claude Code はプロジェクトパスの非英数字を - に置換してフォルダ名にする
#   （例: /mnt/d/Library/フリーランス活動/クラウドワークス
#         -> -mnt-d-Library------------------）
# - transcript に cwd が無い（空セッション等）プロジェクトの逆引きに使う
# - 同一エンコード名に複数キーが衝突する場合は曖昧として $null（使用しない）
###############################################################
function Get-EncodedCwdMap {
    param([string]$ConfigJsonPath)

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($ConfigJsonPath)) { return $map }
    if (-not (Test-Path -LiteralPath $ConfigJsonPath)) { return $map }
    try {
        # 空文字 "" キー対策で -AsHashTable
        $cfg = Get-Content -LiteralPath $ConfigJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashTable
        if ($cfg.ContainsKey('projects') -and $cfg.projects) {
            foreach ($k in $cfg.projects.Keys) {
                $enc = ([string]$k -replace '[^a-zA-Z0-9]', '-')
                if ($map.ContainsKey($enc)) {
                    $map[$enc] = $null   # 衝突 -> 曖昧（使わない）
                } else {
                    $map[$enc] = [string]$k
                }
            }
        }
    } catch {
        # 読めない場合は空マップのまま
    }
    return $map
}

function Get-WindowsClaudeProjects {
    # ~/.claude のパスを解決
    $claudeDir = if ($env:CLAUDE_CONFIG_DIR) {
        $env:CLAUDE_CONFIG_DIR
    } else {
        Join-Path $env:USERPROFILE ".claude"
    }
    $projectsDir = Join-Path $claudeDir "projects"

    if (-not (Test-Path $projectsDir)) {
        return @()
    }

    $projects = @()

    # cwd不明プロジェクトの逆引き表（削除時の history.jsonl / .claude.json 編集は
    # USERPROFILE\.claude.json を対象にするため、それと同じファイルから作る）
    $encodedToCwd = Get-EncodedCwdMap -ConfigJsonPath (Join-Path $env:USERPROFILE ".claude.json")

    Get-ChildItem -Path $projectsDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $folder = $_
        $proj = [ClaudeProject]::new()
        $proj.EncodedFolderName = $folder.Name
        $proj.Environment = "Windows"
        $proj.Distro = $null
        $proj.LastModified = $folder.LastWriteTime

        # JSONL一覧は1回だけ取得して使い回す（同一フォルダへの重複列挙を回避）
        $jsonlFiles = @(Get-ChildItem -Path $folder.FullName -Filter "*.jsonl" -ErrorAction SilentlyContinue)
        # 先頭のJSONLファイルからcwdを取得
        $jsonl = $jsonlFiles | Select-Object -First 1
        $cwd = $null
        if ($jsonl) {
            try {
                # StreamReader を明示的に Dispose してファイルハンドルを確実に解放する。
                # （[System.IO.File]::ReadLines() を foreach で break すると列挙子が
                #   破棄されず StreamReader のハンドルが残り、後続の削除で
                #   「being used by another process」ロックを起こすため）
                $sr = $null
                try {
                    $sr = New-Object System.IO.StreamReader($jsonl.FullName, [System.Text.Encoding]::UTF8)
                    $count = 0
                    while ($null -ne ($line = $sr.ReadLine())) {
                        $count++
                        if ($count -gt 100) { break }
                        if ([string]::IsNullOrWhiteSpace($line)) { continue }
                        try {
                            $parsed = $line | ConvertFrom-Json -ErrorAction Stop
                            if ($parsed.PSObject.Properties['cwd'] -and $parsed.cwd) {
                                $cwdCandidate = [string]$parsed.cwd
                                if (-not [string]::IsNullOrWhiteSpace($cwdCandidate)) {
                                    $cwd = $cwdCandidate
                                    break
                                }
                            }
                        } catch {
                            continue
                        }
                    }
                } finally {
                    if ($sr) { $sr.Dispose() }
                }
            } catch {
                # ファイル読み取り失敗時は $cwd = $null のまま
            }
        }
        if ($cwd) { $proj.Cwd = $cwd }
        if ([string]::IsNullOrEmpty($proj.Cwd)) {
            # transcript に cwd が無い場合は .claude.json から逆引き（一意一致のみ採用）
            if ($encodedToCwd.ContainsKey($proj.EncodedFolderName) -and $encodedToCwd[$proj.EncodedFolderName]) {
                $proj.Cwd = $encodedToCwd[$proj.EncodedFolderName]
            } else {
                $proj.Cwd = "(cwd不明: $($folder.Name))"
            }
        }

        # セッションID群（.jsonlファイル名）/ フルパス / 基底ディレクトリ（列挙結果を再利用）
        $proj.SessionIds = @($jsonlFiles | ForEach-Object { $_.BaseName })
        $proj.Files = @($jsonlFiles | ForEach-Object { $_.FullName })
        $proj.BasePath = $claudeDir

        # サイズ計算（失敗しても0で続行）
        try {
            $sum = (Get-ChildItem -Path $folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
            $proj.SizeBytes = if ($null -eq $sum) { 0 } else { [long]$sum }
        }
        catch {
            $proj.SizeBytes = 0
        }

        $projects += $proj
    }

    return ,$projects
}

###############################################################
# WSL2側のClaude Codeプロジェクトを列挙
# - wsl -d <distro> bash -c "..." で同等情報を取得
# - cwd取得、セッションID列挙、サイズ、最終更新を一括取得
###############################################################
function Get-WslClaudeProjects {
    param(
        [Parameter(Mandatory)]
        [string]$Distro
    )

    # スクリプト先頭で $env:WSL_UTF8 = "1" を設定済みなので、ここでは再設定不要

    # 1. WSL の HOME パスを取得（単純な1コマンド、--exec で軽量起動）
    $wslHome = $null
    try {
        $wslHomeRaw = wsl -d $Distro --exec bash -c 'echo $HOME' 2>$null
        if ($wslHomeRaw) {
            $wslHome = (($wslHomeRaw | Out-String) -replace "`0", "").Trim()
        }
    } catch {
        # 取得失敗時は $wslHome = $null のまま
    }
    if ([string]::IsNullOrWhiteSpace($wslHome)) {
        return @()
    }

    # 2. UNCパス組み立て
    $uncBase = "\\wsl.localhost\$Distro" + ($wslHome -replace '/', '\')
    $projectsDir = Join-Path $uncBase ".claude\projects"

    if (-not (Test-Path -LiteralPath $projectsDir)) {
        return @()
    }

    # 3. プロジェクトフォルダを列挙（Get-WindowsClaudeProjectsと同形のロジック）
    $projects = @()

    # cwd不明プロジェクトの逆引き表（WSL側の ~/.claude.json から作る）
    $encodedToCwd = Get-EncodedCwdMap -ConfigJsonPath (Join-Path $uncBase ".claude.json")

    Get-ChildItem -LiteralPath $projectsDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $folder = $_
        $proj = [ClaudeProject]::new()
        $proj.EncodedFolderName = $folder.Name
        $proj.Environment = "WSL2: $Distro"
        $proj.Distro = $Distro
        $proj.LastModified = $folder.LastWriteTime

        # JSONL一覧は1回だけ取得して使い回す（UNC越しの重複列挙を回避）
        $jsonlFiles = @(Get-ChildItem -LiteralPath $folder.FullName -Filter "*.jsonl" -ErrorAction SilentlyContinue)
        # 先頭のJSONLファイルからcwdを取得（マルチライン走査、最大100行）
        $jsonl = $jsonlFiles | Select-Object -First 1
        $cwd = $null
        if ($jsonl) {
            try {
                # StreamReader を明示的に Dispose してファイルハンドルを確実に解放する。
                # （[System.IO.File]::ReadLines() を foreach で break すると列挙子が
                #   破棄されず StreamReader のハンドルが残り、後続の削除で
                #   「being used by another process」ロックを起こすため）
                $sr = $null
                try {
                    $sr = New-Object System.IO.StreamReader($jsonl.FullName, [System.Text.Encoding]::UTF8)
                    $count = 0
                    while ($null -ne ($line = $sr.ReadLine())) {
                        $count++
                        if ($count -gt 100) { break }
                        if ([string]::IsNullOrWhiteSpace($line)) { continue }
                        try {
                            $parsed = $line | ConvertFrom-Json -ErrorAction Stop
                            if ($parsed.PSObject.Properties['cwd'] -and $parsed.cwd) {
                                $cwdCandidate = [string]$parsed.cwd
                                if (-not [string]::IsNullOrWhiteSpace($cwdCandidate)) {
                                    $cwd = $cwdCandidate
                                    break
                                }
                            }
                        } catch {
                            continue
                        }
                    }
                } finally {
                    if ($sr) { $sr.Dispose() }
                }
            } catch {
                # ファイル読み取り失敗時は $cwd = $null のまま
            }
        }
        if ($cwd) { $proj.Cwd = $cwd }
        if ([string]::IsNullOrEmpty($proj.Cwd)) {
            # transcript に cwd が無い場合は .claude.json から逆引き（一意一致のみ採用）
            if ($encodedToCwd.ContainsKey($proj.EncodedFolderName) -and $encodedToCwd[$proj.EncodedFolderName]) {
                $proj.Cwd = $encodedToCwd[$proj.EncodedFolderName]
            } else {
                $proj.Cwd = "(cwd不明: $($folder.Name))"
            }
        }

        # セッションID群 / フルパス / 基底ディレクトリ(UNC .claude)（列挙結果を再利用）
        $proj.SessionIds = @($jsonlFiles | ForEach-Object { $_.BaseName })
        $proj.Files = @($jsonlFiles | ForEach-Object { $_.FullName })
        $proj.BasePath = (Join-Path $uncBase ".claude")

        # サイズ計算（失敗しても0で続行）
        try {
            $sum = (Get-ChildItem -LiteralPath $folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
            $proj.SizeBytes = if ($null -eq $sum) { 0 } else { [long]$sum }
        }
        catch {
            $proj.SizeBytes = 0
        }

        $projects += $proj
    }

    return ,$projects
}

###############################################################
# 環境名から対応するプロジェクト一覧取得関数にディスパッチ
###############################################################
function Get-ClaudeProjects {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Environment
    )

    if ($Environment.IsWsl) {
        return Get-WslClaudeProjects -Distro $Environment.Distro
    } else {
        return Get-WindowsClaudeProjects
    }
}

###############################################################
# ===== ここから Codex 対応 =====
# Codex は Claude と保存構造が異なる:
#   sessions/YYYY/MM/DD/rollout-<時刻>-<uuid>.jsonl（日付別・プロジェクト別ではない）
#   archived_sessions/ にも同形式の退避分
#   cwd は rollout 先頭の session_meta.payload.cwd
#   session_id は rollout ファイル名末尾の UUID（meta側は版により空）
#   history.jsonl       : {session_id, ts, text}
#   session_index.jsonl : {id, thread_name, updated_at}
# 本ツールでは cwd で束ねて「プロジェクト」として扱い、ファイル系のみ削除する。
###############################################################

# rollout 先頭付近の session_meta.payload を返す（稼働中ファイルも読めるよう共有読み取り）
function Read-CodexSessionMeta {
    param([Parameter(Mandatory)][string]$Path)
    $fs = $null; $sr = $null
    try {
        $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
        for ($i = 0; $i -lt 5; $i++) {
            $line = $sr.ReadLine()
            if ($null -eq $line) { break }
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try {
                $o = $line | ConvertFrom-Json -ErrorAction Stop
                if ($o.type -eq 'session_meta' -and $o.PSObject.Properties['payload']) { return $o.payload }
            } catch { continue }
        }
    } catch {
        # 読めない場合は $null
    } finally {
        if ($sr) { $sr.Dispose() }
        if ($fs) { $fs.Dispose() }
    }
    return $null
}

# rollout ファイル名末尾の UUID を session_id として取り出す
function Get-CodexSessionId {
    param([Parameter(Mandatory)][string]$BaseName)
    if ($BaseName -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})$') {
        return $Matches[1]
    }
    return $null
}

# CODEX_HOME 配下の rollout を cwd で束ねて ClaudeProject[] を返す
# （Windows はローカルパス、WSL は UNC パスを $CodexHome に渡せば同じロジックで動く）
function Build-CodexProjects {
    param(
        [Parameter(Mandatory)][string]$CodexHome,
        [Parameter(Mandatory)][string]$Environment,
        [string]$Distro
    )

    $byCwd = @{}
    foreach ($sub in 'sessions', 'archived_sessions') {
        $root = Join-Path $CodexHome $sub
        if (-not (Test-Path -LiteralPath $root)) { continue }
        Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'rollout-*.jsonl' -ErrorAction SilentlyContinue | ForEach-Object {
            $f = $_
            $meta = Read-CodexSessionMeta -Path $f.FullName
            $cwd = if ($meta -and $meta.PSObject.Properties['cwd'] -and $meta.cwd) { [string]$meta.cwd } else { $null }
            if ([string]::IsNullOrWhiteSpace($cwd)) { $cwd = "(cwd不明)" }
            $sid = Get-CodexSessionId -BaseName $f.BaseName

            if (-not $byCwd.ContainsKey($cwd)) {
                $byCwd[$cwd] = [pscustomobject]@{
                    Files = [System.Collections.Generic.List[string]]::new()
                    Sids  = [System.Collections.Generic.List[string]]::new()
                    Size  = [long]0
                    Last  = [datetime]::MinValue
                }
            }
            $g = $byCwd[$cwd]
            $g.Files.Add($f.FullName)
            if ($sid) { $g.Sids.Add($sid) }
            $g.Size += $f.Length
            if ($f.LastWriteTime -gt $g.Last) { $g.Last = $f.LastWriteTime }
        }
    }

    $projects = @()
    foreach ($cwd in $byCwd.Keys) {
        $g = $byCwd[$cwd]
        $p = [ClaudeProject]::new()
        $p.Kind = "Codex"
        $p.Cwd = $cwd
        $p.EncodedFolderName = ""
        $p.Environment = $Environment
        $p.Distro = $Distro
        $p.BasePath = $CodexHome
        $p.SessionIds = @($g.Sids | Select-Object -Unique)
        $p.Files = @($g.Files)
        $p.LastModified = $g.Last
        $p.SizeBytes = $g.Size
        $projects += $p
    }
    return , $projects
}

function Get-WindowsCodexProjects {
    $cx = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE ".codex" }
    if (-not (Test-Path -LiteralPath $cx)) { return @() }
    return Build-CodexProjects -CodexHome $cx -Environment "Windows" -Distro $null
}

function Get-WslCodexProjects {
    param([Parameter(Mandatory)][string]$Distro)

    # WSL の HOME を取得して UNC を組む（CODEX_HOME 既定は ~/.codex）
    $wslHome = $null
    try {
        $wslHomeRaw = wsl -d $Distro --exec bash -c 'echo $HOME' 2>$null
        if ($wslHomeRaw) { $wslHome = (($wslHomeRaw | Out-String) -replace "`0", "").Trim() }
    } catch { }
    if ([string]::IsNullOrWhiteSpace($wslHome)) { return @() }

    $uncBase = "\\wsl.localhost\$Distro" + ($wslHome -replace '/', '\')
    $cx = Join-Path $uncBase ".codex"
    if (-not (Test-Path -LiteralPath $cx)) { return @() }
    return Build-CodexProjects -CodexHome $cx -Environment "WSL2: $Distro" -Distro $Distro
}

function Get-CodexProjects {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Environment
    )
    if ($Environment.IsWsl) {
        return Get-WslCodexProjects -Distro $Environment.Distro
    } else {
        return Get-WindowsCodexProjects
    }
}

# Codex プロジェクト1件分の削除（ファイル系のみ・SQLiteは対象外）
function Invoke-CodexDeletion {
    param([Parameter(Mandatory)][ClaudeProject]$Project)

    $errors = @()
    $cx = $Project.BasePath
    if ([string]::IsNullOrWhiteSpace($cx)) {
        $errors += "CODEX_HOME が解決できませんでした"
        return $errors
    }

    # session_id 集合
    $sidSet = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($s in $Project.SessionIds) { if ($s) { [void]$sidSet.Add([string]$s) } }

    # 1. rollout ファイル削除（稼働中セッションはロックで失敗しうる）
    foreach ($file in $Project.Files) {
        if (Test-Path -LiteralPath $file) {
            try { Remove-Item -LiteralPath $file -Force -ErrorAction Stop }
            catch { $errors += "rollout削除失敗(Codex起動中の可能性): $file : $_" }
        }
    }

    # 2. history.jsonl から session_id 一致行を除去
    $histPath = Join-Path $cx "history.jsonl"
    if ((Test-Path -LiteralPath $histPath) -and $sidSet.Count -gt 0) {
        try {
            Copy-Item -LiteralPath $histPath -Destination "$histPath.bak" -Force -ErrorAction SilentlyContinue
            $lines = Get-Content -LiteralPath $histPath -Encoding UTF8 -ErrorAction Stop
            $kept = New-Object System.Collections.Generic.List[string]
            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $drop = $false
                try {
                    $o = $line | ConvertFrom-Json -ErrorAction Stop
                    if ($o.PSObject.Properties['session_id'] -and $sidSet.Contains([string]$o.session_id)) { $drop = $true }
                } catch { $drop = $false }
                if (-not $drop) { $kept.Add($line) }
            }
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            $content = ($kept -join "`n"); if ($kept.Count -gt 0) { $content += "`n" }
            [System.IO.File]::WriteAllText($histPath, $content, $utf8NoBom)
        } catch {
            $errors += "history.jsonl編集失敗: $_"
        }
    }

    # 3. session_index.jsonl から id 一致行を除去
    $idxPath = Join-Path $cx "session_index.jsonl"
    if ((Test-Path -LiteralPath $idxPath) -and $sidSet.Count -gt 0) {
        try {
            Copy-Item -LiteralPath $idxPath -Destination "$idxPath.bak" -Force -ErrorAction SilentlyContinue
            $lines = Get-Content -LiteralPath $idxPath -Encoding UTF8 -ErrorAction Stop
            $kept = New-Object System.Collections.Generic.List[string]
            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $drop = $false
                try {
                    $o = $line | ConvertFrom-Json -ErrorAction Stop
                    if ($o.PSObject.Properties['id'] -and $sidSet.Contains([string]$o.id)) { $drop = $true }
                } catch { $drop = $false }
                if (-not $drop) { $kept.Add($line) }
            }
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            $content = ($kept -join "`n"); if ($kept.Count -gt 0) { $content += "`n" }
            [System.IO.File]::WriteAllText($idxPath, $content, $utf8NoBom)
        } catch {
            $errors += "session_index.jsonl編集失敗: $_"
        }
    }

    return $errors
}

# Codex 用 Dry-Run プラン
function New-CodexDeletionPlanText {
    param([Parameter(Mandatory)][ClaudeProject[]]$Projects)

    $sb = New-Object System.Text.StringBuilder
    $sep = ("─" * 60)
    $totalSize = 0L; $totalFiles = 0

    foreach ($p in $Projects) {
        [void]$sb.AppendLine($sep)
        [void]$sb.AppendLine("■ プロジェクト(cwd): $($p.Cwd)")
        [void]$sb.AppendLine("  環境: $($p.Environment)")
        [void]$sb.AppendLine("  CODEX_HOME: $($p.BasePath)")
        [void]$sb.AppendLine("  rolloutファイル数: $($p.Files.Count)  ($(Format-Size -Bytes $p.SizeBytes))")
        [void]$sb.AppendLine("  session_id数: $($p.SessionIds.Count)")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("  削除対象:")
        [void]$sb.AppendLine("    sessions/・archived_sessions/ 内の該当 rollout-*.jsonl（$($p.Files.Count) 件）")
        [void]$sb.AppendLine("    history.jsonl       （session_id 一致行を削除）")
        [void]$sb.AppendLine("    session_index.jsonl （id 一致行を削除）")
        [void]$sb.AppendLine("")
        $totalSize += $p.SizeBytes; $totalFiles += $p.Files.Count
    }

    [void]$sb.AppendLine($sep)
    [void]$sb.AppendLine("◆ 合計: $($Projects.Count) プロジェクト / rollout $totalFiles 件 / $(Format-Size -Bytes $totalSize)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("◆ 温存対象（v1ではファイル系のみ削除）:")
    [void]$sb.AppendLine("    state_5.sqlite / logs_2.sqlite / goals_1.sqlite / memories_1.sqlite（SQLiteは対象外）")
    [void]$sb.AppendLine("    config.toml / auth.json / AGENTS.md など設定類")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("※ 稼働中の Codex セッションは rollout がロックされ削除に失敗します。")
    [void]$sb.AppendLine("   対象プロジェクトの Codex を終了してから実行してください。")

    return $sb.ToString()
}

###############################################################
# 削除プラン生成（Dry-Run）
# 選択されたプロジェクト群に対し、削除対象を人間可読な文字列で返す
###############################################################
function New-DeletionPlanText {
    param(
        [Parameter(Mandatory)]
        [ClaudeProject[]]$Projects,

        # backups/（~/.claude.json 過去スナップショット）も削除するか
        [bool]$IncludeBackups = $false
    )

    $sb = New-Object System.Text.StringBuilder
    $totalSize = 0L
    $sep = ("─" * 60)

    foreach ($p in $Projects) {
        [void]$sb.AppendLine($sep)
        [void]$sb.AppendLine("■ プロジェクト: $($p.Cwd)")
        [void]$sb.AppendLine("  環境: $($p.Environment)")
        [void]$sb.AppendLine("  エンコード済みフォルダ: $($p.EncodedFolderName)")
        [void]$sb.AppendLine("  セッション数: $($p.SessionIds.Count)")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("  削除対象:")
        [void]$sb.AppendLine("    ~/.claude/projects/$($p.EncodedFolderName)/  ($(Format-Size -Bytes $p.SizeBytes))")
        $totalSize += $p.SizeBytes

        if ($p.SessionIds.Count -gt 0) {
            foreach ($sid in $p.SessionIds) {
                [void]$sb.AppendLine("    ~/.claude/file-history/$sid/")
                [void]$sb.AppendLine("    ~/.claude/tasks/$sid*  （旧 todos/ も対象）")
                [void]$sb.AppendLine("    ~/.claude/session-env/$sid/")
                [void]$sb.AppendLine("    ~/.claude/debug/$sid.txt")
            }
        }
        [void]$sb.AppendLine("")
        if ($p.Cwd -like '(cwd不明*') {
            [void]$sb.AppendLine("  ※ cwd不明のため history.jsonl / ~/.claude.json は編集されません（照合不可）")
        } else {
            [void]$sb.AppendLine("  該当行/エントリを除去:")
            [void]$sb.AppendLine("    ~/.claude/history.jsonl   （project/cwd が一致する行を削除）")
            [void]$sb.AppendLine("    ~/.claude.json            （プロジェクトエントリを削除）")
        }
        [void]$sb.AppendLine("")
    }

    [void]$sb.AppendLine($sep)
    [void]$sb.AppendLine("◆ 合計: $($Projects.Count) プロジェクト")
    [void]$sb.AppendLine("  projects/ 配下の合計サイズ: $(Format-Size -Bytes $totalSize)")
    [void]$sb.AppendLine("  （file-history/、tasks/、session-env/、debug/ のサイズは含まず）")
    [void]$sb.AppendLine("")

    if ($IncludeBackups) {
        # 対象環境（distro）ごとに backups/ を全削除する旨を明示
        $distros = @($Projects | ForEach-Object {
            if ($_.Distro) { "WSL2: $($_.Distro)" } else { "Windows" }
        } | Select-Object -Unique)
        [void]$sb.AppendLine("◆ backups/ を全削除（オプション有効・全プロジェクト共通／実害なし）:")
        foreach ($d in $distros) {
            [void]$sb.AppendLine("    [$d] ~/.claude/backups/   （~/.claude.json の過去スナップショット全て）")
        }
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("◆ 温存対象:")
        [void]$sb.AppendLine("    ~/.claude/shell-snapshots/")
        [void]$sb.AppendLine("    ~/.claude/settings.json")
        [void]$sb.AppendLine("    ~/.claude/plugins/")
    } else {
        [void]$sb.AppendLine("◆ 温存対象（公式 claude project purge と同様）:")
        [void]$sb.AppendLine("    ~/.claude/shell-snapshots/")
        [void]$sb.AppendLine("    ~/.claude/backups/")
        [void]$sb.AppendLine("    ~/.claude/settings.json")
        [void]$sb.AppendLine("    ~/.claude/plugins/")
    }

    return $sb.ToString()
}

###############################################################
# Windows側で1プロジェクト分の削除を実行
###############################################################
function Invoke-WindowsProjectDeletion {
    param(
        [Parameter(Mandatory)]
        [ClaudeProject]$Project
    )

    $errors = @()

    $claudeDir = if ($env:CLAUDE_CONFIG_DIR) { $env:CLAUDE_CONFIG_DIR } else { Join-Path $env:USERPROFILE ".claude" }
    $configJsonPath = Join-Path $env:USERPROFILE ".claude.json"

    # 1. projects/<encoded>/ 全体を削除
    $projectFolder = Join-Path $claudeDir "projects\$($Project.EncodedFolderName)"
    if (Test-Path $projectFolder) {
        try {
            Remove-Item -Path $projectFolder -Recurse -Force -ErrorAction Stop
        } catch {
            $errors += "projectsフォルダ削除失敗: $_"
        }
    }

    # 2. セッションIDで紐づく周辺ファイル群を削除
    foreach ($sid in $Project.SessionIds) {
        $targets = @(
            (Join-Path $claudeDir "file-history\$sid")
            (Join-Path $claudeDir "session-env\$sid")
            (Join-Path $claudeDir "debug\$sid.txt")
        )
        foreach ($t in $targets) {
            if (Test-Path $t) {
                try { Remove-Item -Path $t -Recurse -Force -ErrorAction Stop }
                catch { $errors += "${t}: $_" }
            }
        }
        # tasks/（現行）および todos/（旧称・後方互換）から、
        # セッションIDで始まるエントリを削除。
        # 命名は <sid>.json / <sid>-agent-*.json 双方がありうるため $sid* で総当たり。
        foreach ($taskDir in @("tasks", "todos")) {
            $taskBase = Join-Path $claudeDir $taskDir
            if (Test-Path $taskBase) {
                Get-ChildItem -Path $taskBase -Filter "$sid*" -ErrorAction SilentlyContinue | ForEach-Object {
                    try { Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop }
                    catch { $errors += "$($_.FullName): $_" }
                }
            }
        }
    }

    # 3. history.jsonl から該当プロジェクト行を除去
    $historyPath = Join-Path $claudeDir "history.jsonl"
    if (Test-Path $historyPath) {
        try {
            # 書き換え前にバックアップを取得
            Copy-Item -Path $historyPath -Destination "$historyPath.bak" -Force -ErrorAction SilentlyContinue

            $lines = Get-Content -Path $historyPath -Encoding UTF8 -ErrorAction Stop
            $filtered = New-Object System.Collections.Generic.List[string]
            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $keep = $true
                try {
                    $obj = $line | ConvertFrom-Json -ErrorAction Stop
                    # history.jsonl のパスは "project" キー（cwd ではない）。
                    # 旧スキーマ互換で cwd も照合する。
                    # 末尾の / \ 差異で取りこぼさないよう両側を正規化して比較
                    $pcwd = $Project.Cwd.TrimEnd('\', '/')
                    if (
                        ($obj.PSObject.Properties['project'] -and (([string]$obj.project).TrimEnd('\', '/') -eq $pcwd)) -or
                        ($obj.PSObject.Properties['cwd'] -and (([string]$obj.cwd).TrimEnd('\', '/') -eq $pcwd))
                    ) {
                        $keep = $false
                    }
                } catch {
                    # パース不能な行はそのまま残す（破壊回避）
                    $keep = $true
                }
                if ($keep) { $filtered.Add($line) }
            }
            # WriteAllLines は Windows で CRLF になる。WSL版と同様に LF を明示して
            # 改行コードを統一する（Node系ツールの慣例・混在回避）。
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            $content = ($filtered -join "`n")
            if ($filtered.Count -gt 0) { $content += "`n" }
            [System.IO.File]::WriteAllText($historyPath, $content, $utf8NoBom)
        } catch {
            $errors += "history.jsonl編集失敗: $_"
        }
    }

    # 4. ~/.claude.json から該当プロジェクトエントリを削除
    if (Test-Path $configJsonPath) {
        try {
            # 書き換え前にバックアップを取得
            Copy-Item -Path $configJsonPath -Destination "$configJsonPath.bak" -Force -ErrorAction SilentlyContinue

            $raw = Get-Content -Path $configJsonPath -Raw -Encoding UTF8 -ErrorAction Stop
            # .claude.json は空文字 "" キーを含むことがあり、PSCustomObject では
            # 表現できず ConvertFrom-Json が失敗する。-AsHashTable で回避する。
            $config = $raw | ConvertFrom-Json -AsHashTable -ErrorAction Stop
            if ($config.ContainsKey('projects') -and $config.projects) {
                # projects は cwd をキーにしたハッシュテーブル
                if ($config.projects.ContainsKey($Project.Cwd)) {
                    [void]$config.projects.Remove($Project.Cwd)
                    $serialized = $config | ConvertTo-Json -Depth 100
                    $serialized = $serialized -replace "`r`n", "`n"   # WSL版と同様 LF に統一
                    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
                    [System.IO.File]::WriteAllText($configJsonPath, $serialized, $utf8NoBom)
                }
            }
        } catch {
            $errors += "~/.claude.json編集失敗: $_"
        }
    }

    return $errors
}

###############################################################
# WSL2側で1プロジェクト分の削除を実行
###############################################################
function Invoke-WslProjectDeletion {
    param(
        [Parameter(Mandatory)]
        [ClaudeProject]$Project
    )

    $errors = @()
    $distro = $Project.Distro

    # 1. WSL の HOME パスを取得（--exec で軽量起動）
    $wslHome = $null
    try {
        $wslHomeRaw = wsl -d $distro --exec bash -c 'echo $HOME' 2>$null
        if ($wslHomeRaw) {
            $wslHome = (($wslHomeRaw | Out-String) -replace "`0", "").Trim()
        }
    } catch {
        # 取得失敗
    }
    if ([string]::IsNullOrWhiteSpace($wslHome)) {
        $errors += "WSL2 のホームパスが取得できませんでした"
        return $errors
    }

    # 2. UNCパス組み立て
    $uncBase = "\\wsl.localhost\$distro" + ($wslHome -replace '/', '\')
    $claudeDir = Join-Path $uncBase ".claude"
    $projectFolder = Join-Path $claudeDir "projects\$($Project.EncodedFolderName)"
    $historyPath = Join-Path $claudeDir "history.jsonl"
    $configPath = Join-Path $uncBase ".claude.json"

    # 3. プロジェクトフォルダ削除
    if (Test-Path -LiteralPath $projectFolder) {
        try {
            Remove-Item -LiteralPath $projectFolder -Recurse -Force -ErrorAction Stop
        } catch {
            $errors += "WSL側 プロジェクトフォルダ削除失敗: $_"
        }
    }

    # 4. セッションIDで紐づく周辺ファイル群を削除
    foreach ($sid in $Project.SessionIds) {
        $targets = @(
            (Join-Path $claudeDir "file-history\$sid")
            (Join-Path $claudeDir "session-env\$sid")
            (Join-Path $claudeDir "debug\$sid.txt")
        )
        foreach ($t in $targets) {
            if (Test-Path -LiteralPath $t) {
                try { Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction Stop }
                catch { $errors += "WSL側 ${t}: $_" }
            }
        }
        # tasks/（現行）および todos/（旧称・後方互換）から、
        # セッションIDで始まるエントリを削除。
        # 命名は <sid>.json / <sid>-agent-*.json 双方がありうるため $sid* で総当たり。
        # セッションIDは UUID 形式（[a-f0-9-]+ のみ）なので -Filter のワイルドカードに
        # 特殊文字が混入する懸念はない。
        foreach ($taskDir in @("tasks", "todos")) {
            $taskBase = Join-Path $claudeDir $taskDir
            if (Test-Path -LiteralPath $taskBase) {
                Get-ChildItem -LiteralPath $taskBase -Filter "$sid*" -ErrorAction SilentlyContinue | ForEach-Object {
                    try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop }
                    catch { $errors += "WSL側 $($_.FullName): $_" }
                }
            }
        }
    }

    # 5. history.jsonl から該当cwd行を除去
    if (Test-Path -LiteralPath $historyPath) {
        try {
            # 書き換え前にバックアップを取得
            Copy-Item -LiteralPath $historyPath -Destination "$historyPath.bak" -Force -ErrorAction SilentlyContinue

            $lines = Get-Content -LiteralPath $historyPath -Encoding UTF8 -ErrorAction Stop
            $filtered = New-Object System.Collections.Generic.List[string]
            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $keep = $true
                try {
                    $obj = $line | ConvertFrom-Json -ErrorAction Stop
                    # history.jsonl のパスは "project" キー（cwd ではない）。
                    # 旧スキーマ互換で cwd も照合する。
                    # 末尾の / \ 差異で取りこぼさないよう両側を正規化して比較
                    $pcwd = $Project.Cwd.TrimEnd('\', '/')
                    if (
                        ($obj.PSObject.Properties['project'] -and (([string]$obj.project).TrimEnd('\', '/') -eq $pcwd)) -or
                        ($obj.PSObject.Properties['cwd'] -and (([string]$obj.cwd).TrimEnd('\', '/') -eq $pcwd))
                    ) {
                        $keep = $false
                    }
                } catch {
                    # パース不能行はそのまま残す
                    $keep = $true
                }
                if ($keep) { $filtered.Add($line) }
            }

            # WSL側ファイルはLF改行を保つ（CRLFになるとJSONLパーサが壊れる）
            # 全行削除されて $filtered.Count -eq 0 になるケースも正常な結果として
            # 受け入れる（ゼロバイトファイルになる、Claude Code側も再生成可）
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            $content = ($filtered -join "`n")
            if ($filtered.Count -gt 0) { $content += "`n" }
            [System.IO.File]::WriteAllText($historyPath, $content, $utf8NoBom)
        } catch {
            $errors += "WSL側 history.jsonl編集失敗: $_"
        }
    }

    # 6. .claude.json から該当プロジェクトエントリを削除
    if (Test-Path -LiteralPath $configPath) {
        try {
            # 書き換え前にバックアップを取得
            Copy-Item -LiteralPath $configPath -Destination "$configPath.bak" -Force -ErrorAction SilentlyContinue

            $raw = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 -ErrorAction Stop
            # .claude.json は空文字 "" キーを含むことがあり、PSCustomObject では
            # 表現できず ConvertFrom-Json が失敗する。-AsHashTable で回避する。
            $config = $raw | ConvertFrom-Json -AsHashTable -ErrorAction Stop
            if ($config.ContainsKey('projects') -and $config.projects) {
                if ($config.projects.ContainsKey($Project.Cwd)) {
                    [void]$config.projects.Remove($Project.Cwd)
                    $serialized = $config | ConvertTo-Json -Depth 100

                    # WSL側ファイルはLF改行を保つ
                    $serialized = $serialized -replace "`r`n", "`n"

                    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
                    [System.IO.File]::WriteAllText($configPath, $serialized, $utf8NoBom)
                }
            }
        } catch {
            $errors += "WSL側 ~/.claude.json編集失敗: $_"
        }
    }

    return $errors
}

###############################################################
# 環境に応じて削除をディスパッチ
###############################################################
function Invoke-ProjectDeletion {
    param(
        [Parameter(Mandatory)]
        [ClaudeProject]$Project
    )

    if ($Project.Distro) {
        return Invoke-WslProjectDeletion -Project $Project
    } else {
        return Invoke-WindowsProjectDeletion -Project $Project
    }
}

###############################################################
# backups/（~/.claude.json の過去スナップショット）を全削除
# - プロジェクト単位ではなく環境（Windows / 各WSL distro）単位の全削除
# - 公式 claude project purge は温存するが、手動削除しても
#   "Nothing user-facing"（実害なし）とドキュメントに明記。
#   ロールバック起点を失う点のみ留意。
# - $Distro が null なら Windows、指定があればそのWSL distroを対象
###############################################################
function Remove-ClaudeBackups {
    param(
        [string]$Distro
    )

    $errors = @()

    if ($Distro) {
        # WSL: HOME を解決して UNC パスを組み立て
        $wslHome = $null
        try {
            $wslHomeRaw = wsl -d $Distro --exec bash -c 'echo $HOME' 2>$null
            if ($wslHomeRaw) {
                $wslHome = (($wslHomeRaw | Out-String) -replace "`0", "").Trim()
            }
        } catch { }
        if ([string]::IsNullOrWhiteSpace($wslHome)) {
            $errors += "WSL2($Distro) のホームパスが取得できず backups/ を削除できませんでした"
            return $errors
        }
        $uncBase = "\\wsl.localhost\$Distro" + ($wslHome -replace '/', '\')
        $backupsDir = Join-Path $uncBase ".claude\backups"
        if (Test-Path -LiteralPath $backupsDir) {
            try { Remove-Item -LiteralPath $backupsDir -Recurse -Force -ErrorAction Stop }
            catch { $errors += "WSL側 backups/削除失敗: $_" }
        }
    } else {
        # Windows
        $claudeDir = if ($env:CLAUDE_CONFIG_DIR) { $env:CLAUDE_CONFIG_DIR } else { Join-Path $env:USERPROFILE ".claude" }
        $backupsDir = Join-Path $claudeDir "backups"
        if (Test-Path $backupsDir) {
            try { Remove-Item -Path $backupsDir -Recurse -Force -ErrorAction Stop }
            catch { $errors += "backups/削除失敗: $_" }
        }
    }

    return $errors
}

###############################################################
# サイズ表示用フォーマッタ
###############################################################
function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

###############################################################
# ===== セッション単位の閲覧 / 会話履歴抽出 / 個別削除 =====
###############################################################

# 共有読み取りで開いた StreamReader を返す（稼働中の jsonl も読めるように）
function Open-SharedReader {
    param([Parameter(Mandatory)][string]$Path)
    $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    return (New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8))
}

# プロジェクト配下の各セッションのメタ情報（日時/タイトル/件数/サイズ/ID/ファイル）を返す
function Get-SessionInfos {
    param([Parameter(Mandatory)][ClaudeProject]$Project)

    # Codex はタイトルを session_index.jsonl(thread_name) から引けると分かりやすい
    $idxMap = @{}
    if ($Project.Kind -eq 'Codex' -and $Project.BasePath) {
        $idxPath = Join-Path $Project.BasePath "session_index.jsonl"
        if (Test-Path -LiteralPath $idxPath) {
            try {
                foreach ($l in (Get-Content -LiteralPath $idxPath -Encoding UTF8 -ErrorAction Stop)) {
                    if ([string]::IsNullOrWhiteSpace($l)) { continue }
                    try { $o = $l | ConvertFrom-Json -ErrorAction Stop; if ($o.id) { $idxMap[[string]$o.id] = [string]$o.thread_name } } catch {}
                }
            } catch {}
        }
    }

    $infos = @()
    foreach ($file in $Project.Files) {
        if (-not (Test-Path -LiteralPath $file)) { continue }
        $fi = Get-Item -LiteralPath $file -ErrorAction SilentlyContinue
        $id = [System.IO.Path]::GetFileNameWithoutExtension($file)
        if ($Project.Kind -eq 'Codex') { $sid = Get-CodexSessionId -BaseName $id; if ($sid) { $id = $sid } }

        $title = $null; $count = 0
        if ($Project.Kind -eq 'Codex' -and $idxMap.ContainsKey($id) -and $idxMap[$id]) { $title = $idxMap[$id] }

        $sr = $null
        try {
            $sr = Open-SharedReader -Path $file
            while ($null -ne ($line = $sr.ReadLine())) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if ($Project.Kind -eq 'Codex') {
                    if ($line -like '*"type":"response_item"*' -and $line -like '*"type":"message"*' -and ($line -like '*"role":"user"*' -or $line -like '*"role":"assistant"*')) {
                        $count++
                        if (-not $title -and $line -like '*"role":"user"*') {
                            try { $o = $line | ConvertFrom-Json; $t = [string]($o.payload.content | Where-Object { $_.text } | Select-Object -First 1).text; if ($t -and $t -notmatch '^\s*<(environment_context|user_instructions|environment_details)') { $title = $t } } catch {}
                        }
                    }
                } else {
                    if ($line -like '*"type":"user"*' -or $line -like '*"type":"assistant"*') { $count++ }
                    if (-not $title -and $line -like '*"type":"ai-title"*') {
                        try { $o = $line | ConvertFrom-Json; if ($o.aiTitle) { $title = [string]$o.aiTitle } } catch {}
                    }
                    if (-not $title -and $line -like '*"type":"user"*') {
                        try { $o = $line | ConvertFrom-Json; if ($o.message.content -is [string]) { $title = [string]$o.message.content } } catch {}
                    }
                }
            }
        } catch {} finally { if ($sr) { $sr.Dispose() } }

        if ([string]::IsNullOrWhiteSpace($title)) { $title = "(無題)" }
        $title = ($title -replace '\s+', ' ').Trim()
        if ($title.Length -gt 70) { $title = $title.Substring(0, 70) + '…' }

        $infos += [pscustomobject]@{
            Id = $id; Title = $title; Date = $fi.LastWriteTime
            MsgCount = $count; SizeBytes = $fi.Length; File = $file
        }
    }
    return , (@($infos) | Sort-Object Date -Descending)
}

# 1セッションの会話を Markdown へ書き出す（IncludeThinkingで思考も）
function Export-SessionConversation {
    param(
        [Parameter(Mandatory)][ClaudeProject]$Project,
        [Parameter(Mandatory)][pscustomobject]$Info,
        [bool]$IncludeThinking = $false,
        [Parameter(Mandatory)][string]$OutPath
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# 会話履歴")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("- プロジェクト(cwd): $($Project.Cwd)")
    [void]$sb.AppendLine("- 対象: $($Project.Kind) / 環境: $($Project.Environment)")
    [void]$sb.AppendLine("- セッションID: $($Info.Id)")
    [void]$sb.AppendLine("- 最終更新: $($Info.Date.ToString('yyyy-MM-dd HH:mm'))")
    [void]$sb.AppendLine("- タイトル: $($Info.Title)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("---")
    [void]$sb.AppendLine("")

    $emit = {
        param($role, $text)
        if ([string]::IsNullOrWhiteSpace($text)) { return }
        $label = switch ($role) {
            'user' { 'User' }; 'assistant' { 'Assistant' }
            'thinking' { 'Thinking' }; default { $role }
        }
        [void]$sb.AppendLine("## $label")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine([string]$text)
        [void]$sb.AppendLine("")
    }

    $sr = $null
    try {
        $sr = Open-SharedReader -Path $Info.File
        while ($null -ne ($line = $sr.ReadLine())) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $o = $null; try { $o = $line | ConvertFrom-Json -ErrorAction Stop } catch { continue }

            if ($Project.Kind -eq 'Codex') {
                if ($o.type -ne 'response_item') { continue }
                $pt = $o.payload.type
                if ($pt -eq 'message') {
                    $role = [string]$o.payload.role
                    if ($role -ne 'user' -and $role -ne 'assistant') { continue }   # developer/system は除外
                    $txt = (($o.payload.content | Where-Object { $_.text }) | ForEach-Object { [string]$_.text }) -join "`n"
                    # Codex がuserロールで注入する環境/指示ブロックは会話本文ではないので除外
                    if ($role -eq 'user' -and $txt -match '^\s*<(environment_context|user_instructions|environment_details)') { continue }
                    & $emit $role $txt
                } elseif ($pt -eq 'reasoning' -and $IncludeThinking) {
                    $txt = (($o.payload.summary | Where-Object { $_.text }) | ForEach-Object { [string]$_.text }) -join "`n"
                    & $emit 'thinking' $txt
                }
            } else {
                if ($o.type -eq 'user') {
                    $c = $o.message.content
                    if ($c -is [string]) { & $emit 'user' $c }
                    else {
                        $txt = (($c | Where-Object { $_.type -eq 'text' }) | ForEach-Object { [string]$_.text }) -join "`n"
                        & $emit 'user' $txt
                    }
                } elseif ($o.type -eq 'assistant') {
                    $c = $o.message.content
                    if ($c -is [string]) { & $emit 'assistant' $c }
                    else {
                        foreach ($item in $c) {
                            if ($item.type -eq 'text') { & $emit 'assistant' ([string]$item.text) }
                            elseif ($item.type -eq 'thinking' -and $IncludeThinking) { & $emit 'thinking' ([string]$item.thinking) }
                        }
                    }
                }
            }
        }
    } finally { if ($sr) { $sr.Dispose() } }

    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($OutPath, $sb.ToString(), $utf8)
}

# 1セッションだけ削除（ファイル系のみ）
function Remove-SingleSession {
    param(
        [Parameter(Mandatory)][ClaudeProject]$Project,
        [Parameter(Mandatory)][pscustomobject]$Info
    )
    $errors = @()
    $base = $Project.BasePath
    $sid = $Info.Id

    # 本体ファイル（jsonl / rollout）
    if (Test-Path -LiteralPath $Info.File) {
        try { Remove-Item -LiteralPath $Info.File -Force -ErrorAction Stop }
        catch { $errors += "本体削除失敗(セッション起動中の可能性): $($Info.File) : $_" }
    }

    if ($Project.Kind -eq 'Codex') {
        # history.jsonl / session_index.jsonl から session_id 一致行を除去
        foreach ($pair in @(@{Path = (Join-Path $base 'history.jsonl'); Key = 'session_id' }, @{Path = (Join-Path $base 'session_index.jsonl'); Key = 'id' })) {
            $fp = $pair.Path; $key = $pair.Key
            if (-not (Test-Path -LiteralPath $fp)) { continue }
            try {
                Copy-Item -LiteralPath $fp -Destination "$fp.bak" -Force -ErrorAction SilentlyContinue
                $kept = New-Object System.Collections.Generic.List[string]
                foreach ($line in (Get-Content -LiteralPath $fp -Encoding UTF8 -ErrorAction Stop)) {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    $drop = $false
                    try { $o = $line | ConvertFrom-Json -ErrorAction Stop; if ($o.PSObject.Properties[$key] -and ([string]$o.$key -eq $sid)) { $drop = $true } } catch { $drop = $false }
                    if (-not $drop) { $kept.Add($line) }
                }
                $utf8 = New-Object System.Text.UTF8Encoding $false
                $content = ($kept -join "`n"); if ($kept.Count -gt 0) { $content += "`n" }
                [System.IO.File]::WriteAllText($fp, $content, $utf8)
            } catch { $errors += "$([System.IO.Path]::GetFileName($fp))編集失敗: $_" }
        }
    } else {
        # Claude: sid 紐づけの周辺ファイルのみ削除
        # （history.jsonl / .claude.json はプロジェクト単位＝session_idで引けないため対象外）
        foreach ($t in @((Join-Path $base "file-history\$sid"), (Join-Path $base "session-env\$sid"), (Join-Path $base "debug\$sid.txt"))) {
            if (Test-Path -LiteralPath $t) {
                try { Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction Stop } catch { $errors += "${t}: $_" }
            }
        }
        foreach ($td in @('tasks', 'todos')) {
            $tb = Join-Path $base $td
            if (Test-Path -LiteralPath $tb) {
                Get-ChildItem -LiteralPath $tb -Filter "$sid*" -ErrorAction SilentlyContinue | ForEach-Object {
                    try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop } catch { $errors += "$($_.FullName): $_" }
                }
            }
        }
    }
    return $errors
}

###############################################################
# セッション閲覧ダイアログ（行UI・抽出/個別削除）
###############################################################
function Show-SessionBrowser {
    param([Parameter(Mandatory)][ClaudeProject]$Project)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "セッション一覧 - $($Project.Cwd)"
    $dlg.Size = New-Object System.Drawing.Size(860, 520)
    $dlg.MinimumSize = New-Object System.Drawing.Size(640, 380)
    $dlg.StartPosition = "CenterParent"

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "このプロジェクトのセッション（チェックして抽出/削除）:"
    $lbl.Location = New-Object System.Drawing.Point(10, 10)
    $lbl.Size = New-Object System.Drawing.Size(820, 20)
    $dlg.Controls.Add($lbl)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = New-Object System.Drawing.Point(10, 35)
    $lv.Size = New-Object System.Drawing.Size(825, 390)
    $lv.View = [System.Windows.Forms.View]::Details
    $lv.CheckBoxes = $true
    $lv.FullRowSelect = $true
    $lv.GridLines = $true
    $lv.Anchor = "Top,Left,Bottom,Right"
    [void]$lv.Columns.Add("最終更新", 130)
    [void]$lv.Columns.Add("タイトル", 430)
    [void]$lv.Columns.Add("件数", 55)
    [void]$lv.Columns.Add("サイズ", 90)
    [void]$lv.Columns.Add("ID", 110)
    $dlg.Controls.Add($lv)

    $thinkChk = New-Object System.Windows.Forms.CheckBox
    $thinkChk.Text = "thinking(思考)も含めて抽出"
    $thinkChk.Location = New-Object System.Drawing.Point(10, 432)
    $thinkChk.Size = New-Object System.Drawing.Size(220, 24)
    $thinkChk.Anchor = "Bottom,Left"
    $dlg.Controls.Add($thinkChk)

    $exportBtn = New-Object System.Windows.Forms.Button
    $exportBtn.Text = "会話履歴を抽出(.md)"
    $exportBtn.Location = New-Object System.Drawing.Point(380, 430)
    $exportBtn.Size = New-Object System.Drawing.Size(160, 30)
    $exportBtn.BackColor = [System.Drawing.Color]::LightBlue
    $exportBtn.Anchor = "Bottom,Right"
    $dlg.Controls.Add($exportBtn)

    $delBtn = New-Object System.Windows.Forms.Button
    $delBtn.Text = "選択セッションを削除"
    $delBtn.Location = New-Object System.Drawing.Point(550, 430)
    $delBtn.Size = New-Object System.Drawing.Size(160, 30)
    $delBtn.BackColor = [System.Drawing.Color]::LightCoral
    $delBtn.Anchor = "Bottom,Right"
    $dlg.Controls.Add($delBtn)

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "閉じる"
    $closeBtn.Location = New-Object System.Drawing.Point(720, 430)
    $closeBtn.Size = New-Object System.Drawing.Size(115, 30)
    $closeBtn.Anchor = "Bottom,Right"
    $closeBtn.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($closeBtn)

    $script:BrowserInfos = @()
    $populate = {
        $lv.Items.Clear()
        $script:BrowserInfos = @(Get-SessionInfos -Project $Project)
        for ($i = 0; $i -lt $script:BrowserInfos.Count; $i++) {
            $s = $script:BrowserInfos[$i]
            $item = New-Object System.Windows.Forms.ListViewItem($s.Date.ToString("yyyy-MM-dd HH:mm"))
            [void]$item.SubItems.Add($s.Title)
            [void]$item.SubItems.Add([string]$s.MsgCount)
            [void]$item.SubItems.Add((Format-Size -Bytes $s.SizeBytes))
            [void]$item.SubItems.Add($s.Id.Substring(0, [Math]::Min(8, $s.Id.Length)))
            $item.Tag = $i
            [void]$lv.Items.Add($item)
        }
        $lbl.Text = "このプロジェクトのセッション: $($script:BrowserInfos.Count) 件（チェックして抽出/削除）"
    }

    $getChecked = {
        $sel = @()
        foreach ($it in $lv.Items) { if ($it.Checked) { $sel += $script:BrowserInfos[[int]$it.Tag] } }
        return $sel
    }

    # 既定ファイル名を生成: <種別>_<プロジェクト名>_<日時>_<タイトル>_<id8>.md
    $makeFileName = {
        param($s)
        $leaf = ($Project.Cwd.TrimEnd('\', '/') -replace '.*[\\/]', '')
        $leaf = ($leaf -replace '[\\/:*?"<>|]', '_')
        if ([string]::IsNullOrWhiteSpace($leaf)) { $leaf = "session" }
        $title = ($s.Title -replace '[\\/:*?"<>|]', '_') -replace '\s+', '_'
        if ($title.Length -gt 24) { $title = $title.Substring(0, 24) }
        if ([string]::IsNullOrWhiteSpace($title) -or $title -eq '(無題)') { $title = 'untitled' }
        $idShort = $s.Id.Substring(0, [Math]::Min(8, $s.Id.Length))
        $kind = $Project.Kind
        $name = "{0}_{1}_{2}_{3}_{4}.md" -f $kind, $leaf, $s.Date.ToString('yyyyMMdd-HHmm'), $title, $idShort
        # 念のため最終的な不正文字を除去
        return ($name -replace '[\\/:*?"<>|]', '_')
    }

    $exportBtn.Add_Click({
        $sel = & $getChecked
        if ($sel.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show("抽出するセッションをチェックしてください", "情報", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        if ($sel.Count -eq 1) {
            # 1件: ファイル名欄付きの保存ダイアログ（既定名を投入）
            $s = $sel[0]
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Title = "会話履歴の保存（ファイル名を確認）"
            $sfd.Filter = "Markdown (*.md)|*.md|すべてのファイル (*.*)|*.*"
            $sfd.DefaultExt = "md"
            $sfd.AddExtension = $true
            $sfd.FileName = (& $makeFileName $s)
            if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            try {
                Export-SessionConversation -Project $Project -Info $s -IncludeThinking:$thinkChk.Checked -OutPath $sfd.FileName
                [void][System.Windows.Forms.MessageBox]::Show("抽出しました。`n$($sfd.FileName)", "抽出完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [void][System.Windows.Forms.MessageBox]::Show("抽出失敗: $_", "エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            return
        }

        # 複数: 保存先フォルダを選び、各セッションを既定名で書き出し
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "会話履歴(.md)の保存先フォルダを選択（ファイル名は自動）"
        if ($fbd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $outDir = $fbd.SelectedPath
        $ok = 0; $errs = @()
        foreach ($s in $sel) {
            $out = Join-Path $outDir (& $makeFileName $s)
            try { Export-SessionConversation -Project $Project -Info $s -IncludeThinking:$thinkChk.Checked -OutPath $out; $ok++ }
            catch { $errs += "$($s.Id): $_" }
        }
        $msg = "$ok / $($sel.Count) セッションを抽出しました。`n保存先: $outDir"
        if ($errs.Count -gt 0) { $msg += "`n`nエラー:`n" + ($errs -join "`n") }
        [void][System.Windows.Forms.MessageBox]::Show($msg, "抽出完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    $delBtn.Add_Click({
        $sel = & $getChecked
        if ($sel.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show("削除するセッションをチェックしてください", "情報", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $confirm = "以下の $($sel.Count) セッションを削除します。`n`n"
        foreach ($s in $sel) { $confirm += "  ・ $($s.Date.ToString('MM-dd HH:mm'))  $($s.Title)`n" }
        if ($Project.Kind -ne 'Codex') { $confirm += "`n（Claudeの history.jsonl / .claude.json はプロジェクト単位のため対象外）" }
        $confirm += "`nこの操作は取り消せません。続行しますか?"
        $r = [System.Windows.Forms.MessageBox]::Show($confirm, "セッション削除確認", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $ok = 0; $errs = @()
        foreach ($s in $sel) {
            try { $e = Remove-SingleSession -Project $Project -Info $s; if ($e -and $e.Count -gt 0) { $errs += $e } else { $ok++ } }
            catch { $errs += "$($s.Id): $_" }
        }
        $msg = "$ok / $($sel.Count) セッションを削除しました。"
        if ($errs.Count -gt 0) { $msg += "`n`n警告/エラー:`n" + ($errs -join "`n") }
        [void][System.Windows.Forms.MessageBox]::Show($msg, "削除完了", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        & $populate
    })

    & $populate
    [void]$dlg.ShowDialog($form)
}

###############################################################
# 起動中の claude.exe セッションを列挙
# - PEB から cwd を読み、開いているプロジェクトを識別
# - プロジェクトの jsonl 最終更新から「無活動分」を算出（活動の目安）
# - 保護判定: 本体(このツールの親)/Chrome連携/稼働中(無活動5分未満)
###############################################################
function Get-ClaudeSessions {
    param([int]$ActiveThresholdMinutes = 5)

    $base = if ($env:CLAUDE_CONFIG_DIR) {
        Join-Path $env:CLAUDE_CONFIG_DIR "projects"
    } else {
        Join-Path $env:USERPROFILE ".claude\projects"
    }

    # このツール自身の claude.exe 祖先（あれば保護）
    $hostClaude = $null
    try {
        $cur = $PID
        for ($i = 0; $i -lt 12 -and $cur; $i++) {
            $pr = Get-CimInstance Win32_Process -Filter "ProcessId=$cur" -ErrorAction SilentlyContinue
            if (-not $pr) { break }
            if ($pr.Name -eq 'claude.exe') { $hostClaude = [int]$pr.ProcessId; break }
            $cur = $pr.ParentProcessId
        }
    } catch { }

    $list = @()
    Get-CimInstance Win32_Process -Filter "Name='claude.exe'" -ErrorAction SilentlyContinue | ForEach-Object {
        $p = $_
        $isChrome = $p.CommandLine -like '*--chrome-native-host*'

        $cwd = $null
        try { $cwd = [ClaudeProcPeb]::GetCwd([int]$p.ProcessId) } catch { }
        $project = if ($cwd) { ($cwd.TrimEnd('\', '/') -replace '.*[\\/]', '') } else { '(不明)' }

        $idle = $null; $lastAct = $null
        if ($cwd -and -not $isChrome) {
            $enc = (($cwd.TrimEnd('\', '/')) -replace '[^a-zA-Z0-9]', '-')
            $pf = Join-Path $base $enc
            if (Test-Path -LiteralPath $pf) {
                $latest = Get-ChildItem -LiteralPath $pf -Filter '*.jsonl' -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($latest) {
                    $lastAct = $latest.LastWriteTime
                    $idle = [int]((Get-Date) - $latest.LastWriteTime).TotalMinutes
                }
            }
        }

        $protected = $false; $reason = ''
        if ($hostClaude -and [int]$p.ProcessId -eq $hostClaude) { $protected = $true; $reason = '本体(このツールの親)' }
        elseif ($isChrome) { $protected = $true; $reason = 'Chrome連携' }
        elseif ($null -ne $idle -and $idle -lt $ActiveThresholdMinutes) { $protected = $true; $reason = "稼働中(${idle}分)" }

        $list += [pscustomobject]@{
            ProcessId    = [int]$p.ProcessId
            Project      = $project
            Cwd          = $cwd
            LastActivity = $lastAct
            IdleMinutes  = $idle
            IsChrome     = $isChrome
            Protected    = $protected
            Reason       = $reason
        }
    }
    return , $list
}

###############################################################
# セッション管理ダイアログ（保護ガード付きで claude.exe を終了）
###############################################################
function Show-SessionManager {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Claude セッション管理"
    $dlg.Size = New-Object System.Drawing.Size(720, 480)
    $dlg.MinimumSize = New-Object System.Drawing.Size(560, 360)
    $dlg.StartPosition = "CenterParent"

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "起動中の claude.exe セッション（保護対象=灰色はチェック不可）:"
    $lbl.Location = New-Object System.Drawing.Point(10, 10)
    $lbl.Size = New-Object System.Drawing.Size(680, 20)
    $dlg.Controls.Add($lbl)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = New-Object System.Drawing.Point(10, 35)
    $lv.Size = New-Object System.Drawing.Size(685, 360)
    $lv.View = [System.Windows.Forms.View]::Details
    $lv.CheckBoxes = $true
    $lv.FullRowSelect = $true
    $lv.GridLines = $true
    $lv.Anchor = "Top,Left,Bottom,Right"
    [void]$lv.Columns.Add("PID", 70)
    [void]$lv.Columns.Add("プロジェクト", 210)
    [void]$lv.Columns.Add("最終活動", 130)
    [void]$lv.Columns.Add("無活動(分)", 80)
    [void]$lv.Columns.Add("状態", 175)
    $dlg.Controls.Add($lv)

    $statusL = New-Object System.Windows.Forms.Label
    $statusL.Location = New-Object System.Drawing.Point(110, 412)
    $statusL.Size = New-Object System.Drawing.Size(350, 20)
    $statusL.Anchor = "Bottom,Left"
    $dlg.Controls.Add($statusL)

    $refresh = New-Object System.Windows.Forms.Button
    $refresh.Text = "再取得"
    $refresh.Location = New-Object System.Drawing.Point(10, 405)
    $refresh.Size = New-Object System.Drawing.Size(90, 30)
    $refresh.Anchor = "Bottom,Left"
    $dlg.Controls.Add($refresh)

    $killBtn = New-Object System.Windows.Forms.Button
    $killBtn.Text = "選択セッションを終了"
    $killBtn.Location = New-Object System.Drawing.Point(465, 405)
    $killBtn.Size = New-Object System.Drawing.Size(165, 30)
    $killBtn.BackColor = [System.Drawing.Color]::LightCoral
    $killBtn.Anchor = "Bottom,Right"
    $dlg.Controls.Add($killBtn)

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "閉じる"
    $closeBtn.Location = New-Object System.Drawing.Point(635, 405)
    $closeBtn.Size = New-Object System.Drawing.Size(60, 30)
    $closeBtn.Anchor = "Bottom,Right"
    $closeBtn.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($closeBtn)

    # 保護アイテムはチェックさせない
    $lv.Add_ItemCheck({
        param($s, $e)
        $it = $lv.Items[$e.Index]
        if ($it.Tag -and $it.Tag.Protected -and $e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
        }
    })

    $populate = {
        $lv.Items.Clear()
        $sessions = @(Get-ClaudeSessions) | Sort-Object @{ Expression = { if ($null -eq $_.IdleMinutes) { -1 } else { $_.IdleMinutes } }; Descending = $true }
        foreach ($sess in $sessions) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$sess.ProcessId)
            [void]$item.SubItems.Add([string]$sess.Project)
            [void]$item.SubItems.Add($(if ($sess.LastActivity) { $sess.LastActivity.ToString('MM-dd HH:mm') } else { '-' }))
            [void]$item.SubItems.Add($(if ($null -ne $sess.IdleMinutes) { [string]$sess.IdleMinutes } else { '-' }))
            [void]$item.SubItems.Add($(if ($sess.Protected) { "保護: $($sess.Reason)" } else { "終了可" }))
            if ($sess.Protected) { $item.ForeColor = [System.Drawing.Color]::Gray }
            $item.Tag = $sess
            [void]$lv.Items.Add($item)
        }
        $statusL.Text = "$($sessions.Count) セッション（保護 $(@($sessions | Where-Object Protected).Count) / 終了可 $(@($sessions | Where-Object { -not $_.Protected }).Count)）"
    }
    $refresh.Add_Click($populate)

    $killBtn.Add_Click({
        $targets = @()
        foreach ($it in $lv.Items) {
            if ($it.Checked -and $it.Tag -and -not $it.Tag.Protected) { $targets += $it.Tag }
        }
        if ($targets.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "終了対象が選択されていません（保護対象は選べません）",
                "情報",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $msg = "以下の $($targets.Count) セッションを強制終了します。`n`n"
        foreach ($t in $targets) {
            $msg += "  ・ PID $($t.ProcessId)  [$($t.Project)]  無活動 $(if($null -ne $t.IdleMinutes){"$($t.IdleMinutes)分"}else{'不明'})`n"
        }
        $msg += "`n入力途中の未保存分は失われます。続行しますか?"
        $r = [System.Windows.Forms.MessageBox]::Show(
            $msg, "終了確認",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $errs = @(); $ok = 0
        foreach ($t in $targets) {
            try { Stop-Process -Id $t.ProcessId -Force -ErrorAction Stop; $ok++ }
            catch { $errs += "PID $($t.ProcessId): $_" }
        }
        $res = "$ok / $($targets.Count) セッションを終了しました。"
        if ($errs.Count -gt 0) { $res += "`n`nエラー:`n" + ($errs -join "`n") }
        [void][System.Windows.Forms.MessageBox]::Show(
            $res, "完了",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
        & $populate
    })

    & $populate
    [void]$dlg.ShowDialog($form)
}

###############################################################
# GUI構築 - メインフォーム
###############################################################
$form = New-Object System.Windows.Forms.Form
$form.Text = "Claude Code / Codex データクリーナー"
$form.Size = New-Object System.Drawing.Size(1120, 720)
$form.MinimumSize = New-Object System.Drawing.Size(900, 550)
$form.StartPosition = "CenterScreen"

###############################################################
# 上段: 環境選択ComboBox + 再スキャンボタン
###############################################################
$envLabel = New-Object System.Windows.Forms.Label
$envLabel.Text = "環境:"
$envLabel.Location = New-Object System.Drawing.Point(10, 15)
$envLabel.Size = New-Object System.Drawing.Size(50, 20)
$form.Controls.Add($envLabel)

$envCombo = New-Object System.Windows.Forms.ComboBox
$envCombo.Location = New-Object System.Drawing.Point(60, 12)
$envCombo.Size = New-Object System.Drawing.Size(280, 25)
$envCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($envCombo)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "再スキャン"
$refreshButton.Location = New-Object System.Drawing.Point(355, 10)
$refreshButton.Size = New-Object System.Drawing.Size(100, 28)
$refreshButton.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($refreshButton)

# セッション管理ボタン（起動中 claude.exe の確認・終了）
$sessionButton = New-Object System.Windows.Forms.Button
$sessionButton.Text = "セッション管理..."
$sessionButton.Location = New-Object System.Drawing.Point(470, 10)
$sessionButton.Size = New-Object System.Drawing.Size(140, 28)
$sessionButton.Anchor = "Top,Left"
$sessionButton.Add_Click({ Show-SessionManager })
$form.Controls.Add($sessionButton)

# 対象セレクタ（Claude Code / Codex）
$targetLabel = New-Object System.Windows.Forms.Label
$targetLabel.Text = "対象:"
$targetLabel.Location = New-Object System.Drawing.Point(625, 15)
$targetLabel.Size = New-Object System.Drawing.Size(40, 20)
$form.Controls.Add($targetLabel)

$targetCombo = New-Object System.Windows.Forms.ComboBox
$targetCombo.Location = New-Object System.Drawing.Point(665, 12)
$targetCombo.Size = New-Object System.Drawing.Size(150, 25)
$targetCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$targetCombo.Items.Add("Claude Code")
[void]$targetCombo.Items.Add("Codex")
$targetCombo.SelectedIndex = 0
$form.Controls.Add($targetCombo)

# セッション抽出/削除ボタン（選択プロジェクトのセッション一覧を開く）
$sessBrowseButton = New-Object System.Windows.Forms.Button
$sessBrowseButton.Text = "セッション抽出/削除..."
$sessBrowseButton.Location = New-Object System.Drawing.Point(825, 10)
$sessBrowseButton.Size = New-Object System.Drawing.Size(170, 28)
$sessBrowseButton.Anchor = "Top,Left"
$form.Controls.Add($sessBrowseButton)

###############################################################
# 中段左: プロジェクト一覧 ListView（チェックボックス + 詳細表示）
###############################################################
$listLabel = New-Object System.Windows.Forms.Label
$listLabel.Text = "プロジェクト一覧（チェックして選択）:"
$listLabel.Location = New-Object System.Drawing.Point(10, 50)
$listLabel.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($listLabel)

$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 75)
$listView.Size = New-Object System.Drawing.Size(620, 470)
$listView.View = [System.Windows.Forms.View]::Details
$listView.CheckBoxes = $true
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Anchor = "Top,Left,Bottom"
[void]$listView.Columns.Add("プロジェクトパス (cwd)", 340)
[void]$listView.Columns.Add("最終更新", 130)
[void]$listView.Columns.Add("サイズ", 80)
[void]$listView.Columns.Add("セッション", 55)
$form.Controls.Add($listView)

###############################################################
# 全選択 / 全解除ボタン
###############################################################
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "全選択"
$selectAllButton.Location = New-Object System.Drawing.Point(10, 555)
$selectAllButton.Size = New-Object System.Drawing.Size(75, 28)
$selectAllButton.Anchor = "Bottom,Left"
$selectAllButton.Add_Click({
    foreach ($item in $listView.Items) { $item.Checked = $true }
})
$form.Controls.Add($selectAllButton)

$deselectAllButton = New-Object System.Windows.Forms.Button
$deselectAllButton.Text = "全解除"
$deselectAllButton.Location = New-Object System.Drawing.Point(90, 555)
$deselectAllButton.Size = New-Object System.Drawing.Size(75, 28)
$deselectAllButton.Anchor = "Bottom,Left"
$deselectAllButton.Add_Click({
    foreach ($item in $listView.Items) { $item.Checked = $false }
})
$form.Controls.Add($deselectAllButton)

###############################################################
# 中段右: 削除プラン表示 RichTextBox
###############################################################
$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Text = "削除プラン (Dry-Run):"
$previewLabel.Location = New-Object System.Drawing.Point(640, 50)
$previewLabel.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($previewLabel)

$previewBox = New-Object System.Windows.Forms.RichTextBox
$previewBox.Location = New-Object System.Drawing.Point(640, 75)
$previewBox.Size = New-Object System.Drawing.Size(450, 470)
$previewBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$previewBox.ReadOnly = $true
$previewBox.WordWrap = $false
$previewBox.ScrollBars = "Both"
$previewBox.Anchor = "Top,Left,Bottom,Right"
$form.Controls.Add($previewBox)

###############################################################
# 下段: Dry-Run / 削除実行 ボタン
###############################################################
$dryRunButton = New-Object System.Windows.Forms.Button
$dryRunButton.Text = "Dry-Run（削除プラン表示）"
$dryRunButton.Location = New-Object System.Drawing.Point(640, 555)
$dryRunButton.Size = New-Object System.Drawing.Size(200, 30)
$dryRunButton.BackColor = [System.Drawing.Color]::LightBlue
$dryRunButton.Anchor = "Bottom,Left"
$form.Controls.Add($dryRunButton)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "削除実行"
$deleteButton.Location = New-Object System.Drawing.Point(960, 555)
$deleteButton.Size = New-Object System.Drawing.Size(130, 30)
$deleteButton.BackColor = [System.Drawing.Color]::LightCoral
$deleteButton.Anchor = "Bottom,Right"
$form.Controls.Add($deleteButton)

###############################################################
# backups/ 全削除オプション（既定OFF）
# ~/.claude.json の過去スナップショットを環境単位で全削除する。
# プロジェクト選択とは独立した全体操作（実害なし）。
###############################################################
$backupsCheck = New-Object System.Windows.Forms.CheckBox
$backupsCheck.Text = "backups/ も全削除（.claude.json 過去スナップショット／全プロジェクト共通）"
$backupsCheck.Location = New-Object System.Drawing.Point(175, 560)
$backupsCheck.Size = New-Object System.Drawing.Size(455, 22)
$backupsCheck.Checked = $false
$backupsCheck.Anchor = "Bottom,Left"
$form.Controls.Add($backupsCheck)

###############################################################
# ステータスバー
###############################################################
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "準備完了"
[void]$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

###############################################################
# イベントハンドラ
###############################################################

# プロジェクト再スキャン
$refreshButton.Add_Click({
    if ($envCombo.SelectedIndex -lt 0) { return }
    $selectedEnv = $script:Environments[$envCombo.SelectedIndex]

    $statusLabel.Text = "スキャン中: $($selectedEnv.Name) ..."
    $form.Refresh()
    $listView.Items.Clear()
    $previewBox.Text = ""

    # 再スキャン時はソート状態とヘッダ矢印をリセット
    $script:SortColumn = -1
    $script:SortAscending = $true
    for ($c = 0; $c -lt $listView.Columns.Count; $c++) {
        $listView.Columns[$c].Text = $listView.Columns[$c].Text -replace ' [▲▼]$', ''
    }

    $isCodex = ($targetCombo.SelectedItem -eq "Codex")
    try {
        if ($isCodex) {
            $script:CurrentProjects = Get-CodexProjects -Environment $selectedEnv
        } else {
            $script:CurrentProjects = Get-ClaudeProjects -Environment $selectedEnv
        }
    } catch {
        $script:CurrentProjects = @()
        [System.Windows.Forms.MessageBox]::Show(
            "プロジェクト取得中にエラー: $_",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }

    # backups/ オプションは Claude 専用（Codex時は無効化）
    $backupsCheck.Enabled = -not $isCodex
    if ($isCodex) { $backupsCheck.Checked = $false }

    for ($i = 0; $i -lt $script:CurrentProjects.Count; $i++) {
        $p = $script:CurrentProjects[$i]
        $item = New-Object System.Windows.Forms.ListViewItem($p.Cwd)
        [void]$item.SubItems.Add($p.LastModified.ToString("yyyy-MM-dd HH:mm"))
        [void]$item.SubItems.Add((Format-Size -Bytes $p.SizeBytes))
        [void]$item.SubItems.Add([string]$p.SessionIds.Count)
        $item.Tag = $i  # CurrentProjects へのインデックス
        [void]$listView.Items.Add($item)
    }

    # 警告があれば優先表示、なければ通常のスキャン結果サマリ
    if ($script:LastScanWarning) {
        $statusLabel.Text = $script:LastScanWarning
        $script:LastScanWarning = $null
    } else {
        $statusLabel.Text = "$($selectedEnv.Name): $($script:CurrentProjects.Count) プロジェクトを検出"
    }
})

# 環境変更時に自動で再スキャン
$envCombo.Add_SelectedIndexChanged({
    $refreshButton.PerformClick()
})

# 対象(Claude/Codex)変更時も再スキャン
$targetCombo.Add_SelectedIndexChanged({
    $refreshButton.PerformClick()
})

# 選択プロジェクトのセッション一覧を開く（行ダブルクリック / ボタン共通）
# チェック（チェックボックス）を優先し、無ければ行ハイライトを使う
$openSessionBrowser = {
    $p = $null
    $checked = @($listView.Items | Where-Object { $_.Checked })
    if ($checked.Count -eq 1) {
        $p = $script:CurrentProjects[[int]$checked[0].Tag]
    } elseif ($checked.Count -gt 1) {
        [void][System.Windows.Forms.MessageBox]::Show("セッションを開くプロジェクトは1つだけにしてください（チェックを1つに）", "情報", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    } elseif ($listView.SelectedItems.Count -ge 1) {
        $p = $script:CurrentProjects[[int]$listView.SelectedItems[0].Tag]
    } else {
        [void][System.Windows.Forms.MessageBox]::Show("プロジェクトを1つ選択してください（チェック、または行をクリック）", "情報", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    Show-SessionBrowser -Project $p
    # セッション削除が行われている可能性があるので一覧を更新
    $refreshButton.PerformClick()
}
$listView.Add_DoubleClick($openSessionBrowser)
$sessBrowseButton.Add_Click($openSessionBrowser)

# 列ヘッダクリックでソート（実データ型で昇降順・チェック状態は保持）
$listView.Add_ColumnClick({
    param($s, $e)
    $col = $e.Column
    if ($script:SortColumn -eq $col) {
        $script:SortAscending = -not $script:SortAscending
    } else {
        $script:SortColumn = $col
        $script:SortAscending = $true
    }

    # 現在の項目とチェック状態を退避（Tag=CurrentProjectsインデックス）
    $items = @($listView.Items)
    if ($items.Count -eq 0) { return }
    $checked = @{}
    foreach ($it in $items) { if ($it.Checked) { $checked[[int]$it.Tag] = $true } }

    # 実データ型でソート（0:cwd文字列, 1:最終更新日時, 2:サイズbytes, 3:セッション数）
    $sorted = $items | Sort-Object -Descending:(-not $script:SortAscending) -Property @{
        Expression = {
            $p = $script:CurrentProjects[[int]$_.Tag]
            switch ($col) {
                1 { $p.LastModified }
                2 { $p.SizeBytes }
                3 { [int]$p.SessionIds.Count }
                default { [string]$p.Cwd }
            }
        }
    }

    $listView.BeginUpdate()
    $listView.Items.Clear()
    foreach ($it in $sorted) { [void]$listView.Items.Add($it) }
    # チェック状態を復元
    foreach ($it in $listView.Items) { if ($checked.ContainsKey([int]$it.Tag)) { $it.Checked = $true } }
    # ヘッダに ▲/▼ を表示
    for ($c = 0; $c -lt $listView.Columns.Count; $c++) {
        $base = $listView.Columns[$c].Text -replace ' [▲▼]$', ''
        if ($c -eq $col) { $base += $(if ($script:SortAscending) { ' ▲' } else { ' ▼' }) }
        $listView.Columns[$c].Text = $base
    }
    $listView.EndUpdate()
})

# Dry-Run表示
$dryRunButton.Add_Click({
    $selected = @()
    foreach ($item in $listView.Items) {
        if ($item.Checked) {
            $idx = [int]$item.Tag
            $selected += $script:CurrentProjects[$idx]
        }
    }

    if ($selected.Count -eq 0) {
        $previewBox.Text = "（プロジェクトが選択されていません）"
        $statusLabel.Text = "選択なし"
        return
    }

    if ($targetCombo.SelectedItem -eq "Codex") {
        $previewBox.Text = New-CodexDeletionPlanText -Projects $selected
    } else {
        $previewBox.Text = New-DeletionPlanText -Projects $selected -IncludeBackups $backupsCheck.Checked
    }
    $statusLabel.Text = "$($selected.Count) プロジェクトのDry-Runを表示"
})

# 削除実行
$deleteButton.Add_Click({
    $selected = @()
    foreach ($item in $listView.Items) {
        if ($item.Checked) {
            $idx = [int]$item.Tag
            $selected += $script:CurrentProjects[$idx]
        }
    }

    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "削除対象が選択されていません",
            "警告",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # backups/ 削除対象の環境（distro）を算出（Windows は null）
    $backupTargets = @()
    if ($backupsCheck.Checked) {
        $backupTargets = @($selected | ForEach-Object { $_.Distro } | Select-Object -Unique)
    }

    # 確認ダイアログ
    $confirmText = "以下の $($selected.Count) プロジェクトを完全削除します。`n`n"
    foreach ($p in $selected) {
        $confirmText += "  ・ $($p.Cwd)`n"
    }
    if ($backupsCheck.Checked) {
        $confirmText += "`n★ backups/（~/.claude.json の過去スナップショット）も全削除します。`n"
        $confirmText += "   対象環境: " + (($backupTargets | ForEach-Object {
            if ($_) { "WSL2: $_" } else { "Windows" }
        }) -join ", ") + "`n"
    }
    $confirmText += "`nこの操作は取り消せません。続行しますか?"

    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmText,
        "削除確認",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )

    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        $statusLabel.Text = "キャンセル"
        return
    }

    # 削除実行（プロジェクト毎にループ）
    $allErrors = @()
    $successCount = 0
    for ($i = 0; $i -lt $selected.Count; $i++) {
        $p = $selected[$i]
        $statusLabel.Text = "削除中 ($($i + 1)/$($selected.Count)): $($p.Cwd)"
        $form.Refresh()

        try {
            if ($p.Kind -eq "Codex") {
                $errors = Invoke-CodexDeletion -Project $p
            } else {
                $errors = Invoke-ProjectDeletion -Project $p
            }
            if ($errors -and $errors.Count -gt 0) {
                $allErrors += "[$($p.Cwd)]"
                $allErrors += $errors
            } else {
                # エラーが無かった場合のみ成功としてカウント（部分失敗を成功に数えない）
                $successCount++
            }
        } catch {
            $allErrors += "[$($p.Cwd)] 削除実行失敗: $_"
        }
    }

    # backups/ 全削除（オプション有効時のみ・環境単位で一度ずつ）
    $backupRemoved = 0
    foreach ($d in $backupTargets) {
        $envName = if ($d) { "WSL2: $d" } else { "Windows" }
        $statusLabel.Text = "backups/ 削除中: $envName"
        $form.Refresh()
        try {
            $bErrors = Remove-ClaudeBackups -Distro $d
            if ($bErrors -and $bErrors.Count -gt 0) {
                $allErrors += "[backups/ $envName]"
                $allErrors += $bErrors
            } else {
                $backupRemoved++
            }
        } catch {
            $allErrors += "[backups/ $envName] 削除失敗: $_"
        }
    }

    # 結果表示
    $resultMsg = "$successCount / $($selected.Count) プロジェクトを削除しました。"
    if ($backupsCheck.Checked) {
        $resultMsg += "`nbackups/ を $backupRemoved / $($backupTargets.Count) 環境で削除しました。"
    }
    if ($allErrors.Count -gt 0) {
        $resultMsg += "`n`n以下の警告/エラーがありました:`n"
        $resultMsg += ($allErrors -join "`n")
    }

    [System.Windows.Forms.MessageBox]::Show(
        $resultMsg,
        "削除完了",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    # 一覧を再読込
    $refreshButton.PerformClick()
})

###############################################################
# 初期化 - 環境一覧をComboBoxへ投入
###############################################################
$statusLabel.Text = "環境を検出中..."
$script:Environments = @(Get-AvailableEnvironments)
$envCombo.Items.Clear()
foreach ($e in $script:Environments) {
    [void]$envCombo.Items.Add($e.Name)
}
if ($envCombo.Items.Count -gt 0) {
    $envCombo.SelectedIndex = 0  # SelectedIndexChangedイベントで自動的に初回スキャンが走る
}
$statusLabel.Text = "$($script:Environments.Count) 個の環境を検出"

###############################################################
# フォーム表示
###############################################################
[void]$form.ShowDialog()