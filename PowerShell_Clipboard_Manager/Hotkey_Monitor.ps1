#requires -version 5.1

<#
.SYNOPSIS
    Hotkey Monitor Process
    ホットキー監視プロセス
.DESCRIPTION
    Monitors configurable hotkey and notifies main process when detected
    設定可能なホットキーを監視し、検出時にメインプロセスへ通知
.NOTES
    PowerShell: 5.1+
    Author: Professional implementation for public release
#>

# Set UTF-8 encoding to prevent character corruption
# UTF-8エンコーディングを設定して文字化けを防ぐ
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {
    # ISE environment doesn't support Console class - continue without encoding change
    # ISE環境ではConsoleクラスがサポートされていない - エンコーディング変更なしで続行
    Write-Verbose "Console encoding setting skipped / エンコーディング設定スキップ: $($_.Exception.Message)"
}

#region Hotkey Configuration / ホットキー設定

# --- ユーザー設定可能なホットキー設定 ---
# ========================================
# User-configurable hotkey settings
# ユーザー設定可能なホットキー設定
# ========================================

# Default: Ctrl+Shift+Space (Recommended for ISE environment)
# デフォルト: Ctrl+Shift+Space（ISE環境推奨）
[uint32]$HOTKEY_MODIFIERS = 0x0002 -bor 0x0004  # MOD_CONTROL | MOD_SHIFT
[uint32]$HOTKEY_KEY = 0x20  # VK_SPACE
[string]$HOTKEY_DESCRIPTION = "Ctrl+Shift+Space"

# Alternative hotkey configurations (uncomment to use)
# 代替ホットキー設定（使用する場合はコメントを外す）

# Example 1: Win+Shift+C (Avoids Windows Terminal conflict)
# 例1: Win+Shift+C（Windows Terminalとの競合を回避）
# $HOTKEY_MODIFIERS = 0x0008 -bor 0x0004  # MOD_WIN | MOD_SHIFT
# $HOTKEY_KEY = 0x43  # 'C' key
# $HOTKEY_DESCRIPTION = "Win+Shift+C"

# Example 2: Ctrl+Alt+C (Common combination)
# 例2: Ctrl+Alt+C（一般的な組み合わせ）
# $HOTKEY_MODIFIERS = 0x0002 -bor 0x0001  # MOD_CONTROL | MOD_ALT
# $HOTKEY_KEY = 0x43  # 'C' key
# $HOTKEY_DESCRIPTION = "Ctrl+Alt+C"

# Example 3: Alt+Shift+V (V for View)
# 例3: Alt+Shift+V（Vは表示の頭文字）
# $HOTKEY_MODIFIERS = 0x0001 -bor 0x0004  # MOD_ALT | MOD_SHIFT
# $HOTKEY_KEY = 0x56  # 'V' key
# $HOTKEY_DESCRIPTION = "Alt+Shift+V"

# Example 4: F10 key alone (Function key example)
# 例4: F10キー単独（ファンクションキー使用例）
# $HOTKEY_MODIFIERS = 0x0000  # No modifiers / 修飾キーなし
# $HOTKEY_KEY = 0x79  # F10
# $HOTKEY_DESCRIPTION = "F10"

# Key code reference / キーコード参考:
# 0x41-0x5A: A-Z
# 0x30-0x39: 0-9
# 0x70-0x7B: F1-F12
# 0x20: Space
# 0x0D: Enter
# 0x1B: Escape
# 0x09: Tab

# Modifier reference / 修飾キー参考:
# 0x0001: Alt (MOD_ALT)
# 0x0002: Ctrl (MOD_CONTROL)
# 0x0004: Shift (MOD_SHIFT)
# 0x0008: Win (MOD_WIN)

#endregion

#region Constants and Globals / 定数とグローバル変数

# --- アプリケーション定数・リトライ設定 ---
# Application constants / アプリケーション定数
[string]$script:PIPE_NAME = "PSClipboardPipe"
[int]$script:HOTKEY_ID = 1
[int]$script:PIPE_TIMEOUT_MS = 1000
[bool]$script:is_running = $true

# Retry settings for pipe connection / パイプ接続のリトライ設定
[int]$script:PIPE_RETRY_COUNT = 3
[int]$script:PIPE_RETRY_DELAY_MS = 500

#endregion

#region Windows API Definition / Windows API定義

# --- グローバルホットキー用Windows API定義 ---
# Windows API for global hotkey functionality
# グローバルホットキー機能用のWindows API
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    
    public class HotkeyAPI {
        // Register a global hotkey / グローバルホットキーを登録
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        
        // Unregister a global hotkey / グローバルホットキーを解除
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        
        // Get message from message queue / メッセージキューからメッセージを取得
        [DllImport("user32.dll")]
        public static extern int GetMessage(out WindowsMessage lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);
        
        // Modifier key constants / 修飾キー定数
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_CONTROL = 0x0002;
        public const uint MOD_SHIFT = 0x0004;
        public const uint MOD_WIN = 0x0008;
        
        // Virtual key codes (commonly used) / 仮想キーコード（よく使うもの）
        public const uint VK_SPACE = 0x20;
        public const uint VK_RETURN = 0x0D;
        public const uint VK_ESCAPE = 0x1B;
        public const uint VK_TAB = 0x09;
        
        // Windows messages / Windowsメッセージ
        public const uint WM_HOTKEY = 0x0312;
        
        // Windows message structure / Windowsメッセージ構造体
        [StructLayout(LayoutKind.Sequential)]
        public struct WindowsMessage {
            public IntPtr hwnd;      // Window handle / ウィンドウハンドル
            public uint message;     // Message identifier / メッセージ識別子
            public IntPtr wParam;    // Message-specific parameter / メッセージ固有パラメータ
            public IntPtr lParam;    // Message-specific parameter / メッセージ固有パラメータ
            public uint time;        // Time message was posted / メッセージ投稿時刻
            public int x;            // Cursor X position / カーソルX座標
            public int y;            // Cursor Y position / カーソルY座標
        }
    }
"@ -ErrorAction SilentlyContinue

#endregion

#region Core Functions / コア関数

# --- メインプロセスへSHOWコマンド送信（パイプ通信） ---
<#
.SYNOPSIS
    Send SHOW command to main process via named pipe
    名前付きパイプ経由でメインプロセスにSHOWコマンドを送信
.DESCRIPTION
    Connects to the main process pipe server and requests window display
    メインプロセスのパイプサーバーに接続してウィンドウ表示を要求
#>
function Send-ShowCommand {
    [CmdletBinding()]
    param()
    
    for ($retry = 0; $retry -lt $script:PIPE_RETRY_COUNT; $retry++) {
        try {
            Write-Host "[Hotkey] Connecting to pipe / パイプ接続中... (Attempt $($retry + 1)/$script:PIPE_RETRY_COUNT)" -ForegroundColor Cyan
            
            # Create pipe client / パイプクライアントを作成
            $pipe = New-Object System.IO.Pipes.NamedPipeClientStream(
                ".",  # Local machine / ローカルマシン
                $script:PIPE_NAME,
                [System.IO.Pipes.PipeDirection]::InOut
            )
            
            # Connect with timeout / タイムアウト付きで接続
            $pipe.Connect($script:PIPE_TIMEOUT_MS)
            
            if ($pipe.IsConnected) {
                Write-Host "[Hotkey] Connected, sending command / 接続成功、コマンド送信" -ForegroundColor Green
                
                # Send command / コマンドを送信
                $writer = New-Object System.IO.StreamWriter($pipe)
                $writer.WriteLine("SHOW")
                $writer.Flush()
                
                # Receive response / 応答を受信
                $reader = New-Object System.IO.StreamReader($pipe)
                $response = $reader.ReadLine()
                
                if ($response -eq "OK") {
                    Write-Host "[Hotkey] Window display requested / ウィンドウ表示要求成功" -ForegroundColor Green
                }
                
                # Cleanup / クリーンアップ
                $writer.Close()
                $reader.Close()
                $pipe.Close()
                $pipe.Dispose()
                
                return $true
            } else {
                Write-Warning "[Hotkey] Failed to connect / 接続失敗"
            }
        }
        catch {
            if ($retry -eq ($script:PIPE_RETRY_COUNT - 1)) {
                Write-Warning "[Hotkey] Pipe communication failed / パイプ通信失敗: $($_.Exception.Message)"
            } else {
                Write-Verbose "Retrying after delay / 遅延後にリトライ..."
                Start-Sleep -Milliseconds $script:PIPE_RETRY_DELAY_MS
            }
        }
    }
    return $false
}

# --- グローバルホットキー登録 ---
<#
.SYNOPSIS
    Register global hotkey with Windows
    Windowsにグローバルホットキーを登録
.DESCRIPTION
    Uses RegisterHotKey API to register system-wide hotkey
    RegisterHotKey APIを使用してシステム全体のホットキーを登録
#>
function Register-Hotkey {
    [CmdletBinding()]
    param()
    
    try {
        # Register hotkey with user-configured settings
        # ユーザー設定値でホットキーを登録
        $result = [HotkeyAPI]::RegisterHotKey(
            [IntPtr]::Zero,
            $script:HOTKEY_ID,
            $HOTKEY_MODIFIERS,
            $HOTKEY_KEY
        )
        
        if ($result) {
            Write-Host "✅ Hotkey registered successfully / ホットキー登録成功: $HOTKEY_DESCRIPTION" -ForegroundColor Green
            return $true
        } else {
            Write-Host "❌ Failed to register hotkey / ホットキー登録失敗: $HOTKEY_DESCRIPTION" -ForegroundColor Red
            Write-Host "   Another application may be using this hotkey" -ForegroundColor Red
            Write-Host "   他のアプリケーションが使用している可能性があります" -ForegroundColor Red
            Write-Host ""
            Write-Host "【How to set alternative hotkey / 代替ホットキーの設定方法】" -ForegroundColor Yellow
            Write-Host "1. Open this script in a text editor / このスクリプトをテキストエディタで開く" -ForegroundColor Yellow
            Write-Host "2. Check hotkey settings at lines 23-67 / 23-67行目のホットキー設定を確認" -ForegroundColor Yellow
            Write-Host "3. Uncomment desired configuration / 使用したい設定のコメント(#)を外す" -ForegroundColor Yellow
            Write-Host "4. Comment out default settings / デフォルト設定をコメントアウト" -ForegroundColor Yellow
            Write-Host "5. Run the script again / スクリプトを再実行" -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Error "[Hotkey] Registration error / 登録エラー: $($_.Exception.Message)"
        return $false
    }
}

# --- グローバルホットキー解除 ---
<#
.SYNOPSIS
    Unregister global hotkey
    グローバルホットキーを登録解除
#>
function Unregister-Hotkey {
    [CmdletBinding()]
    param()
    
    try {
        [HotkeyAPI]::UnregisterHotKey([IntPtr]::Zero, $script:HOTKEY_ID)
        Write-Host "Hotkey unregistered / ホットキー登録解除: $HOTKEY_DESCRIPTION" -ForegroundColor Yellow
    }
    catch {
        Write-Warning "[Hotkey] Unregister error / 解除エラー: $($_.Exception.Message)"
    }
}

# --- ホットキー検出用メッセージループ開始 ---
<#
.SYNOPSIS
    Start message loop for hotkey detection
    ホットキー検出用のメッセージループを開始
.DESCRIPTION
    Runs Windows message loop to detect WM_HOTKEY messages
    WindowsメッセージループでWM_HOTKEYメッセージを検出
#>
function Start-MessageLoop {
    [CmdletBinding()]
    param()
    
    Write-Host "[Hotkey] Message loop starting / メッセージループ開始" -ForegroundColor Cyan
    Write-Host "Waiting... / 待機中... ($HOTKEY_DESCRIPTION to activate / で反応)" -ForegroundColor Gray
    
    # Create message structure / メッセージ構造体を作成
    $windowsMessage = New-Object HotkeyAPI+WindowsMessage
    
    # Message loop / メッセージループ
    while ($script:is_running -and [HotkeyAPI]::GetMessage(
        [ref]$windowsMessage,
        [IntPtr]::Zero,
        0,
        0
    )) {
        # Check for WM_HOTKEY message / WM_HOTKEYメッセージをチェック
        if ($windowsMessage.message -eq [HotkeyAPI]::WM_HOTKEY) {
            $timestamp = Get-Date -Format "HH:mm:ss.fff"
            Write-Host "🔔 [$timestamp] Hotkey detected / ホットキー検出！ $HOTKEY_DESCRIPTION" -ForegroundColor Cyan
            
            # Notify main process / メインプロセスへ通知
            Send-ShowCommand
        }
    }
    
    Write-Host "[Hotkey] Message loop ended / メッセージループ終了" -ForegroundColor Yellow
}

# --- 終了時リソースクリーンアップ ---
<#
.SYNOPSIS
    Cleanup resources on exit
    終了時にリソースをクリーンアップ
#>
function Invoke-Cleanup {
    [CmdletBinding()]
    param()
    
    Write-Host "`nShutting down / 終了処理中..." -ForegroundColor Yellow
    Unregister-Hotkey
    $script:is_running = $false
}

#endregion

#region Main Execution / メイン実行

# --- Ctrl+Cハンドリング・クリーンアップ登録 ---
# Handle Ctrl+C gracefully / Ctrl+Cを適切に処理
try {
    [Console]::TreatControlCAsInput = $false
} catch {
    # Ignore in ISE environment / ISE環境では無視
}

# Register cleanup on exit / 終了時のクリーンアップを登録
Register-EngineEvent PowerShell.Exiting -Action { Invoke-Cleanup } | Out-Null

# --- 起動情報表示 ---
# Display startup information / 起動情報を表示
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Hotkey Monitor Process" -ForegroundColor Cyan
Write-Host " ホットキー監視プロセス" -ForegroundColor Cyan
Write-Host " Version 1.1 - No changes needed" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Configured hotkey / 設定ホットキー: $HOTKEY_DESCRIPTION" -ForegroundColor White
Write-Host "Pipe name / パイプ名: $script:PIPE_NAME" -ForegroundColor Gray
Write-Host ""
Write-Host "NOTE: This file does not require modifications" -ForegroundColor Green
Write-Host "注: このファイルは修正不要です" -ForegroundColor Green
Write-Host ""

# --- ホットキー登録と監視開始 ---
# Register hotkey and start monitoring / ホットキー登録と監視開始
if (Register-Hotkey) {
    # Start message loop / メッセージループ開始
    Start-MessageLoop
} else {
    Write-Host ""
    Write-Host "Failed to register hotkey, exiting..." -ForegroundColor Red
    Write-Host "ホットキー登録に失敗したため終了します" -ForegroundColor Red
    Write-Host "Please configure an alternative hotkey as described above" -ForegroundColor Yellow
    Write-Host "上記の代替ホットキーを設定してから再実行してください" -ForegroundColor Yellow
    exit 1
}

# --- 正常終了時クリーンアップ ---
# Cleanup on normal exit / 正常終了時のクリーンアップ
Invoke-Cleanup

#endregion
