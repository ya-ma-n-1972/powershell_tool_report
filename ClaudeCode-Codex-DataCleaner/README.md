# ClaudeCode-Codex-DataCleaner

Claude Code / OpenAI Codex CLI のプロジェクト永続化データクリーナー — Windows/WSL2両対応 GUI

Claude Code および OpenAI Codex CLI が `~/.claude/`・`~/.codex/` 配下に蓄積するプロジェクト単位の永続化データ（会話トランスクリプト・履歴・タスク・ファイル編集履歴など）を、Windows Forms の GUI から選択的にクリーニングする PowerShell 単一ファイルツールです。外部依存はありません。

公式の `claude project purge` 相当の削除射程をカバーしつつ、セッション単位の会話履歴抽出（Markdown）や稼働中プロセスの安全管理など、公式コマンドにはない機能を備えます。

> [!WARNING]
> **本ツールはファイルを完全削除します。削除されたデータは復元できません。**
> - 必ず **Dry-Run（削除プラン表示）** で対象を確認してから「削除実行」してください。
> - 重要なデータは事前にバックアップしてください。
> - 本ソフトウェアは無保証です。データ損失を含むいかなる損害についても作者は責任を負いません。**自己責任でご利用ください。**

## 特徴

- **Windows/WSL2両対応** - ローカルの `~/.claude`・`~/.codex` も WSL2 環境も1つのツールで（全ディストロ自動検出）
- **Claude Code / Codex 両対応** - 上段「対象」セレクタで切替
- **Dry-Run** - 実際に削除されるファイル・ディレクトリ・JSONL行を事前確認
- **セッション単位の会話履歴抽出** - Markdown でエクスポート（thinking/reasoning も任意で含む）
- **稼働中プロセスの安全管理** - 起動中の `claude.exe` を保護判定付きで終了
- **PowerShellネイティブ** - 追加インストール不要・管理者権限不要

## クイックスタート

```powershell
# 実行（管理者権限不要）
.\ClaudeCode-Codex-DataCleaner.ps1
```

実行ポリシーでブロックされる場合:

```powershell
powershell -ExecutionPolicy Bypass -File ".\ClaudeCode-Codex-DataCleaner.ps1"
```

### 基本的な使い方

1. **環境を選択** - Windows または WSL2 ディストロ
2. **対象を選択** - Claude Code / Codex
3. **プロジェクトをチェック** - 一覧から削除したいものを選択
4. **Dry-Run** - 削除プランを確認
5. **削除実行** - 確認ダイアログ（既定「いいえ」）の後に削除

## 動作環境

- Windows 10 / 11
- PowerShell 5.1 以降（Windows 標準）または PowerShell 7+
- .NET Framework / System.Windows.Forms（Windows 標準）
- WSL2（WSL 機能使用時のみ・任意）
- 環境変数 `CLAUDE_CONFIG_DIR` / `CODEX_HOME` 設定時はそちらを優先

## 削除対象 / 温存対象

### Claude Code（プロジェクト単位）

- **削除:** `projects/<encoded>/`（トランスクリプト含む）、`file-history/<sid>/`、`tasks/<sid>*`（旧 `todos/` も対象）、`session-env/<sid>/`、`debug/<sid>.txt`、`history.jsonl` の該当行、`~/.claude.json` の該当エントリ
- **温存:** `shell-snapshots/`、`backups/`（オプションで削除可）、`settings.json`、`plugins/`

### Codex（cwd 単位 / v1 はファイル系のみ）

- **削除:** `sessions/`・`archived_sessions/` 内の該当 `rollout-*.jsonl`、`history.jsonl` の `session_id` 一致行、`session_index.jsonl` の `id` 一致行
- **温存:** `state_5.sqlite` / `logs_2.sqlite` / `goals_1.sqlite` / `memories_1.sqlite` などの SQLite、`config.toml` / `auth.json` / `AGENTS.md` 等の設定

## 安全策と注意事項

- JSONL / JSON ファイルの編集前に `.bak` を自動作成
- パース不能な JSONL 行はそのまま保持（破壊回避）
- WSL 側ファイルの書き戻しは LF 改行を強制（CRLF 混入防止）
- 稼働中の Claude/Codex セッションは対象ファイルがロックされ、削除に失敗する場合があります（終了してから実行してください）
- **Codex の `state_5.sqlite` は温存される**ため、rollout 削除後にインデックスと実体の不整合（`codex resume` 等に痕跡が残る）が生じることがあります。詳細は [DOC/特徴と機能一覧](DOC/ClaudeCodeデータクリーナー_特徴と機能一覧.md) を参照

## ドキュメント

- [特徴と機能一覧](DOC/ClaudeCodeデータクリーナー_特徴と機能一覧.md)
- [コード分析レポート](DOC/ClaudeCodeデータクリーナー_コード分析レポート.md)

## ライセンス

MIT License

## 貢献

Issue・Pull Request歓迎です
