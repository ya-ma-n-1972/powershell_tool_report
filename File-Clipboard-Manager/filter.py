"""
filter.py - .gitignore パーサー
PowerShellから呼び出されるファイルフィルタリングスクリプト

通常モード: python filter.py <フォルダパス> [--no-gitignore]
テストモード: python filter.py (引数なし)
"""

import sys
import os
import json


def load_settings():
    """filter.pyと同じディレクトリにあるsettings.jsonを読み込む"""
    settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.json')
    with open(settings_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def get_all_files(folder_path):
    """フォルダ内のファイルを再帰的に取得"""
    all_files = []
    for root, dirs, files in os.walk(folder_path):
        # 仕様: 先頭が . のディレクトリは常に除外（.git, .github, .vscode等）
        # この除外は .gitignore の有効/無効に関わらず常時適用される
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        for file in files:
            full_path = os.path.join(root, file)
            all_files.append(full_path)
    return all_files


def filter_by_extension(files, exclude_extensions):
    """除外拡張子に該当しないファイルのみ残す（ネガティブリスト方式）"""
    result = []
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext not in exclude_extensions:
            result.append(f)
    return result


def filter_by_filename(files, exclude_filenames):
    """除外ファイル名に該当しないファイルのみ残す"""
    result = []
    for f in files:
        basename = os.path.basename(f)
        if basename not in exclude_filenames:
            result.append(f)
    return result


def filter_by_gitignore(folder_path, files):
    """pathspecで.gitignoreルールを適用"""
    import pathspec

    gitignore_path = os.path.join(folder_path, '.gitignore')
    with open(gitignore_path, 'r', encoding='utf-8') as f:
        patterns = f.read()

    spec = pathspec.PathSpec.from_lines('gitwildmatch', patterns.splitlines())

    result = []
    for file_path in files:
        # フォルダからの相対パスで判定
        rel_path = os.path.relpath(file_path, folder_path)
        # Windowsパス区切りをスラッシュに変換
        rel_path_unix = rel_path.replace('\\', '/')
        if not spec.match_file(rel_path_unix):
            result.append(file_path)

    return result


def normal_mode(folder_path, use_gitignore):
    """通常モード: PowerShellから呼び出し"""
    settings = load_settings()
    exclude_extensions = set(settings.get('exclude_extensions', []))
    exclude_filenames = set(settings.get('exclude_filenames', []))

    if not os.path.isdir(folder_path):
        print(f"ERROR: フォルダが見つかりません: {folder_path}", file=sys.stderr)
        sys.exit(1)

    if use_gitignore:
        gitignore_path = os.path.join(folder_path, '.gitignore')
        if not os.path.exists(gitignore_path):
            print(f"ERROR: .gitignoreが見つかりません: {gitignore_path}", file=sys.stderr)
            sys.exit(1)

    all_files = get_all_files(folder_path)

    if use_gitignore:
        all_files = filter_by_gitignore(folder_path, all_files)

    all_files = filter_by_extension(all_files, exclude_extensions)
    all_files = filter_by_filename(all_files, exclude_filenames)

    for f in all_files:
        print(f)


def test_mode():
    """テストモード: 単独起動でpathspecと.gitignoreの動作確認"""
    print("=== filter.py テストモード ===")
    print()

    # pathspecチェック
    print("[1] pathspec ライブラリの確認...")
    try:
        import pathspec
        print(f"    OK: pathspec {pathspec.__version__} が利用可能です")
    except ImportError:
        print("    NG: pathspec が見つかりません")
        print("    → pip install pathspec を実行してください")
        sys.exit(1)

    print()

    # .gitignoreパスの入力
    print("[2] .gitignore のパスを入力してください")
    print("    例: C:\\projects\\myapp\\.gitignore")
    gitignore_path = input("    パス: ").strip().strip('"')

    if not os.path.exists(gitignore_path):
        print(f"    NG: ファイルが見つかりません: {gitignore_path}")
        sys.exit(1)

    print(f"    OK: ファイルを確認しました")
    print()

    # .gitignore読み込みとフィルタリング結果表示
    print("[3] .gitignore の読み込みと適用...")
    try:
        import pathspec
        folder_path = os.path.dirname(gitignore_path)
        with open(gitignore_path, 'r', encoding='utf-8') as f:
            patterns = f.read()

        spec = pathspec.PathSpec.from_lines('gitwildmatch', patterns.splitlines())
        print(f"    OK: パターンを読み込みました")
        print()

        # フィルタリング結果の表示
        print("[4] フィルタリング結果")
        settings = load_settings()
        exclude_extensions = set(settings.get('exclude_extensions', []))
        exclude_filenames = set(settings.get('exclude_filenames', []))

        all_files = get_all_files(folder_path)
        after_gitignore = filter_by_gitignore(folder_path, all_files)
        after_extension = filter_by_extension(after_gitignore, exclude_extensions)
        after_filename = filter_by_filename(after_extension, exclude_filenames)

        print(f"    スキャン総数          : {len(all_files)} ファイル")
        print(f"    .gitignore適用後      : {len(after_gitignore)} ファイル")
        print(f"    拡張子・ファイル名除外後（最終）: {len(after_filename)} ファイル")
        print()
        print("    対象ファイル一覧:")
        for f in after_filename:
            rel = os.path.relpath(f, folder_path)
            print(f"      {rel}")

    except Exception as e:
        print(f"    NG: エラーが発生しました: {e}")
        sys.exit(1)

    print()
    print("=== テスト完了 ===")


if __name__ == '__main__':
    args = sys.argv[1:]

    if len(args) == 0:
        # テストモード
        test_mode()
    else:
        # 通常モード
        folder_path = args[0]
        use_gitignore = '--no-gitignore' not in args
        normal_mode(folder_path, use_gitignore)