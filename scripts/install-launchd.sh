#!/bin/bash
# launchd エージェントをインストールするスクリプト

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
LAUNCHD_DIR="$PROJECT_DIR/launchd"
USER_AGENTS_DIR="$HOME/Library/LaunchAgents"
PLIST_NAME="com.timesheettool.batch.plist"

echo "=== TimesheetTool launchd インストーラ ==="
echo "プロジェクト: $PROJECT_DIR"
echo ""

# logsディレクトリ作成
mkdir -p "$PROJECT_DIR/logs"
echo "✅ logsディレクトリ作成"

# LaunchAgentsディレクトリ作成
mkdir -p "$USER_AGENTS_DIR"

# 既存のエージェントを停止・削除
label="${PLIST_NAME%.plist}"
if launchctl list | grep -q "$label"; then
    echo "🔄 既存の $label を停止中..."
    launchctl unload "$USER_AGENTS_DIR/$PLIST_NAME" 2>/dev/null || true
fi

# plistファイルをコピー
cp "$LAUNCHD_DIR/$PLIST_NAME" "$USER_AGENTS_DIR/"
echo "✅ plistファイルをコピー"

# エージェントを登録
launchctl load "$USER_AGENTS_DIR/$PLIST_NAME"
echo "✅ launchdエージェントを登録"

echo ""
echo "=== インストール完了 ==="
echo ""
echo "スケジュール: 毎日 8:00 に当月分を自動処理"
echo ""
echo "確認コマンド:"
echo "  launchctl list | grep timesheettool"
echo ""
echo "手動実行（テスト）:"
echo "  cd $PROJECT_DIR"
echo "  python3 Timesheet_Tool.py              # 当月を処理"
echo "  python3 Timesheet_Tool.py 202501       # 指定月を処理"
echo "  DRY_RUN=1 python3 Timesheet_Tool.py    # ドライラン"
echo ""
echo "ログ確認:"
echo "  tail -f $PROJECT_DIR/logs/timesheet.log"
