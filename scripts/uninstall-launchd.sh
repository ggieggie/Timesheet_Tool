#!/bin/bash
# launchd エージェントをアンインストールするスクリプト

set -e

USER_AGENTS_DIR="$HOME/Library/LaunchAgents"
PLIST_NAME="com.timesheettool.batch.plist"

echo "=== TimesheetTool launchd アンインストーラ ==="

# エージェントを停止・削除
label="${PLIST_NAME%.plist}"
if launchctl list | grep -q "$label"; then
    echo "🔄 $label を停止中..."
    launchctl unload "$USER_AGENTS_DIR/$PLIST_NAME" 2>/dev/null || true
fi
if [ -f "$USER_AGENTS_DIR/$PLIST_NAME" ]; then
    rm "$USER_AGENTS_DIR/$PLIST_NAME"
    echo "✅ $PLIST_NAME を削除"
fi

echo ""
echo "=== アンインストール完了 ==="
