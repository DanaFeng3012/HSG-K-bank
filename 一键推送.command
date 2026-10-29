#!/bin/zsh
# ========== 一键推送 K Bank Workshop 看板到 GitHub Pages ==========
# 使用方法：
#   1. 在浏览器看板里点「📤 导出数据」（下载 JSON 到 Downloads 文件夹）
#   2. 双击这个文件，自动替换 data.json + push 到 GitHub
# =================================================================

set -e
REPO="/Users/danafeng/Documents/trae_projects/K Bank"
cd "$REPO"

echo ""
echo "🔄 K Bank Workshop 一键发布"
echo "============================"

# ① 自动从 Downloads 找最新导出的 JSON
#    看板导出文件名格式："K Bank看板完整数据_YYYY-MM-DD.json" 或旧的 "K Bank Workshop看板数据_..."
LATEST_JSON=""
CANDIDATES=(
  "$HOME/Downloads/K Bank看板完整数据_"*.json
  "$HOME/Downloads/K Bank Workshop看板数据_"*.json
  "$HOME/Downloads/"*"Workshop"*.json
  "$HOME/Downloads/"*"K.Bank"*.json
)
for f in "${CANDIDATES[@]}"; do
  if [[ -f "$f" ]]; then
    if [[ -z "$LATEST_JSON" || "$f" -nt "$LATEST_JSON" ]]; then
      LATEST_JSON="$f"
    fi
  fi
done

if [[ -n "$LATEST_JSON" ]]; then
  echo "📥 发现最新导出数据：$(basename "$LATEST_JSON")"
  # 新格式：{ items: [...], dicts: {...} } → 直接复制（看板 init 会自动识别）
  # 旧格式：纯数组 → 也直接复制
  cp "$LATEST_JSON" "$REPO/data.json"
  echo "✅ 已替换 data.json"
else
  echo "⚠️  没找到导出的 JSON，使用仓库里已有的 data.json"
  echo "   （如果刚在看板里改了内容，请先点「📤 导出数据」再推送）"
fi

# ② 复制到中文文件名
if [[ -f "$REPO/index.html" ]]; then
  cp "$REPO/index.html" "$REPO/K Bank Workshop追踪看板.html" 2>/dev/null || true
fi

# ③ Git add + commit + push
echo ""
echo "📤 Git 推送..."
git add -A
if git diff --cached --quiet; then
  echo "ℹ️  没有新变更，跳过 commit"
else
  MSG="更新看板：$(date '+%Y-%m-%d %H:%M')"
  git commit -m "$MSG" >/dev/null
  echo "✅ Commit: $MSG"
fi
git push origin main 2>&1 | tail -3

echo ""
echo "=========================================="
echo "🎉 推送完成！"
echo ""
echo "🌐 公网链接（1-2 分钟后自动刷新）："
echo "   https://danafeng3012.github.io/HSG-K-bank/index.html"
echo ""
echo "（现在可以关掉这个窗口了）"
echo ""
