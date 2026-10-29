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
#    看板导出文件名格式："K Bank Workshop看板数据_YYYY-MM-DD.json"
#    或者任何 .json 只要包含 "Workshop" 都认
LATEST_JSON=""
CANDIDATES=(
  "$HOME/Downloads/K Bank Workshop看板数据_"*.json
  "$HOME/Downloads/"*"Workshop"*.json
  "$HOME/Downloads/"*"K.Bank"*.json
)
for f in "${CANDIDATES[@]}"; do
  if [[ -f "$f" ]]; then
    # 选修改时间最新的
    if [[ -z "$LATEST_JSON" || "$f" -nt "$LATEST_JSON" ]]; then
      LATEST_JSON="$f"
    fi
  fi
done

if [[ -n "$LATEST_JSON" ]]; then
  echo "📥 发现最新导出数据：$(basename "$LATEST_JSON")"
  # 替换 repo 里的 data.json
  cp "$LATEST_JSON" "$REPO/data.json"
  echo "✅ 已替换 data.json"
else
  echo "⚠️  没找到导出的 JSON，使用仓库里已有的 data.json"
  echo "   （如果刚在看板里改了内容，请先点「📤 导出数据」再推送）"
fi

# ② 更新 index.html 里的 CONFIG.seed（让离线用户也能看到最新数据）
if command -v node >/dev/null 2>&1 || [[ -x /tmp/node-tar/node-v22.11.0-darwin-arm64/bin/node ]]; then
  NODE_BIN=$(command -v node 2>/dev/null || echo "/tmp/node-tar/node-v22.11.0-darwin-arm64/bin/node")
  echo "🔧 更新 index.html 的 CONFIG.seed..."
  $NODE_BIN -e "
const fs = require('fs');
const htmlPath = '$REPO/index.html';
const jsonPath = '$REPO/data.json';
let html = fs.readFileSync(htmlPath, 'utf-8');
const seed = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
const jsSeed = JSON.stringify(seed, null, 2).replace(/\"/g, (m) => m);
// 替换 seed: [ ... ] 整个数组（从 "seed: [" 到匹配的 "],"）
const startIdx = html.indexOf('seed: [');
if (startIdx >= 0) {
  let depth = 0, i = startIdx;
  while (i < html.length) {
    if (html[i] === '[') depth++;
    else if (html[i] === ']') { depth--; if (depth === 0) break; }
    i++;
  }
  html = html.slice(0, startIdx + 6) + jsSeed + html.slice(i + 1);
  fs.writeFileSync(htmlPath, html);
  console.log('   ✅ index.html seed 已同步');
}
" 2>/dev/null && echo "   ✅ 完成" || echo "   ⚠️  跳过 index.html 更新（node 不可用或格式变化）"
fi

# ③ 复制到中文文件名（如果存在的话）
CHINESE_HTML="$REPO/K Bank Workshop追踪看板.html"
if [[ -f "$REPO/index.html" ]]; then
  cp "$REPO/index.html" "$CHINESE_HTML" 2>/dev/null || true
fi

# ④ Git add + commit + push
echo ""
echo "📤 Git 推送..."
git add -A
# 只有有变更才 commit
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

__tr_native_ec=$?; pwd -P >| '/Users/danafeng/Library/Application Support/Trae CN/ModularData/ai-agent/command-execution/3b8a88d7540d479a02b90a2a365a82cc/cwd.txt'; exit "$__tr_native_ec"