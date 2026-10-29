#!/bin/zsh
cd "/Users/danafeng/Documents/trae_projects/K Bank"
echo "🔄 正在推送最新看板到 GitHub..."
git add -A
msg="${1:-更新看板：$(date '+%Y-%m-%d %H:%M')}"
git commit -m "$msg" 2>/dev/null || echo "（无新变更，跳过 commit）"
git push origin main
echo ""
echo "✅ 推送完成！1-2 分钟后公网链接自动刷新："
echo "   https://danafeng3012.github.io/HSG-K-bank/index.html"
echo ""
read -p "按回车关闭..."
