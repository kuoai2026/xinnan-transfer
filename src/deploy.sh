#!/bin/bash
# 重建「倉庫調貨單」門市端頁面並推上 GitHub Pages。
# 由 cron 呼叫（早上 08:30、中午 12:50，接在網翼庫存下載之後）。
set -euo pipefail

PY=/Library/Frameworks/Python.framework/Versions/3.14/bin/python3
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$HOME/repos/xinnan-transfer"

"$PY" "$HERE/build_transfer_page.py"

cd "$REPO"
git add -A
if git diff --cached --quiet; then
  echo "無變更，不推送"
  exit 0
fi
git -c user.email=kuoai2026@gmail.com -c user.name="kuoai2026" \
  commit -q -m "rebuild $(date '+%Y-%m-%d %H:%M')"
GIT_SSH_COMMAND="ssh -i $HOME/.ssh/xinnan-transfer-deploy -o IdentitiesOnly=yes" \
  git push -q origin main
echo "已推送 $(git rev-parse --short HEAD)"
