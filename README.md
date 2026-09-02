# xinnan-transfer

信男藥局 · 倉庫調貨單（門市端）。

- 網址：https://kuoai2026.github.io/xinnan-transfer/
- `index.html` 由 `~/我的雲端硬碟/scripts/transfer_page/build_transfer_page.py` 產生，請勿手動編輯。
- 每天 08:30 / 12:50 由 cron 重建並自動推送（見 `scripts/transfer_page/deploy.sh`）。
- 後端：綁在「藥局調貨表單」試算表的 Apps Script（`scripts/transfer_page/Code.gs`）。
