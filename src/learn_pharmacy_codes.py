"""從「調貨單_明細」備註欄學藥局商品編號（7位數字）→ 寫入 pharmacy_codes_extra.json。
對照表 csv 已有的貨號不動；有衝突（同貨號出現不同碼）只列出不寫入。
用法：python3 learn_pharmacy_codes.py [--dry]"""
import csv, glob, json, os, re, sys, collections
import gspread
HERE = os.path.dirname(os.path.abspath(__file__))
HOME = os.path.expanduser("~")
KEY = "1em9eYPdK8ErWPDH_kVJwx8cT9zk-qiBjIW2_qt4TogY"
EXTRA = os.path.join(HERE, "pharmacy_codes_extra.json")
CSV_GLOB = os.path.join(HOME, "我的雲端硬碟/蝦皮獲利計算表/倉庫貨號對照表-*.csv")
FIXES = {"B0022XS01": "B0022XS015"}

gc = gspread.service_account(filename=os.path.join(HOME, "我的雲端硬碟/scripts/service_account.json"))
rows = gc.open_by_key(KEY).worksheet("調貨單_明細").get_all_values()[1:]
base = {}
files = sorted(glob.glob(CSV_GLOB))
if files:
    with open(files[-1], encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            s = (r.get("倉庫貨號") or "").strip()
            if s: base[FIXES.get(s, s)] = (r.get("藥局商品編號") or "").strip()
extra = json.load(open(EXTRA, encoding="utf-8")) if os.path.exists(EXTRA) else {}
seen = collections.defaultdict(collections.Counter)
for r in rows:
    sku = r[4].strip()
    if sku:
        for m in re.findall(r"(?<!\d)(\d{7})(?!\d)", r[9]): seen[sku][m] += 1
added = {}
for sku, c in seen.items():
    if len(c) > 1: print("衝突，略過", sku, dict(c)); continue
    code = next(iter(c))
    if sku in base:
        if base[sku] != code: print("與對照表不同，略過", sku, base[sku], code)
        continue
    if extra.get(sku) != code: added[sku] = code
print(f"新學到 {len(added)} 筆", added)
if added and "--dry" not in sys.argv:
    extra.update(added)
    json.dump(dict(sorted(extra.items())), open(EXTRA, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
