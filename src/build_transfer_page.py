#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
產生「倉庫調貨單」門市端頁面 index.html。

資料來源：
  1. ~/我的雲端硬碟/產品進貨單/庫存盤點表/庫存盤點表_YYYY-MM-DD.xlsx （最新一份原始下載）
       欄位：主貨號 / 系統商品名 / 選項貨號 / 儲位 / 重量 / 規格一 / 規格二 / 可售庫存 / 總庫存 ...
  2. ~/我的雲端硬碟/蝦皮獲利計算表/蝦皮獲利計算表_含稅版_updated.xlsx → 「📦 商品列表」
       貨號 → 商品名稱（當作別名，提高搜尋命中）

輸出：<repo>/index.html （由 template.html 內嵌 catalog + config 產生）

設定檔：transfer_page/config.json
  { "endpoint": "...Apps Script /exec URL...", "token": "...", "persons": ["小賴",...],
    "repo_dir": "~/repos/xinnan-transfer" }
"""
import json, re, sys, os, glob, datetime, html

HERE = os.path.dirname(os.path.abspath(__file__))
HOME = os.path.expanduser("~")
INV_DIR   = os.path.join(HOME, "我的雲端硬碟/產品進貨單/庫存盤點表")
SHOPEE_XLSX = os.path.join(HOME, "我的雲端硬碟/蝦皮獲利計算表/蝦皮獲利計算表_含稅版_updated.xlsx")
TEMPLATE  = os.path.join(HERE, "template.html")
CONFIG    = os.path.join(HERE, "config.json")

CAT_BY_PREFIX = [
    ("PW", "充電線材"), ("PG", "益智玩具"), ("BD", "入浴球"), ("NO", "入浴球"),
    ("A", "保健食品"), ("B", "醫療口罩"), ("C", "傷口護理"), ("D", "營養補充"),
    ("E", "嬰幼兒清潔"), ("F", "生活用品"), ("G", "3C配件"), ("H", "個人護理"),
    ("I", "食品飲品"), ("J", "成人用品"),
]

def log(*a): print(*a, file=sys.stderr)

def latest_inventory():
    pat = re.compile(r"庫存盤點表_(\d{4}-\d{2}-\d{2})\.xlsx$")
    best = None
    for p in glob.glob(os.path.join(INV_DIR, "庫存盤點表_*.xlsx")):
        m = pat.search(os.path.basename(p))
        if not m:
            continue
        if best is None or m.group(1) > best[1]:
            best = (p, m.group(1))
    if not best:
        sys.exit("找不到 庫存盤點表_YYYY-MM-DD.xlsx")
    return best

_BRACKET = re.compile(r"[【\[（(〔「〈][^】\])）〕」〉]*[】\])）〕」〉]")
_SEO_STOP = [
    "醫療口罩", "醫用口罩", "醫療級", "醫用", "醫療", "口罩", "台灣製造", "台灣製", "台灣",
    "3D立體口罩", "3D立體", "4D立體", "獨家", "聯名", "限量", "現貨出清", "現貨",
    "免運", "優惠", "賣場", "搶先上市", "最新", "新款", "新色", "熱銷", "特價", "促銷",
    "獨立包裝", "獨立包", "單片包", "雙鋼印", "MD雙鋼印", "MIT", "鼻樑壓條", "彈性耳繩",
    "奈米抑菌", "舒適透氣", "舒適", "透氣", "親膚", "親子", "系列", "款式", "多款", "多色",
    "彩色", "素色", "新年", "過年", "虎年", "兔年", "龍年", "蛇年", "聖誕", "聖誕節",
    "母親節", "情人節", "萬聖節", "端午", "中秋", "節慶", "虎爺", "國旗", "特殊",
    "洗白牛仔", "買就送", "加價購", "組合", "賣場搜尋", "另有", "現貨供應",
]
_BRAND_BAD = ("口罩", "聯名", "限量", "免運", "優惠", "獨家", "贈", "買就送", "熱銷", "新",
              "搶先", "現貨", "促銷", "加價", "組合", "折扣", "藥局直營", "信男藥局",
              "藥局", "守護天使", "媽媽最愛", "限定")
_AGE = ("嬰幼", "幼幼", "幼童", "兒童", "中童", "大童", "成人", "小顏", "小臉", "加大", "XL")
_TYPE = ("平面", "立體", "全彩", "蝶型", "蝶形", "魚口", "KF94", "KN95", "N95", "鈔票",
         "呼吸", "活性碳", "泡泡", "耳繩", "耳掛")


def series_label(sysname):
    brands = re.findall(r"[【\[〔「]([^】\]〕」]{1,8})[】\]〕」]", sysname)
    brand = next((b for b in brands if not any(x in b for x in _BRAND_BAD)), "")
    s = _BRACKET.sub(" ", sysname)
    for w in _SEO_STOP:
        s = s.replace(w, " ")
    for ch in "★☆｜∣|•‧·":
        s = s.replace(ch, " ")
    raw = [t for t in re.split(r"[\s／/,、\-]+", s)
           if t and t not in ("入", "盒", "/盒", "版", "型")]
    seen, toks = set(), []
    for t in raw:
        if t not in seen:
            seen.add(t)
            toks.append(t)
    ages = [t for t in toks if any(a in t for a in _AGE)][:2]
    types = [t for t in toks if any(a in t for a in _TYPE)][:1]
    rest = [t for t in toks if t not in ages and t not in types]
    picked = ([brand] if brand else []) + rest[:1] + ages + types
    if not picked:
        picked = rest[:3] or toks[:3]
    label = " ".join(dict.fromkeys(picked)).strip()
    label = re.sub(r"^[^\w一-鿿]+", "", label)
    label = re.sub(r"\s+", " ", label.replace("!", "").replace("『", "").replace("』", "")
                   .replace("{", "").replace("}", "").replace("☘", "").replace("︎", "")).strip()
    if not label or len(label) > 24:
        # 退回：品牌 + 第一段中文
        cjk = re.findall(r"[一-鿿A-Za-z0-9]{2,10}", _BRACKET.sub(" ", sysname))
        label = " ".join(dict.fromkeys(([brand] if brand else []) + cjk[:2]))[:24] or sysname
    return label


def category(sku):
    s = (sku or "").upper()
    for pre, name in CAT_BY_PREFIX:
        if s.startswith(pre):
            return name
    return "其他"

def load_aliases():
    import openpyxl
    if not os.path.exists(SHOPEE_XLSX):
        log("警告：找不到蝦皮獲利計算表，略過別名")
        return {}
    wb = openpyxl.load_workbook(SHOPEE_XLSX, read_only=True, data_only=True)
    ws = wb["📦 商品列表"] if "📦 商品列表" in wb.sheetnames else wb[wb.sheetnames[0]]
    alias = {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        if len(row) < 3:
            continue
        sku, name = row[1], row[2]
        if not sku or not name:
            continue
        sku = str(sku).strip()
        name = str(name).strip()
        if sku and name and sku not in alias:
            alias[sku] = name
    wb.close()
    return alias

def norm(v):
    if v is None:
        return ""
    return str(v).strip()

def load_catalog(inv_path, alias):
    import openpyxl
    wb = openpyxl.load_workbook(inv_path, read_only=True, data_only=True)
    ws = wb["庫存盤點"] if "庫存盤點" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {name: i for i, name in enumerate(header)}
    def col(r, key):
        i = idx.get(key)
        return r[i] if i is not None and i < len(r) else None

    series_map = {}     # 系統商品名 -> list of items
    series_alias = {}   # 系統商品名 -> set of 蝦皮品名 (供搜尋，系列層級)
    order = []
    n_items = 0
    for r in rows:
        sku = norm(col(r, "選項貨號"))
        sysname = norm(col(r, "系統商品名"))
        g1 = norm(col(r, "規格一"))
        g2 = norm(col(r, "規格二"))
        avail = col(r, "可售庫存")
        if not sku or not sysname:
            continue
        try:
            avail = int(avail) if avail is not None and str(avail).strip() != "" else 0
        except (TypeError, ValueError):
            avail = 0
        disp = g1 or g2 or sysname
        if g1 and g2 and g2 not in ("-", ""):
            disp = f"{g1}／{g2}"
        item = {"s": sku, "g": disp, "v": avail}
        if sysname not in series_map:
            series_map[sysname] = []
            series_alias[sysname] = set()
            order.append(sysname)
        series_map[sysname].append(item)
        a = alias.get(sku)
        if a:
            series_alias[sysname].add(a)
        n_items += 1
    wb.close()

    catalog = []
    for sysname in order:
        items = series_map[sysname]
        cat = category(items[0]["s"])
        # 系列層級搜尋別名：蝦皮品名（去重、限長），只給搜尋用不顯示
        ali = " ".join(sorted(series_alias[sysname]))[:300]
        entry = {"n": sysname, "d": series_label(sysname), "c": cat, "i": items}
        if ali:
            entry["x"] = ali
        catalog.append(entry)
    return catalog, n_items

def main():
    inv_path, inv_date = latest_inventory()
    inv_mtime = datetime.datetime.fromtimestamp(os.path.getmtime(inv_path))
    log(f"庫存來源：{os.path.basename(inv_path)}  (mtime {inv_mtime:%Y-%m-%d %H:%M})")

    alias = load_aliases()
    log(f"別名：{len(alias)} 筆")
    catalog, n_items = load_catalog(inv_path, alias)
    log(f"catalog：{len(catalog)} 系列 / {n_items} 品項")

    with open(CONFIG, encoding="utf-8") as f:
        cfg = json.load(f)

    config_json = {
        "endpoint": cfg.get("endpoint", ""),
        "token": cfg.get("token", ""),
        "persons": cfg.get("persons", ["小賴", "瓊如", "柔柔", "阿霞"]),
    }
    now = datetime.datetime.now()
    build_info = {
        "stock": f"{inv_date}",
        "built": f"{now:%m/%d %H:%M}",
    }

    with open(TEMPLATE, encoding="utf-8") as f:
        tpl = f.read()

    out = (tpl
           .replace("__CONFIG_JSON__", json.dumps(config_json, ensure_ascii=False))
           .replace("__BUILD_INFO_JSON__", json.dumps(build_info, ensure_ascii=False))
           .replace("__CATALOG_JSON__", json.dumps(catalog, ensure_ascii=False, separators=(",", ":"))))

    repo_dir = os.path.expanduser(cfg.get("repo_dir", os.path.join(HOME, "repos/xinnan-transfer")))
    os.makedirs(repo_dir, exist_ok=True)
    out_path = os.path.join(repo_dir, "index.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(out)
    size_kb = os.path.getsize(out_path) / 1024
    log(f"寫出 {out_path}  ({size_kb:.0f} KB)")
    print(out_path)

if __name__ == "__main__":
    main()
