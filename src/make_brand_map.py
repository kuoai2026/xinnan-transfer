#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
產生 src/brand_map.json —— 口罩貨號(B開頭) 的「4位產品線碼 → 品牌」對照。

做法：掃蝦皮獲利計算表「📦 商品列表」，每個 B 開頭 4 位碼取品名裡的品牌關鍵字、
多數決。再套 OVERRIDES 修正混雜的碼。輸出可手改的 JSON。

平常不用跑；只有蝦皮獲利表新增品牌、或發現分類怪怪時，跑一次再手修 JSON。
"""
import openpyxl, re, collections, json, os

HOME = os.path.expanduser("~")
SHOPEE = os.path.join(HOME, "我的雲端硬碟/蝦皮獲利計算表/蝦皮獲利計算表_含稅版_updated.xlsx")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "repos",
                   "xinnan-transfer", "src", "brand_map.json")
OUT = os.path.abspath(os.path.expanduser(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "brand_map.json")))

# 門市會叫貨的口罩品牌關鍵字（順序＝優先序，長/獨特的放前面）
BRANDS = [
    "安心罩護", "鼻恩恩", "幸福物語", "吉伊卡哇", "健康天使", "愛貝恩", "舒膚康", "挺立舒",
    "郡昱", "興安", "昌明", "匠心", "中衛", "聚泰", "德冠", "凱上", "凱馺", "億宏", "上好",
    "星業", "佑合", "佑和", "水舞", "華淨", "艾爾絲", "明基", "盛籐", "天心", "新寵兒",
    "睿昱", "摩戴舒", "活力安", "艾可兒", "丞威", "沐慷", "康丞", "索菲亞", "萊潔", "月池",
    "優美特", "3M", "宇宙", "卡娜赫拉", "大甲媽", "MissMix", "怡安", "荷康", "BNN", "YOKU",
    "A.O.K",
]

# 混雜碼手動修正（多數決判不準的）
OVERRIDES = {
    "B0031": "幸福物語",   # 幸福物語 V美型 / 艾爾絲混賣場
    "B0032": "幸福物語",
    "B0033": "艾爾絲",
    "B0034": "吉伊卡哇",
    "B0072": "德冠",        # DG / 大甲媽 都是德冠
    "B0091": "昌明",
    "B0092": "安心罩護",    # 安心罩護就是昌明的牌子，門市講「安心罩護」
    "B0122": "MissMix",
    "B0132": "盛籐",
    "B0143": "凱上",
    "B0202": "億宏",
    "B0353": "優美特",
    "B0362": "普潔",
    "B0442": "幸福物語",
    "B0502": "迪士尼",
    "B0011": "BNN",
    "B0012": "BNN",
}


def brand_of(name):
    for b in BRANDS:
        if b in name:
            return b
    m = re.match(r"^[\[【〔]([^\]】〕]{1,6})", name)
    if m:
        return m.group(1)
    return re.split(r"[｜|／/ 　\-∣]", name)[0][:6]


def main():
    wb = openpyxl.load_workbook(SHOPEE, read_only=True, data_only=True)
    ws = wb["📦 商品列表"]
    by4 = collections.defaultdict(collections.Counter)
    for r in ws.iter_rows(min_row=4, values_only=True):
        if not (r[1] and r[2]):
            continue
        sku, name = str(r[1]).strip(), str(r[2]).strip()
        m = re.match(r"^B(\d{4})", sku)
        if not m:
            continue
        by4["B" + m.group(1)][brand_of(name)] += 1

    codes = {}
    review = []
    for code, c in sorted(by4.items()):
        if code in OVERRIDES:
            codes[code] = OVERRIDES[code]
            continue
        top, n = c.most_common(1)[0]
        total = sum(c.values())
        codes[code] = top
        if n / total < 0.8 or total < 2:
            review.append({"code": code, "picked": top, "counts": dict(c.most_common(4))})

    out = {
        "_說明": "口罩貨號前4碼 → 品牌。可手改。keywords 是比對品名的兜底清單。",
        "codes": codes,
        "keywords": BRANDS,
        "_待人工複查": review,
    }
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"寫出 {OUT}  ({len(codes)} 碼, {len(review)} 待複查)")


if __name__ == "__main__":
    main()
