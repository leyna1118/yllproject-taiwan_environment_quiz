#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 data.xlsx 轉成網頁用的 data/quiz-data.js

用法（在專案根目錄執行）：
    python tools/xlsx_to_js.py

需要的分頁與欄位：
  地區分類題 : 題號 / 情境 / 選項 / 對應地區類型 / 備註（備註寫「平手決勝題」的那題會被當成平手決勝依據）
  好玩活動題 : 地區 / 題號 / 題目 / 選項 / 對應主線任務（行政區）/ 知識派(-2)↔實踐派(+2) / 深度(-2)↔廣度(+2)
  好玩活動   : 地區類型 / 行政區 / 環保類別 / 行動／資源名稱 / 一句話說明 / 連結 / 狀態 / 備註 / 選擇（標 V 者才會被採用）
  延伸資源   : 地區類型 / 行政區 / 環保類別 / 行動／資源名稱 / 一句話說明 / 連結 / 狀態 / 備註（備註裡的「搭配主線任務：XXX」用來對應主線任務）

轉出來的 data/quiz-data.js 會定義 window.QUIZ_DATA，index.html 直接用 <script> 載入，
不需要架本機伺服器，雙擊 index.html 就能跑。
"""

import json
import os
import re
import sys
import unicodedata

try:
    import openpyxl
except ImportError:
    sys.exit("請先安裝 openpyxl：pip install openpyxl")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(ROOT, "data.xlsx")
OUT = os.path.join(ROOT, "data", "quiz-data.js")

AREAS = ["海邊", "城市", "鄉間", "山林"]

# 環保類別 → 結果頁標題用的行動主題
THEME_TITLES = {
    "海洋與海岸保育": "淨海之旅",
    "森林山林與生物多樣性": "山林尋跡之旅",
    "水資源與河川湖泊保育": "溯溪護水之旅",
    "氣候變遷與能源轉型": "低碳行動之旅",
    "循環經濟與減廢": "減塑生活之旅",
    "空氣品質與污染防治": "深呼吸之旅",
    "友善農業與食農教育": "食農體驗之旅",
    "永續城鄉與綠色生活": "慢城漫遊之旅",
    "動物保護與野生動物救援": "守護動物之旅",
    "環境教育與公民行動": "公民行動之旅",
}

# 環保類別 → 結果頁標題底下的一句呼應文案
TAGLINES = {
    "海洋與海岸保育": "這趟旅程，你想留給這片海的，是乾淨，不是垃圾",
    "森林山林與生物多樣性": "走慢一點，這片森林會讓你看見更多",
    "水資源與河川湖泊保育": "順著水的方向走，你會看見它從哪裡來",
    "氣候變遷與能源轉型": "每一度電、每一趟車，都是你為氣候投下的一票",
    "循環經濟與減廢": "這趟旅程，你可以什麼都不留下",
    "空氣品質與污染防治": "深呼吸一口，然後想想它為什麼乾淨",
    "友善農業與食農教育": "吃進嘴裡的每一口，都是一次選擇",
    "永續城鄉與綠色生活": "老房子還在，是因為有人願意留住它",
    "動物保護與野生動物救援": "牠們本來就住在這裡，我們只是路過",
    "環境教育與公民行動": "知道之後，就很難假裝不知道了",
}


def norm(v):
    """儲存格 → 去頭尾空白的字串；空值回傳空字串。"""
    if v is None:
        return ""
    return unicodedata.normalize("NFKC", str(v)).strip()


def num(v):
    if v is None or v == "":
        return 0
    return int(float(v))


def sheet_rows(wb, name, ncheck):
    ws = wb[name]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in row[:ncheck]):
            continue
        out.append(row)
    return out


def longest_common_substr(a, b):
    if not a or not b:
        return 0
    prev = [0] * (len(b) + 1)
    best = 0
    for i in range(1, len(a) + 1):
        cur = [0] * (len(b) + 1)
        for j in range(1, len(b) + 1):
            if a[i - 1] == b[j - 1]:
                cur[j] = prev[j - 1] + 1
                best = max(best, cur[j])
        prev = cur
    return best


def short_name(region):
    """把『苗栗淺山（通霄／苑裡等）』縮成『苗栗淺山』，給結果頁標題用。"""
    return re.sub(r"[（(].*?[)）]", "", region).strip() or region


def build():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    problems = []

    # ---------- Step 1：地區分類題 ----------
    step1, cur = [], None
    for row in sheet_rows(wb, "地區分類題", 4):
        qno, stem, opt, area, note = (norm(row[i]) for i in range(5))
        if qno:
            cur = {"id": qno, "stem": stem, "options": [], "tiebreak": "平手決勝" in note}
            step1.append(cur)
        if cur is None:
            continue
        if stem and not cur["stem"]:
            cur["stem"] = stem
        if note and "平手決勝" in note:
            cur["tiebreak"] = True
        if opt:
            if area not in AREAS:
                problems.append(f"[地區分類題] {cur['id']} 的對應地區類型「{area}」不在四種類型裡")
            cur["options"].append({"text": opt, "area": area})

    for q in step1:
        if len(q["options"]) != 4:
            problems.append(f"[地區分類題] {q['id']} 有 {len(q['options'])} 個選項（應為 4 個）")
    if not any(q["tiebreak"] for q in step1):
        problems.append("[地區分類題] 沒有任何一題標記「平手決勝題」，平手時將改用最後一題決定")

    # ---------- 主線任務（好玩活動，只取選擇欄標 V 的） ----------
    activities = {a: [] for a in AREAS}
    for row in sheet_rows(wb, "好玩活動", 6):
        area = norm(row[0])
        chosen = norm(row[8]) if len(row) > 8 else ""
        if area not in AREAS or chosen.upper() not in ("V", "Ｖ"):
            continue
        category = norm(row[2])
        if category not in THEME_TITLES:
            problems.append(f"[好玩活動] {norm(row[1])} 的環保類別「{category}」不在十大類裡")
        activities[area].append({
            "region": norm(row[1]),
            "short": short_name(norm(row[1])),
            "category": category,
            "theme": THEME_TITLES.get(category, "綠色旅遊"),
            "tagline": TAGLINES.get(category, ""),
            "name": norm(row[3]),
            "desc": norm(row[4]),
            "url": norm(row[5]),
        })

    for area in AREAS:
        if len(activities[area]) != 4:
            problems.append(f"[好玩活動] {area} 有 {len(activities[area])} 個主線任務（應為 4 個）")

    # ---------- Step 2：好玩活動題 ----------
    step2 = {a: [] for a in AREAS}
    cur_area = cur_q = None
    for row in sheet_rows(wb, "好玩活動題", 7):
        area, qno, stem, opt, key = (norm(row[i]) for i in range(5))
        know, depth = num(row[5]), num(row[6])
        if area:
            cur_area = area
        if qno:
            cur_q = {"id": qno, "stem": stem, "options": []}
            step2.setdefault(cur_area, []).append(cur_q)
        if cur_q is None:
            continue
        if stem and not cur_q["stem"]:
            cur_q["stem"] = stem
        if not opt:
            continue
        # 題目用簡稱（花蓮），主線任務用全名（花蓮縣新城鄉）→ 在同一個地區類型裡比對
        match = None
        for act in activities.get(cur_area, []):
            if key and (key in act["region"] or act["region"].startswith(key)):
                match = act
                break
        if match is None:
            problems.append(f"[好玩活動題] {cur_area} {cur_q['id']} 的「{key}」對不到任何主線任務")
        cur_q["options"].append({
            "text": opt,
            "key": match["region"] if match else key,
            "know": know,
            "depth": depth,
        })

    for area in AREAS:
        qs = step2.get(area, [])
        if len(qs) != 4:
            problems.append(f"[好玩活動題] {area} 有 {len(qs)} 題（應為 4 題）")
        for q in qs:
            if len(q["options"]) != 4:
                problems.append(f"[好玩活動題] {area} {q['id']} 有 {len(q['options'])} 個選項（應為 4 個）")
                continue
            axis = "know" if any(o["know"] for o in q["options"]) else "depth"
            vals = sorted(o[axis] for o in q["options"])
            if vals != [-2, -1, 1, 2]:
                problems.append(
                    f"[好玩活動題] {area} {q['id']} 的分數分佈是 {vals}（每題四個選項應該是 +2/+1/-1/-2 各一個）")
            other = "depth" if axis == "know" else "know"
            if any(o[other] for o in q["options"]):
                problems.append(f"[好玩活動題] {area} {q['id']} 同一題同時給了兩個維度的分數")

    # ---------- 延伸資源 ----------
    resources = {}
    for row in sheet_rows(wb, "延伸資源", 6):
        area = norm(row[0])
        note = norm(row[7]) if len(row) > 7 else ""
        m = re.search(r"搭配主線任務[：:]\s*([^；;]+)", note)
        label = m.group(1).strip() if m else ""
        # 用「同地區內字串重疊最長」的方式對應回主線任務
        best, best_score = None, 0
        for act in activities.get(area, []):
            score = longest_common_substr(re.sub(r"[（(].*?[)）]", "", label), act["name"])
            score += longest_common_substr(label, act["region"]) * 0.5
            if score > best_score:
                best, best_score = act, score
        if best is None or best_score < 2:
            problems.append(f"[延伸資源] {area}「{norm(row[3])}」的備註對不到主線任務（備註寫：{label or '（空白）'}）")
            continue
        resources.setdefault(best["region"], []).append({
            "region": norm(row[1]),
            "category": norm(row[2]),
            "name": norm(row[3]),
            "desc": norm(row[4]),
            "url": norm(row[5]),
        })

    for area in AREAS:
        for act in activities[area]:
            n = len(resources.get(act["region"], []))
            if n == 0:
                problems.append(f"[延伸資源] {act['region']}（{act['name']}）沒有任何延伸資源")

    data = {
        "areas": AREAS,
        "step1": step1,
        "step2": step2,
        "activities": activities,
        "resources": resources,
    }

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("// 這個檔案由 tools/xlsx_to_js.py 從 data.xlsx 自動產生，請不要直接編輯。\n")
        f.write("// 要改題目或資源請改 data.xlsx，然後重新執行： python tools/xlsx_to_js.py\n")
        f.write("window.QUIZ_DATA = ")
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write(";\n")

    print(f"已寫出 {os.path.relpath(OUT, ROOT)}")
    print(f"  地區分類題 {len(step1)} 題")
    for area in AREAS:
        print(f"  {area}：{len(step2.get(area, []))} 題、"
              f"{len(activities[area])} 個主線任務、"
              f"{sum(len(resources.get(a['region'], [])) for a in activities[area])} 筆延伸資源")

    if problems:
        print("\n⚠ 資料有以下問題，請回 data.xlsx 修正：")
        for p in problems:
            print("   -", p)
        return 1
    print("\n✅ 資料檢查全部通過")
    return 0


if __name__ == "__main__":
    sys.exit(build())
