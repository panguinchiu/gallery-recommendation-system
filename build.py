"""從 商業推薦_metadata_v2.xlsx 與 訓練圖片/ 產生網頁資料 data.js。

用法：python3 build.py
"""
import json
import os
import re

import openpyxl

ROOT = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = "訓練圖片"


def norm(s):
    return re.sub(r"\s+", "", s or "")


def split_name(stem):
    """把檔名拆成 (編號, 名稱)，編號可能在前或在後。"""
    m = re.match(r"^([\d-]+)\s*(.+)$", stem)
    if m:
        return m.group(1), m.group(2)
    m = re.match(r"^(.+?)([\d-]+)$", stem)
    if m:
        return m.group(2), m.group(1)
    return "", stem


def load_metadata():
    wb = openpyxl.load_workbook(os.path.join(ROOT, "商業推薦_metadata_v2.xlsx"), read_only=True)
    ws = wb["印花Metadata"]
    rows = ws.iter_rows(values_only=True)
    header = list(next(rows))
    col = {h: i for i, h in enumerate(header) if h}

    def pick(r, *names):
        return [r[col[n]] for n in names if r[col[n]]]

    meta = {}
    for r in rows:
        if not r[col["印花編號"]]:
            continue
        code, name = str(r[col["印花編號"]]).strip(), str(r[col["印花名稱"]]).strip()
        meta[norm(code + name)] = {
            "brand": r[col["品牌代表性"]],
            "themes": pick(r, "主題標籤1", "主題標籤2", "主題標籤3"),
            "styles": pick(r, "風格標籤1", "風格標籤2", "風格標籤3"),
            "art": pick(r, "畫風類型1", "畫風類型2", "畫風類型3"),
            "density": r[col["圖案密度"]],
            "taiwan": r[col["台灣元素強度"]],
            "colors": pick(r, "適合色系1", "適合色系2", "適合色系3", "適合色系4"),
            "recolor": r[col["改色容易度"]],
            "audience": pick(r, "適合客群1", "適合客群2", "適合客群3"),
            "story": r[col["印花故事或設計描述"]],
            "pitch": r[col["推薦話術"]],
        }
    return meta


def main():
    meta = load_metadata()
    items = []
    for fn in sorted(os.listdir(os.path.join(ROOT, IMG_DIR))):
        stem, ext = os.path.splitext(fn)
        if ext.lower() not in (".jpg", ".jpeg", ".png"):
            continue
        code, name = split_name(stem)
        item = {"code": code, "name": name.strip(), "img": f"{IMG_DIR}/{fn}"}
        m = meta.get(norm(code + name)) or meta.get(norm(stem))
        if m:
            item.update(m)
        items.append(item)

    tagged = sum(1 for i in items if "themes" in i)
    with open(os.path.join(ROOT, "data.js"), "w", encoding="utf-8") as f:
        f.write("window.PATTERNS = ")
        json.dump(items, f, ensure_ascii=False, indent=1)
        f.write(";\n")
    print(f"{len(items)} 張圖，{tagged} 張有標籤")


if __name__ == "__main__":
    main()
