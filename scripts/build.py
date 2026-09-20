# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""統計トピックス No.149 のテキストと人口推計 Excel からデータを抽出し index.html を生成する。"""
import json
import re
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
PDF_TEXT = RAW / "topi149_01.txt"
XLSX_PATH = RAW / "05k2024-3.xlsx"
TEMPLATE = Path(__file__).resolve().parent / "template.html"
OUT_PATH = ROOT / "index.html"

AGE_COLUMNS = ["total", "under15", "age15_64", "age65", "age70", "age75",
               "age80", "age85", "age90", "age95", "age100"]

TREND_RE = re.compile(
    r"^\s*(\d{4})年?\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+"
    r"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$",
    re.M,
)

COUNTRY_RE = re.compile(r"^\s*(\d{1,2})\s+(\S+)\s+(\d+)\s+(\d+)\s+([\d.]+)\s*$", re.M)


def parse_trend(text):
    """表2: 65歳以上人口及び割合の推移 (1950-2050)。人口は万人、割合は%。"""
    rows = []
    for m in TREND_RE.finditer(text):
        year, total, p65, p70, p75, p80, r65, r70, r75, r80 = m.groups()
        rows.append({
            "year": int(year), "total": int(total),
            "pop65": int(p65), "pop70": int(p70), "pop75": int(p75), "pop80": int(p80),
            "rate65": float(r65), "rate70": float(r70), "rate75": float(r75), "rate80": float(r80),
        })
    if len(rows) != 22:
        raise ValueError(f"推移表の行数が想定外です: {len(rows)} 行 (想定 22 行)")
    return rows


def parse_age_detail(text, year=2026):
    """表1: 指定年の年齢区分別人口 (万人) と割合 (%) を男女別に抽出する。"""
    lines = text.splitlines()
    start = next((i for i, l in enumerate(lines) if l.strip() == f"{year}年"), None)
    if start is None:
        raise ValueError(f"表1の{year}年ブロックが見つかりません")
    pop, ratio = {}, {}
    for line in lines[start + 1:]:
        if re.match(r"^\s*\d{4}年\s*$", line):
            break
        m = re.match(r"^\s*(男女計|男|女)\s+(.+)$", line)
        if not m:
            continue
        label, numstr = m.group(1), m.group(2)
        nums = numstr.split()
        if len(nums) != len(AGE_COLUMNS):
            raise ValueError(f"表1の列数が想定外です: {year}年 {label} {len(nums)} 列")
        (ratio if "." in numstr else pop)[label] = [
            float(n) if "." in numstr else int(n) for n in nums
        ]
    if len(pop) != 3 or len(ratio) != 3:
        raise ValueError(f"表1の{year}年ブロックの行が不足: pop={list(pop)}, ratio={list(ratio)}")
    return {"population": pop, "ratio": ratio}


def parse_countries(text):
    """表3: 65歳以上人口の割合 上位10か国 (2026年)。人口は万人。"""
    start = text.find("上位 10 か国")
    if start == -1:
        raise ValueError("表3の見出しが見つかりません")
    end = text.find("資料", start)
    rows = []
    for m in COUNTRY_RE.finditer(text[start:end]):
        rank, name, total, pop65, rate65 = m.groups()
        rows.append({"rank": int(rank), "name": name,
                     "total": int(total), "pop65": int(pop65), "rate65": float(rate65)})
    if len(rows) != 10:
        raise ValueError(f"国際比較表の行数が想定外です: {len(rows)} 行 (想定 10 行)")
    return rows


def parse_prefectures(xlsx_path=XLSX_PATH):
    """人口推計 第3表 (2024-10-01 現在、都道府県・年齢3区分) から高齢化率を算出。単位は千人。"""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["第3表"]
    national, prefs = None, []
    for row in ws.iter_rows(values_only=True):
        a, b, c = row[0], row[1], row[2]
        under15, working, elderly, elderly75 = row[4], row[5], row[6], row[7]
        if isinstance(a, str) and a.strip().startswith("全国"):
            national = {"name": "全国", "under15": under15, "age15_64": working,
                        "pop65": elderly, "pop75": elderly75}
        elif isinstance(b, str) and re.fullmatch(r"\d{2}", b.strip()) and isinstance(c, str):
            name = c.strip().replace("　", "")
            total = under15 + working + elderly
            prefs.append({
                "code": b.strip(), "name": name,
                "under15": under15, "age15_64": working,
                "pop65": elderly, "pop75": elderly75, "total": total,
                "rate65": round(elderly / total * 100, 1),
                "rate75": round(elderly75 / total * 100, 1),
            })
    if national is None or len(prefs) != 47:
        raise ValueError(f"都道府県データの抽出に失敗: national={national is not None}, 件数={len(prefs)}")
    return {"national": national, "prefectures": prefs}


def build_data():
    """全データを統合し、公式既知値 (統計トピックス No.149) でアサートする。"""
    text = PDF_TEXT.read_text(encoding="utf-8")
    trend = parse_trend(text)
    age = parse_age_detail(text, 2026)
    countries = parse_countries(text)
    pref = parse_prefectures()

    i65, i75 = AGE_COLUMNS.index("age65"), AGE_COLUMNS.index("age75")
    t2026 = next(r for r in trend if r["year"] == 2026)
    prev = next(r for r in trend if r["year"] == 2025)
    assert trend[0]["year"] == 1950 and trend[0]["rate65"] == 4.9
    assert trend[-1]["year"] == 2050 and trend[-1]["rate65"] == 37.1
    assert t2026["pop65"] == 3624 and t2026["rate65"] == 29.6
    assert age["population"]["男女計"][i65] == 3624
    assert age["ratio"]["男女計"][i65] == 29.6
    assert age["population"]["男女計"][i75] == 2168
    assert age["ratio"]["男女計"][i75] == 17.7
    assert countries[0]["name"] == "日本" and countries[0]["rate65"] == 29.6
    assert next(p for p in pref["prefectures"] if p["name"] == "秋田県")["rate65"] == 39.5
    assert pref["national"]["pop65"] == 36243

    national_total = (pref["national"]["under15"] + pref["national"]["age15_64"]
                      + pref["national"]["pop65"])
    detail_keys = AGE_COLUMNS[3:]  # age65..age100
    return {
        "asof": "2026年9月15日現在",
        "prefAsof": "2024年10月1日現在",
        "summary": {
            "pop65": t2026["pop65"], "rate65": t2026["rate65"],
            "pop65Diff": t2026["pop65"] - prev["pop65"],
            "rate65Diff": round(t2026["rate65"] - prev["rate65"], 1),
            "pop75": age["population"]["男女計"][i75],
            "rate75": age["ratio"]["男女計"][i75],
            "male65": age["population"]["男"][i65],
            "female65": age["population"]["女"][i65],
        },
        "trend": trend,
        "ageDetail": {
            "labels": ["65歳以上", "70歳以上", "75歳以上", "80歳以上",
                       "85歳以上", "90歳以上", "95歳以上", "100歳以上"],
            "male": [age["population"]["男"][AGE_COLUMNS.index(k)] for k in detail_keys],
            "female": [age["population"]["女"][AGE_COLUMNS.index(k)] for k in detail_keys],
        },
        "prefectures": pref["prefectures"],
        "nationalRate65": round(pref["national"]["pop65"] / national_total * 100, 1),
        "countries": countries,
    }


def generate_html(data):
    template = TEMPLATE.read_text(encoding="utf-8")
    return template.replace("__DATA__", json.dumps(data, ensure_ascii=False))


def main():
    html = generate_html(build_data())
    OUT_PATH.write_text(html, encoding="utf-8")
    print(f"OK: {OUT_PATH} ({OUT_PATH.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
