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
