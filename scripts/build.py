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
