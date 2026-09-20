# 日本の高齢者人口ダッシュボード 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 総務省統計局の最新データ (統計トピックス No.149、2026-09-20公表) から高齢者人口データを抽出し、ダッシュボード型の静的 index.html を生成して GitHub Pages で公開する。

**Architecture:** `scripts/fetch.py` (stdlib のみ) が PDF/Excel を `data/raw/` にダウンロードし pdftotext でテキスト化 → `scripts/build.py` (PEP723: openpyxl) がテキストと Excel をパースして公式既知値でアサート → `scripts/template.html` に JSON を埋め込み自己完結 `index.html` を生成 → gh CLI で public リポジトリ `kklab-kimi-older-people` を作成し Pages (main / root) を有効化。

**Tech Stack:** Python 3 + uv run (PEP723 インライン依存: openpyxl)、pdftotext (poppler)、pytest (テスト)、Chart.js v4 (CDN)、gh CLI (認証済み: katzkawai)。

**仕様書:** `docs/superpowers/specs/2026-09-21-elderly-population-dashboard-design.md`

**検証済みの事実 (実データで確認済み):**
- `data/raw/topi149_01.txt` (pdftotext -layout 済み): 表2 推移は22行 (1950,1955,...,2020,2025,2026,2030,...,2050)。2026年行: `12262 3624 2911 2168 1288 | 29.6 23.7 17.7 10.5`
- 表3 国際比較: 「上位 10 か国」文字列の後に10行。1位 日本 29.6、2位 イタリア 25.6
- 表1: `2026年` 単独行の後、男女計/男/女 の人口行 (整数11列) → 割合行 (小数11列) → `2025年` 行で終了。列順: 総人口, 15歳未満, 15〜64歳, 65歳以上, 70歳以上, 75歳以上, 80歳以上, 85歳以上, 90歳以上, 95歳以上, 100歳以上
- `data/raw/05k2024-3.xlsx`: シート名「第3表」。A列「全国…」が全国行、B列が '01'〜'47' の2桁文字列で C列に県名 (全角空白パディング)。E,F,G,H列 = 15歳未満, 15〜64歳, 65歳以上, うち75歳以上。**単位は千人** (全国 65歳以上 = 36243)。秋田県: 354/896 = 39.5%
- 要約カード用: 前年差 = 3624 − 3622 = +2万人、高齢化率差 = 29.6 − 29.4 = +0.2pt (推移表の2025年行は確定値ベース改訂済みで 3622)

---

### Task 1: データ取得スクリプト scripts/fetch.py

**Files:**
- Create: `scripts/fetch.py`

- [ ] **Step 1: fetch.py を作成**

```python
#!/usr/bin/env python3
"""総務省統計局の公開データを data/raw/ にダウンロードし、PDF をテキスト化する。"""
import subprocess
import urllib.request
from pathlib import Path

RAW = Path(__file__).resolve().parents[1] / "data" / "raw"

DOWNLOADS = [
    ("https://www.stat.go.jp/data/topics/pdf/topi149_01.pdf", "topi149_01.pdf", 50_000),
    ("https://www.stat.go.jp/data/jinsui/2024np/zuhyou/05k2024-3.xlsx", "05k2024-3.xlsx", 10_000),
]


def download(url, dest, min_size):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as res:
        if res.status != 200:
            raise RuntimeError(f"ダウンロード失敗: {url} (HTTP {res.status})")
        data = res.read()
    if len(data) < min_size:
        raise RuntimeError(f"ファイルが小さすぎます: {dest} ({len(data)} bytes)")
    dest.write_bytes(data)
    print(f"OK: {dest.name} ({len(data):,} bytes)")


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    for url, name, min_size in DOWNLOADS:
        download(url, RAW / name, min_size)
    subprocess.run(
        ["pdftotext", "-layout", str(RAW / "topi149_01.pdf"), str(RAW / "topi149_01.txt")],
        check=True,
    )
    print(f"OK: topi149_01.txt ({(RAW / 'topi149_01.txt').stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 実行して成功を確認**

Run: `cd /home/katzkawai/kklab-kimi-older-people && python3 scripts/fetch.py`
Expected: `OK: topi149_01.pdf (...)`, `OK: 05k2024-3.xlsx (...)`, `OK: topi149_01.txt (...)` の3行。stdlib のみなので uv 不要。

- [ ] **Step 3: コミット**

```bash
git add scripts/fetch.py && git commit -m "feat: 総務省データのダウンロードスクリプト"
```

---

### Task 2: 推移表パーサ parse_trend (TDD)

**Files:**
- Create: `scripts/build.py`
- Test: `tests/test_build.py`

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_build.py`:

```python
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import build  # noqa: E402

RAW = Path(__file__).resolve().parents[1] / "data" / "raw"


@pytest.fixture(scope="module")
def text():
    return (RAW / "topi149_01.txt").read_text(encoding="utf-8")


def test_parse_trend_row_count(text):
    assert len(build.parse_trend(text)) == 22


def test_parse_trend_known_values(text):
    by_year = {r["year"]: r for r in build.parse_trend(text)}
    assert by_year[1950]["rate65"] == 4.9
    assert by_year[2026]["pop65"] == 3624
    assert by_year[2026]["rate65"] == 29.6
    assert by_year[2050]["rate65"] == 37.1
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `cd /home/katzkawai/kklab-kimi-older-people && uv run --with pytest python -m pytest tests/ -v`
Expected: FAIL (ModuleNotFoundError: No module named 'build')

- [ ] **Step 3: build.py に parse_trend を実装**

`scripts/build.py`:

```python
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
```

- [ ] **Step 4: テストが通ることを確認**

Run: `uv run --with pytest --with openpyxl python -m pytest tests/ -v`
Expected: 2 passed

- [ ] **Step 5: コミット**

```bash
git add scripts/build.py tests/test_build.py && git commit -m "feat: 推移表(表2)のパーサ"
```

---

### Task 3: 表1 (年齢区分) と表3 (国際比較) のパーサ (TDD)

**Files:**
- Modify: `scripts/build.py`
- Test: `tests/test_build.py`

- [ ] **Step 1: 失敗するテストを追加**

`tests/test_build.py` に追記:

```python
def test_parse_age_detail(text):
    age = build.parse_age_detail(text, 2026)
    i65 = build.AGE_COLUMNS.index("age65")
    i75 = build.AGE_COLUMNS.index("age75")
    assert age["population"]["男女計"][i65] == 3624
    assert age["ratio"]["男女計"][i65] == 29.6
    assert age["population"]["男女計"][i75] == 2168
    assert age["ratio"]["男女計"][i75] == 17.7
    assert age["population"]["男"][i65] == 1571
    assert age["population"]["女"][i65] == 2054


def test_parse_countries(text):
    countries = build.parse_countries(text)
    assert len(countries) == 10
    assert countries[0] == {"rank": 1, "name": "日本", "total": 12262, "pop65": 3624, "rate65": 29.6}
    assert countries[1]["name"] == "イタリア" and countries[1]["rate65"] == 25.6
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `uv run --with pytest --with openpyxl python -m pytest tests/ -v`
Expected: 2 failed (AttributeError: module 'build' has no attribute 'parse_age_detail' / 'parse_countries')、2 passed

- [ ] **Step 3: build.py に2関数を実装**

`scripts/build.py` の `TREND_RE` 定義の後に追記:

```python
COUNTRY_RE = re.compile(r"^\s*(\d{1,2})\s+(\S+)\s+(\d+)\s+(\d+)\s+([\d.]+)\s*$", re.M)
```

`parse_trend` の後に追記:

```python
def parse_age_detail(text, year=2026):
    """表1: 指定年の年齢区分別人口 (万人) と割合 (%) を男女別に抽出する。"""
    lines = text.splitlines()
    start = next(i for i, l in enumerate(lines) if l.strip() == f"{year}年")
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
```

注意: テキスト中に「表３」は2回出現する (本文参照と表題)。`"上位 10 か国"` で探すのはその回避策。変更しないこと。

- [ ] **Step 4: テストが通ることを確認**

Run: `uv run --with pytest --with openpyxl python -m pytest tests/ -v`
Expected: 4 passed

- [ ] **Step 5: コミット**

```bash
git add scripts/build.py tests/test_build.py && git commit -m "feat: 表1(年齢区分)・表3(国際比較)のパーサ"
```

---

### Task 4: 都道府県 Excel パーサ parse_prefectures (TDD)

**Files:**
- Modify: `scripts/build.py`
- Test: `tests/test_build.py`

- [ ] **Step 1: 失敗するテストを追加**

`tests/test_build.py` に追記:

```python
def test_parse_prefectures():
    pref = build.parse_prefectures(RAW / "05k2024-3.xlsx")
    assert len(pref["prefectures"]) == 47
    akita = next(p for p in pref["prefectures"] if p["name"] == "秋田県")
    assert akita["rate65"] == 39.5
    assert pref["national"]["pop65"] == 36243  # 単位: 千人
    assert pref["national"]["pop65"] // 10 == 3624
```

- [ ] **Step 2: テストが失敗することを確認**

Run: `uv run --with pytest --with openpyxl python -m pytest tests/ -v`
Expected: 1 failed (AttributeError: ... 'parse_prefectures')、4 passed

- [ ] **Step 3: build.py に実装**

`scripts/build.py` の `parse_countries` の後に追記:

```python
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
```

- [ ] **Step 4: テストが通ることを確認**

Run: `uv run --with pytest --with openpyxl python -m pytest tests/ -v`
Expected: 5 passed

- [ ] **Step 5: コミット**

```bash
git add scripts/build.py tests/test_build.py && git commit -m "feat: 都道府県別高齢化率のパーサ"
```

---

### Task 5: データ統合 build_data + template.html + index.html 生成

**Files:**
- Modify: `scripts/build.py`
- Create: `scripts/template.html`

- [ ] **Step 1: build.py に build_data / generate_html / main を追記**

`scripts/build.py` の末尾に追記:

```python
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
        "national2024": pref["national"],
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
```

- [ ] **Step 2: scripts/template.html を作成**

```html
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>日本の高齢者人口ダッシュボード | 総務省統計局のデータより</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
  :root { --ink:#1f2937; --sub:#6b7280; --accent:#c2410c; --blue:#2563eb; --line:#e5e7eb; --bg:#f8fafc; }
  * { box-sizing: border-box; }
  body { margin:0; font-family:"Hiragino Kaku Gothic ProN","Hiragino Sans","Noto Sans JP",Meiryo,sans-serif; color:var(--ink); background:var(--bg); }
  header { background:#fff; border-bottom:1px solid var(--line); padding:24px 16px; }
  .wrap { max-width:1080px; margin:0 auto; }
  h1 { font-size:1.5rem; margin:0 0 6px; }
  .asof { color:var(--sub); font-size:.85rem; }
  .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:12px; margin:20px 0; }
  .card { background:#fff; border:1px solid var(--line); border-radius:10px; padding:16px; }
  .card .label { font-size:.8rem; color:var(--sub); }
  .card .value { font-size:1.8rem; font-weight:700; margin:4px 0; }
  .card .sub2 { font-size:.85rem; color:var(--accent); }
  .panel { background:#fff; border:1px solid var(--line); border-radius:10px; padding:16px; margin-bottom:20px; }
  .panel h2 { font-size:1.05rem; margin:0 0 8px; }
  .panel .note { font-size:.78rem; color:var(--sub); margin:8px 0 0; }
  .chart-box { position:relative; }
  #box-trend { height:340px; }
  #box-pref { height:1000px; }
  #box-age, #box-world { height:320px; }
  footer { color:var(--sub); font-size:.8rem; padding:24px 16px 48px; }
  footer a { color:var(--blue); }
</style>
</head>
<body>
<header><div class="wrap">
  <h1>日本の高齢者人口ダッシュボード</h1>
  <div class="asof">総務省統計局「統計からみた我が国の高齢者」(統計トピックス No.149、令和8年9月20日公表) — <span id="asof"></span>推計</div>
</div></header>
<main class="wrap">
  <section class="cards" id="cards"></section>
  <div class="panel">
    <h2>65歳以上人口と高齢化率の推移（1950年〜2050年）</h2>
    <div class="chart-box" id="box-trend"><canvas id="trend"></canvas></div>
    <p class="note">2030年以降（破線）は国立社会保障・人口問題研究所「日本の将来推計人口（令和5年推計）」出生中位・死亡中位推計。</p>
  </div>
  <div class="panel">
    <h2>都道府県別 高齢化率（65歳以上人口の割合）</h2>
    <div class="chart-box" id="box-pref"><canvas id="pref"></canvas></div>
    <p class="note">総務省統計局「人口推計」<span id="pref-asof"></span>（確報）。全国は 29.3%。</p>
  </div>
  <div class="panel">
    <h2>高齢者人口の内訳（男女別）</h2>
    <div class="chart-box" id="box-age"><canvas id="age"></canvas></div>
    <p class="note">「○歳以上」は累計の階級。単位は万人。</p>
  </div>
  <div class="panel">
    <h2>国際比較：65歳以上人口の割合（人口4000万以上の39か国中 上位10か国）</h2>
    <div class="chart-box" id="box-world"><canvas id="world"></canvas></div>
    <p class="note">日本は人口推計の2026年9月15日現在。他国は国連 World Population Prospects 2024 に基づく2026年7月1日現在の推計値。</p>
  </div>
</main>
<footer><div class="wrap">
  <p>出典: <a href="https://www.stat.go.jp/data/topics/topi1490.html">総務省統計局 統計トピックス No.149「統計からみた我が国の高齢者」（令和8年9月20日公表）</a> ／ <a href="https://www.stat.go.jp/data/jinsui/2024np/index.html">人口推計（2024年10月1日現在）</a></p>
  <p>本ページは非公式の可視化です。都道府県別データのみ基準日が異なります（<span id="pref-asof2"></span>）。</p>
</div></footer>
<script>
const DATA = __DATA__;

document.getElementById("asof").textContent = DATA.asof;
document.getElementById("pref-asof").textContent = DATA.prefAsof;
document.getElementById("pref-asof2").textContent = DATA.prefAsof;

const S = DATA.summary;
const cards = [
  { label: "65歳以上人口", value: S.pop65.toLocaleString() + "万人", sub: "（" + DATA.asof + "）" },
  { label: "高齢化率（総人口に占める割合）", value: S.rate65 + "%", sub: "過去最高" },
  { label: "65歳以上人口の前年差", value: (S.pop65Diff >= 0 ? "+" : "") + S.pop65Diff + "万人", sub: "高齢化率は " + (S.rate65Diff >= 0 ? "+" : "") + S.rate65Diff.toFixed(1) + "pt" },
  { label: "75歳以上人口", value: S.pop75.toLocaleString() + "万人", sub: "総人口の " + S.rate75 + "%" },
];
document.getElementById("cards").innerHTML = cards.map(c =>
  `<div class="card"><div class="label">${c.label}</div><div class="value">${c.value}</div><div class="sub2">${c.sub}</div></div>`
).join("");

const years = DATA.trend.map(r => r.year);
const splitIdx = years.indexOf(2030);
const dashed = ctx => ctx.p1DataIndex >= splitIdx ? [5, 4] : undefined;
const pop = key => DATA.trend.map(r => r[key]);
new Chart(document.getElementById("trend"), {
  type: "line",
  data: { labels: years, datasets: [
    { label: "65歳以上人口", data: pop("pop65"), borderColor: "#c2410c", backgroundColor: "#c2410c", yAxisID: "y", pointRadius: 2, tension: .2, segment: { borderDash: dashed } },
    { label: "うち70歳以上", data: pop("pop70"), borderColor: "#ea8b55", backgroundColor: "#ea8b55", yAxisID: "y", pointRadius: 2, tension: .2, segment: { borderDash: dashed } },
    { label: "うち75歳以上", data: pop("pop75"), borderColor: "#f2c194", backgroundColor: "#f2c194", yAxisID: "y", pointRadius: 2, tension: .2, segment: { borderDash: dashed } },
    { label: "高齢化率（右軸）", data: DATA.trend.map(r => r.rate65), borderColor: "#2563eb", backgroundColor: "#2563eb", yAxisID: "y1", borderWidth: 2, pointRadius: 0, tension: .2, segment: { borderDash: dashed } },
  ]},
  options: {
    responsive: true, maintainAspectRatio: false,
    interaction: { mode: "index", intersect: false },
    plugins: { legend: { position: "bottom" } },
    scales: {
      y: { title: { display: true, text: "人口（万人）" }, suggestedMin: 0, suggestedMax: 4000 },
      y1: { position: "right", title: { display: true, text: "割合（%）" }, min: 0, max: 40, grid: { drawOnChartArea: false } },
      x: { ticks: { autoSkip: true, maxTicksLimit: 16 } },
    },
  },
});

const prefs = [...DATA.prefectures].sort((a, b) => b.rate65 - a.rate65);
new Chart(document.getElementById("pref"), {
  type: "bar",
  data: { labels: prefs.map(p => p.name), datasets: [{
    data: prefs.map(p => p.rate65),
    backgroundColor: prefs.map((p, i) => i === 0 ? "#c2410c" : "#2563eb"),
  }]},
  options: {
    indexAxis: "y", responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.parsed.x}%（65歳以上 ${(c.raw && prefs[c.dataIndex].pop65 / 10).toLocaleString()}万人）` } } },
    scales: { x: { title: { display: true, text: "%" }, suggestedMax: 40 } },
  },
});

const A = DATA.ageDetail;
new Chart(document.getElementById("age"), {
  type: "bar",
  data: { labels: A.labels, datasets: [
    { label: "男", data: A.male, backgroundColor: "#2563eb" },
    { label: "女", data: A.female, backgroundColor: "#f59e0b" },
  ]},
  options: {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { position: "bottom" } },
    scales: { y: { title: { display: true, text: "人口（万人）" } } },
  },
});

const W = DATA.countries;
new Chart(document.getElementById("world"), {
  type: "bar",
  data: { labels: W.map(c => c.name), datasets: [{
    data: W.map(c => c.rate65),
    backgroundColor: W.map(c => c.name === "日本" ? "#c2410c" : "#94a3b8"),
  }]},
  options: {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.parsed.y}%` } } },
    scales: { y: { max: 32, title: { display: true, text: "%" } } },
  },
});
</script>
</body>
</html>
```

- [ ] **Step 3: ビルドを実行**

Run: `cd /home/katzkawai/kklab-kimi-older-people && uv run scripts/build.py`
Expected: アサーション全通過のうえ `OK: /home/katzkawai/kklab-kimi-older-people/index.html (... bytes)`

- [ ] **Step 4: 生成物を検証**

Run: `grep -c "3624" index.html && grep -o '"rate65": 29.6' index.html | head -1`
Expected: 1以上のカウントと `"rate65": 29.6` の出力

- [ ] **Step 5: コミット**

```bash
git add scripts/build.py scripts/template.html index.html data/raw/05k2024-3.xlsx data/raw/topi149_01.txt && git commit -m "feat: ダッシュボード index.html の生成"
```

---

### Task 6: ブラウザ表示の目視確認

**Files:**
- なし (検証のみ)

- [ ] **Step 1: ヘッドレス Chrome でスクリーンショット**

Run:
```bash
google-chrome --headless --disable-gpu --no-sandbox --window-size=1280,2600 --virtual-time-budget=10000 --screenshot=/tmp/dashboard.png file:///home/katzkawai/kklab-kimi-older-people/index.html
```
Expected: `/tmp/dashboard.png` が生成される (Chart.js は CDN から取得するためネットワーク必須)

- [ ] **Step 2: スクリーンショットを確認**

ReadMediaFile で `/tmp/dashboard.png` を開き、以下を目視確認:
- 要約カード4枚: 3,624万人 / 29.6% / +2万人 / 2,168万人
- 推移グラフ: 1950→2050、2030以降が破線
- 都道府県ランキング: 秋田県が最上位 (オレンジ)
- 内訳グラフ: 女 > 男
- 国際比較: 日本が最上位 (オレンジ)

表示崩れがあれば template.html を修正して `uv run scripts/build.py` を再実行し、再確認する。

- [ ] **Step 3: コミット (修正があった場合のみ)**

```bash
git add scripts/template.html index.html && git commit -m "fix: 表示調整"
```

---

### Task 7: README.md

**Files:**
- Create: `README.md`

- [ ] **Step 1: README.md を作成**

````markdown
# 日本の高齢者人口ダッシュボード

総務省統計局の公開データから、日本の高齢者 (65歳以上) 人口を可視化したダッシュボードです。

**公開ページ:** https://katzkawai.github.io/kklab-kimi-older-people/

## データ出典

- [統計トピックス No.149「統計からみた我が国の高齢者」](https://www.stat.go.jp/data/topics/topi1490.html) (総務省統計局、令和8年9月20日公表、2026年9月15日現在推計) — 全国の高齢者人口・推移・国際比較
- [人口推計 (2024年10月1日現在)](https://www.stat.go.jp/data/jinsui/2024np/index.html) (総務省統計局) — 都道府県別の年齢3区分人口

本ページは非公式の可視化です。都道府県別データのみ基準日 (2024年10月1日現在) が異なります。

## 内容

- 要約カード: 65歳以上人口 3,624万人、高齢化率 29.6% など
- 長期推移 (1950年〜2050年、将来推計含む)
- 都道府県別 高齢化率ランキング
- 高齢者人口の内訳 (男女別)
- 国際比較 (人口4000万以上の39か国中 上位10か国)

## 再生成手順

```bash
python3 scripts/fetch.py   # データを data/raw/ にダウンロード (stdlib のみ)
uv run scripts/build.py    # data/raw/ から index.html を生成 (依存は PEP723 で自動解決)
```

テスト:

```bash
uv run --with pytest --with openpyxl python -m pytest tests/ -v
```

## ファイル構成

- `index.html` — 生成された公開ページ (自己完結、Chart.js は CDN)
- `scripts/fetch.py` — データダウンロード
- `scripts/build.py` — データ抽出・検証・HTML生成
- `scripts/template.html` — ページテンプレート
- `data/raw/` — 元データ (PDF/Excel と pdftotext 出力)
- `tests/test_build.py` — パーサのテスト (公式既知値との照合)
````

- [ ] **Step 2: コミット**

```bash
git add README.md && git commit -m "docs: README"
```

---

### Task 8: GitHub リポジトリ作成と Pages 公開

**Files:**
- なし (公開作業)

- [ ] **Step 1: リポジトリ作成 & push**

Run: `cd /home/katzkawai/kklab-kimi-older-people && gh repo create kklab-kimi-older-people --public --source . --remote origin --push`
Expected: `✓ Created repository katzkawai/kklab-kimi-older-people on github.com` → push 完了

- [ ] **Step 2: GitHub Pages を有効化 (main / root)**

Run: `gh api repos/katzkawai/kklab-kimi-older-people/pages -X POST -F "source[branch]=main" -F "source[path]=/" `
Expected: JSON レスポンスに `"status": "built"` または `"building"`

- [ ] **Step 3: 公開を確認**

Run: `sleep 60; curl -s -o /dev/null -w "%{http_code}\n" https://katzkawai.github.io/kklab-kimi-older-people/`
Expected: `200` (404 の場合は60秒待って再試行。Pages の初回ビルドに数分かかることがある)

- [ ] **Step 4: 公開ページのスクリーンショット確認**

Run: `google-chrome --headless --disable-gpu --no-sandbox --window-size=1280,2600 --virtual-time-budget=10000 --screenshot=/tmp/dashboard-live.png https://katzkawai.github.io/kklab-kimi-older-people/`
ReadMediaFile で `/tmp/dashboard-live.png` を確認 (Task 6 と同じ内容が表示されていること)

- [ ] **Step 5: ブレインストーム用サーバーの停止 (後片付け)**

Run: `/home/katzkawai/.kimi-code/plugins/managed/superpowers/skills/brainstorming/scripts/stop-server.sh /home/katzkawai/kklab-kimi-older-people/.superpowers/brainstorm/32593-1789940326`
Expected: サーバー停止の出力 (`.superpowers/` は .gitignore 済みでリポジトリには含まれない)
