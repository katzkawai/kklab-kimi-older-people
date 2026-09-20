# 日本の高齢者人口ダッシュボード

総務省統計局の公開データから、日本の高齢者 (65歳以上) 人口を可視化したダッシュボードです。

**公開ページ:** https://katzkawai.org/kklab-kimi-older-people/

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
