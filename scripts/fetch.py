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
