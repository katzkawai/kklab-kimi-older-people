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
