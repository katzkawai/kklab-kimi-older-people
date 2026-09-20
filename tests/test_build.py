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
