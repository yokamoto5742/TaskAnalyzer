from datetime import date

import pytest

from utils.date_period import PERIOD_NAMES, get_period_range, parse_date


@pytest.mark.parametrize("period_name, expected", [
    ('2月', (date(2024, 2, 1), date(2024, 2, 29))),
    ('12月', (date(2024, 12, 1), date(2024, 12, 31))),
    ('第1四半期', (date(2024, 1, 1), date(2024, 3, 31))),
    ('第4四半期', (date(2024, 10, 1), date(2024, 12, 31))),
    ('上半期', (date(2024, 1, 1), date(2024, 6, 30))),
    ('下半期', (date(2024, 7, 1), date(2024, 12, 31))),
    ('一年', (date(2024, 1, 1), date(2024, 12, 31))),
])
def test_get_period_range_january_start(period_name, expected):
    assert get_period_range(2024, period_name, 1) == expected


@pytest.mark.parametrize("period_name, expected", [
    ('4月', (date(2025, 4, 1), date(2025, 4, 30))),
    ('3月', (date(2026, 3, 1), date(2026, 3, 31))),
    ('第1四半期', (date(2025, 4, 1), date(2025, 6, 30))),
    ('第4四半期', (date(2026, 1, 1), date(2026, 3, 31))),
    ('上半期', (date(2025, 4, 1), date(2025, 9, 30))),
    ('下半期', (date(2025, 10, 1), date(2026, 3, 31))),
    ('一年', (date(2025, 4, 1), date(2026, 3, 31))),
])
def test_get_period_range_april_start(period_name, expected):
    assert get_period_range(2025, period_name, 4) == expected


def test_period_names_are_all_supported():
    for period_name in PERIOD_NAMES:
        get_period_range(2025, period_name, 4)


def test_parse_date():
    assert parse_date(' 2025/1/5 ') == date(2025, 1, 5)


@pytest.mark.parametrize("text", ['2025-01-05', '20250105', '2025/02/30', ''])
def test_parse_date_invalid(text):
    with pytest.raises(ValueError):
        parse_date(text)
