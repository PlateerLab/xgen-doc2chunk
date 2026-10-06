import datetime as dt

import pytest

from xgen_doc2chunk.core.processor.excel_helper.cell_format import format_cell_value, is_date_format


@pytest.mark.parametrize(
    "value, fmt, expected",
    [
        (0.153, "0.0%", "15.3%"),
        (0.5, "0%", "50%"),
        (1234567, "#,##0", "1,234,567"),
        (1234567, '#,##0"원"', "1,234,567원"),
        (-1500, "#,##0;(#,##0)", "(1,500)"),
        (1500, "#,##0;(#,##0)", "1,500"),
        (1234.5, "[$₩-412]#,##0.00", "₩1,234.50"),
        (12345678, '#,##0,"천원"', "12,346천원"),
        (42, "00000", "00042"),
        (2.5, "0.00", "2.50"),
        (0.125, "0.00", "0.13"),
        (1234.5, "General", "1234.5"),
        (3.0, None, "3"),
        (12345.678, "0.00E+00", "1.23E+04"),
        (True, None, "TRUE"),
        (dt.datetime(2026, 10, 6), "yyyy-mm-dd", "2026-10-06"),
        (dt.datetime(2026, 10, 6), 'yyyy"년" m"월" d"일"', "2026년 10월 6일"),
        (dt.datetime(2026, 10, 6), "yyyy-mm-dd (aaa)", "2026-10-06 (화)"),
        (dt.datetime(2026, 10, 6, 14, 5), "yyyy-mm-dd hh:mm", "2026-10-06 14:05"),
        (dt.time(14, 5), "h:mm AM/PM", "2:05 PM"),
        (dt.time(0, 30), "h:mm AM/PM", "12:30 AM"),
        (45000, "yyyy-mm-dd", "2023-03-15"),
    ],
)
def test_display_format(value, fmt, expected):
    assert format_cell_value(value, fmt) == expected


def test_date_detection():
    assert is_date_format("yyyy-mm-dd")
    assert is_date_format("h:mm AM/PM")
    assert not is_date_format("#,##0")
    assert not is_date_format('#,##0"일"')
