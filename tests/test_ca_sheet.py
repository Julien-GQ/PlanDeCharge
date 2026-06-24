from mise_en_forme_charge.ca_sheet import (
    _aggregate_by_period,
    _month_key_from_row,
    _week_key_from_row,
)


def test_weekly_aggregation_uses_caprev_lcde() -> None:
    rows = [
        {
            "DT_LIV_CONFIRMEE": "05/01/2026",
            "TYPE": "MANCHE LONGUE",
            "CAPREV_LCDE": "10",
        },
        {
            "DT_LIV_CONFIRMEE": "06/01/2026",
            "TYPE": "MANCHE LONGUE",
            "CAPREV_LCDE": "2,5",
        },
        {
            "DT_LIV_CONFIRMEE": "08/01/2026",
            "TYPE": "POCHE INTERIEURE",
            "CAPREV_LCDE": 3,
        },
    ]

    week = _week_key_from_row(rows[0])
    assert week is not None

    ordered_types, matrix, totals = _aggregate_by_period(rows, [week], _week_key_from_row)

    assert ordered_types == ["MANCHE LONGUE", "POCHE INTERIEURE"]
    assert matrix["MANCHE LONGUE"][week] == 12.5
    assert matrix["POCHE INTERIEURE"][week] == 3.0
    assert totals[week] == 15.5


def test_monthly_aggregation_groups_all_rows_in_same_month() -> None:
    rows = [
        {"DT_LIV_CONFIRMEE": "15/02/2026", "TYPE": "AUTRE", "CAPREV_LCDE": 1},
        {"DT_LIV_CONFIRMEE": "28/02/2026", "TYPE": "AUTRE", "CAPREV_LCDE": 4},
    ]

    month_key = _month_key_from_row(rows[0])
    assert month_key is not None

    ordered_types, matrix, totals = _aggregate_by_period(rows, [month_key], _month_key_from_row)

    assert ordered_types == ["AUTRE"]
    assert matrix["AUTRE"][month_key] == 5.0
    assert totals[month_key] == 5.0
