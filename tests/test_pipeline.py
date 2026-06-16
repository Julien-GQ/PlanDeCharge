from datetime import date
from pathlib import Path

from mise_en_forme_charge.pipeline import (
    _is_row_kept,
    _ordered_columns,
    _parse_profile,
    _transform_row,
)


def test_transform_adds_iso_week_and_type() -> None:
    row = {
        "DT_LIV_CONFIRMEE": "16/02/2026",
        "DESIGNATION": "B-TOILE MANCHON PL.1650X1700MM",
    }

    rules = {
        "SEMAINE": {"source": "DT_LIV_CONFIRMEE", "transform": "iso_week"},
        "TYPE": {"source": "DESIGNATION", "transform": "first_word"},
    }

    transformed = _transform_row(row, rules)

    assert transformed["SEMAINE"] == date(2026, 2, 16).isocalendar().week
    assert transformed["TYPE"] == "B-TOILE"


def test_filter_keeps_only_annoeull_with_mor_of_and_not_soldfab_o() -> None:
    valid = {
        "MOR_ATELIER": "ANNOEULL",
        "MOR_OF": "U4977/01",
        "OF_SOLDFAB": "N",
    }
    rules = [
        {"column": "MOR_ATELIER", "operator": "equals", "value": "ANNOEULL"},
        {"column": "MOR_OF", "operator": "not_empty"},
        {
            "column": "OF_SOLDFAB",
            "operator": "not_equals",
            "value": "O",
            "allow_empty": True,
        },
    ]

    assert _is_row_kept(valid, rules)

    wrong_site = dict(valid, MOR_ATELIER="PROVIN")
    assert not _is_row_kept(wrong_site, rules)

    missing_mor_of = dict(valid, MOR_OF="")
    assert not _is_row_kept(missing_mor_of, rules)

    soldfab_o = dict(valid, OF_SOLDFAB="O")
    assert not _is_row_kept(soldfab_o, rules)


def test_ordered_columns_starts_with_requested_sequence() -> None:
    headers = [
        "DATECDE",
        "DT_LIV_CONFIRMEE",
        "COMMANDE",
        "NOMREDCLI_CDE",
        "DESIGNATION",
        "RESTE_A_LIV_UV",
        "MOR_OF",
        "OF_SOLDFAB",
        "ARTICLE",
    ]
    rows = [
        {
            "SEMAINE": 7,
            "DT_LIV_CONFIRMEE": "16/02/2026",
            "COMMANDE": "U4977",
            "MOR_OF": "U4977/01",
            "NOMREDCLI_CDE": "CHOQUENET",
            "RESTE_A_LIV_UV": 138,
            "ARTICLE": "HCH1792",
            "DESIGNATION": "B-TOILE MANCHON",
            "TYPE": "B-TOILE",
            "OF_SOLDFAB": "N",
        }
    ]

    cols = _ordered_columns(
        headers,
        rows,
        preferred_order=[
            "SEMAINE",
            "DT_LIV_CONFIRMEE",
            "MOR_OF",
            "NOMREDCLI_CDE",
            "RESTE_A_LIV_UV",
            "ARTICLE",
            "DESIGNATION",
            "TYPE",
            "CPAYS_CDE",
        ],
        drop_columns={"COMMANDE", "OF_SOLDFAB"},
    )

    assert cols[:8] == [
        "SEMAINE",
        "DT_LIV_CONFIRMEE",
        "MOR_OF",
        "NOMREDCLI_CDE",
        "RESTE_A_LIV_UV",
        "ARTICLE",
        "DESIGNATION",
        "TYPE",
    ]


def test_parse_profile_1_json() -> None:
    profile = _parse_profile(Path("parametre/profile_1.json"))
    assert profile.name == "parametre_1"
    assert "DT_LIV_CONFIRMEE" in profile.expected_columns
    assert profile.output_sheet == "Charge_Global"
