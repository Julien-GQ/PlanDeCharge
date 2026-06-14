from mise_en_forme_charge.formatter import format_charge


def test_format_charge_default() -> None:
    assert format_charge(12.3456) == "12.35 kN"


def test_format_charge_custom() -> None:
    assert format_charge(3, unit="N", decimals=0) == "3 N"
