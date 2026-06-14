"""Fonctions de formatage pour les valeurs de charge."""


def format_charge(value: float, unit: str = "kN", decimals: int = 2) -> str:
    """Retourne une representation standardisee d'une charge."""
    return f"{value:.{decimals}f} {unit}"
