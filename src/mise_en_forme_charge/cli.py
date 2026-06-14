"""Point d'entree CLI du projet."""

from .formatter import format_charge


def main() -> None:
    """Affiche un exemple simple de mise en forme."""
    print(format_charge(12.3456))


if __name__ == "__main__":
    main()
