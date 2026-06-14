"""Point d'entree CLI du projet."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from tkinter import Tk, filedialog

try:
    from .pipeline import build_v1_sheet
except ImportError:
    # Fallback when launched as a plain script: python src/.../cli.py
    from pipeline import build_v1_sheet


def _pick_input_file() -> Path | None:
    root = Tk()
    root.withdraw()
    try:
        selected = filedialog.askopenfilename(
            title="Selectionner le fichier source",
            filetypes=[
                ("Fichiers Excel", "*.xlsm *.xlsx"),
                ("Fichiers texte", "*.txt"),
                ("Tous les fichiers", "*.*"),
            ],
        )
    except KeyboardInterrupt:
        selected = ""
    finally:
        root.destroy()

    if not selected:
        return None
    return Path(selected)


def main() -> None:
    """Lance la generation de la feuille V1."""
    parser = argparse.ArgumentParser(description="Generation de feuille via profil JSON")
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Fichier source (.xlsm, .xlsx ou .txt). Si absent, ouverture de l'explorateur.",
    )
    parser.add_argument(
        "--profile",
        type=Path,
        default=Path("parametre/profile_1.json"),
        help="Profil JSON de transformation",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Fichier de sortie (optionnel). Si absent sur Excel, la feuille est ajoutee dans le fichier source.",
    )
    parser.add_argument(
        "--sheet",
        default="V1_ANNOEULL",
        help="Nom de la feuille de sortie",
    )
    parser.add_argument(
        "--no-picker",
        action="store_true",
        help="Desactive l'ouverture de l'explorateur si --input n'est pas fourni",
    )
    args = parser.parse_args()

    input_path = args.input
    if input_path is None and not args.no_picker:
        input_path = _pick_input_file()
        if input_path is None:
            print("Operation annulee: aucun fichier selectionne.")
            return
    if input_path is None:
        raise ValueError("Fichier source requis: utiliser --input ou l'explorateur")

    result = build_v1_sheet(input_path, args.output, args.profile, args.sheet)
    print(f"Profil utilise: {args.profile}")
    print(f"Source lue: {input_path}")
    print(f"Lignes source: {result.rows_total}")
    print(f"Lignes retenues: {result.rows_kept}")
    print(f"Fichier genere: {result.output_path}")
    print(f"Feuille generee: {result.output_sheet}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Operation interrompue par l'utilisateur.")
        sys.exit(1)
    except ValueError as exc:
        print(f"Erreur: {exc}")
        sys.exit(1)
