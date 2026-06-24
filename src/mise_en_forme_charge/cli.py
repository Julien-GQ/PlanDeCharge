"""Point d'entree CLI du projet."""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from tkinter import BOTH, Button, Tk, TclError, Text, filedialog

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


def _format_filters_for_display(profile_path: Path) -> str:
    try:
        with profile_path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except Exception as exc:
        return f"Impossible de lire le profil: {exc}"

    filters = payload.get("filters", [])
    if not isinstance(filters, list) or not filters:
        return "Aucun filtre defini dans le profil."

    op_labels = {
        "equals": "egal a",
        "not_equals": "different de",
        "not_empty": "non vide",
    }

    lines = [
        "Filtres appliques sur le tableau de base pour creer la page GLOBAL:",
        "",
    ]
    for idx, rule in enumerate(filters, start=1):
        col = str(rule.get("column", "")).strip() or "<colonne inconnue>"
        op = str(rule.get("operator", "")).strip().lower()
        value = str(rule.get("value", "")).strip()
        op_label = op_labels.get(op, op)

        if op == "not_empty":
            detail = f"{idx}. {col}: {op_label}"
        elif value != "":
            detail = f"{idx}. {col}: {op_label} {value}"
        else:
            detail = f"{idx}. {col}: {op_label}"

        if bool(rule.get("allow_empty", False)):
            detail += " (valeur vide autorisee)"
        lines.append(detail)

    return "\n".join(lines)


def _show_global_filters_window(profile_path: Path) -> None:
    message = _format_filters_for_display(profile_path)

    root = Tk()
    root.title("Filtres page GLOBAL")
    root.geometry("760x380")
    root.resizable(True, True)

    text = Text(root, wrap="word", font=("Segoe UI", 10))
    text.insert("1.0", message)
    text.configure(state="disabled")
    text.pack(fill=BOTH, expand=True, padx=12, pady=(12, 8))

    close_button = Button(root, text="Continuer", command=root.destroy)
    close_button.pack(pady=(0, 12))

    # Met la fenetre au premier plan, puis retire le mode topmost.
    def _focus_window() -> None:
        try:
            root.deiconify()
            root.lift()
            root.attributes("-topmost", True)
            root.focus_force()
            root.after(300, lambda: root.attributes("-topmost", False))
        except TclError:
            return

    root.after(50, _focus_window)

    # Evite un blocage infini si la GUI n'est pas affichable sur la session.
    root.after(30000, root.destroy)
    root.mainloop()


def _open_generated_workbook(path: Path) -> None:
    if not path.exists():
        return
    try:
        os.startfile(path)  # type: ignore[attr-defined]
    except Exception:
        print(f"Classeur genere (ouverture auto impossible): {path}")


def _export_pdf(workbook_path: Path, pdf_output: Path | None = None) -> Path:
    try:
        import win32com.client as win32  # type: ignore
    except ImportError as exc:
        raise ValueError(
            "Export PDF indisponible: installer pywin32 (pip install pywin32)."
        ) from exc

    pdf_path = pdf_output or workbook_path.with_suffix(".pdf")

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(workbook_path.resolve()))
        wb.Save()
        wb.ExportAsFixedFormat(0, str(pdf_path.resolve()))
    except Exception as exc:
        raise ValueError(
            "Export PDF impossible. Verifier que le classeur n'est pas ouvert dans Excel, puis relancer."
        ) from exc
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()

    return pdf_path


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
        default="Global",
        help="Nom de la feuille de sortie",
    )
    parser.add_argument(
        "--no-picker",
        action="store_true",
        help="Desactive l'ouverture de l'explorateur si --input n'est pas fourni",
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="Exporte un PDF multi-pages des feuilles principales apres generation",
    )
    parser.add_argument(
        "--pdf-output",
        type=Path,
        default=None,
        help="Chemin du PDF de sortie (optionnel, utilise avec --pdf)",
    )
    parser.add_argument(
        "--show-filters-window",
        action="store_true",
        help="Affiche une fenetre Tkinter listant les filtres utilises pour GLOBAL",
    )
    parser.add_argument(
        "--no-filters-window",
        action="store_true",
        help="Desactive l'affichage automatique de la fenetre des filtres",
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

    should_show_filters = args.show_filters_window or not args.no_filters_window
    if should_show_filters:
        try:
            print("Affichage des filtres GLOBAL: fermez la fenetre pour continuer...")
            _show_global_filters_window(args.profile)
        except Exception:
            # En environnement sans affichage, on garde un fallback console.
            print(_format_filters_for_display(args.profile))

    result = build_v1_sheet(input_path, args.output, args.profile, args.sheet)
    print(f"Profil utilise: {args.profile}")
    print(f"Source lue: {input_path}")
    print(f"Lignes source: {result.rows_total}")
    print(f"Lignes retenues: {result.rows_kept}")
    print(f"Fichier genere: {result.output_path}")
    print(f"Feuille generee: {result.output_sheet}")

    _open_generated_workbook(result.output_path)

    if args.pdf:
        pdf_path = _export_pdf(result.output_path, args.pdf_output)
        print(f"PDF genere: {pdf_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Operation interrompue par l'utilisateur.")
        sys.exit(1)
    except ValueError as exc:
        print(f"Erreur: {exc}")
        sys.exit(1)
