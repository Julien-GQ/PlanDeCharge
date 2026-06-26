"""Utilitaires de lecture et presentation des filtres de profil."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def load_profile_filters(profile_path: Path) -> list[dict[str, Any]]:
    """Charge la liste des filtres depuis un profil JSON."""
    try:
        with profile_path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except Exception as exc:  # pragma: no cover - depend du systeme de fichiers
        raise ValueError(f"Impossible de lire le profil: {exc}") from exc

    filters = payload.get("filters", [])
    if not isinstance(filters, list) or not filters:
        return []
    return [rule for rule in filters if isinstance(rule, dict)]


def default_atelier_from_filters(filters: list[dict[str, Any]]) -> str:
    """Retourne l'atelier par defaut defini dans les filtres, s'il existe."""
    for rule in filters:
        if str(rule.get("column", "")).strip().upper() != "MOR_ATELIER":
            continue
        if str(rule.get("operator", "")).strip().lower() != "equals":
            continue
        return str(rule.get("value", "")).strip()
    return ""


def preferred_atelier(ateliers: list[str], fallback: str) -> str:
    """Priorise Annoeullin/Annoeull, puis la valeur fallback, puis le premier atelier."""
    for atelier in ateliers:
        normalized = atelier.strip().upper()
        if normalized in {"ANNOEULLIN", "ANNOEULL"}:
            return atelier

    if fallback in ateliers:
        return fallback
    return ateliers[0] if ateliers else ""


def format_filters_for_display(
    profile_path: Path,
    selected_atelier: str | None = None,
) -> str:
    """Formate les filtres pour affichage utilisateur."""
    try:
        filters = load_profile_filters(profile_path)
    except ValueError as exc:
        return str(exc)

    if not filters:
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
        if col.upper() == "MOR_ATELIER" and selected_atelier:
            value = selected_atelier
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
