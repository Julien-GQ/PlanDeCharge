# Mise en forme charge

Base de projet Python avec structure `src/`, tests et outils de qualite.

## Demarrage rapide

1. Creer/activer l'environnement virtuel:
   - PowerShell: `python -m venv .venv` puis `./.venv/Scripts/Activate.ps1`
2. Installer les dependances de dev:
   - `pip install -e .[dev]`
3. Lancer les tests:
   - `pytest -q`
4. Verifier le style:
   - `ruff check .` puis `black --check .`
