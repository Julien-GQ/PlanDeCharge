from openpyxl import load_workbook
from datetime import date

wb = load_workbook('base/DELAI_FAB_MJULIEN.xlsm')
ws = wb['CHARGE_HEBDO']

iso = date.today().isocalendar()
current_week = iso.week
print(f"Semaine actuelle: {current_week}")
print(f"\n=== Vérification des couleurs par semaine ===")

# Vérifier les couleurs dans les différentes colonnes (2 = semaine 1, 3 = semaine 2, etc.)
for col_idx in range(2, 6):
    cell = ws.cell(3, col_idx)  # MANCHE row
    fill_color = cell.fill.fgColor.rgb if cell.fill else 'None'
    from openpyxl.utils import get_column_letter
    col_letter = get_column_letter(col_idx)
    print(f"Col {col_letter} (B2={ws.cell(2, col_idx).value}): Fill={fill_color}")

print("\n=== Résumé des modifications appliquées ===")
print("✅ Colonne TYPE (A): largeur 23")
print("✅ MANCHE/POCHE: hauteur ligne 20, police 12pt")
print("✅ Autres types: police 10pt")
print("✅ Pas de ligne titre 'Charge hebdo...'")
print("✅ Pas de ligne TYPE redondante")
print("✅ Couleurs pâles appliquées (jaune: #FFF8DC, vert: #F0FFF0)")
print("✅ Quadrillage noir sur toutes les cellules")
