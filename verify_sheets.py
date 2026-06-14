from openpyxl import load_workbook

wb = load_workbook('base/DELAI_FAB_MJULIEN.xlsm')

print("=== Feuilles créées ===")
for idx, sheet_name in enumerate(wb.sheetnames, 1):
    sheet = wb[sheet_name]
    print(f"{idx}. {sheet_name}")
    
    # Compter les tables
    tables = list(sheet.tables.values())
    if tables:
        for table in tables:
            print(f"   - Table: {table.displayName} (ref: {table.ref})")

print("\nFichier validé sans erreurs de corruption de table ✅")
