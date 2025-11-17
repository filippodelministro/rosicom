from openpyxl import load_workbook

# Carica il file Excel
file_path = "source/ALL_ROSTERS.xlsx"
wb = load_workbook(filename=file_path, data_only=True)  # data_only=True legge i valori, non le formule

# Seleziona il primo foglio (puoi usare anche wb['NomeFoglio'])
ws = wb.active

# Legge i valori delle celle A4 e G4
squadra_1 = ws['A5'].value
squadra_2 = ws['F5'].value

print("Squadra 1:", squadra_1)
print("Squadra 2:", squadra_2)