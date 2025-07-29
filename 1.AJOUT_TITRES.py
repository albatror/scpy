import openpyxl

# GTA HS paie mai_synthese Chargement du fichier Excel - Fichier à modifier selon le mois
wb = openpyxl.load_workbook('GTA HS PAIE JUILLET 2025.xlsx')
sheet = wb.active

# Ajout des noms de colonne pour les colonnes specifiques
sheet['K1'] = 'Categories'
sheet['L1'] = 'Dates'
sheet['O1'] = 'Etat'

# GTA HS paie mai_synthese Sauvegarde des modifications dans un nouveau fichier - DT
wb.save('GTA HS PAIE JUILLET 2025.xlsx')
