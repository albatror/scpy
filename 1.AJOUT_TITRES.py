import openpyxl
import os

def get_excel_filename(pattern):
    """Lire le nom du fichier Excel correspondant au pattern depuis excel_filenames.txt"""
    if os.path.exists('excel_filenames.txt'):
        with open('excel_filenames.txt', 'r') as f:
            filenames = [line.strip() for line in f.readlines()]
        
        # Chercher le fichier correspondant au pattern
        for filename in filenames:
            if pattern.lower() in filename.lower():
                return filename
    
    return None

# GTA HS paie - Récupération automatique du nom de fichier
gta_file_name = get_excel_filename('gta hs paie')
if not gta_file_name:
    print("Erreur : Impossible de trouver le fichier 'GTA HS PAIE'")
    exit(1)

print(f"Chargement du fichier Excel : {gta_file_name}")
wb = openpyxl.load_workbook(gta_file_name)
sheet = wb.active

# Ajout des noms de colonne pour les colonnes specifiques
sheet['K1'] = 'Categories'
sheet['L1'] = 'Dates'
sheet['O1'] = 'Etat'

# GTA HS paie - Sauvegarde des modifications dans le même fichier
wb.save(gta_file_name)
print(f"Fichier {gta_file_name} mis à jour avec succès.")