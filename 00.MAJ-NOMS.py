import tkinter as tk
from tkinter import filedialog, simpledialog
import os
import re

# Liste des fichiers de script
script_files = [
    "1.AJOUT_TITRES.py", 
    "2.SORTIE.py", 
    "4.HS.py", 
    "5.ASTREINTES.py", 
    "6.PERMANENCES.py", 
    "7.INTERVENTIONS.py"
]

# Liste des motifs de fichiers Excel
excel_file_patterns = [
    "ETAT DES HEURES SUPPLEMENTAIRES SIEGE *.xlsx",
    "GTA HS PAIE *.xlsx",
    "INDEMNITE ASTREINTE SIEGE *.xlsx",
    "INDEMNITE PERMANENCE SIEGE *.xlsx",
    "INTERVENTION ASTREINTE SIEGE *.xlsx"
]

def get_month():
    root = tk.Tk()
    root.withdraw()
    month = simpledialog.askstring("Sélection du Mois", "Entrez le mois (ex: JUIN 2025):")
    return month

def get_excel_files():
    root = tk.Tk()
    root.withdraw()
    excel_files = []
    for pattern in excel_file_patterns:
        file_path = filedialog.askopenfilename(
            title=f"Sélectionnez le fichier {pattern}",
            filetypes=[("Fichiers Excel", "*.xlsx")],
            initialdir=os.getcwd()
        )
        if file_path:
            excel_files.append((pattern, file_path))
    return excel_files

def update_script_files(excel_files, month):
    for script_file in script_files:
        try:
            with open(script_file, "r") as f:
                content = f.read()
        except FileNotFoundError:
            print(f"Erreur : Le fichier {script_file} est introuvable.")
            continue
        except Exception as e:
            print(f"Erreur lors de l'ouverture de {script_file}: {e}")
            continue
        
        original_content = content
        for pattern, file_path in excel_files:
            file_name = os.path.basename(file_path)
            content = re.sub(r"['\"].*" + pattern.replace("*", ".*") + r"['\"]", f"'{file_name}'", content)
        
        # Mettre à jour le mois dans les scripts
        content = re.sub(r"mois = '.*'", f"mois = '{month}'", content)

        if content != original_content:
            print(f"Modifications dans {script_file}:")
            print(f"Avant: {original_content}")
            print(f"Après: {content}")
        
            try:
                with open(script_file, "w") as f:
                    f.write(content)
            except Exception as e:
                print(f"Erreur lors de l'écriture dans {script_file}: {e}")

    print("Les fichiers de script ont été mis à jour avec succès.")

def main():
    month = get_month()
    if month:
        excel_files = get_excel_files()
        if excel_files:
            update_script_files(excel_files, month)
            print("Les fichiers de script ont été mis à jour avec succès.")
        else:
            print("Aucun fichier n'a été sélectionné. Les scripts n'ont pas été modifiés.")
    else:
        print("Aucun mois n'a été sélectionné. Les scripts n'ont pas été modifiés.")

if __name__ == "__main__":
    main()
