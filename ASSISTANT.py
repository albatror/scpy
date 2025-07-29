import tkinter as tk
from tkinter import ttk, scrolledtext, simpledialog, messagebox, filedialog
import subprocess
import os
import sys
import importlib
import time
import glob
import re
import ast

scripts = [
    "000.INSTALL-MODULES.py",
    "1.AJOUT_TITRES.py",
    "2.SORTIE.py",
    "3.REORG-SORTIE.py",
    "4.HS.py",
    "5.ASTREINTES.py",
    "6.PERMANENCES.py",
    "7.INTERVENTIONS.py"
]

required_modules = ['openpyxl', 'pandas', 're']
EXCEL_FILENAMES_FILE = "excel_filenames.txt"

def parse_categories_from_sortie():
    """Parse categories directly from 2.SORTIE.py file"""
    try:
        with open("2.SORTIE.py", "r", encoding="utf-8") as f:
            content = f.read()
        
        # Using ast to safely parse category lists
        tree = ast.parse(content)
        categories = {}
        
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name):
                        if target.id.startswith("categories_"):
                            try:
                                value = ast.literal_eval(node.value)
                                categories[target.id] = value
                            except (ValueError, SyntaxError):
                                pass
        
        return categories
    except Exception as e:
        print(f"Error parsing 2.SORTIE.py: {e}")
        # Default categories if parsing fails
        return {
            'categories_heures_sup': ['HS_NORMALES_INF14', 'HS_NORMALES_SUP14', 'HS_NUIT_INF14', 'HS_NUIT_SUP14', 'HS_DIM_ET_JF_INF14', 'HS_DIM_ET_JF_SUP14'],
            'categories_astreintes': ['AST_SEMAINE', 'AST_SEM_CALENDAIRE', 'AST_SEM_CALEND_MAJO', 'AST_DJF', 'AST_DJF_MAJO', 'AST_WEEKEND', 'AST_NUIT'],
            'categories_permanences': ['PERM_DJF_MAJO', 'PERM_DJF', 'PERM_SAMEDI_MAJO', 'PERM_WEEKEND'],
            'categories_interventions': ['INT_SEMAINE', 'INT_NUIT', 'INT_SAMEDI', 'INT_DJF']
        }

# Parse categories from 2.SORTIE.py
parsed_categories = parse_categories_from_sortie()

class CategoryEditor(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Éditeur de Catégories")
        self.geometry("600x700")
        
        # Dict to store category lists
        self.category_lists = {
            'Heures Supplémentaires': parsed_categories.get('categories_heures_sup', []),
            'Astreintes': parsed_categories.get('categories_astreintes', []),
            'Permanences': parsed_categories.get('categories_permanences', []),
            'Interventions': parsed_categories.get('categories_interventions', [])
        }
        
        self.create_widgets()
    
    def create_widgets(self):
        # Notebook (tabbed interface)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs for each category type
        self.tabs = {}
        self.listboxes = {}
        for category_type, categories in self.category_lists.items():
            tab = ttk.Frame(self.notebook)
            self.notebook.add(tab, text=category_type)
            self.tabs[category_type] = tab
            
            # Listbox for categories
            listbox = tk.Listbox(tab, width=50)
            listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
            self.listboxes[category_type] = listbox
            
            # Populate listbox
            for category in categories:
                listbox.insert(tk.END, category)
            
            # Frame for buttons
            button_frame = tk.Frame(tab)
            button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
            
            # Add Category Button
            add_btn = tk.Button(button_frame, text="Ajouter Catégorie", 
                                command=lambda t=category_type: self.add_category(t))
            add_btn.pack(side=tk.LEFT, padx=5)
            
            # Remove Category Button
            remove_btn = tk.Button(button_frame, text="Supprimer Catégorie", 
                                   command=lambda t=category_type: self.remove_category(t))
            remove_btn.pack(side=tk.LEFT, padx=5)
            
            # Edit Category Button (NEW)
            edit_btn = tk.Button(button_frame, text="Éditer Catégorie", 
                                 command=lambda t=category_type: self.edit_category(t))
            edit_btn.pack(side=tk.LEFT, padx=5)
        
        # Save Button
        save_btn = tk.Button(self, text="Sauvegarder", command=self.save_categories)
        save_btn.pack(side=tk.BOTTOM, pady=10)
    
    def add_category(self, category_type):
        # Prompt for new category name
        new_category = simpledialog.askstring(
            "Ajouter Catégorie", 
            f"Entrez le nom de la nouvelle catégorie pour {category_type}:"
        )
        
        if new_category and new_category.strip():
            # Get the listbox for this tab
            listbox = self.listboxes[category_type]
            
            # Check if category already exists
            if new_category in listbox.get(0, tk.END):
                messagebox.showwarning("Doublon", "Cette catégorie existe déjà.")
                return
            
            # Add to listbox and category list
            listbox.insert(tk.END, new_category)
            self.category_lists[category_type].append(new_category)
    
    def remove_category(self, category_type):
        # Get the listbox for this tab
        listbox = self.listboxes[category_type]
        
        # Get selected category
        selected = listbox.curselection()
        
        if not selected:
            messagebox.showwarning("Sélection", "Veuillez sélectionner une catégorie à supprimer.")
            return
        
        # Remove from listbox and category list
        index = selected[0]
        category = listbox.get(index)
        listbox.delete(index)
        self.category_lists[category_type].remove(category)
    
    def edit_category(self, category_type):
        # Get the listbox for this tab
        listbox = self.listboxes[category_type]
        
        # Get selected category
        selected = listbox.curselection()
        
        if not selected:
            messagebox.showwarning("Sélection", "Veuillez sélectionner une catégorie à éditer.")
            return
        
        # Get the current category
        index = selected[0]
        current_category = listbox.get(index)
        
        # Prompt for new category name
        new_category = simpledialog.askstring(
            "Éditer Catégorie", 
            f"Modifier la catégorie {current_category} pour {category_type}:",
            initialvalue=current_category
        )
        
        if new_category and new_category.strip() and new_category != current_category:
            # Check if new category already exists
            if new_category in listbox.get(0, tk.END):
                messagebox.showwarning("Doublon", "Cette catégorie existe déjà.")
                return
            
            # Update listbox and category list
            listbox.delete(index)
            listbox.insert(index, new_category)
            
            # Update in the category list
            category_list = self.category_lists[category_type]
            category_list[category_list.index(current_category)] = new_category
    
    def save_categories(self):
        # Prepare the updated category definitions
        category_def = "# Definition des categories pour chaque type\n"
        mapping = {
            'Heures Supplémentaires': 'categories_heures_sup',
            'Astreintes': 'categories_astreintes',
            'Permanences': 'categories_permanences',
            'Interventions': 'categories_interventions'
        }
        
        # Write to 2.SORTIE.py
        try:
            with open("2.SORTIE.py", "r", encoding="utf-8") as f:
                content = f.readlines()
            
            # Remove existing category definitions
            content = [line for line in content if not any(key in line for key in mapping.values())]
            
            # Append new category definitions
            for category_type, list_name in mapping.items():
                content.append(f"{list_name} = {repr(self.category_lists[category_type])}\n")
            
            with open("2.SORTIE.py", "w", encoding="utf-8") as f:
                f.writelines(content)
            
            messagebox.showinfo("Succès", "Catégories mises à jour avec succès!")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de sauvegarder : {str(e)}")

class ScriptExecutor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Interface de Traitement des données")
        self.geometry("800x640")

        # Check and remove excel_filenames.txt at startup
        self._check_and_remove_excel_filenames()
        
        self._check_and_remove_sortie_file()
        self._check_and_rename_hs_files()
        self.create_widgets()

    def _check_and_remove_excel_filenames(self):
        """Check and remove excel_filenames.txt file if it exists"""
        if os.path.exists(EXCEL_FILENAMES_FILE):
            try:
                os.remove(EXCEL_FILENAMES_FILE)
                print(f"Le fichier {EXCEL_FILENAMES_FILE} a été supprimé.")
            except Exception as e:
                print(f"Erreur lors de la suppression de {EXCEL_FILENAMES_FILE} : {e}")
        else:
            print(f"Le fichier {EXCEL_FILENAMES_FILE} n'existe pas.")

    def _check_and_remove_sortie_file(self):
        """Check and remove SORTIE.xlsx file if it exists"""
        if os.path.exists("SORTIE.xlsx"):
            try:
                os.remove("SORTIE.xlsx")
                print("Le fichier SORTIE.xlsx a été supprimé.")
            except Exception as e:
                print(f"Erreur lors de la suppression de SORTIE.xlsx : {e}")
        else:
            print("Le fichier SORTIE.xlsx n'existe pas.")

    def _check_and_rename_hs_files(self):
        """Check and rename HS files"""
        # Regular expression to match two patterns:
        # 1. "Copie de HS paie [month] [year]_synthèse.xlsx"
        # 2. "HS paie [month] [year]_synthèse.xlsx"
        pattern = r"(?:Copie de )?HS paie ([a-zéû]+) (20\d\d)_synthèse\.xlsx"
        
        # Mapping of French month names to uppercase
        month_mapping = {
            "janvier": "JANVIER", "fevrier": "FEVRIER", "février": "FEVRIER", 
            "mars": "MARS", "avril": "AVRIL", "mai": "MAI", 
            "juin": "JUIN", "juillet": "JUILLET", "aout": "AOUT", 
            "août": "AOUT", "septembre": "SEPTEMBRE", "octobre": "OCTOBRE", 
            "novembre": "NOVEMBRE", "decembre": "DECEMBRE", "décembre": "DECEMBRE"
        }
        
        # Check all Excel files in the current directory
        for filename in glob.glob("*.xlsx"):
            match = re.match(pattern, filename, re.IGNORECASE)
            if match:
                month, year = match.groups()
                # Convert month to uppercase using mapping
                month_upper = month_mapping.get(month.lower(), month.upper())
                new_filename = f"GTA HS PAIE {month_upper} {year}.xlsx"
                
                try:
                    os.rename(filename, new_filename)
                    print(f"Fichier renommé: {filename} → {new_filename}")
                except Exception as e:
                    print(f"Erreur lors du renommage de {filename}: {e}")

    def create_widgets(self):
        # Titre pour la fenêtre des fichiers
        tk.Label(self, text="FICHIERS EN COURS DE TRAITEMENT", font=("Arial", 12, "bold")).pack(pady=5)

        # Fenêtre de visualisation des fichiers
        self.file_visualization = tk.Text(self, height=10, width=80)
        self.file_visualization.pack(pady=10)

        # Espace entre les fenêtres
        tk.Frame(self, height=20).pack()

        # Titre pour la fenêtre de journal
        tk.Label(self, text="JOURNAL", font=("Arial", 12, "bold")).pack(pady=5)
        self.console_output = scrolledtext.ScrolledText(self, height=15, width=80)
        self.console_output.pack(pady=10)

        # Barre de progression
        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        # Frame pour les boutons
        button_frame = tk.Frame(self)
        button_frame.pack(pady=5)

        # Bouton CHARGER FICHIERS EXCEL
        tk.Button(button_frame, text="CHARGER FICHIERS EXCEL", 
                  command=self.load_excel_files).pack(side=tk.LEFT, padx=10)

        # Bouton START
        tk.Button(button_frame, text="START", command=self.start_scripts).pack(side=tk.LEFT, padx=10)

        # Bouton EDITER LES CATEGORIES
        tk.Button(button_frame, text="EDITER LES CATEGORIES", 
                  command=self.open_category_editor).pack(side=tk.LEFT, padx=10)

        # Charger les noms de fichiers existants
        self.load_excel_filenames()

    def load_excel_files(self):
        """Permet à l'utilisateur de sélectionner manuellement les fichiers Excel"""
        excel_files = []
        
        # Types de fichiers attendus
        file_types = [
            ("ETAT DES HEURES SUPPLEMENTAIRES", "etat des heures supplementaires"),
            ("GTA HS PAIE", "gta hs paie"),
            ("INDEMNITE ASTREINTE", "indemnite astreinte"),
            ("INDEMNITE PERMANENCE", "indemnite permanence"),
            ("INTERVENTION ASTREINTE", "intervention astreinte")
        ]
        
        for display_name, pattern in file_types:
            file_path = filedialog.askopenfilename(
                title=f"Sélectionnez le fichier {display_name}",
                filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")],
                initialdir=os.getcwd()
            )
            
            if file_path:
                filename = os.path.basename(file_path)
                excel_files.append(filename)
                self.log(f"Fichier sélectionné : {filename}")
            else:
                self.log(f"Aucun fichier sélectionné pour {display_name}")
                return
        
        if excel_files:
            self.save_excel_filenames(excel_files)
            self.load_excel_filenames()
            self.log("Tous les fichiers Excel ont été chargés avec succès!")
        else:
            self.log("Aucun fichier n'a été sélectionné.")

    def start_scripts(self):
        # Vérifier si les fichiers Excel ont été chargés
        if not os.path.exists(EXCEL_FILENAMES_FILE):
            messagebox.showwarning("Fichiers manquants", 
                                   "Veuillez d'abord charger les fichiers Excel en cliquant sur 'CHARGER FICHIERS EXCEL'.")
            return
        
        self.clear_outputs()
        self.progress["maximum"] = len(scripts)
        
        for i, script in enumerate(scripts):
            self.log(f"Exécution de {script}...")
            
            if script == "000.INSTALL-MODULES.py":
                if not self.check_modules():
                    self.run_script(script)
            else:
                self.run_script(script)
            
            self.progress["value"] = i + 1
            self.update_idletasks()

        self.log("Tous les scripts ont été exécutés.")

    def open_category_editor(self):
        # Ouvrir la fenêtre d'édition des catégories
        CategoryEditor(self)

    def check_modules(self):
        missing_modules = []
        for module in required_modules:
            try:
                importlib.import_module(module)
            except ImportError:
                missing_modules.append(module)
        
        if missing_modules:
            self.log(f"Modules manquants : {', '.join(missing_modules)}")
            return False
        else:
            self.log("Tous les modules requis sont installés.")
            return True

    def run_script(self, script_name):
        try:
            result = subprocess.run(
                [sys.executable, script_name],
                capture_output=True, text=True, check=True
            )
            self.log(result.stdout)
            return result.stdout.splitlines() if result.stdout else []
        except subprocess.CalledProcessError as e:
            self.log(f"Erreur lors de l'exécution de {script_name}: {e.stderr}")
            return []

    def clear_outputs(self):
        self.console_output.delete(1.0, tk.END)
        self.progress["value"] = 0

    def log(self, message):
        self.console_output.insert(tk.END, f"{message}\n")
        self.console_output.see(tk.END)
        self.update_idletasks()

    def save_excel_filenames(self, filenames):
        with open(EXCEL_FILENAMES_FILE, 'w') as f:
            for filename in filenames:
                f.write(f"{filename}\n")

    def load_excel_filenames(self):
        self.file_visualization.delete(1.0, tk.END)
        if os.path.exists(EXCEL_FILENAMES_FILE):
            with open(EXCEL_FILENAMES_FILE, 'r') as f:
                content = f.read()
                self.file_visualization.insert(tk.END, content)
        else:
            self.file_visualization.insert(tk.END, "Aucun fichier Excel chargé.\nCliquez sur 'CHARGER FICHIERS EXCEL' pour commencer.")

if __name__ == "__main__":
    app = ScriptExecutor()
    app.mainloop()