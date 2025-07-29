import pandas as pd
import sys
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

# Print script arguments for debugging
print(f"Script arguments: {sys.argv}")

# R1 Chargement du fichier Excel - Récupération automatique du nom
excel_file = get_excel_filename('gta hs paie')
if not excel_file:
    print("Erreur : Impossible de trouver le fichier 'GTA HS PAIE'")
    sys.exit(1)

print(f"Attempting to read Excel file: {excel_file}")

try:
    df = pd.read_excel(excel_file)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    sys.exit(1)

# Filtrer les lignes de la colonne O (Etat) est egale a C
df = df[df['Etat'] == 'C']

# Fonction pour transformer la date au format mois annee en francais et avec une majuscule
def format_date(date_str):
    month_names = ['JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE']
    try:
        date = pd.to_datetime(date_str)
        month_index = date.month - 1
        month_name = month_names[month_index]
        year = date.year
        return f"{month_name} {year}"
    except Exception as e:
        print(f"Error formatting date {date_str}: {e}")
        return str(date_str)

# Definition des categories pour chaque type
categories_heures_sup = ['HS_NORMALES_INF14', 'HS_NORMALES_SUP14', 'HS_NUIT_INF14', 'HS_NUIT_SUP14', 'HS_DIM_ET_JF_INF14', 'HS_DIM_ET_JF_SUP14']
categories_astreintes = ['AST_SEMAINE', 'AST_SEM_CALENDAIRE', 'AST_SEM_CALEND_MAJO', 'AST_DJF', 'AST_DJF_MAJO', 'AST_WEEKEND', 'AST_NUIT']
categories_permanences = ['PERM_DJF_MAJO', 'PERM_DJF', 'PERM_SAMEDI_MAJO', 'PERM_WEEKEND']
categories_interventions = ['INT_SEMAINE', 'INT_NUIT', 'INT_SAMEDI', 'INT_DJF']

# Fonction pour filtrer les donnees en fonction de la categorie
def filter_data(df, cat_list, cat_name):
    # print(f"Filtering data for category: {cat_name}")
    # print(f"Unique categories in DataFrame: {df['Categories'].unique()}")
    filtered_df = df[df['Categories'].isin(cat_list)].copy()
    
    # Debug print
    print(f"Number of rows for {cat_name}: {len(filtered_df)}")
    
    if filtered_df.empty:
        print(f"Warning: No data found for category {cat_name}")
    
    filtered_df['Categorie'] = cat_name
    return filtered_df

# Filtrer les donnees pour chaque type avec gestion des erreurs
print("Starting data filtering...")
try:
    heures_sup_data = filter_data(df, categories_heures_sup, 'Heures Supplémentaires')
    astreintes_data = filter_data(df, categories_astreintes, 'Astreintes')
    permanences_data = filter_data(df, categories_permanences, 'Permanences')
    interventions_data = filter_data(df, categories_interventions, 'Interventions')
except Exception as e:
    print(f"Error during data filtering: {e}")
    sys.exit(1)

# Regrouper les donnees par matricule et mois annee
def format_dates_and_group(data):
    if data.empty:
        print("Empty DataFrame passed to format_dates_and_group")
        return pd.DataFrame(columns=['Matricule', 'Dates', 'Mois Année'])
    
    # Vérifier les colonnes requises
    required_columns = ['Matricule', 'Dates']
    missing_columns = [col for col in required_columns if col not in data.columns]
    
    if missing_columns:
        print(f"Missing columns: {missing_columns}")
        print(f"Available columns: {list(data.columns)}")
        raise ValueError(f"Missing required columns: {missing_columns}")
    
    # Convert dates to strings before processing
    data['Dates_str'] = data['Dates'].astype(str)
    data['Mois Année'] = data['Dates'].apply(format_date)
    
    # Regroupement avec gestion des erreurs
    try:
        grouped_data = (
            data.groupby(['Matricule'])
            .agg({
                'Dates_str': lambda x: '|'.join(x.astype(str).unique()),
                'Mois Année': lambda x: '|'.join(x.astype(str).unique())
            })
            .reset_index()
        )
        # Rename 'Dates_str' to 'Dates' for consistency
        grouped_data = grouped_data.rename(columns={'Dates_str': 'Dates'})
        return grouped_data
    except Exception as e:
        print(f"Error during grouping: {e}")
        # Print more debugging info
        print(f"Data types in DataFrame: {data.dtypes}")
        return pd.DataFrame(columns=['Matricule', 'Dates', 'Mois Année'])

# Traiter les données pour chaque catégorie avec gestion des erreurs
print("Processing data for each category...")
try:
    heures_sup_data = format_dates_and_group(heures_sup_data)
    astreintes_data = format_dates_and_group(astreintes_data)
    permanences_data = format_dates_and_group(permanences_data)
    interventions_data = format_dates_and_group(interventions_data)
except Exception as e:
    print(f"Error during data processing: {e}")
    sys.exit(1)

# Ecrire les donnees dans un nouveau fichier Excel
try:
    with pd.ExcelWriter('SORTIE.xlsx') as writer:
        heures_sup_data.to_excel(writer, sheet_name='HEURES SUPPLEMENTAIRES', index=False)
        astreintes_data.to_excel(writer, sheet_name='ASTREINTES', index=False)
        permanences_data.to_excel(writer, sheet_name='PERMANENCES', index=False)
        interventions_data.to_excel(writer, sheet_name='INTERVENTIONS', index=False)
    print("SORTIE.xlsx created successfully.")
except Exception as e:
    print(f"Error writing Excel file: {e}")
    sys.exit(1)
