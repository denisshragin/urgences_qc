#from google.colab import auth
#from oauth2client.client import GoogleCredentials
import pandas as pd
import gspread
import math


#auth.authenticate_user()
#gc=gspread.authorize(GoogleCredentials.get_application_default())

url_urgences="https://www.msss.gouv.qc.ca/professionnels/statistiques/documents/urgences/Releve_horaire_urgences_7jours.csv"

gs_urgences = "https://docs.google.com/spreadsheets/d/1iJQK9JN7POOmMpO2G9VsyKuo1iI4bzNcdJJ3T6xBSIg/edit#gid=0"

json_filename = "qc-urgences-9418e0f4958b.json"

# Add header to sheets in workbook

def add_header_to_sheets(json_filename, workbook_url, list_of_sheet_names, column_header):
	gc = gspread.service_account(filename=json_filename)
	workbook=gc.open_by_url(workbook_url)
	for sheet_name in list_of_sheet_names:
		worksheet = workbook.worksheet(sheet_name)
		worksheet.append_row(column_header)



# Clear workbook

def clear_workbook(json_filename, workbook_url, list_of_sheet_names):
	gc = gspread.service_account(filename=json_filename)
	workbook=gc.open_by_url(workbook_url)
	for sheet_name in list_of_sheet_names:
		worksheet = workbook.worksheet(sheet_name)
		worksheet.clear()


df = pd.read_csv(url_urgences, parse_dates=[' Mise_a_jour'], encoding='latin1')
# df.head(5)
# df.columns



# Создание списка названий листов в рабочей книге Urgences (url_urgences) 

urgences_df_header = list(df)
urgences_df_header_stripped = []
for item in urgences_df_header:
  item_stripped = item.strip()
  urgences_df_header_stripped.append(item_stripped)

list_of_sheet_names = urgences_df_header_stripped[:-2]
list_of_sheet_names.append('Taux_occupation')

# print(urgences_df_header_stripped)
# print(list_of_sheet_names)



# Создание списка названий столбцов (118 столбцов)

nom_installation = df[' Nom_installation ']
list_installations=list(nom_installation)
column_header=[" Heure_de_l'extraction_(image) ", ' Mise_a_jour'] + list_installations
# print(column_header)


# update spreadsheets with today's data

gc = gspread.service_account(filename=json_filename)

wb_urgences=gc.open_by_url(gs_urgences)

data_nom_etablissement = {}
data_nom_installation = {}
data_no_permis_installation = {}
data_civieres_fonctionnelles = {}
data_civieres_occupees = {}
data_patients_24_h = {}
data_patients_48_h = {}
data_taux_occupation = {}

list_of_data_dict = [data_nom_etablissement, data_nom_installation, data_no_permis_installation, data_civieres_fonctionnelles, data_civieres_occupees, data_patients_24_h, data_patients_48_h, data_taux_occupation]

nom_etablissement_to_append = {}
nom_installation_to_append = {}
no_permis_installation_to_append = {}
civieres_fonctionnelles_to_append = {}
civieres_occupees_to_append = {}
patients_24_h_to_append = {}
patients_48_h_to_append = {}
taux_occupation_to_append = {}

list_of_dict_to_append = [nom_etablissement_to_append, nom_installation_to_append, no_permis_installation_to_append, civieres_fonctionnelles_to_append, civieres_occupees_to_append, patients_24_h_to_append, patients_48_h_to_append,taux_occupation_to_append]

merged_list = list(zip(list_of_data_dict, list_of_dict_to_append))
#print(merged_list)

heure = str(df.iloc[0, -2])
mise_a_jour = str(df.iloc[0, -1])

for data_dict in list_of_data_dict:
  data_dict[" Heure_de_l'extraction_(image) "] = heure
  data_dict[" Mise_a_jour"] = mise_a_jour


for index, row in df.iterrows():
  data_nom_etablissement[row[' Nom_installation '].strip()] = row['Nom_etablissement ']
  data_nom_installation[row[' Nom_installation '].strip()] = row[' Nom_installation ']
  data_no_permis_installation[row[' Nom_installation '].strip()] = row[' No_permis_installation ']
  data_civieres_fonctionnelles[row[' Nom_installation '].strip()] = row[' Nombre_de_civieres_fonctionnelles ']
  data_civieres_occupees[row[' Nom_installation '].strip()] = row[' Nombre_de_civieres_occupees ']
  data_patients_24_h[row[' Nom_installation '].strip()] = row[' Nombre_de_patients_sur_civiere_plus_de_24_heures ']
  data_patients_48_h[row[' Nom_installation '].strip()] = row[' Nombre_de_patients_sur_civiere_plus_de_48_heures ']
  if row[' Nombre_de_civieres_occupees '].isnumeric() and row[' Nombre_de_civieres_fonctionnelles '].isnumeric():
    data_taux_occupation[row[' Nom_installation '].strip()] = math.trunc(int(row[' Nombre_de_civieres_occupees ']) / int(row[' Nombre_de_civieres_fonctionnelles '])*100)
  else:
    data_taux_occupation[row[' Nom_installation '].strip()] = 'NaN'

for tuple_item in merged_list:
  for key in column_header:
    if key in tuple_item[0]:
      tuple_item[1][key] = tuple_item[0][key]

sheet_vs_dict_to_append = dict(zip(list_of_sheet_names, list_of_dict_to_append))
#print(sheet_vs_dict_to_append)


for sheet in list_of_sheet_names:
  ws_urgences = wb_urgences.worksheet(sheet)
  data_to_append_list = list((sheet_vs_dict_to_append[sheet].values()))
  #print(data_to_append_list)
  ws_urgences.append_row(data_to_append_list)

