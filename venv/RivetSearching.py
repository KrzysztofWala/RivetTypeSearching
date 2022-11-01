import pandas as pd
import openpyxl
import xlsxwriter
# from RivetsFunctions import *
import RivetsFunctions as rf



rf.ImportParamCSV()

# Openig excel files
material_file = 'Material.xlsx'
spots_file = 'Spot_list.xlsx'
df_m = pd.read_excel(material_file)
df_s = pd.read_excel(spots_file)

# material_file = '1_Nieten_Material.xlsx'
# spots_file = '2_U1G1_working.xlsx'
# df_m = pd.read_excel(material_file)
# df_s = pd.read_excel(spots_file)

# Creating dictionary variables
dfm_dict = {}
dfm_dict_temp = {}
dfs_dict = {}
Spot_material_dict={}


# Extracting the necessary data from full data frame
dfm=(df_m[['rivet_type','m1', 't1', 'm2', 't2', 'm3', 't3']])
dfs=(df_s[['Name','m1', 't1', 'm2', 't2', 'm3', 't3']])

# dfm=(df_m[[ 'Niettyp', 'Matrizentyp','BeMi','Material','Blechdicke','Material2','Blechdicke2','Material3','Blechdicke3']])
# dfs=(df_s[['Name','Material1', 'Thickness1', 'Material2', 'Thickness2', 'Material3', 'Thickness3']])

# Creating material dictionary - dfm_dict{} - collection all rivet type assigned to matal sheets combinations
# Iteration on dfm rows
for row in dfm.itertuples():
    dfm_dict_temp = {}
    # Adding combination meterial:tickness to dictionary
    for i in range(2,7,2):
    # for i in range(4, 9, 2):
        # Checking if excel cell is not empty
        if pd.isnull(row[i])==False:
            dfm_dict_temp[row[i]] = row[i+1]
    # Alphabetical sort
    dfm_dict_temp=dict(sorted(dfm_dict_temp.items()))
    # Checking if dfm_dict_temp is not empty
    if dfm_dict_temp != {}:
        # Adding dfm_dict_temp if not empty to dfm_dict
        dfm_dict[row[1]]=dfm_dict_temp

# Iteration on all rivets in spots list
for row in dfs.itertuples():
    # Creating material dictionary for each spot - dfs_dict
    dfs_dict = {}
    for i in range(2,7,2):
        # Checking if excel cell is not empty
        if pd.isnull(row[i])==False:
            dfs_dict[row[i]] = row[i+1]
    # Alphabetical sort
    dfs_dict = dict(sorted(dfs_dict.items()))
    # Searching for rivet type (dfs_dict) in meterial dictionary (dfm_dict)
    for key, value in dfm_dict.items():
        if value==dfs_dict:
            # Creating dictionary with Spot name and matching material
            Spot_material_dict[row[1]] = key


# Save result to file - dictionary (Spot_name:Rivet_type)
workbook = xlsxwriter.Workbook('Spot_list_RESULT.xlsx')
worksheet = workbook.add_worksheet()
l=0
for key, value in Spot_material_dict.items():
    worksheet.write(l, 0, key)
    worksheet.write(l, 1, value)
    l = l + 1
workbook.close()