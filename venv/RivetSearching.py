import pandas as pd
import openpyxl
import xlsxwriter
# from RivetsFunctions import *
import RivetsFunctions as rf

Spot_material_dict={}
dfm_list = []
dfm_list_temp = []
dfs_list = []

# Import data
dfs, dfm = rf.ImportParam()

licz=0
# Creating material dictionary - dfm_dict{} - collection all rivet type assigned to metal sheets combinations
# Iteration on dfm rows
for row in dfm.itertuples():
    dfm_list_temp = []
    # Adding combination meterial:tickness to list
    for i in range(2,7,2):
        # Checking if excel cell is not empty
        if pd.isnull(row[i])==False:
            dfm_list_temp.append([row[i],row[i+1]])
    # Alphabetical sort
    dfm_list_temp.sort()

    # Checking if dfm_list_temp is not empty
    if len(dfm_list_temp) > 0:
        # Adding dfm_list_temp if not empty to dfm_list
        dfm_list.append([row[1],dfm_list_temp])

# Iteration on all rivets in spots list
for row in dfs.itertuples():
    # Creating material dictionary for each spot - dfs_dict
    dfs_list = []
    for i in range(2,7,2):

        # Checking if excel cell is not empty
        if pd.isnull(row[i])==False:
            dfs_list.append([row[i], row[i + 1]])
    # Alphabetical sort
    dfs_list.sort()

    # # Searching for rivet type (dfs_dict) in meterial dictionary (dfm_list)
    match=False
    for x in range (0, len(dfm_list)):
        if dfs_list==dfm_list[x][1]:
            Spot_material_dict[row[1]] = dfm_list[x][0]
            match=True
    if match==False:
        Spot_material_dict[row[1]] = 'not found'


# Save result to file - dictionary (Spot_name:Rivet_type)
workbook = xlsxwriter.Workbook('Spot_list_RESULT.xlsx')
worksheet = workbook.add_worksheet()
l=0
for key, value in Spot_material_dict.items():
    worksheet.write(l, 0, key)
    worksheet.write(l, 1, value)
    l = l + 1
workbook.close()