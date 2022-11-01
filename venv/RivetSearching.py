import pandas as pd
import openpyxl
import xlsxwriter
# from RivetsFunctions import *
import RivetsFunctions as rf

# Creating dictionary variables
dfm_dict = {}
dfm_dict_temp = {}
dfs_dict = {}
Spot_material_dict={}

# Import data
dfs, dfm = rf.ImportParam()
print(dfs)
print(dfm)


# Creating material dictionary - dfm_dict{} - collection all rivet type assigned to metal sheets combinations
# Iteration on dfm rows
for row in dfm.itertuples():
    dfm_dict_temp = {}
    # Adding combination meterial:tickness to dictionary
    for i in range(2,7,2):
        # Checking if excel cell is not empty
        if pd.isnull(row[i])==False:
            dfm_dict_temp[row[i]] = row[i+1]
    # Alphabetical sort
    dfm_dict_temp=dict(sorted(dfm_dict_temp.items()))
    # Checking if dfm_dict_temp is not empty
    if dfm_dict_temp != {}:
        # Adding dfm_dict_temp if not empty to dfm_dict
        dfm_dict[row[1]]=dfm_dict_temp

print(dfm_dict)

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
