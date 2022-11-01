import pandas as pd
import openpyxl
import xlsxwriter

df = pd.DataFrame({'name': ['Raphael', 'Donatello']})
df.to_csv('out.csv', index=False)
# df.to_csv('your.csv', index=False)


df_csv = pd.read_csv('out.csv')
# print(df_csv)
namef = df_csv.columns[0]
print('file name:', namef)
material = df_csv.iloc[[0,1]]
print('kolumns: \n', material)

# values.tolist()

print(df_csv[namef].values.tolist())