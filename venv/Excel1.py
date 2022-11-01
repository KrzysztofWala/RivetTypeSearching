import pandas as pd
import openpyxl

# https://pythonbasics.org/pandas-iterate-dataframe/

excel_file='test_excel.xlsx'
###############################
#opening
df = pd.read_excel(excel_file, sheet_name=None)
print(df['Arkusz2'])

###############################
# Write DataFrame to Excel file
###############################
df = pd.DataFrame([[11, 21, 31], [12, 22, 32], [31, 32, 33]],
                   index=['one', 'two', 'three'], columns=['a', 'b', 'c'])

# print(df)
df.to_excel('pandas_to_excel.xlsx', sheet_name='new_sheet_name')
df2 = df[['a', 'c']]
# print(df2)

with pd.ExcelWriter('pandas_to_excel.xlsx') as writer:
    df.to_excel(writer, sheet_name='sheet1')
    df2.to_excel(writer, sheet_name='sheet2')

###############################
df = pd.DataFrame({'age': [20, 32], 'state': ['NY', 'CA'], 'point': [64, 92]},
                  index=['Alice', 'Bob'])
print(df)
##### Iteracja po klumnach
# for column_name in df:
#     print(type(column_name))
#     print(column_name)
#     print('------\n')

##### Iteracja po zawartości kolumny
# for column_name2, item in df.items():
#     print(type(column_name2))
#     print(column_name2)
#     print('~~~~~~')
#
#     print(type(item))
#     print(item)
#     print(item[0])
#     print('------')

##### Iteracja po zawartości wiersza
for row in df.itertuples():
    print(type(row))
    print(row)
    print('------')

    print(row[0])
    print(row.point)
    print('------\n')



###############################
# Write DataFrame to Excel file
###############################