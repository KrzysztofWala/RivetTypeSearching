import pandas as pd
import openpyxl
# https://www.youtube.com/watch?v=DCDe29sIKcE&list=PL-osiE80TeTsWmV9i9c58mdDCSskIFdDS&index=5
# https://github.com/CoreyMSchafer/code_snippets/tree/master/Python/Pandas

excel_file='test_excel.xlsx'
###############################
#opening
#df = pd.read_excel(excel_file, sheet_name=None)
df = pd.read_excel(excel_file)


# print(df.shape)
# print(df.info())
# print(df.head())
# wypisanie kolumny 'area1'   <class 'pandas.core.series.Series'>
print(df['area1'])
# print(df.area1)    # zamienna składnia
print(type(df['area1']))

# wyciąganie poszczególnych kolumn
print(df['x1'])
print(df[['x1','y1']])
print(type(df[['x1','y1']]))

#wypisanie wszystkich klumn
print(df.columns)

# dostęp do [0] wiersza
print("-------")

# dostęp do wierszy
print(df.iloc[[0,2]])
print("-------")

# dostęp do pojedyńczej wartości
print(df.iloc[0,3])

print("-------")
# dostęp do [wierszy] [kolumn]
print(df.iloc[[0,2,3],[1,2]])
print(df.loc[[1,2,3],'Pass'])
print("-------")

# Porównanie iloc / loc  dla loc można używać nazwy kolumny
print(df.iloc[[0,1],3])
print(df.loc[[0,1],'x1'])
print("-------")

# Zliczanie wystąpień
print("-------")
print(df['Pass'].value_counts())

# Wybieranie zakresu
print("-------")
print(df.loc[0:2,'x1':'Pass'])

# Wybieranie indexu
# print("-------")
# df.set_index('Spot', inplace=True)
# print(df.loc['s1'])
# print("-------")
# print(df.loc['s1', 'Pass'])
# print(df.loc[0]  > nie działa, zmieniony index
# df.reset_index(inplace=True) > reset indexu
# Wczytanie df z zaznaczonym indexem
# df = pd.read_excel(excel_file, index_col="Spot")
# Sortowanei przez index
# df.sort_index()
# print(df)


# Filtrowanie Part4
print("-------")
print(df['Pass']=='yes')
filt = (df['Pass']=='yes')
print("-------")
print(df[filt])
print(df.loc[filt])
print("-------")
print(df.loc[filt, 'x1'])

print("-------")
filt = (df['Pass']=='yes') & (df['area1']=='UB2')
print(df[filt])
print("-------")
print(df.loc[filt, 'robot1'])
# Pokazuje niepasujące wyniki / ~ negacja
print("-------")
print(df.loc[~filt, 'robot1'])

print("-------")
x_plus = (df['x1'] > 0)
print(df[x_plus])

print("-------")

areas = ['UB1', 'UB2']
filt_area = df['area1'].isin(areas)
print(df.loc[filt_area])

# Wyszukiwanie fragmentu tekstu
print("-------")
filt_str = df['area1'].str.contains('UB', na = False)
print(df.loc[filt_str, 'area1'])

# Modyfikacja danych (Part5)
# Zmian nazw kolumn
print(df.columns)
df.columns=['Xname1', 'Xarea1', 'Xrobot1', 'x1', 'y1', 'z1', 'XSpot', 'XPass']
print(df.columns)
df.columns = [x.upper() for x in df.columns]
print(df.columns)
df.columns = df.columns.str.replace('X','H_')
print(df.columns)
df.columns = ['name1', 'area1', 'robot1', 'x1', 'y1', 'z1', 'Spot', 'Pass']
print(df.columns)

# Zmian nazw pojedyńczych wartości
print("---------------------")
# df.loc[0] = ['x1', 's2,.....]
print(df.loc[2,'area1'])
df.loc[2,'area1']='UB12'
print(df.loc[2,'area1'])
print("---------------------")
df.loc[2,['x1','y1']]=[999,999]
print(df.loc[2])

# Inna metoda podmiany
df.at[2,'area1']='UB12'

filt_pass = (df['Pass'] == 'yes' )
df.loc[filt_pass, 'Pass']= 'tak'
print(df)

# Metody: apply / map / applymap / replace
print("---------------------")
# Można zastosować funkcje do każdego pola
print(df['Spot'].apply(len))

print("---------------------")
def up_pass(word):
    return word.upper()
df['Pass'] = df['Pass'].apply(up_pass)
print(df)
# * lambda
df['Pass'] = df['Pass'].apply(lambda x: x.lower())

print("---------------------")
print(df.apply(len))
print(df.apply(len, axis = 'columns'))

# applymap   ????
# print("---------------------")
# print(  df.applymap(str.lower)  )

# replace
df['area1']=df['area1'].replace({'UC': 'UCchuj'})
print( df)

# Dodawanie usuwanie klumn/wierszy (Part6)
# Add
print("---------------------")
df['full_name'] = df['area1'] + "_" + df['robot1']
print( df)
# Remove
print("---------------------")
df.drop(columns=['area1', 'robot1'], inplace=True)
print( df)
# Split
print("----------hhh-----------")
df[['new_area','new_robot']] = df['full_name'].str.split('_', expand=True)
print( df)
# Remove
print("---------------------")
df=df.drop(index = 6)
print(df)
print("---------------------")
filt=df['new_area'] == 'UCchuj'
df=df.drop(index = df[filt].index)
print(df)

# Zapis (Part11)
print("---------------------")
filt = (df['Pass'] == 'tak')
pass_yes_df = df.loc[filt]
print(pass_yes_df)
pass_yes_df.to_excel('pass_yes_df.xlsx')