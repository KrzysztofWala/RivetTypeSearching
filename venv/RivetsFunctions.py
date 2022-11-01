import pandas as pd
import openpyxl
import xlsxwriter

def ImportParam():
    # Reading excel data file
    df_excel = pd.read_excel('Data.xlsx')

    # Import information about spots
    SpotsFileName = df_excel.columns[0]
    df_s_param=df_excel[SpotsFileName].values.tolist()

    # Import information about material
    MaterialFileName = df_excel.columns[1]
    df_m_param = df_excel[MaterialFileName].values.tolist()

    # Openig excel files
    df_m = pd.read_excel(MaterialFileName)
    df_s = pd.read_excel(SpotsFileName)

    # Extracting the necessary data, described in Data.xlsx from data frame
    dfm = (df_m[df_m_param])
    dfs = (df_s[df_s_param])

    return (dfs, dfm)



