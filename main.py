import pandas as pd
import openpyxl
from datetime import datetime

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
pd.set_option('max_colwidth', 520)

file_1 = './excel_files/Пример №1.xlsx'
file_2 = './excel_files/Пример №2.xlsx'
file_3 = './excel_files/Пример_3_20201022.xlsx'


def sample_1():
    df = pd.read_excel(file_1, sheet_name='Светофор №5', skiprows=0, usecols='A:R')
    df_names = df.head(2)
    df_data = df.loc[5:]
    df_n = df_names.T
    df_n.ffill(axis=0, inplace=True)
    df_n.fillna('', inplace=True)

    df_n['name'] = df_n[0] + ' ' + df_n[1]
    col_one_list = df_n['name'].tolist()

    print(col_one_list)
    df_data.columns = col_one_list
    print(df)
    print(df_data)
    # df.dropna(inplace=True)
    # return df


def sample_2():
    wb = openpyxl.load_workbook(filename=file_1, read_only=True)
    ws = wb['Светофор №5']
    data_rows = []
    for row in ws['B2':'R78']:
        data_cols = []
        for cell in row:
            data_cols.append(cell.value)
        data_rows.append(data_cols)
    df = pd.DataFrame(data_rows)
    df_names = df.head(3)

    df_data = df.loc[4:]
    df_n = df_names.T
    df_n.ffill(axis=0, inplace=True)
    df_n.fillna('', inplace=True)

    df_n['name'] = df_n[0] + ' ' + df_n[1] + ' (' + df_n[2] + ')'
    # df_n['name'] = df[[0, 1, 2]].agg(" ".join, axis=1)
    df_n['name'].replace(to_replace=r'\n', value='', inplace=True, regex=True)
    # df_n['name'].replace(to_replace=r'  ()', value='', inplace=True, regex=True)
    # print(df_n)
    # exit()

    col_one_list = df_n['name'].tolist()
    # print(df_data)
    # exit()
    df_data.columns = col_one_list
    # print(df)
    df_unpivot = pd.melt(df_data, id_vars=col_one_list[0], value_vars=col_one_list[1:],
                         var_name='Params', value_name='Values')
    df_unpivot.rename(columns={col_one_list[0]: "Subject"}, inplace=True)
    df_unpivot['time_stamp'] = datetime.now()
    print(df_unpivot)

    df_unpivot.to_csv('out.csv', index=False, sep='|')
    # print(df)
    return df


def sample_3():
    wb = openpyxl.load_workbook(filename=file_2, read_only=True)
    ws = wb['Лист1']
    data_rows = []
    for row in ws['A16':'J60']:
        data_cols = []
        for cell in row:
            data_cols.append(cell.value)
        data_rows.append(data_cols)
    df = pd.DataFrame(data_rows)
    return df


if __name__ == '__main__':
    sample_2()
    print('-' * 30)
    # print(sample_3().head(10))
