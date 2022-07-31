import pandas as pd
import openpyxl

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

file_1 = './excel_files/Пример №1.xlsx'
file_2 = './excel_files/Пример №2.xlsx'
file_3 = './excel_files/Пример_3_20201022.xlsx'



def test1():
    df = pd.read_excel(file_1, sheet_name='Светофор №5', skiprows=4, nrows=18, usecols='A:R')
    df.dropna(inplace=True)
    return df


def test2():
    wb = openpyxl.load_workbook(filename=file_1, read_only=True)
    ws = wb['Светофор №5']
    data_rows = []
    for row in ws['B7':'R78']:
        data_cols = []
        for cell in row:
            data_cols.append(cell.value)
        data_rows.append(data_cols)
    df = pd.DataFrame(data_rows)
    return df


if __name__ == '__main__':
    print(test1().head())
    print('-' * 30)
    print(test2().head())