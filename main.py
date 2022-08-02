from parce_excel import *
import pandas as pd

pd.set_option('display.max_columns', 500)
pd.set_option('display.max_rows', 500)
pd.set_option('display.width', 1000)
pd.set_option('max_colwidth', 520)

file_1 = './excel_files/Пример №1.xlsx'
file_2 = './excel_files/Пример №2.xlsx'
file_3 = './excel_files/Пример_3_20201022.xlsx'


def main():
    # обработка Пример №1.xlsx
    print(sample_1(file_1, 'Светофор №5', start_range='B2', end_range='R78').head(15))
    print(sample_1(file_1, 'Светофор №5', start_range='B84', end_range='R96').head(15))

    # обработка Пример №3.xlsx
    print(sample_3(file_3, 'РФ', start_range='B2', end_range='N88').head(15))


if __name__ == '__main__':
    main()
