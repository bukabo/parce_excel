from datetime import datetime
import pandas as pd
import openpyxl

pd.set_option('display.max_columns', 500)
pd.set_option('display.max_rows', 500)
pd.set_option('display.width', 1000)
pd.set_option('max_colwidth', 120)


def get_data_from_excel_range(file, wb_name, start_range, end_range):
    """
    Функция получает данные из диапазона в excel файле.

    :param file: исходный excel файл
    :param wb_name: имя листа
    :param start_range: стратовая ячейка диапазона
    :param end_range: конечная ячейка диапазона
    :return: raw dataframe
    """
    wb = openpyxl.load_workbook(filename=file, read_only=True)
    ws = wb[wb_name]
    data_rows = []
    for row in ws[start_range:end_range]:
        data_cols = []
        for cell in row:
            data_cols.append(cell.value)
        data_rows.append(data_cols)
    df = pd.DataFrame(data_rows)
    return df


def sample_1(file, wb_name, start_range='A1', end_range='B2'):
    """
    'Cхлопывает' шапку и преобразует данные в плоскую таблицу.

    :param file: исходный excel файл
    :param wb_name: имя листа
    :param start_range: стратовая ячейка диапазона
    :param end_range: конечная ячейка диапазона
    :return: итоговый DataFrame
    """

    df = get_data_from_excel_range(file, wb_name, start_range, end_range)

    # получаем шапку из столбцов для преобразования
    df_names = df.head(3)

    # получаем данные
    df_data = df.loc[4:]

    # транспонируем шапку, заполняем значения вниз, объединяем в один столбец (будущие заголовки)
    df_names = df_names.T
    df_names.ffill(axis=0, inplace=True)
    df_names.fillna('', inplace=True)
    df_names['name'] = df_names[0] + ' ' + df_names[1] + ' (' + df_names[2] + ')'
    df_names['name'].replace(to_replace=r'\n', value='', inplace=True, regex=True)

    # собираем заголовки в список, применяем новые заголовки к датафрейму с данными
    col_one_list = df_names['name'].tolist()
    df_data.columns = col_one_list

    # анпивот датафрейма с данными
    df_unpivot = pd.melt(df_data, id_vars=col_one_list[0], value_vars=col_one_list[1:],
                         var_name='Params', value_name='Values')
    df_unpivot.rename(columns={col_one_list[0]: "Subject"}, inplace=True)

    df_unpivot['time_stamp'] = datetime.now()
    return df_unpivot


def sample_3(file, wb_name, start_range='A1', end_range='B2'):
    """
    'Cхлопывает' шапку и преобразует данные в плоскую таблицу.

    :param file: исходный excel файл
    :param wb_name: имя листа
    :param start_range: стратовая ячейка диапазона
    :param end_range: конечная ячейка диапазона
    :return: итоговый DataFrame
    """

    # получаем дату отчета из имени файла
    date_string = file.split('_')[-1].split('.')[0]
    report_date = datetime.strptime(date_string, '%Y%m%d').date()

    df = get_data_from_excel_range(file, wb_name, start_range, end_range)

    df_names = df.head(4)
    df_names = df_names.T
    df_names.ffill(axis=0, inplace=True)
    df_names.fillna('', inplace=True)
    df_names['name'] = df_names[0] + ' ' + df_names[1] + ' ' + df_names[2] + ' (' + df_names[3] + ')'
    df_names['name'].replace(to_replace=r'\n', value='', inplace=True, regex=True)

    df_data = df.loc[5:]

    # собираем заголовки в список, применяем новые заголовки к датафрейму с данными
    col_one_list = df_names['name'].tolist()
    df_data.columns = col_one_list

    # анпивот датафрейма с данными
    df_unpivot = pd.melt(df_data, id_vars=col_one_list[0], value_vars=col_one_list[1:],
                         var_name='Params', value_name='Values')
    df_unpivot.rename(columns={col_one_list[0]: "Subject"}, inplace=True)

    df_unpivot['report_date'] = report_date
    df_unpivot['time_stamp'] = datetime.now()

    return df_unpivot

def sample_2(file, wb_name, start_range='A1', end_range='B2'):
    """
    'Cхлопывает' шапку и преобразует данные в плоскую таблицу.

    :param file: исходный excel файл
    :param wb_name: имя листа
    :param start_range: стратовая ячейка диапазона
    :param end_range: конечная ячейка диапазона
    :return: итоговый DataFrame
    """

    df = get_data_from_excel_range(file, wb_name, start_range, end_range)

    df_names = df.head(3)
    df_names = df_names.T
    df_names[0].ffill(axis=0, inplace=True)
    df_names.fillna('', inplace=True)
    df_names['name'] = df_names[0] + ' ' + df_names[1] + df_names[2]
    df_names['name'].replace(to_replace=r'\n', value='', inplace=True, regex=True)

    df_data = df.loc[5:]

    # собираем заголовки в список, применяем новые заголовки к датафрейму с данными
    col_one_list = df_names['name'].tolist()
    df_data.columns = col_one_list



    # анпивот датафрейма с данными
    df_unpivot = pd.melt(df_data, id_vars=[col_one_list[0], col_one_list[1]], value_vars=col_one_list[2:],
                         var_name='Params', value_name='Values')
    df_unpivot.rename(columns={col_one_list[0]: "Code", col_one_list[1]: "Product"}, inplace=True)

    # df_unpivot['report_date'] = report_date
    df_unpivot['time_stamp'] = datetime.now()

    print(df_unpivot.head(10))
    return df_unpivot


if __name__ == '__main__':
    file_2 = '../excel_files/Пример №2.xlsx'
    sample_2(file_2, 'Лист1', start_range='A16', end_range='J60')