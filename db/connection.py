import psycopg2
from psycopg2 import Error
import pandas as pd
from constant import auth


def con(auth):
    try:
    # Подключение к существующей базе данных
    connection = psycopg2.connect(user=auth["user"],
                                  # пароль, который указали при установке PostgreSQL
                                  password=auth["password"],
                                  host=auth["host"],
                                  port=auth["port"],
                                  database=auth["database"])

    except (Exception, Error) as error:
    connection = False
    print("Ошибка при работе с PostgreSQL", error)
    return connection


def connection_test():
    # Распечатать сведения о PostgreSQL
    cursor = connection.cursor()
    print("Информация о сервере PostgreSQL")
    print(connection.get_dsn_parameters(), "\n")
    # Выполнение SQL-запроса
    cursor.execute("SELECT version();")
    # Получить результат
    record = cursor.fetchone()
    print("Вы подключены к - ", record, "\n")
    cursor.close()


if __name__ == '__main__':
    con(auth)
