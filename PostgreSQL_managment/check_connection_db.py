import os

import psycopg2
from dotenv import load_dotenv

load_dotenv()


# def connect_to_db():
#     """Функция соединения с базой PostgreSQL"""
#     conn = None
#     try:
#         print('Соединение с базой…')
#         conn = psycopg2.connect(
#                                 host=os.getenv('DB_HOST'),
#                                 database=os.getenv('DB_NAME'),
#                                 user=os.getenv('DB_USERNAME'),
#                                 password=os.getenv('DB_PASSWORD'))
#         cursor = conn.cursor()
#         cursor.execute(
#             'SELECT version();'
#         )
#         print(f"Данные о сервере PostgreSQL: {cursor.fetchone()}")
#     except Exception as error:
#         print(f"Ошибка при соединение к базе: {error}")
#     finally:
#         if conn:
#             cursor.close()
#             print('Курсор закрыт.')
#             conn.close()
#     return 'Все хорошо, соединение с базой установлено!'
#
#
# print(connect_to_db())

def connect_to_db():
    """Функция соединения с базой PostgreSQL"""
    conn = None
    try:
        print('Соединение с базой…')
        conn = psycopg2.connect(
                                host=os.getenv('DB_HOST'),
                                database=os.getenv('DB_NAME'),
                                user=os.getenv('DB_USERNAME'),
                                password=os.getenv('DB_PASSWORD'))
        conn.autocommit = True
    except Exception as error:
        print(f"Ошибка при соединение к базе: {error}")
    print('Соединение с базой PostgreSQL - установлено!')
    return conn.cursor()

