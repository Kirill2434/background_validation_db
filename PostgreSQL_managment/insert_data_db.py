from PostgreSQL_managment.check_connection_db import connect_to_db


def insert_data_to_db():
    """Функция добавления данных в таблицу PostgreSQL"""
    cursor = connect_to_db()
    try:
        cursor.execute(
            ''' 
                INSERT INTO test_2 ("col-2") 
                VALUES ('Тестовые данные')
            '''
        )
        cursor.close()
    except Exception as error:
        print(f"Ошибка при заполении данными: {error}")
    return 'Данные упешно добавлены!'


print(insert_data_to_db())
