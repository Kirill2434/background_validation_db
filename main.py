from PostgreSQL_managment import create_table_db


def main_manager(argument):
    create = create_table_db
    if argument == 'Создать':
        return create
    elif argument == 'Наполнить':
        pass

