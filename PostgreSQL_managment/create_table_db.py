from PostgreSQL_managment.check_connection_db import connect_to_db


def create_table():
    """Функция создания таблицы в базе PostgreSQL"""
    cursor = connect_to_db()
    try:
        cursor.execute(
            ''' 
            CREATE TABLE IF NOT EXISTS public.test_2
            (
                col_id integer NOT NULL GENERATED ALWAYS AS IDENTITY ( INCREMENT 1 START 1 MINVALUE 1 MAXVALUE 100 CACHE 1 ),
                "col-2" text COLLATE pg_catalog."default",
                CONSTRAINT test_2_pkey PRIMARY KEY (col_id)
            )
            
            TABLESPACE pg_default;
            
            ALTER TABLE IF EXISTS public.test_2
                OWNER to postgres;
            '''
        )
        cursor.close()
    except Exception as error:
        print(f"Ошибка при создании таблицы: {error}")
    return 'Таблица создана упешно!'


print(create_table())
