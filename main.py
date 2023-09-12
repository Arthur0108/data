import psycopg2
import pyexcel_ods
from datetime import datetime

db_params = {
    "dbname": "sql",
    "user": "postgres",
    "password": "01081990psql",
    "host": "localhost",
    "port": "5432"
}

data = pyexcel_ods.get_data("birthdays.ods")
sheet = data['Лист1']

connection = psycopg2.connect(**db_params)
cursor = connection.cursor()

create_table_query = """
CREATE TABLE IF NOT EXISTS birthdays (
    id SERIAL PRIMARY KEY,
    celebrate VARCHAR(255),
    birth_date TEXT,
    name VARCHAR(255),
    position VARCHAR(255),
    UNIQUE (celebrate, birth_date, name, position)
)
"""
cursor.execute(create_table_query)

for row in sheet:
    if len(row) >= 4:
        celebrate = row[0]
        birth_date_str = row[1]

        birth_year = int(birth_date_str[-2:])
        if birth_year < 100:
            birth_year += 1900
        birth_date_str = birth_date_str[:-2] + str(birth_year)

        birth_date = datetime.strptime(birth_date_str, "%d.%m.%Y").date().strftime("%d.%m.%Y")
        name = row[2]
        position = row[3]

        insert_query = """
            INSERT INTO birthdays (celebrate, birth_date, name, position)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (celebrate, birth_date, name, position) DO NOTHING
        """
        cursor.execute(insert_query, (celebrate, birth_date, name, position))
    else:
        pass

connection.commit()

cursor.close()
connection.close()
