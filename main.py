from docx import Document
import psycopg2


def create_table_if_not_exists():
    conn = psycopg2.connect('postgresql://user:password@localhost:5432/db_name')
    cursor = conn.cursor()

    create_table_query = f'''
        CREATE TABLE IF NOT EXISTS birthdays (
            id SERIAL PRIMARY KEY,
            birth_date TEXT,
            full_name VARCHAR,
            position TEXT
        );
    '''

    cursor.execute(create_table_query)
    conn.commit()
    cursor.close()
    conn.close()


def read_tables(filename):
    doc = Document(filename)
    line = []
    data = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text not in line:
                    line.append(cell.text)
            data.append(line)
            line = []
    return data


def write_to_db(data):
    conn = psycopg2.connect('postgresql://user:password@localhost:5432/db_name')
    cursor = conn.cursor()
    for line in data[1:]:
        line_new = []
        for i in line:
            i = i.replace("\n", " ")
            line_new.append(i)
        if len(line_new) >= 4:  # Проверка на длину списка перед обращением к индексам
            cursor.execute('INSERT INTO birthdays (birth_date, full_name, position) VALUES (%s, %s, %s)',
                           [line_new[1], line_new[2], line_new[3]])
        conn.commit()
    cursor.close()
    conn.close()


if __name__ == '__main__':
    create_table_if_not_exists()
    parsed_data = read_tables('BirthdayCelebrants.docx')
    write_to_db(parsed_data)
