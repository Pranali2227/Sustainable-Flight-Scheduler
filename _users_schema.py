import server

connection = server.get_pool().acquire()
cursor = connection.cursor()
for table in ("USERS",):
    cursor.execute("SELECT COLUMN_ID, COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE FROM USER_TAB_COLUMNS WHERE TABLE_NAME = :table_name ORDER BY COLUMN_ID", {"table_name": table})
    print(table, cursor.fetchall())
    cursor.execute("SELECT CONSTRAINT_NAME, CONSTRAINT_TYPE FROM USER_CONSTRAINTS WHERE TABLE_NAME = :table_name", {"table_name": table})
    print("constraints", cursor.fetchall())
cursor.close()
connection.close()
