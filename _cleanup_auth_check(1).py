import server

with server.connection() as connection:
    with connection.cursor() as cursor:
        cursor.execute(f"DELETE FROM {server.AUTH_TABLE} WHERE EMAIL = :email", {"email": "northstar-ui-verification@example.invalid"})
        connection.commit()
        print("temporary UI verification account removed:", cursor.rowcount == 1)
