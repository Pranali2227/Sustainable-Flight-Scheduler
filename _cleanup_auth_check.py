
import server

emails = ("northstar-auth-check@example.invalid", "northstar-auth-check-2@example.invalid")
with server.connection() as connection:
    with connection.cursor() as cursor:
        cursor.execute(f"DELETE FROM {server.AUTH_TABLE} WHERE EMAIL IN (:first, :second)", {"first": emails[0], "second": emails[1]})
        connection.commit()
        print("temporary verification accounts removed:", cursor.rowcount)
