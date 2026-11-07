import server

names = {
    "AIRCRAFT", "BOOKING", "CREW", "DISRUPTION", "EMISSIONS", "FLIGHT",
    "FLIGHT_CREW", "LOYALTY_TIER", "PASSENGER", "SPEC_CARGO",
}
connection = server.get_pool().acquire()
cursor = connection.cursor()
cursor.execute("SELECT USER, SYS_CONTEXT('USERENV','SERVICE_NAME'), SYS_CONTEXT('USERENV','CON_NAME') FROM DUAL")
print("Oracle session user/service/container:", cursor.fetchone())
cursor.execute("SELECT TABLE_NAME FROM USER_TABLES")
print("matching owned tables:", sorted(row[0] for row in cursor.fetchall() if row[0] in names))
cursor.execute("SELECT OWNER, TABLE_NAME FROM ALL_TABLES")
print("matching visible tables:", sorted((owner, table) for owner, table in cursor.fetchall() if table in names))
cursor.close()
connection.close()
