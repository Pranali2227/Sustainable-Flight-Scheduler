import os

import oracledb
from dotenv import dotenv_values

settings = dotenv_values(".env")
user = settings.get("ORACLE_USER")
password = settings.get("ORACLE_PASSWORD")
if not user or not password:
    raise SystemExit("Local Oracle username/password are not configured.")

for dsn in ("localhost:1521/XE", "localhost:1521/SYS$USERS", "localhost:1521/XEPDB1"):
    try:
        with oracledb.connect(user=user, password=password, dsn=dsn) as connection:
            with connection.cursor() as cursor:
                cursor.execute("SELECT SYS_CONTEXT('USERENV','SERVICE_NAME'), SYS_CONTEXT('USERENV','CON_NAME') FROM DUAL")
                service, container = cursor.fetchone()
                cursor.execute("SELECT TABLE_NAME FROM USER_TABLES")
                tables = {row[0] for row in cursor.fetchall()}
                matching = sorted(tables & {"AIRCRAFT", "CREW", "FLIGHT"})
            print(f"{dsn}: connected service={service}, container={container}, matching tables={matching}")
    except oracledb.Error as exc:
        error = exc.args[0]
        print(f"{dsn}: connection failed code={getattr(error, 'code', 'unknown')}")
