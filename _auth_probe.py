import traceback

import server
from fastapi import Response

try:
    server.ensure_auth_table()
    print("auth table setup: ok")
except Exception as exc:
    print("auth table setup:", type(exc).__name__, str(exc))
    traceback.print_exc()

test_email = "northstar-auth-check-2@example.invalid"
try:
    response = Response()
    result = server.signup({"name": "Local Test User", "email": test_email, "password": "TemporaryCheckPass42!"}, response)
    print("signup result:", result, "cookie set:", bool(response.headers.get("set-cookie")))
    login_response = Response()
    print("login result:", server.login({"email": test_email, "password": "TemporaryCheckPass42!"}, login_response), "cookie set:", bool(login_response.headers.get("set-cookie")))
except Exception as exc:
    print("login:", type(exc).__name__, str(exc))
    traceback.print_exc()
finally:
    with server.connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(f"DELETE FROM {server.AUTH_TABLE} WHERE EMAIL = :email", {"email": test_email})
            conn.commit()
            print("disposable test account removed:", cursor.rowcount == 1)
