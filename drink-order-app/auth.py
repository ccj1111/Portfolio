import os
from functools import wraps

from flask import Response, request


def _check_auth(username: str, password: str) -> bool:
    expected_user = os.environ.get("ADMIN_USERNAME", "admin")
    expected_pass = os.environ.get("ADMIN_PASSWORD", "admin123")
    return username == expected_user and password == expected_pass


def _unauthorized():
    return Response(
        "需要管理者權限才能存取此頁面。",
        401,
        {"WWW-Authenticate": 'Basic realm="drink-order-admin"'},
    )


def requires_admin(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        auth = request.authorization
        if not auth or not _check_auth(auth.username, auth.password):
            return _unauthorized()
        return view(*args, **kwargs)

    return wrapped
