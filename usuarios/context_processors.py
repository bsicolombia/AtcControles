
from .roles import (
    ADMINISTRADOR,
    ALMACENISTA,
    DIGITADOR,
    COORDINADOR,
    tiene_rol,
)


def roles(request):
    user = request.user

    es_admin = False
    es_coordinador = False
    es_almacenista = False
    es_digitador = False
    puede_ver_almacen = False
    rol_actual = ""

    if user.is_authenticated:
        es_admin = (
            user.is_superuser
            or tiene_rol(user, ADMINISTRADOR)
        )

        es_coordinador = tiene_rol(user, COORDINADOR)
        es_almacenista = tiene_rol(user, ALMACENISTA)
        es_digitador = tiene_rol(user, DIGITADOR)

        # Administrador y coordinador tienen acceso completo.
        puede_ver_almacen = (
            es_admin
            or es_coordinador
            or (es_almacenista and not es_digitador)
        )

        if es_admin:
            rol_actual = ADMINISTRADOR
        elif es_coordinador:
            rol_actual = COORDINADOR
        elif es_digitador:
            rol_actual = DIGITADOR
        elif es_almacenista:
            rol_actual = ALMACENISTA
        else:
            rol_actual = "Sin rol"

    return {
        "es_admin": es_admin,
        "es_coordinador": es_coordinador,
        "es_almacenista": es_almacenista,
        "es_digitador": es_digitador,
        "puede_ver_almacen": puede_ver_almacen,
        "rol_actual": rol_actual,
    }