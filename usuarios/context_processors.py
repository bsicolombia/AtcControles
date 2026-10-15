from .roles import ADMINISTRADOR, ALMACENISTA, tiene_rol


def roles(request):
    user = request.user
    es_admin = tiene_rol(user, ADMINISTRADOR)
    es_almacenista = tiene_rol(user, ALMACENISTA)
    if user.is_authenticated:
        rol = "Administrador" if es_admin else "Almacenista" if es_almacenista else "Sin rol"
    else:
        rol = ""
    return {"es_admin": es_admin, "es_almacenista": es_almacenista, "rol_actual": rol}
