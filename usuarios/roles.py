"""Roles del sistema. Se implementan con los Grupos de Django."""
ADMINISTRADOR = "Administrador"
ALMACENISTA = "Almacenista"
TODOS = (ADMINISTRADOR, ALMACENISTA)


def tiene_rol(user, *roles):
    if not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    return user.groups.filter(name__in=roles).exists()
