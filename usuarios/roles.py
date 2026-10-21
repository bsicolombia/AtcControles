"""Roles del sistema. Se implementan con los Grupos de Django."""
ADMINISTRADOR = "Administrador"
ALMACENISTA = "Almacenista"
COORDINADOR = "Coordinador"
SUPERVISOR = "Supervisor"
TECNICO = "Tecnico"
DIGITADOR = "Digitador"

TODOS = (ADMINISTRADOR)

ROLES_ADMIN = (ADMINISTRADOR,)
ROLES_COORDINACION = (ADMINISTRADOR, COORDINADOR)
ROLES_ALMACEN = (ADMINISTRADOR, COORDINADOR, ALMACENISTA)
ROLES_DIGITADOR = (ADMINISTRADOR, COORDINADOR, ALMACENISTA,DIGITADOR)
ROLES_SAP = (ADMINISTRADOR, COORDINADOR)

def tiene_rol(user, *roles):
    """Comprueba si el usuario pertenece a alguno de los roles indicados."""
    if not user.is_authenticated:
        return False

    # El superusuario tiene acceso administrativo.
    if user.is_superuser:
        return True

    return user.groups.filter(name__in=roles).exists()
