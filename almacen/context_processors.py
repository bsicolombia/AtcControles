# almacen/context_processors.py
#
# Calcula las inconsistencias PENDIENTES para mostrarlas en la burbuja
# flotante (_burbuja_inconsistencias.html) en todas las pantallas.

from django.db.models import Count

from usuarios.roles import ROLES_ALMACEN, tiene_rol

from .etiquetas import ETIQUETAS_ERRORES
from .models import Inconsistencia


def inconsistencias_burbuja(request):
    user = getattr(request, "user", None)

    if not user or not user.is_authenticated:
        return {}

    # Solo quien puede abrir la lista de inconsistencias ve la burbuja.
    if not (user.is_superuser or any(tiene_rol(user, r) for r in ROLES_ALMACEN)):
        return {}

    # Si ya está en la lista de inconsistencias no hace falta la burbuja.
    match = getattr(request, "resolver_match", None)
    if match and match.url_name == "inconsistencias":
        return {}

    filas = (
        Inconsistencia.objects
        .filter(estado=Inconsistencia.Estado.PENDIENTE)
        .values("tipo")
        .annotate(n=Count("id"))
        .order_by("-n")
    )

    tipos = [
        {
            "tipo": f["tipo"],
            "etiqueta": ETIQUETAS_ERRORES.get(f["tipo"], f["tipo"]),
            "n": f["n"],
        }
        for f in filas
    ]

    total = sum(t["n"] for t in tipos)
    if not total:
        return {}

    return {
        "burbuja_inc": {
            "total": total,
            "tipos": tipos[:6],
            "otros": max(len(tipos) - 6, 0),
        }
    }