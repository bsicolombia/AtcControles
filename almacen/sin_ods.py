import re

from django.db.models import Count, Q, Sum
from django.db.models.functions import Coalesce


def ordenes_sin_ods_qs():
    """Operaciones cuyo número es un temporal (prefijo + dígitos, ej. RSO0001)."""
    from operaciones.models import Operacion, TipoOperacion

    prefijos = list(
        TipoOperacion.objects.exclude(prefijo_sin_ods="")
        .values_list("prefijo_sin_ods", flat=True).distinct()
    )
    if not prefijos:
        return Operacion.objects.none()

    cond = Q()
    for p in prefijos:
        cond |= Q(numero_orden__regex=rf"^{re.escape(p)}[0-9]+$")
    return Operacion.objects.filter(cond)


def con_cantidades(qs):
    """Agrega n_materiales (líneas) y piezas (suma de cantidades instaladas)."""
    from .models import Asignacion

    instalada = Q(materiales__estado=Asignacion.Estado.INSTALADA)
    return qs.annotate(
        n_materiales=Count("materiales", filter=instalada),
        piezas=Coalesce(Sum("materiales__cantidad", filter=instalada), 0),
    )