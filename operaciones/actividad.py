from .models import RegistroActividad


def _ip(request):
    if request is None:
        return ""
    fwd = request.META.get("HTTP_X_FORWARDED_FOR", "")
    return (fwd.split(",")[0].strip() if fwd else request.META.get("REMOTE_ADDR", ""))[:45]


def registrar(request, accion, referencia="", detalle="", tecnico=None,
              datos=None, usuario=None):
    """Una línea en la consola. Llamar DESPUÉS del transaction.atomic, no dentro."""
    u = usuario or (request.user if request is not None and request.user.is_authenticated else None)
    RegistroActividad.objects.create(
        usuario=u,
        usuario_nombre=(u.get_full_name() or u.username) if u else "",
        accion=accion,
        referencia=str(referencia)[:60],
        tecnico=tecnico,
        detalle=detalle or "",
        datos=datos or {},          # solo valores simples (str, int, bool, listas)
        ip=_ip(request),
    )