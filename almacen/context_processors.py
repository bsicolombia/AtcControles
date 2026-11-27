# almacen/context_processors.py
#
# Datos para las tres burbujas flotantes:
#   _burbuja_inconsistencias.html -> inconsistencias_burbuja  (burbuja_inc)
#   _burbuja_sap.html             -> sap_burbuja              (burbuja_sap)
#   _burbuja_almacen.html         -> almacen_burbuja          (burbuja_alm)

import logging
from django.db.models.functions import Trim, Upper
from django.core.cache import cache
from django.db.models import Count
from operaciones.models import CargaInventario, CargaStock, RegistroStock
from usuarios.roles import ROLES_ALMACEN, tiene_rol
from .config_sap import CODIGOS_EXCLUIDOS_VANTI
from .etiquetas import ETIQUETAS_ERRORES
from .models import Inconsistencia
from operaciones.models import CargaInventario

log = logging.getLogger(__name__)

CLAVE_SAP = "burbuja_sap_v1"
CLAVE_ALMACEN = "burbuja_almacen_v1"
SEGUNDOS_CACHE = 30


def _puede_ver(user):
    return bool(
        user and user.is_authenticated
        and (user.is_superuser or any(tiene_rol(user, r) for r in ROLES_ALMACEN))
    )
from usuarios.roles import ROLES_ALMACEN, ROLES_DIGITADOR, tiene_rol

CLAVE_SIN_ODS = "burbuja_sin_ods_v1"      # junto a CLAVE_STOCK
CLAVE_STOCK = "burbuja_stock_v2"

def invalidar_burbuja_sin_ods():
    """Llamar al crear, editar o asignar ODS a una orden."""
    cache.delete(CLAVE_SIN_ODS)

def invalidar_burbujas_conciliacion():
    """Llamar después de una carga SAP, IQ09 o MB52 para refrescar al instante."""
    cache.delete_many([CLAVE_SAP, CLAVE_ALMACEN, CLAVE_STOCK, CLAVE_SIN_ODS])


# ------------------------------------------------------------------
# 1) Inconsistencias (sin cambios)
# ------------------------------------------------------------------
def inconsistencias_burbuja(request):
    user = getattr(request, "user", None)

    if not user or not user.is_authenticated:
        return {}

    # Solo quien puede abrir la lista de inconsistencias ve la burbuja.
    if not _puede_ver(user):
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


# ------------------------------------------------------------------
# 2) Instalación: CICB vs SAP
# ------------------------------------------------------------------
def _calcular_sap():
    # Imports locales para evitar importaciones circulares al arrancar Django
    from operaciones import services_sap
    from operaciones.models import RegistroSAP

    E = RegistroSAP.Estado

    # Mismos criterios que la pantalla "Conciliación CICB vs SAP"
    pendientes = services_sap.instalados_sin_sap()
    n_pend = pendientes.count() if hasattr(pendientes, "query") else len(pendientes)

    datos = {
        "pendientes": n_pend,
        "ok": RegistroSAP.objects.filter(estado=E.OK).count(),
        "error": RegistroSAP.objects.filter(estado=E.ERROR, vigente=True).count(),
    }
    datos["total"] = datos["pendientes"] + datos["error"]

    # A dónde lleva el botón: lo más urgente primero
    if datos["error"]:
        datos["destino"] = "ERROR"
    elif datos["pendientes"]:
        datos["destino"] = "PENDIENTE"
    else:
        datos["destino"] = ""

    datos["hay_datos"] = any((datos["pendientes"], datos["ok"], datos["error"]))
    return datos


def sap_burbuja(request):
    if not _puede_ver(getattr(request, "user", None)):
        return {}

    # Ya está viendo la conciliación SAP: no hace falta la burbuja.
    match = getattr(request, "resolver_match", None)
    if match and match.url_name == "conciliacion_sap":
        return {}

    try:
        datos = cache.get_or_set(CLAVE_SAP, _calcular_sap, SEGUNDOS_CACHE)
    except Exception:
        # Un error en la burbuja nunca debe tumbar una página
        log.exception("No se pudo calcular la burbuja de instalación SAP")
        return {}

    return {"burbuja_sap": datos} if datos["hay_datos"] else {}


# ------------------------------------------------------------------
# 3) Almacén: inventario cargado vs sistema
# ------------------------------------------------------------------
def _calcular_almacen():
    # Última carga de cada categoría
    ultimas = {}
    for c in (
        CargaInventario.objects.select_related("categoria")
        .order_by("-fecha", "-pk")[:50]
    ):
        ultimas.setdefault(c.categoria_id, c)

    categorias = [
        {
            "pk": c.pk,
            "categoria": c.categoria.nombre,
            "fecha": c.fecha,
            "ok": c.filas_ok,
            "error": c.filas_error,
            "faltantes": c.faltantes_archivo,
        }
        for c in ultimas.values()
    ]

    if not categorias:
        return {"hay_datos": False}

    total = sum(a["error"] + a["faltantes"] for a in categorias)

    # A dónde lleva el botón: primero errores, luego faltantes, luego lo que coincide
    destino = next(
        ((a["pk"], "errores") for a in categorias if a["error"]),
        None,
    ) or next(
        ((a["pk"], "faltantes") for a in categorias if a["faltantes"]),
        None,
    ) or (categorias[0]["pk"], "ok")

    return {
        "categorias": categorias,
        "total": total,
        "destino_pk": destino[0],
        "destino_ver": destino[1],
        "hay_datos": True,
    }


def almacen_burbuja(request):
    if not _puede_ver(getattr(request, "user", None)):
        return {}

    # Ya está viendo el detalle de una carga: no hace falta la burbuja.
    match = getattr(request, "resolver_match", None)
    if match and match.url_name == "carga_inventario_detalle":
        return {}

    try:
        datos = cache.get_or_set(CLAVE_ALMACEN, _calcular_almacen, SEGUNDOS_CACHE)
    except Exception:
        log.exception("No se pudo calcular la burbuja de almacén")
        return {}

    return {"burbuja_alm": datos} if datos["hay_datos"] else {}

# ------------------------------------------------------------------
# 4) Stock por material: MB52 vs Five
# ------------------------------------------------------------------
def _calcular_stock():
    carga = CargaStock.objects.order_by("-fecha", "-pk").first()
    if not carga:
        return {"hay_datos": False}

    FALTANTE = RegistroStock.TipoError.NO_EN_MB52

    # Mismo criterio que CargaStockDetalleView._base(): solo materiales VANTI
    base = (
        carga.registros
        .annotate(_cod=Upper(Trim("codigo_producto")))
        .exclude(_cod__in=CODIGOS_EXCLUIDOS_VANTI)
    )

    errores = (
        base.filter(resultado=RegistroStock.Resultado.ERROR)
        .exclude(tipo_error=FALTANTE)
    )

    n_ok = base.filter(resultado=RegistroStock.Resultado.OK).count()
    n_error = errores.count()
    n_faltantes = base.filter(tipo_error=FALTANTE).count()
    n_total = base.count()

    etiquetas = dict(RegistroStock.TipoError.choices)
    tipos = [
        {
            "codigo": x["tipo_error"],
            "label": etiquetas.get(x["tipo_error"], x["tipo_error"]),
            "n": x["n"],
        }
        for x in (
            errores.values("tipo_error")
            .annotate(n=Count("id"))
            .order_by("-n")
        )
    ]

    # Lo pendiente = diferencias + faltantes (lo que el detalle muestra como problema)
    total_problemas = n_error + n_faltantes

    return {
        "pk": carga.pk,
        "fecha": carga.fecha,
        "ok": n_ok,
        "error": n_error,
        "faltantes": n_faltantes,
        "total_materiales": n_total,
        "total": total_problemas,
        "tipos": tipos[:4],
        "otros": max(len(tipos) - 4, 0),
        "destino_ver": (
            "errores" if n_error
            else "faltantes" if n_faltantes
            else "ok"
        ),
        "hay_datos": True,
    }

def stock_burbuja(request):
    if not _puede_ver(getattr(request, "user", None)):
        return {}

    # Ya está viendo el detalle de una carga MB52: no hace falta la burbuja.
    match = getattr(request, "resolver_match", None)
    if match and match.url_name == "carga_stock_detalle":
        return {}

    try:
        datos = cache.get_or_set(CLAVE_STOCK, _calcular_stock, SEGUNDOS_CACHE)
    except Exception:
        log.exception("No se pudo calcular la burbuja de stock MB52")
        return {}

    return {"burbuja_stk": datos} if datos["hay_datos"] else {}

def _puede_ver_sin_ods(user):
    # Esta pantalla la usan también los digitadores, no solo almacén
    roles = tuple(ROLES_ALMACEN) + tuple(ROLES_DIGITADOR)
    return bool(
        user and user.is_authenticated
        and (user.is_superuser or any(tiene_rol(user, r) for r in roles))
    )


def _calcular_sin_ods():
    from .sin_ods import con_cantidades, ordenes_sin_ods_qs

    filas = list(
        con_cantidades(ordenes_sin_ods_qs()).values_list(
            "tecnico_id", "tecnico__first_name", "tecnico__last_name",
            "tecnico__username", "piezas",
        )
    )
    if not filas:
        return {"hay_datos": False}

    por_tec = {}
    for tid, nombre, apellido, usuario, _piezas in filas:
        t = por_tec.setdefault(tid, {
            "id": tid, "nombre": f"{nombre} {apellido}".strip() or usuario, "n": 0,
        })
        t["n"] += 1
    tecnicos = sorted(por_tec.values(), key=lambda t: (-t["n"], t["nombre"]))

    return {
        "total": len(filas),
        "piezas": sum(f[4] for f in filas),
        "sin_material": sum(1 for f in filas if not f[4]),
        "tecnicos": tecnicos[:5],
        "otros": max(len(tecnicos) - 5, 0),
        "hay_datos": True,
    }


def sin_ods_burbuja(request):
    if not _puede_ver_sin_ods(getattr(request, "user", None)):
        return {}

    match = getattr(request, "resolver_match", None)
    if match and match.url_name == "ordenes_sin_ods":
        return {}

    try:
                datos = cache.get_or_set(CLAVE_SIN_ODS, _calcular_sin_ods, 5)
    except Exception:
        log.exception("No se pudo calcular la burbuja de órdenes sin ODS")
        return {}

    return {"burbuja_sods": datos} if datos["hay_datos"] else {}