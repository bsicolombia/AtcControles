"""
operaciones/resumen.py

Resumen de ejecución por técnico: año -> mes -> día -> órdenes.
Para cada técnico y período cuenta:
    ej  = ejecutadas      (estados que instalan: EJECUTADA, EFECTIVA)
    no  = no ejecutadas   (estados negativos: NO_EJECUTADA, CANCELADA, ANULADA)
    tot = total de órdenes (incluye también CREADA, PENDIENTE y ASIGNADA)

Los estados salen del propio modelo (Operacion.ESTADOS_QUE_INSTALAN y
Operacion.ESTADOS_NEGATIVOS), así que si cambia el modelo no hay que tocar nada aquí.
"""
import calendar

from django.db.models import Count, Q
from django.db.models.functions import ExtractDay, ExtractMonth

from .models import Operacion

# ---------------------------------------------------------------------
# Estados (tomados del modelo)
# ---------------------------------------------------------------------
ESTADOS_EJECUTADOS = sorted(e.value for e in Operacion.ESTADOS_QUE_INSTALAN)
ESTADOS_NO_EJECUTADOS = sorted(e.value for e in Operacion.ESTADOS_NEGATIVOS)

# Campo por el que se agrupa por día y mes. Es el mismo que muestra el listado.
CAMPO_FECHA = "fecha_operacion"

MESES = [
    (1, "Enero"), (2, "Febrero"), (3, "Marzo"), (4, "Abril"),
    (5, "Mayo"), (6, "Junio"), (7, "Julio"), (8, "Agosto"),
    (9, "Septiembre"), (10, "Octubre"), (11, "Noviembre"), (12, "Diciembre"),
]
MESES_CORTO = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
               "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

RESULTADOS_VALIDOS = ("EJECUTADAS", "NO_EJECUTADAS", "PENDIENTES")


# ---------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------
def url_con(g, **cambios):
    """Devuelve '?...' conservando los filtros actuales y aplicando cambios.
    Un valor None o '' elimina ese parámetro."""
    p = g.copy()
    p.pop("page", None)
    p.pop("detalle", None)
    for k, v in cambios.items():
        if v in (None, ""):
            p.pop(k, None)
        else:
            p[k] = str(v)
    q = p.urlencode()
    return "?" + q if q else "?"


def filtrar_resultado(qs, resultado):
    """Aplica el filtro 'Resultado' al listado de órdenes."""
    if resultado == "EJECUTADAS":
        return qs.filter(estado__in=ESTADOS_EJECUTADOS)
    if resultado == "NO_EJECUTADAS":
        return qs.filter(estado__in=ESTADOS_NO_EJECUTADOS)
    if resultado == "PENDIENTES":
        return qs.exclude(estado__in=ESTADOS_EJECUTADOS + ESTADOS_NO_EJECUTADOS)
    return qs


# ---------------------------------------------------------------------
# Resumen
# ---------------------------------------------------------------------
def resumen_ejecucion(base, g, anio, mes):
    """
    base : queryset de Operacion con búsqueda, estado, tipo, técnico y anomalía
           YA aplicados, pero SIN fechas, SIN 'resultado' y SIN annotate
           de agregados (por ejemplo, antes de anotar n_inc).
    g    : request.GET
    anio : int
    mes  : int o None
           None -> matriz técnico x meses del año
           1-12 -> matriz técnico x días de ese mes
    """
    qs = base.filter(**{f"{CAMPO_FECHA}__year": anio})

    if mes:
        qs = qs.filter(**{f"{CAMPO_FECHA}__month": mes})
        campo = ExtractDay(CAMPO_FECHA)
        claves = list(range(1, calendar.monthrange(anio, mes)[1] + 1))
        columnas = [
            {
                "label": str(d),
                "url": url_con(g, anio=anio, mes=mes, fecha=f"{anio}-{mes:02d}-{d:02d}"),
            }
            for d in claves
        ]
        nivel = "dia"
        titulo = f"{dict(MESES)[mes]} {anio}"
        url_volver = url_con(g, anio=anio, mes=None, fecha=None)
    else:
        campo = ExtractMonth(CAMPO_FECHA)
        claves = list(range(1, 13))
        columnas = [
            {"label": MESES_CORTO[m - 1], "url": url_con(g, anio=anio, mes=m, fecha=None)}
            for m in claves
        ]
        nivel = "mes"
        titulo = str(anio)
        url_volver = ""

    datos = (
        qs.annotate(k=campo)
        .values("tecnico_id", "tecnico__first_name", "tecnico__last_name",
                "tecnico__username", "k")
        .annotate(
            tot=Count("id"),
            ej=Count("id", filter=Q(estado__in=ESTADOS_EJECUTADOS)),
            no=Count("id", filter=Q(estado__in=ESTADOS_NO_EJECUTADOS)),
        )
        .order_by("tecnico__first_name", "tecnico__username", "k")
    )

    por_tecnico = {}
    for d in datos:
        nombre = (
            f"{d['tecnico__first_name'] or ''} {d['tecnico__last_name'] or ''}".strip()
            or d["tecnico__username"]
            or "Sin técnico"
        )
        t = por_tecnico.setdefault(d["tecnico_id"], {"nombre": nombre, "cuentas": {}})
        t["cuentas"][d["k"]] = (d["ej"], d["no"], d["tot"])

    filas = []
    totales = [{"ej": 0, "no": 0, "tot": 0} for _ in claves]

    for tid, t in por_tecnico.items():
        celdas = []
        for i, k in enumerate(claves):
            ej, no, tot = t["cuentas"].get(k, (0, 0, 0))
            totales[i]["ej"] += ej
            totales[i]["no"] += no
            totales[i]["tot"] += tot

            if nivel == "mes":
                kw = dict(tecnico=tid, anio=anio, mes=k, fecha=None)
            else:
                kw = dict(tecnico=tid, anio=anio, mes=mes,
                          fecha=f"{anio}-{mes:02d}-{k:02d}")

            celdas.append({
                "ej": ej,
                "no": no,
                "tot": tot,
                "url_ej": url_con(g, resultado="EJECUTADAS", **kw) if ej else "",
                "url_no": url_con(g, resultado="NO_EJECUTADAS", **kw) if no else "",
                "url_tot": url_con(g, resultado=None, **kw) if tot else "",
            })

        total_ej = sum(c["ej"] for c in celdas)
        total_no = sum(c["no"] for c in celdas)
        total_tot = sum(c["tot"] for c in celdas)
        kw_tot = dict(tecnico=tid, anio=anio, mes=mes or None, fecha=None)

        filas.append({
            "nombre": t["nombre"],
            "celdas": celdas,
            "total_ej": total_ej,
            "total_no": total_no,
            "total_tot": total_tot,
            "url_ej": url_con(g, resultado="EJECUTADAS", **kw_tot) if total_ej else "",
            "url_no": url_con(g, resultado="NO_EJECUTADAS", **kw_tot) if total_no else "",
            "url_tot": url_con(g, resultado=None, **kw_tot) if total_tot else "",
        })

    return {
        "nivel": nivel,
        "anio": anio,
        "titulo": titulo,
        "url_volver": url_volver,
        "columnas": columnas,
        "filas": filas,
        "totales": totales,
        "total_ej": sum(t["ej"] for t in totales),
        "total_no": sum(t["no"] for t in totales),
        "total_tot": sum(t["tot"] for t in totales),
    }