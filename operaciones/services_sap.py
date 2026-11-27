import hashlib
import re
import unicodedata
from collections import defaultdict
from datetime import timedelta
from decimal import Decimal, InvalidOperation

from django.db import transaction
from django.db.models import Exists, OuterRef
from django.utils import timezone
from openpyxl import load_workbook

from almacen.models import Asignacion, Producto
from operaciones.models import CargaSAP, Operacion, RegistroSAP

E = RegistroSAP.Estado

COLUMNAS_CANTIDAD = ("ctd_en_um_entrada", "cantidad")


class ErrorArchivo(Exception):
    pass


# ... CATALOGO, TIPOS_ERROR, ORDEN_SEV, meta(), tipos_por_severidad() (sin cambios) ...

CATALOGO = {
    
    "ORDEN_NO_EJECUTADA": ("Orden no ejecutada en CICB", "ALTA",
        "SAP reporta trabajo en una orden que CICB tiene como no ejecutada, cancelada o anulada. "
        "Revisar la anomalía registrada y corregir donde esté mal."),
    
    "UNIDAD_NO_EXISTE": ("Unidad no existe en CICB","ALTA",
        "Verificar el código de unidad o registrar su ingreso en CICB.",),
    "TIPO_DIFIERE": ("Tipo de operación distinto", "BAJA",
        "Serial, orden y cuenta coinciden. Corregir el tipo donde esté mal."),
    "TIPO_INVALIDO": ("Tipo de operación inválido", "BAJA",
        "Usar Instalación, Retiro o Mantenimiento en el Excel y recargar."),
    "VARIOS_DIFIEREN": ("Varios datos distintos", "MEDIA",
        "Probablemente es otra instalación. Investigar el caso completo."),
    "IDENTIFICADOR_INCORRECTO": ("Identificador equivocado", "MEDIA",
        "Medidores se reportan por número de serie; reguladores y válvulas por código de unidad."),
    "FALTA_CODIGO_UNIDAD": ("Falta el código de unidad", "BAJA",
        "SAP debe enviar el código de unidad (ej. REG-001). Corregir el archivo y recargar."),
    "FALTA_DATO": ("Fila incompleta", "BAJA",
        "Cada fila necesita orden, cuenta, tipo y: serie, código de unidad, o código de producto + cantidad."),
    "FALTA_SERIE": ("Falta el número de serie", "BAJA",
        "SAP debe enviar el serial. Corregir el archivo y recargar."),
    "CANTIDAD_INVALIDA": ("Cantidad inválida", "BAJA",
        "Corregir la cantidad en el Excel (entero positivo)."),
    "DUPLICADA_EN_ARCHIVO": ("Fila repetida en el archivo", "BAJA",
        "Eliminar la fila repetida del Excel."),
    "PRODUCTO_NO_EXISTE": ("Producto no existe en CICB", "MEDIA",
        "Verificar el código. Si es válido, crearlo en el catálogo."),

    # --- Tus casos 1 y 2 y sus variantes ---
    "NO_EXISTE_EN_CICB": ("Ni medidor ni orden existen", "ALTA",
        "SAP instaló algo que CICB no conoce. Verificar físicamente y contra la remisión."),
    "SERIE_NO_EXISTE": ("Medidor no existe en CICB", "ALTA",
        "Contrastar con la remisión. Corregir el serial en SAP o registrar el ingreso en CICB."),
    "INSTALADO_SAP_EN_ALMACEN": ("Instalado en SAP, disponible en CICB", "ALTA",
        "Salió sin asignación. Regularizar en CICB: asignar e instalar."),
    "ASIGNADO_NO_INSTALADO": ("Instalado en SAP, asignado en CICB", "MEDIA",
        "El técnico no reportó. Registrar la instalación en CICB."),
    "ESTADO_INVALIDO": ("Instalado en SAP, dañado/baja en CICB", "ALTA",
        "Revisar con auditoría: se dio de baja algo que sí se usó, o SAP está mal."),
    "NO_INSTALADO": ("Medidor no instalado en CICB", "MEDIA",     # heredado
        "Revisar el estado de la unidad en CICB."),
    "SIN_ORDEN_EN_CICB": ("Instalado sin orden en CICB", "MEDIA",
        "Corregir la asignación en CICB para ligarla a su orden."),

    # --- Tus casos 3 y 4 ---
    "SERIE_DIFIERE": ("Orden y cuenta coinciden, medidor distinto", "MEDIA",
        "Probable typo o cambio de medidor. Confirmar con el técnico y corregir donde esté mal."),
    "ORDEN_DIFIERE": ("Orden distinta", "BAJA",
        "Serial y cuenta coinciden. Corregir la orden en el sistema que esté mal."),
    "CUENTA_DIFIERE": ("Cuenta contrato distinta", "BAJA",
        "Serial y orden coinciden. Corregir la cuenta en el sistema que esté mal."),
    "ORDEN_Y_CUENTA_DIFIEREN": ("Orden y cuenta distintas", "MEDIA",
        "Probablemente es otra instalación. Investigar el caso."),
    "PRODUCTO_DIFIERE": ("Producto distinto al de CICB", "MEDIA",
        "Corregir el código de producto donde esté mal."),
    "SERIE_DUP_CICB": ("Serie repetida en CICB", "ALTA",
        "Hay dos unidades con la misma serie. Corregir el inventario."),

    # --- Productos por cantidad ---
    "ORDEN_NO_EXISTE": ("Orden no existe en CICB", "MEDIA",
        "Registrar la operación o verificar el número de orden."),
    "SIN_MATERIAL_EN_ORDEN": ("Orden sin ese material", "MEDIA",
        "La orden existe pero no tiene ese material instalado en CICB."),
    "CANTIDAD_DIFIERE": ("Cantidad distinta", "MEDIA",
        "Revisar la instalación en CICB o el reporte de SAP."),
}

TIPOS_ERROR = {k: v[0] for k, v in CATALOGO.items()}   # compatibilidad con tus vistas
ORDEN_SEV = {"ALTA": 0, "MEDIA": 1, "BAJA": 2}


def meta(codigo):
    etiqueta, sev, accion = CATALOGO.get(codigo, (codigo, "MEDIA", ""))
    return {"etiqueta": etiqueta, "severidad": sev, "accion": accion}


def tipos_por_severidad(sev):
    return [k for k, v in CATALOGO.items() if v[1] == sev]


# ---------------------------------------------------------------- lectura
def _s(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _col(nombre):
    """'Ctd.en UM entrada' -> 'ctd_en_um_entrada'."""
    t = unicodedata.normalize("NFD", str(nombre or ""))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]+", "_", t).strip("_")


def _num(v):
    """Acepta -1, '-1', '1.234,50', '2-' (signo al final, estilo SAP)."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float, Decimal)):
        return Decimal(str(v))
    t = str(v).strip().replace(" ", "")
    if t.endswith("-"):
        t = "-" + t[:-1]
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return Decimal(t)
    except InvalidOperation:
        return None


def _hash(archivo):
    h = hashlib.sha256()
    archivo.seek(0)
    for chunk in archivo.chunks():
        h.update(chunk)
    archivo.seek(0)
    return h.hexdigest()


def leer_excel(archivo):
    """Devuelve (filas_con_orden, cantidad_de_filas_sin_orden)."""
    try:
        wb = load_workbook(archivo, read_only=True, data_only=True)
    except Exception as exc:
        raise ErrorArchivo("El archivo no es un Excel .xlsx válido.") from exc

    filas = wb.active.iter_rows(values_only=True)
    try:
        encabezado = next(filas)
    except StopIteration:
        raise ErrorArchivo("El archivo está vacío.")

    cols = [_col(c) for c in encabezado]
    col_cant = next((c for c in COLUMNAS_CANTIDAD if c in cols), None)

    faltan = [c for c in ("orden", "material") if c not in cols]
    if not col_cant:
        faltan.append("Ctd.en UM entrada")
    if faltan:
        raise ErrorArchivo("Faltan columnas obligatorias: " + ", ".join(faltan) + ".")

    data, sin_orden = [], 0
    for n, fila in enumerate(filas, start=2):
        if not any(c not in (None, "") for c in fila):
            continue
        d = dict(zip(cols, fila))

        orden = _s(d.get("orden"))
        if not orden or set(orden) == {"0"}:
            sin_orden += 1          # traslados, ajustes, etc.
            continue

        data.append({
            "fila": n,
            "orden": orden,
            "codigo": _s(d.get("material")),
            "descripcion": _s(d.get("descripcion_material")),
            "documento": _s(d.get("documento_material")),
            "cantidad_raw": d.get(col_cant),
        })
    return data, sin_orden


def agrupar(filas):
    """Une las filas por (orden, material). Devuelve (grupos, grupos_neteados_a_cero)."""
    grupos = {}
    for f in filas:
        g = grupos.setdefault((f["orden"], f["codigo"]), {
            "fila": f["fila"], "orden": f["orden"], "codigo": f["codigo"],
            "descripcion": f["descripcion"], "suma": Decimal(0),
            "invalida": "", "documentos": set(), "movimientos": 0,
        })
        g["movimientos"] += 1
        if f["documento"]:
            g["documentos"].add(f["documento"])
        cant = _num(f["cantidad_raw"])
        if cant is None:
            g["invalida"] = _s(f["cantidad_raw"]) or "(vacía)"
        else:
            g["suma"] += cant

    resultado, neteados = [], 0
    for g in grupos.values():
        g["neto"] = -g["suma"]                 # consumo negativo -> instalado positivo
        if not g["invalida"] and g["neto"] == 0:
            neteados += 1                      # consumo + reversión = nada
            continue
        g["documentos"] = sorted(g["documentos"])
        resultado.append(g)
    return resultado, neteados


def _clave(g):
    return hashlib.sha1(f"{g['orden']}|{g['codigo']}".encode()).hexdigest()


def _cantidad_entera(g):
    n = g["neto"]
    return int(n) if n > 0 and n == n.to_integral_value() else None


# ---------------------------------------------------------------- contexto
class _Contexto:
    def __init__(self, grupos):
        ordenes = {g["orden"] for g in grupos}
        codigos = {g["codigo"] for g in grupos if g["codigo"]}

        self.productos = {
            p.codigo_producto: p
            for p in Producto.objects.filter(codigo_producto__in=codigos)
        }
        self.ops = {
            o.numero_orden: o
            for o in Operacion.objects
            .select_related("anomalia", "tecnico")
            .filter(numero_orden__in=ordenes)
        }

        # Instalado en CICB por (orden, producto): medidores, reguladores y cantidades
        self.five = defaultdict(int)
        EA = Asignacion.Estado
        for a in (
            Asignacion.objects
            .filter(operacion__numero_orden__in=ordenes)
            .exclude(estado=EA.DEVUELTA)
            .select_related("operacion")
        ):
            q = a.cantidad if a.estado == EA.INSTALADA else a.cantidad_instalada
            self.five[(a.operacion.numero_orden, a.producto_id)] += q or 0


# ---------------------------------------------------------------- conciliación
def _item(label, five, sap, ok):
    return {"label": label, "five": five or "", "sap": sap or "", "ok": bool(ok)}


def _res(estado, tipo, diag, comp, op=None, uni=None, prod=None):
    return {"estado": estado, "tipo": tipo, "diag": diag,
            "comp": comp, "op": op, "uni": uni, "prod": prod}


def conciliar(g, ctx):
    ERR = E.ERROR
    orden, cod = g["orden"], g["codigo"]
    op = ctx.ops.get(orden)

    if not cod:
        return _res(ERR, "FALTA_DATO", "La fila trae orden pero no código de material.", {}, op)

    prod = ctx.productos.get(cod)
    nombre = g["descripcion"] or cod

    if g["invalida"]:
        return _res(ERR, "CANTIDAD_INVALIDA",
                    f"La cantidad '{g['invalida']}' de {nombre} no es un número válido.",
                    {}, op, None, prod)

    cant = _cantidad_entera(g)
    if cant is None:
        return _res(ERR, "CANTIDAD_INVALIDA",
                    f"Los movimientos de {nombre} en la orden {orden} suman {g['neto']} "
                    "(debe ser un entero positivo de consumo).", {}, op, None, prod)

    if prod is None:
        return _res(ERR, "PRODUCTO_NO_EXISTE",
                    f"El material {cod} ({g['descripcion']}) no existe en CICB. "
                    "Verifique que el código del producto sea el mismo de SAP.", {}, op)

    if op is None:
        return _res(ERR, "ORDEN_NO_EXISTE",
                    f"La orden {orden} no existe en CICB.", {}, None, None, prod)

    if op.estado in Operacion.ESTADOS_NEGATIVOS:
        anom = f" Anomalía: {op.anomalia}." if op.anomalia_id else ""
        return _res(ERR, "ORDEN_NO_EJECUTADA",
                    f"SAP consumió material en la orden {orden}, pero en CICB está "
                    f"{op.get_estado_display().lower()}.{anom}", {}, op, None, prod)

    five = ctx.five.get((orden, prod.pk), 0)
    comp = {
        "orden": _item("Orden", op.numero_orden, orden, True),
        "material": _item("Material", prod.codigo_producto, cod, True),
        "cantidad": _item("Cantidad", five, cant, five == cant),
    }

    if five == 0:
        return _res(ERR, "SIN_MATERIAL_EN_ORDEN",
                    f"SAP consumió {cant} × {prod.nombre_completo} en la orden {orden}, "
                    "pero CICB no tiene ese material instalado en ella.",
                    comp, op, None, prod)

    if five != cant:
        return _res(ERR, "CANTIDAD_DIFIERE",
                    f"{prod.nombre_completo} en la orden {orden}: "
                    f"CICB={five} / SAP={cant}.", comp, op, None, prod)

    return _res(E.OK, "", "Todo coincide.", comp, op, None, prod)


# ---------------------------------------------------------------- vigencia
def _quien(previo, nuevo):
    sap = previo.cantidad != nuevo.cantidad
    a = (previo.comparacion or {}).get("cantidad", {}).get("five")
    b = (nuevo.comparacion or {}).get("cantidad", {}).get("five")
    five = a != b
    if sap and five:
        return "AMBOS"
    return "SAP" if sap else "CICB"


def _actualizar_vigencia(carga, registros):
    """Un error anterior del mismo (orden, material) deja de ser vigente;
    si ahora está OK, queda resuelto."""
    nuevos = {r.clave_unica: r for r in registros if r.estado != E.IGNORADO}
    if not nuevos:
        return

    tocados = []
    for ant in (
        RegistroSAP.objects
        .filter(estado=E.ERROR, vigente=True, clave_unica__in=nuevos)
        .exclude(carga=carga)
    ):
        nuevo = nuevos[ant.clave_unica]
        ant.vigente = False
        if nuevo.estado == E.OK:
            ant.resuelto = True
            ant.resuelto_en = carga
            ant.resuelto_por = _quien(ant, nuevo)
        tocados.append(ant)

    RegistroSAP.objects.bulk_update(
        tocados, ["vigente", "resuelto", "resuelto_en", "resuelto_por"])


# ---------------------------------------------------------------- carga
@transaction.atomic
def procesar_carga(archivo, usuario):
    filas, sin_orden = leer_excel(archivo)
    grupos, neteados = agrupar(filas)
    if not grupos:
        raise ErrorArchivo("El archivo no tiene filas con número de orden.")

    ctx = _Contexto(grupos)
    carga = CargaSAP.objects.create(
        archivo=archivo, usuario=usuario, hash_archivo=_hash(archivo))

    claves = {_clave(g) for g in grupos}
    previas = {
        r.clave_unica: r
        for r in RegistroSAP.objects
        .filter(estado=E.OK, clave_unica__in=claves).order_by("pk")
    }

    registros, reemplazadas = [], []
    ok = err = ign = 0

    for g in grupos:
        clave = _clave(g)
        cant = _cantidad_entera(g)
        previa = previas.get(clave)

        reg = RegistroSAP(
            carga=carga, fila=g["fila"],
            numero_orden=g["orden"], codigo_producto=g["codigo"],
            cantidad=cant or 0,
            cantidad_sap_original=format(g["neto"].normalize(), "f"),
            clave_unica=clave,
        )

        if previa and previa.cantidad == cant:
            # Mismo consumo ya conciliado en otra carga
            reg.estado = E.IGNORADO
            reg.registro_original_id = previa.pk
            reg.operacion_id = previa.operacion_id
            reg.tecnico_id = previa.tecnico_id
            reg.producto_id = previa.producto_id
            reg.cuenta_contrato = previa.cuenta_contrato
            reg.comparacion = previa.comparacion
            reg.diagnostico = "Ya fue cargada y conciliada antes; no se vuelve a procesar."
        else:
            if previa:                       # cambió la cantidad en SAP
                reemplazadas.append(previa)
            r = conciliar(g, ctx)
            reg.estado, reg.tipo_error, reg.diagnostico = r["estado"], r["tipo"], r["diag"]
            reg.comparacion = {**r["comp"], "documentos": g["documentos"]}
            reg.operacion, reg.producto = r["op"], r["prod"]
            reg.tecnico = r["op"].tecnico if r["op"] else None
            reg.cuenta_contrato = r["op"].numero_cuenta_contrato if r["op"] else ""

        reg.error = reg.diagnostico if reg.estado == E.ERROR else ""
        reg.vigente = reg.estado == E.ERROR
        registros.append(reg)

        if reg.estado == E.OK:
            ok += 1
        elif reg.estado == E.ERROR:
            err += 1
        else:
            ign += 1

    RegistroSAP.objects.bulk_create(registros, batch_size=500)

    # El OK anterior con otra cantidad se reemplaza (evita sumarlo dos veces)
    for p in reemplazadas:
        p.estado = E.IGNORADO
        p.vigente = False
        p.diagnostico = (f"{p.diagnostico} [Reemplazado por la carga #{carga.pk}: "
                         "cambió la cantidad en SAP.]").strip()
    RegistroSAP.objects.bulk_update(reemplazadas, ["estado", "vigente", "diagnostico"])

    _actualizar_vigencia(carga, registros)

    carga.total_filas, carga.filas_ok = len(registros), ok
    carga.filas_error, carga.filas_ignoradas = err, ign
    carga.save(update_fields=["total_filas", "filas_ok", "filas_error", "filas_ignoradas"])

    carga.filas_sin_orden = sin_orden        # solo informativo, no se guarda
    carga.grupos_neteados = neteados
    return carga


# ---------------------------------------------------------------- pendientes
DIAS_GRACIA_SAP = 3


def instalados_sin_sap():
    """Material instalado en CICB (orden + producto) que ninguna carga ha conciliado OK."""
    reportado = RegistroSAP.objects.filter(
        estado=E.OK,
        operacion_id=OuterRef("operacion_id"),
        producto_id=OuterRef("producto_id"),
    )
    return (
        Asignacion.objects
        .filter(estado=Asignacion.Estado.INSTALADA, operacion__isnull=False)
        .annotate(reportado=Exists(reportado))
        .filter(reportado=False)
        .select_related("producto", "unidad_serializada", "operacion", "tecnico")
        .order_by("fecha_instalacion")
    )


def limite_gracia():
    return timezone.localdate() - timedelta(days=DIAS_GRACIA_SAP)
