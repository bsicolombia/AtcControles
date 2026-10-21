import hashlib
from collections import defaultdict
from django.db import transaction
from openpyxl import load_workbook
from django.db.models import Q
from almacen.models import ( 
    Asignacion, Producto, UnidadSerializada,
)
import difflib
from operaciones.models import Operacion, CargaSAP, RegistroSAP
from django.utils import timezone
from datetime import timedelta
import unicodedata

OBLIGATORIAS = (
    "numero_orden",
    "cuenta_contrato",
    "tipo_operacion",
)

class ErrorArchivo(Exception):
    pass

def _quien(previo, nuevo):
    """¿Quién corrigió? Si cambió la fila de SAP, la clave cambia; si no, se arregló Five."""
    sap = previo.clave_unica != nuevo.clave_unica
    five = any(
        (previo.comparacion or {}).get(k, {}).get("five") != v.get("five")
        for k, v in (nuevo.comparacion or {}).items()
    )
    return "AMBOS" if sap and five else "SAP" if sap else "FIVE"


def _actualizar_vigencia(carga, registros):
    """Cierra o reemplaza los errores abiertos de cargas anteriores que esta carga vuelve a tocar."""
    E = RegistroSAP.Estado
    series = {r.numero_serie for r in registros if r.numero_serie}
    ordenes = {r.numero_orden for r in registros if r.numero_orden}
    codigos = {r.codigo_unidad for r in registros if r.codigo_unidad}
    claves = {r.clave_unica for r in registros}

    abiertos = list(
        RegistroSAP.objects
        .filter(estado=E.ERROR, vigente=True)
        .exclude(carga=carga)
        .filter(
            Q(numero_serie__in=series)
            | Q(numero_orden__in=ordenes)
            | Q(codigo_unidad__in=codigos)
            | Q(clave_unica__in=claves)
        )
    )
    if not abiertos:
        return

    def llaves(r):
        ks = [
            ("c", r.clave_unica),
            ("oc", r.numero_orden),
        ]

        if r.numero_serie:
            ks.append(("s", r.numero_serie))

        if r.codigo_unidad:
            ks.append(("u", r.codigo_unidad))

        return ks

    indice = defaultdict(list)
    for p in abiertos:
        for k in llaves(p):
            indice[k].append(p)

    tocados = {}
    # Primero las filas OK, para que "resuelto" gane sobre "reemplazado por otro error"
    for n in sorted(registros, key=lambda r: r.estado != E.OK):
        if n.estado == E.IGNORADO:
            continue
        for k in llaves(n):
            for p in indice.get(k, []):
                tocados.setdefault(p.pk, (p, n))

    for p, n in tocados.values():
        p.vigente = False
        if n.estado == E.OK:
            p.resuelto, p.resuelto_en, p.resuelto_por = True, carga, _quien(p, n)

    RegistroSAP.objects.bulk_update(
        [p for p, _ in tocados.values()],
        ["vigente", "resuelto", "resuelto_en", "resuelto_por"],
    )
CATALOGO = {
    
    "UNIDAD_NO_EXISTE": ("Unidad no existe en Five","ALTA",
        "Verificar el código de unidad o registrar su ingreso en Five.",),
    "TIPO_DIFIERE": ("Tipo de operación distinto", "BAJA",
        "Serial, orden y cuenta coinciden. Corregir el tipo donde esté mal."),
    "TIPO_INVALIDO": ("Tipo de operación inválido", "BAJA",
        "Usar Instalación, Retiro o Mantenimiento en el Excel y recargar."),
    "VARIOS_DIFIEREN": ("Varios datos distintos", "MEDIA",
        "Probablemente es otra instalación. Investigar el caso completo."),
    # codigo: (etiqueta, severidad, acción sugerida)
    "FALTA_DATO": ("Fila incompleta", "BAJA",
        "Completar orden y código de producto en el Excel y recargar."),
    "FALTA_SERIE": ("Falta el número de serie", "BAJA",
        "SAP debe enviar el serial. Corregir el archivo y recargar."),
    "CANTIDAD_INVALIDA": ("Cantidad inválida", "BAJA",
        "Corregir la cantidad en el Excel (entero positivo)."),
    "DUPLICADA_EN_ARCHIVO": ("Fila repetida en el archivo", "BAJA",
        "Eliminar la fila repetida del Excel."),
    "PRODUCTO_NO_EXISTE": ("Producto no existe en Five", "MEDIA",
        "Verificar el código. Si es válido, crearlo en el catálogo."),

    # --- Tus casos 1 y 2 y sus variantes ---
    "NO_EXISTE_EN_FIVE": ("Ni medidor ni orden existen", "ALTA",
        "SAP instaló algo que Five no conoce. Verificar físicamente y contra la remisión."),
    "SERIE_NO_EXISTE": ("Medidor no existe en Five", "ALTA",
        "Contrastar con la remisión. Corregir el serial en SAP o registrar el ingreso en Five."),
    "INSTALADO_SAP_EN_ALMACEN": ("Instalado en SAP, disponible en Five", "ALTA",
        "Salió sin asignación. Regularizar en Five: asignar e instalar."),
    "ASIGNADO_NO_INSTALADO": ("Instalado en SAP, asignado en Five", "MEDIA",
        "El técnico no reportó. Registrar la instalación en Five."),
    "ESTADO_INVALIDO": ("Instalado en SAP, dañado/baja en Five", "ALTA",
        "Revisar con auditoría: se dio de baja algo que sí se usó, o SAP está mal."),
    "NO_INSTALADO": ("Medidor no instalado en Five", "MEDIA",     # heredado
        "Revisar el estado de la unidad en Five."),
    "SIN_ORDEN_EN_FIVE": ("Instalado sin orden en Five", "MEDIA",
        "Corregir la asignación en Five para ligarla a su orden."),

    # --- Tus casos 3 y 4 ---
    "SERIE_DIFIERE": ("Orden y cuenta coinciden, medidor distinto", "MEDIA",
        "Probable typo o cambio de medidor. Confirmar con el técnico y corregir donde esté mal."),
    "ORDEN_DIFIERE": ("Orden distinta", "BAJA",
        "Serial y cuenta coinciden. Corregir la orden en el sistema que esté mal."),
    "CUENTA_DIFIERE": ("Cuenta contrato distinta", "BAJA",
        "Serial y orden coinciden. Corregir la cuenta en el sistema que esté mal."),
    "ORDEN_Y_CUENTA_DIFIEREN": ("Orden y cuenta distintas", "MEDIA",
        "Probablemente es otra instalación. Investigar el caso."),
    "PRODUCTO_DIFIERE": ("Producto distinto al de Five", "MEDIA",
        "Corregir el código de producto donde esté mal."),
    "SERIE_DUP_FIVE": ("Serie repetida en Five", "ALTA",
        "Hay dos unidades con la misma serie. Corregir el inventario."),

    # --- Productos por cantidad ---
    "ORDEN_NO_EXISTE": ("Orden no existe en Five", "MEDIA",
        "Registrar la operación o verificar el número de orden."),
    "SIN_MATERIAL_EN_ORDEN": ("Orden sin ese material", "MEDIA",
        "La orden existe pero no tiene ese material instalado en Five."),
    "CANTIDAD_DIFIERE": ("Cantidad distinta", "MEDIA",
        "Revisar la instalación en Five o el reporte de SAP."),
}

TIPOS_ERROR = {k: v[0] for k, v in CATALOGO.items()}   # compatibilidad con tus vistas
ORDEN_SEV = {"ALTA": 0, "MEDIA": 1, "BAJA": 2}


def meta(codigo):
    etiqueta, sev, accion = CATALOGO.get(codigo, (codigo, "MEDIA", ""))
    return {"etiqueta": etiqueta, "severidad": sev, "accion": accion}


def tipos_por_severidad(sev):
    return [k for k, v in CATALOGO.items() if v[1] == sev]


def _tipo(raw):
    """'Instalación ' -> 'INSTALACION'. Devuelve None si no es un tipo válido."""
    t = unicodedata.normalize("NFD", raw or "")
    t = "".join(c for c in t if unicodedata.category(c) != "Mn").strip().upper()
    return t if t in Operacion.TipoOperacion.values else None

# ---------------------------------------------------------------- lectura
def _s(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _hash(archivo):
    h = hashlib.sha256()
    archivo.seek(0)
    for chunk in archivo.chunks():
        h.update(chunk)
    archivo.seek(0)
    return h.hexdigest()


def leer_excel(archivo):
    try:
        wb = load_workbook(archivo, read_only=True, data_only=True)
    except Exception as exc:
        raise ErrorArchivo("El archivo no es un Excel .xlsx válido.") from exc
    filas = wb.active.iter_rows(values_only=True)
    try:
        encabezado = next(filas)
    except StopIteration:
        raise ErrorArchivo("El archivo está vacío.")

    cols = [_s(c).lower().replace(" ", "_") for c in encabezado]
    cols = ["cuenta_contrato" if c == "numero_cuenta_contrato" else c for c in cols]
    faltan = [c for c in OBLIGATORIAS if c not in cols]
    if faltan:
        raise ErrorArchivo("Faltan columnas obligatorias: " + ", ".join(faltan) + ".")

    data = []
    for n, fila in enumerate(filas, start=2):
        if not any(c not in (None, "") for c in fila):
            continue
        d = dict(zip(cols, fila))
        data.append({
            "fila": n,
            "orden": _s(d.get("numero_orden")),
            "cuenta": _s(d.get("cuenta_contrato")),
            "tipo_raw": _s(d.get("tipo_operacion")),
            "serie": _s(d.get("numero_serie")),
            "codigo_unidad": _s(d.get("codigo_unidad")),
            "cantidad_raw": _s(d.get("cantidad")),
        })
    return data


def _cantidad(raw):
    if raw == "":
        return 1
    try:
        n = float(raw.replace(",", "."))
        return int(n) if n.is_integer() and n >= 0 else None
    except ValueError:
        return None


def _clave(f):
    base = (
        f"{f['orden']}|{f['cuenta']}|"
        f"{_tipo(f['tipo_raw']) or f['tipo_raw']}|"
        f"{f['serie']}|{f['codigo_unidad']}"
    )

    if not f["serie"] and not f["codigo_unidad"]:
        base += f"|{f['cantidad_raw']}"

    return hashlib.sha1(base.encode()).hexdigest()


# ---------------------------------------------------------------- contexto
class _Contexto:
    def __init__(self, filas):
        codigos_unidad = {
            f["codigo_unidad"]
            for f in filas
            if f["codigo_unidad"]
        }
        series = {f["serie"] for f in filas if f["serie"]}
        ordenes = {f["orden"] for f in filas if f["orden"]}
        E = Asignacion.Estado

        self.ops = {o.numero_orden: o
                    for o in Operacion.objects.filter(numero_orden__in=ordenes)}
        self.unidades = defaultdict(list)
        for u in UnidadSerializada.objects.filter(
            numero_serie__in=series
        ).select_related("producto"):
            self.unidades[u.numero_serie].append(u)
        
        self.unidades_codigo = {
            u.codigo: u
            for u in UnidadSerializada.objects.filter(
                codigo__in=codigos_unidad
            ).select_related("producto")
        }

        self.asig_unidad = {}
        self.series_orden = defaultdict(list)
        asigs = (
            Asignacion.objects
            .filter(unidad_serializada__isnull=False)
            .filter(Q(unidad_serializada__numero_serie__in=series)
                    | Q(operacion__numero_orden__in=ordenes))
            .exclude(estado=E.DEVUELTA)
            .select_related("operacion", "unidad_serializada")
            .order_by("id")
        )
        for a in asigs:
            prev = self.asig_unidad.get(a.unidad_serializada_id)
            if prev is None or not (prev.estado == E.INSTALADA and a.estado != E.INSTALADA):
                self.asig_unidad[a.unidad_serializada_id] = a
            if a.operacion_id and a.unidad_serializada.numero_serie:
                self.series_orden[a.operacion.numero_orden].append(
                    a.unidad_serializada.numero_serie)

        # Por cantidad: total instalado en Five por orden (todos los productos sin serie)
        self.cantidad = {}
        for a in (
            Asignacion.objects
            .filter(unidad_serializada__isnull=True, operacion__numero_orden__in=ordenes)
            .exclude(estado=E.DEVUELTA)
            .select_related("operacion")
        ):
            q = a.cantidad if a.estado == E.INSTALADA else a.cantidad_instalada
            k = a.operacion.numero_orden
            self.cantidad[k] = self.cantidad.get(k, 0) + (q or 0)


# ---------------------------------------------------------------- conciliación
def _item(label, five, sap, ok):
    return {"label": label, "five": five or "", "sap": sap or "", "ok": bool(ok)}


def _res(estado, tipo, diag, comp, op=None, uni=None, prod=None):
    return {"estado": estado, "tipo": tipo, "diag": diag,
            "comp": comp, "op": op, "uni": uni, "prod": prod}


def _tipo_final(malos):
    if len(malos) >= 2:
        return "VARIOS_DIFIEREN"
    return {"orden": "ORDEN_DIFIERE", "cuenta_contrato": "CUENTA_DIFIERE",
            "tipo_operacion": "TIPO_DIFIERE", "cantidad": "CANTIDAD_DIFIERE"}[malos[0]]


def _detalle(comp, malos):
    return "; ".join(
        f"{comp[k]['label'].lower()}: Five={comp[k]['five'] or '—'} / SAP={comp[k]['sap'] or '—'}"
        for k in malos)
def _conciliar_codigo_unidad(f, ctx, tipo):
    codigo = f["codigo_unidad"]
    orden = f["orden"]
    cuenta = f["cuenta"]

    ERR = RegistroSAP.Estado.ERROR
    E_U = UnidadSerializada.Estado

    unidad = ctx.unidades_codigo.get(codigo)
    op_sap = ctx.ops.get(orden)

    if unidad is None:
        return _res(
            ERR,
            "UNIDAD_NO_EXISTE",
            f"El código de unidad {codigo} no existe en Five.",
            {},
            op_sap,
        )

    asignaciones = (
        unidad.asignaciones
        .select_related("operacion", "tecnico")
        .exclude(estado=Asignacion.Estado.DEVUELTA)
        .order_by("id")
    )

    asig = None
    for a in asignaciones:
        if asig is None or not (
            asig.estado == Asignacion.Estado.INSTALADA
            and a.estado != Asignacion.Estado.INSTALADA
        ):
            asig = a

    op_five = asig.operacion if asig and asig.operacion_id else None
    instalado = unidad.estado == E_U.INSTALADO

    ref = op_five or op_sap
    f_ord = ref.numero_orden if ref else ""
    f_cta = ref.numero_cuenta_contrato if ref else ""
    f_tip = ref.get_tipo_operacion_display() if ref else ""

    comp = {
        "unidad": _item(
            "Código de unidad",
            unidad.codigo,
            codigo,
            True,
        ),
        "orden": _item(
            "Orden", f_ord, orden, f_ord == orden
        ),
        "cuenta_contrato": _item(
            "Cuenta contrato", f_cta, cuenta, f_cta == cuenta
        ),
        "tipo_operacion": _item(
            "Tipo de operación",
            f_tip,
            Operacion.TipoOperacion(tipo).label,
            ref is not None and ref.tipo_operacion == tipo,
        ),
    }

    if not instalado:
        if unidad.estado == E_U.DISPONIBLE:
            error = "INSTALADO_SAP_EN_ALMACEN"
        elif unidad.estado == E_U.ASIGNADO:
            error = "ASIGNADO_NO_INSTALADO"
        else:
            error = "ESTADO_INVALIDO"

        return _res(
            ERR,
            error,
            f"SAP reporta la unidad {codigo} instalada, "
            f"pero en Five está {unidad.get_estado_display()}.",
            comp,
            op_five or op_sap,
            unidad,
            unidad.producto,
        )

    if op_five is None:
        return _res(
            ERR,
            "SIN_ORDEN_EN_FIVE",
            f"La unidad {codigo} está instalada en Five "
            f"sin orden asociada.",
            comp,
            None,
            unidad,
            unidad.producto,
        )

    malos = [
        k for k in ("orden", "cuenta_contrato", "tipo_operacion")
        if not comp[k]["ok"]
    ]

    if malos:
        return _res(
            ERR,
            _tipo_final(malos),
            f"La unidad {codigo} difiere en {_detalle(comp, malos)}.",
            comp,
            op_five,
            unidad,
            unidad.producto,
        )

    return _res(
        RegistroSAP.Estado.OK,
        "",
        "Todo coincide.",
        comp,
        op_five,
        unidad,
        unidad.producto,
    )

def _conciliar_serie(f, ctx, tipo):
    serie, orden, cuenta = f["serie"], f["orden"], f["cuenta"]
    ERR, E_U = RegistroSAP.Estado.ERROR, UnidadSerializada.Estado
    unidades = ctx.unidades.get(serie, [])
    op_sap = ctx.ops.get(orden)

    if len(unidades) > 1:
        cods = ", ".join(u.codigo for u in unidades)
        return _res(ERR, "SERIE_DUP_FIVE",
                    f"La serie {serie} está repetida {len(unidades)} veces en Five "
                    f"(códigos {cods}). Corrija el inventario.", {}, op_sap)

    unidad = unidades[0] if unidades else None
    asig = ctx.asig_unidad.get(unidad.pk) if unidad else None
    op_five = asig.operacion if asig and asig.operacion_id else None
    prod = unidad.producto if unidad else None
    instalado = bool(unidad and unidad.estado == E_U.INSTALADO)

    cerca = difflib.get_close_matches(serie, ctx.series_orden.get(orden, []), n=1, cutoff=0.75)
    pista = f" ¿Quiso decir {cerca[0]}?" if cerca else ""

    ref = op_five if unidad else op_sap
    f_med = unidad.numero_serie if unidad else ", ".join(ctx.series_orden.get(orden, []))
    f_ord = ref.numero_orden if ref else ""
    f_cta = ref.numero_cuenta_contrato if ref else ""
    f_tip = ref.get_tipo_operacion_display() if ref else ""
    tipo_txt = Operacion.TipoOperacion(tipo).label

    comp = {
        "medidor": _item("Medidor", f_med, serie, instalado),
        "orden": _item("Orden", f_ord, orden, f_ord and f_ord == orden),
        "cuenta_contrato": _item("Cuenta contrato", f_cta, cuenta, f_cta and f_cta == cuenta),
        "tipo_operacion": _item("Tipo de operación", f_tip, tipo_txt,
                                ref is not None and ref.tipo_operacion == tipo),
    }

    # ---- el medidor no existe en Five ----
    if unidad is None:
        if op_sap is None:
            return _res(ERR, "NO_EXISTE_EN_FIVE",
                        f"Ni el medidor {serie} ni la orden {orden} existen en Five.", comp)
        if op_sap.numero_cuenta_contrato == cuenta:
            return _res(ERR, "SERIE_DIFIERE",
                        f"La orden {orden} y la cuenta {cuenta} coinciden con Five, pero el "
                        f"medidor no: SAP reporta {serie} y Five tiene "
                        f"{f_med or 'ningún medidor'} en esa orden.{pista}", comp, op_sap)
        return _res(ERR, "SERIE_NO_EXISTE",
                    f"El medidor {serie} no existe en Five. La orden {orden} sí existe, pero "
                    f"con otra cuenta (Five: {op_sap.numero_cuenta_contrato}, "
                    f"SAP: {cuenta}).{pista}", comp, op_sap)

    # ---- el medidor existe, pero no está instalado ----
    if not instalado:
        if unidad.estado == E_U.DISPONIBLE:
            t, d = "INSTALADO_SAP_EN_ALMACEN", (
                f"SAP reporta el medidor {serie} instalado, pero en Five está DISPONIBLE "
                f"en almacén (nunca fue asignado).")
        elif unidad.estado == E_U.ASIGNADO:
            t, d = "ASIGNADO_NO_INSTALADO", (
                f"SAP reporta el medidor {serie} instalado, pero en Five sigue asignado "
                f"al técnico sin registrar la instalación.")
        else:
            t, d = "ESTADO_INVALIDO", (
                f"SAP reporta el medidor {serie} instalado, pero en Five está "
                f"'{unidad.get_estado_display()}'.")
        return _res(ERR, t, d, comp, op_five, unidad, prod)

    if op_five is None:
        return _res(ERR, "SIN_ORDEN_EN_FIVE",
                    f"El medidor {serie} figura instalado en Five pero sin orden asociada. "
                    f"SAP lo reporta en la orden {orden}.", comp, None, unidad, prod)

    # ---- instalado: comparar orden, cuenta y tipo ----
    malos = [k for k in ("orden", "cuenta_contrato", "tipo_operacion") if not comp[k]["ok"]]
    if not malos:
        return _res(RegistroSAP.Estado.OK, "", "Todo coincide.", comp, op_five, unidad, prod)
    return _res(ERR, _tipo_final(malos),
                f"El medidor {serie} coincide, pero difiere en {_detalle(comp, malos)}.",
                comp, op_five, unidad, prod)


def _conciliar_cantidad(f, ctx, tipo):
    orden, cuenta = f["orden"], f["cuenta"]
    ERR = RegistroSAP.Estado.ERROR
    cant = _cantidad(f["cantidad_raw"])
    op = ctx.ops.get(orden)

    if cant is None:
        return _res(ERR, "CANTIDAD_INVALIDA",
                    f"La cantidad '{f['cantidad_raw']}' no es un número entero válido.", {}, op)
    if op is None:
        return _res(ERR, "ORDEN_NO_EXISTE", f"La orden {orden} no existe en Five.", {})

    five = ctx.cantidad.get(orden)
    if five is None:
        return _res(ERR, "SIN_MATERIAL_EN_ORDEN",
                    f"La orden {orden} existe, pero no tiene material por cantidad "
                    f"instalado en Five.", {}, op)

    comp = {
        "orden": _item("Orden", op.numero_orden, orden, True),
        "cuenta_contrato": _item("Cuenta contrato", op.numero_cuenta_contrato, cuenta,
                                 op.numero_cuenta_contrato == cuenta),
        "tipo_operacion": _item("Tipo de operación", op.get_tipo_operacion_display(),
                                Operacion.TipoOperacion(tipo).label,
                                op.tipo_operacion == tipo),
        "cantidad": _item("Cantidad", str(five), str(cant), five == cant),
    }
    malos = [k for k in ("cuenta_contrato", "tipo_operacion", "cantidad") if not comp[k]["ok"]]
    if not malos:
        return _res(RegistroSAP.Estado.OK, "", "Todo coincide.", comp, op)
    return _res(ERR, _tipo_final(malos),
                f"La orden {orden} difiere en {_detalle(comp, malos)}.", comp, op)


def conciliar(f, ctx):
    ERR = RegistroSAP.Estado.ERROR
    op = ctx.ops.get(f["orden"])

    if not f["orden"] or not f["cuenta"]:
        return _res(
            ERR,
            "FALTA_DATO",
            "La fila no trae número de orden o cuenta contrato.",
            {},
            op,
        )

    tipo = _tipo(f["tipo_raw"])

    if tipo is None:
        return _res(
            ERR,
            "TIPO_INVALIDO",
            f"El tipo de operación '{f['tipo_raw']}' no es válido.",
            {},
            op,
        )

    if f["serie"] and f["codigo_unidad"]:
        return _res(
            ERR,
            "FALTA_DATO",
            "Indique número de serie o código de unidad, no ambos.",
            {},
            op,
        )

    if f["serie"]:
        return _conciliar_serie(f, ctx, tipo)

    if f["codigo_unidad"]:
        return _conciliar_codigo_unidad(f, ctx, tipo)

    if not f["cantidad_raw"]:
        return _res(
            ERR,
            "FALTA_DATO",
            "La fila debe traer número de serie, código de unidad "
            "o cantidad.",
            {},
            op,
        )

    return _conciliar_cantidad(f, ctx, tipo)

# ---------------------------------------------------------------- carga
@transaction.atomic
def procesar_carga(archivo, usuario):
    filas = leer_excel(archivo)
    if not filas:
        raise ErrorArchivo("El archivo no tiene filas con datos.")

    ctx = _Contexto(filas)
    carga = CargaSAP.objects.create(
        archivo=archivo, usuario=usuario, hash_archivo=_hash(archivo),
    )

    claves = {_clave(f) for f in filas}
    previas = dict(
        RegistroSAP.objects
        .filter(estado=RegistroSAP.Estado.OK, clave_unica__in=claves)
        .values_list("clave_unica", "pk")
    )

    E = RegistroSAP.Estado
    vistas, registros = set(), []
    ok = err = ign = 0

    for f in filas:
        clave = _clave(f)
        cant = _cantidad(f["cantidad_raw"])
        reg = RegistroSAP(
            carga=carga,
            fila=f["fila"],
            numero_orden=f["orden"],
            codigo_unidad=f["codigo_unidad"],
            numero_serie=f["serie"],
            cuenta_contrato=f["cuenta"],
            cantidad=cant if cant is not None else 1,
            cantidad_sap_original=f["cantidad_raw"],
            clave_unica=clave,
        )

        if clave in vistas:
            reg.estado, reg.tipo_error = E.ERROR, "DUPLICADA_EN_ARCHIVO"
            reg.diagnostico = "Esta fila está repetida en el mismo archivo."
        elif clave in previas:
            reg.estado = E.IGNORADO
            reg.registro_original_id = previas[clave]
            reg.diagnostico = "Ya fue cargada y conciliada antes; no se vuelve a procesar."
        else:
            r = conciliar(f, ctx)
            reg.estado, reg.tipo_error, reg.diagnostico = r["estado"], r["tipo"], r["diag"]
            reg.comparacion = r["comp"]
            reg.operacion, reg.unidad_serializada, reg.producto = r["op"], r["uni"], r["prod"]
            reg.tecnico = r["op"].tecnico if r["op"] else None
        reg.error = reg.diagnostico if reg.estado == E.ERROR else ""
        reg.vigente = reg.estado == E.ERROR          # <-- NUEVO

        vistas.add(clave)
        registros.append(reg)
        if reg.estado == E.OK:
            ok += 1
        elif reg.estado == E.ERROR:
            err += 1
        else:
            ign += 1

    RegistroSAP.objects.bulk_create(registros, batch_size=500)
    _actualizar_vigencia(carga, registros)           # <-- NUEVO
    carga.total_filas, carga.filas_ok = len(registros), ok
    carga.filas_error, carga.filas_ignoradas = err, ign
    carga.save(update_fields=["total_filas", "filas_ok", "filas_error", "filas_ignoradas"])
    return carga

DIAS_GRACIA_SAP = 3     # SAP suele tardar días en reportar; ajústalo


def instalados_sin_sap():
    """Medidores instalados en Five que ninguna carga ha conciliado como OK."""
    reportados = (
        RegistroSAP.objects
        .filter(estado=RegistroSAP.Estado.OK)
        .exclude(numero_serie="")
        .values("numero_serie")
    )
    return (
        Asignacion.objects
        .filter(estado=Asignacion.Estado.INSTALADA,
                unidad_serializada__isnull=False,
                unidad_serializada__numero_serie__isnull=False)
        .exclude(unidad_serializada__numero_serie__in=reportados)
        .select_related("producto", "unidad_serializada", "operacion", "tecnico")
        .order_by("fecha_instalacion")
    )


def limite_gracia():
    return timezone.localdate() - timedelta(days=DIAS_GRACIA_SAP)

