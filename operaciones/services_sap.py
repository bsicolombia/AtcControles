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
    """Determina si la corrección vino de SAP, Five o ambos."""

    sap = previo.clave_unica != nuevo.clave_unica

    anterior = previo.comparacion or {}
    actual = nuevo.comparacion or {}

    five = any(
        anterior.get(k, {}).get("five")
        != actual.get(k, {}).get("five")
        for k in set(anterior) | set(actual)
    )

    if sap and five:
        return "AMBOS"

    if sap:
        return "SAP"

    return "FIVE"
def _familia(r):
    """Distingue el tipo de material dentro de una misma orden."""
    if r.numero_serie:
        return "S"
    if r.codigo_unidad:
        return "U"
    return f"P:{r.codigo_producto}"

def _actualizar_vigencia(carga, registros):
    """Actualiza únicamente errores anteriores que corresponden
    al mismo identificador o al mismo producto por cantidad.
    """

    E = RegistroSAP.Estado

    series = {r.numero_serie for r in registros if r.numero_serie}
    codigos = {r.codigo_unidad for r in registros if r.codigo_unidad}
    claves = {r.clave_unica for r in registros}

    ordenes_cantidad = {
        r.numero_orden
        for r in registros
        if not r.numero_serie
        and not r.codigo_unidad
        and r.codigo_producto
    }

    abiertos = list(
        RegistroSAP.objects
        .filter(estado=E.ERROR, vigente=True)
        .exclude(carga=carga)
        .filter(
            Q(numero_serie__in=series)
            | Q(codigo_unidad__in=codigos)
            | Q(clave_unica__in=claves)
            | Q(
                numero_orden__in=ordenes_cantidad,
            )
            (
    Q(numero_serie__isnull=True) |
    Q(numero_serie="")
),
(
    Q(codigo_unidad__isnull=True) |
    Q(codigo_unidad="")
),
        )
    )

    if not abiertos:
        return

    def llaves(r):
        resultado = [("c", r.clave_unica)]

        if r.numero_serie:
            resultado.append(("s", r.numero_serie))

        elif r.codigo_unidad:
            resultado.append(("u", r.codigo_unidad))

        elif r.codigo_producto:
            resultado.append((
                "p",
                r.numero_orden,
                r.codigo_producto,
            ))

        return resultado

    indice = defaultdict(list)

    for anterior in abiertos:
        for llave in llaves(anterior):
            indice[llave].append(anterior)

    tocados = {}

    # Los registros OK tienen prioridad sobre los errores.
    for nuevo in sorted(
        registros,
        key=lambda r: r.estado != E.OK
    ):
        if nuevo.estado == E.IGNORADO:
            continue

        for llave in llaves(nuevo):
            for anterior in indice.get(llave, []):
                tocados.setdefault(
                    anterior.pk,
                    (anterior, nuevo)
                )

    for anterior, nuevo in tocados.values():

        anterior.vigente = False

        if nuevo.estado == E.OK:
            anterior.resuelto = True
            anterior.resuelto_en = carga
            anterior.resuelto_por = _quien(anterior, nuevo)

    RegistroSAP.objects.bulk_update(
        [anterior for anterior, _ in tocados.values()],
        [
            "vigente",
            "resuelto",
            "resuelto_en",
            "resuelto_por",
        ],
    )

    def llaves(r):
        ks = [("c", r.clave_unica), ("oc", r.numero_orden, _familia(r))]
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
    
    "ORDEN_NO_EJECUTADA": ("Orden no ejecutada en Five", "ALTA",
        "SAP reporta trabajo en una orden que Five tiene como no ejecutada, cancelada o anulada. "
        "Revisar la anomalía registrada y corregir donde esté mal."),
    
    "UNIDAD_NO_EXISTE": ("Unidad no existe en Five","ALTA",
        "Verificar el código de unidad o registrar su ingreso en Five.",),
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
            "codigo": _s(d.get("codigo_producto")),
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
        return int(n) if n.is_integer() and n > 0 else None
    except ValueError:
        return None


def _clave(f):
    base = (
        f"{f['orden']}|{f['cuenta']}|{_tipo(f['tipo_raw']) or f['tipo_raw']}|"
        f"{f['serie']}|{f['codigo_unidad']}"
    )
    if not f["serie"] and not f["codigo_unidad"]:
        base += f"|{f['codigo']}|{f['cantidad_raw']}"
    return hashlib.sha1(base.encode()).hexdigest()

# ---------------------------------------------------------------- contexto
class _Contexto:
    def __init__(self, filas):
        series = {f["serie"] for f in filas if f["serie"]}
        codigos_unidad = {f["codigo_unidad"] for f in filas if f["codigo_unidad"]}
        ordenes = {f["orden"] for f in filas if f["orden"]}
        codigos_prod = {f["codigo"] for f in filas if f["codigo"]}
        E = Asignacion.Estado

        self.productos = {
            p.codigo_producto: p
            for p in Producto.objects.filter(codigo_producto__in=codigos_prod)
        }
        self.ops = {
            o.numero_orden: o
            for o in Operacion.objects
            .select_related("anomalia", "tecnico")
            .filter(numero_orden__in=ordenes)
        }

        # Se busca cada identificador por serie Y por código, para detectar
        # cuando SAP mandó un identificador en la columna equivocada.
        buscar = series | codigos_unidad
        self.por_serie = defaultdict(list)
        self.por_codigo = {}
        for u in UnidadSerializada.objects.filter(
            Q(numero_serie__in=buscar) | Q(codigo__in=buscar)
        ).select_related("producto"):
            if u.numero_serie:
                self.por_serie[u.numero_serie].append(u)
            self.por_codigo[u.codigo] = u

        ids = ({u.pk for l in self.por_serie.values() for u in l}
               | {u.pk for u in self.por_codigo.values()})

        self.asig_unidad = {}
        self.series_orden = defaultdict(list)    # pista de typo en seriales
        self.codigos_orden = defaultdict(list)   # pista de typo en códigos
        asigs = (
            Asignacion.objects
            .filter(unidad_serializada__isnull=False)
            .filter(Q(unidad_serializada_id__in=ids)
                    | Q(operacion__numero_orden__in=ordenes))
            .exclude(estado=E.DEVUELTA)
            .select_related("operacion", "unidad_serializada")
            .order_by("id")
        )
        for a in asigs:
            prev = self.asig_unidad.get(a.unidad_serializada_id)
            if prev is None or not (prev.estado == E.INSTALADA and a.estado != E.INSTALADA):
                self.asig_unidad[a.unidad_serializada_id] = a
            if a.operacion_id:
                orden = a.operacion.numero_orden
                self.codigos_orden[orden].append(a.unidad_serializada.codigo)
                if a.unidad_serializada.numero_serie:
                    self.series_orden[orden].append(a.unidad_serializada.numero_serie)

        # Por cantidad: instalado en Five por (orden, producto)
        self.cantidad = {}
        for a in (
            Asignacion.objects
            .filter(unidad_serializada__isnull=True, operacion__numero_orden__in=ordenes)
            .exclude(estado=E.DEVUELTA)
            .select_related("operacion")
        ):
            q = a.cantidad if a.estado == E.INSTALADA else a.cantidad_instalada
            k = (a.operacion.numero_orden, a.producto_id)
            self.cantidad[k] = self.cantidad.get(k, 0) + (q or 0)


# ---------------------------------------------------------------- conciliación
def _item(label, five, sap, ok):
    return {"label": label, "five": five or "", "sap": sap or "", "ok": bool(ok)}
def _res(estado, tipo, diag, comp, op=None, uni=None, prod=None):
    return {
        "estado": estado,
        "tipo": tipo,
        "diag": diag,
        "comp": comp,
        "op": op,
        "uni": uni,
        "prod": prod,
    }

def _tipo_final(malos):
    malos = set(malos)

    if {"orden", "cuenta_contrato"}.issubset(malos):
        return "ORDEN_Y_CUENTA_DIFIEREN"

    if len(malos) >= 2:
        return "VARIOS_DIFIEREN"

    equivalencias = {
        "orden": "ORDEN_DIFIERE",
        "cuenta_contrato": "CUENTA_DIFIERE",
        "tipo_operacion": "TIPO_DIFIERE",
        "cantidad": "CANTIDAD_DIFIERE",
    }

    return equivalencias.get(next(iter(malos)), "VARIOS_DIFIEREN")


def _tipo_final(malos):
    if len(malos) >= 2:
        return "VARIOS_DIFIEREN"
    return {"orden": "ORDEN_DIFIERE", "cuenta_contrato": "CUENTA_DIFIERE",
            "tipo_operacion": "TIPO_DIFIERE", "cantidad": "CANTIDAD_DIFIERE"}[malos[0]]


def _detalle(comp, malos):
    return "; ".join(
        f"{comp[k]['label'].lower()}: Five={comp[k]['five'] or '—'} / SAP={comp[k]['sap'] or '—'}"
        for k in malos)
def _conciliar_unidad(f, ctx, tipo):
    por_serie = bool(f["serie"])
    ident = f["serie"] if por_serie else f["codigo_unidad"]
    nombre = "Medidor" if por_serie else "Unidad"
    orden, cuenta = f["orden"], f["cuenta"]
    ERR, E_U = RegistroSAP.Estado.ERROR, UnidadSerializada.Estado
    op_sap = ctx.ops.get(orden)

    if por_serie:
        lista = ctx.por_serie.get(ident, [])
        if len(lista) > 1:
            cods = ", ".join(u.codigo for u in lista)
            return _res(ERR, "SERIE_DUP_FIVE",
                        f"La serie {ident} está repetida {len(lista)} veces en Five "
                        f"(códigos {cods}). Corrija el inventario.", {}, op_sap)
        unidad = lista[0] if lista else None
        cruzada = ctx.por_codigo.get(ident)                # ¿mandaron un código como serial?
        pistas = ctx.series_orden.get(orden, [])
    else:
        unidad = ctx.por_codigo.get(ident)
        cruz = ctx.por_serie.get(ident, [])
        cruzada = cruz[0] if cruz else None                # ¿mandaron un serial como código?
        pistas = ctx.codigos_orden.get(orden, [])

    # ---- identificador en la columna equivocada ----
    if unidad is None and cruzada is not None:
        if por_serie:
            msg = (f"'{ident}' es el código de unidad de {cruzada.producto.nombre_completo}, "
                   f"no un número de serie. Repórtelo en la columna codigo_unidad.")
        else:
            msg = (f"'{ident}' es el número de serie de un medidor "
                   f"({cruzada.codigo}), no un código de unidad. "
                   f"Repórtelo en la columna numero_serie.")
        return _res(ERR, "IDENTIFICADOR_INCORRECTO", msg, {}, op_sap, cruzada, cruzada.producto)

    if unidad is not None and not por_serie and unidad.producto.requiere_numero_serie:
        return _res(ERR, "IDENTIFICADOR_INCORRECTO",
                    f"{unidad.codigo} es un medidor: debe reportarse por número de serie "
                    f"({unidad.numero_serie or 'sin serial en Five'}), no por código de unidad.",
                    {}, op_sap, unidad, unidad.producto)

    cerca = difflib.get_close_matches(ident, pistas, n=1, cutoff=0.75)
    pista = f" ¿Quiso decir {cerca[0]}?" if cerca else ""

    asig = ctx.asig_unidad.get(unidad.pk) if unidad else None
    op_five = asig.operacion if asig and asig.operacion_id else None
    prod = unidad.producto if unidad else None
    instalado = bool(unidad and unidad.estado == E_U.INSTALADO)

    ref = op_five if unidad else op_sap
    f_ident = ((unidad.numero_serie if por_serie else unidad.codigo) if unidad
               else ", ".join(pistas))
    f_ord = ref.numero_orden if ref else ""
    f_cta = ref.numero_cuenta_contrato if ref else ""
    f_tip = ref.get_tipo_operacion_display() if ref else ""

    comp = {
        "medidor": _item(nombre, f_ident, ident, instalado),
        "orden": _item("Orden", f_ord, orden, f_ord and f_ord == orden),
        "cuenta_contrato": _item("Cuenta contrato", f_cta, cuenta, f_cta and f_cta == cuenta),
        "tipo_operacion": _item("Tipo de operación", f_tip,
                                Operacion.TipoOperacion(tipo).label,
                                ref is not None and ref.tipo_operacion == tipo),
    }
    # ---- no existe en Five ----
    if unidad is None:
        if op_sap is None:
            return _res(ERR, "NO_EXISTE_EN_FIVE",
                        f"Ni el identificador {ident} ni la orden {orden} existen en Five.",
                        comp)

        if not por_serie:
            return _res(ERR, "UNIDAD_NO_EXISTE",
                        f"El código de unidad {ident} no existe en Five.{pista}",
                        comp, op_sap)

        if op_sap.numero_cuenta_contrato == cuenta:
            return _res(ERR, "SERIE_DIFIERE",
                        f"La orden {orden} y la cuenta {cuenta} coinciden con Five, pero el "
                        f"medidor no: SAP reporta {ident} y Five tiene "
                        f"{f_ident or 'ningún medidor'} en esa orden.{pista}", comp, op_sap)

        return _res(ERR, "SERIE_NO_EXISTE",
                    f"El medidor {ident} no existe en Five. La orden {orden} sí existe, pero "
                    f"con otra cuenta (Five: {op_sap.numero_cuenta_contrato}, "
                    f"SAP: {cuenta}).{pista}", comp, op_sap)

    # ---- existe pero no está instalado ----
    if not instalado:
        if unidad.estado == E_U.DISPONIBLE:
            t, d = "INSTALADO_SAP_EN_ALMACEN", (
                f"SAP reporta {nombre.lower()} {ident} instalado, pero en Five está "
                f"DISPONIBLE en almacén (nunca fue asignado).")
        elif unidad.estado == E_U.ASIGNADO:
            t, d = "ASIGNADO_NO_INSTALADO", (
                f"SAP reporta {nombre.lower()} {ident} instalado, pero en Five sigue "
                f"asignado al técnico sin registrar la instalación.")
        else:
            t, d = "ESTADO_INVALIDO", (
                f"SAP reporta {nombre.lower()} {ident} instalado, pero en Five está "
                f"'{unidad.get_estado_display()}'.")
        return _res(ERR, t, d, comp, op_five, unidad, prod)

    if op_five is None:
        return _res(ERR, "SIN_ORDEN_EN_FIVE",
                    f"{nombre} {ident} figura instalado en Five pero sin orden asociada. "
                    f"SAP lo reporta en la orden {orden}.", comp, None, unidad, prod)

    malos = [k for k in ("orden", "cuenta_contrato", "tipo_operacion") if not comp[k]["ok"]]
    if not malos:
        return _res(RegistroSAP.Estado.OK, "", "Todo coincide.", comp, op_five, unidad, prod)
    return _res(ERR, _tipo_final(malos),
                f"{nombre} {ident} coincide, pero difiere en {_detalle(comp, malos)}.",
                comp, op_five, unidad, prod)


def _conciliar_cantidad(f, ctx, tipo):
    orden, cuenta = f["orden"], f["cuenta"]
    ERR = RegistroSAP.Estado.ERROR
    op = ctx.ops.get(orden)
    prod = ctx.productos.get(f["codigo"])

    if prod is None:
        return _res(ERR, "PRODUCTO_NO_EXISTE",
                    f"El código de producto {f['codigo']} no existe en Five.", {}, op)

    if prod.es_serializado:
        if prod.requiere_numero_serie:
            return _res(ERR, "FALTA_SERIE",
                        f"{prod.nombre} es un medidor: SAP debe reportarlo por numero_serie, "
                        f"no por código de producto y cantidad.", {}, op, None, prod)
        return _res(ERR, "FALTA_CODIGO_UNIDAD",
                    f"{prod.nombre} se reporta por código de unidad (ej. {prod.categoria.prefijo}-001), "
                    f"no por código de producto y cantidad.", {}, op, None, prod)

    cant = _cantidad(f["cantidad_raw"])
    if cant is None:
        return _res(ERR, "CANTIDAD_INVALIDA",
                    f"La cantidad '{f['cantidad_raw']}' no es un entero positivo.",
                    {}, op, None, prod)
    if op is None:
        return _res(ERR, "ORDEN_NO_EXISTE", f"La orden {orden} no existe en Five.",
                    {}, None, None, prod)

    five = ctx.cantidad.get((orden, prod.pk))
    if five is None:
        return _res(ERR, "SIN_MATERIAL_EN_ORDEN",
                    f"La orden {orden} existe, pero no tiene {prod.nombre} instalado en Five.",
                    {}, op, None, prod)

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
        return _res(RegistroSAP.Estado.OK, "", "Todo coincide.", comp, op, None, prod)
    return _res(ERR, _tipo_final(malos),
                f"{prod.nombre} en la orden {orden} difiere en {_detalle(comp, malos)}.",
                comp, op, None, prod)


def conciliar(f, ctx):
    ERR = RegistroSAP.Estado.ERROR
    op = ctx.ops.get(f["orden"])

    if not f["orden"] or not f["cuenta"]:
        return _res(ERR, "FALTA_DATO",
                    "La fila no trae número de orden o cuenta contrato.", {}, op)

    tipo = _tipo(f["tipo_raw"])
    if tipo is None:
        return _res(ERR, "TIPO_INVALIDO",
                    f"El tipo de operación '{f['tipo_raw']}' no es válido "
                    f"(use Instalación, Retiro o Mantenimiento).", {}, op)

    if f["serie"] and f["codigo_unidad"]:
        return _res(ERR, "FALTA_DATO",
                    "Indique número de serie o código de unidad, no ambos.", {}, op)
        
    if op and op.estado in Operacion.ESTADOS_NEGATIVOS:
        anom = f" Anomalía: {op.anomalia}." if op.anomalia_id else ""
        return _res(
            ERR, "ORDEN_NO_EJECUTADA",
            f"SAP reporta la orden {f['orden']}, pero en Five está "
            f"{op.get_estado_display().lower()}.{anom}",
            {}, op,
        )
        
    if f["serie"] or f["codigo_unidad"]:
        return _conciliar_unidad(f, ctx, tipo)

    if f["codigo"]:
        return _conciliar_cantidad(f, ctx, tipo)

    return _res(ERR, "FALTA_DATO",
                "La fila debe traer numero_serie, codigo_unidad, o codigo_producto con cantidad.",
                {}, op)

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
    previas = {
        r.clave_unica: r
        for r in RegistroSAP.objects
        .filter(estado=RegistroSAP.Estado.OK, clave_unica__in=claves)
        .order_by("pk")
    }

    E = RegistroSAP.Estado
    vistas, registros = set(), []
    ok = err = ign = 0

    for f in filas:
        clave = _clave(f)
        cant = _cantidad(f["cantidad_raw"])
        es_unidad = bool(f["serie"] or f["codigo_unidad"])
        reg = RegistroSAP(
            carga=carga,
            fila=f["fila"],
            numero_orden=f["orden"],
            codigo_producto=f["codigo"],                      # <-- faltaba
            codigo_unidad=f["codigo_unidad"],
            numero_serie=f["serie"],
            cuenta_contrato=f["cuenta"],
            cantidad=1 if es_unidad else (cant if cant is not None else 1),
            cantidad_sap_original=f["cantidad_raw"],
            clave_unica=clave,
        )

        if clave in vistas:
            reg.estado, reg.tipo_error = E.ERROR, "DUPLICADA_EN_ARCHIVO"
            reg.diagnostico = "Esta fila está repetida en el mismo archivo."
        elif clave in previas:
            orig = previas[clave]
            reg.estado = E.IGNORADO
            reg.registro_original_id = orig.pk
            reg.operacion_id = orig.operacion_id
            reg.tecnico_id = orig.tecnico_id
            reg.producto_id = orig.producto_id
            reg.unidad_serializada_id = orig.unidad_serializada_id
            reg.comparacion = orig.comparacion
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
    """
    Unidades serializadas (medidores, reguladores, válvulas) instaladas en Five
    que ninguna carga ha conciliado como OK. Los productos por cantidad NO se
    exigen: son opcionales en SAP.
    """
    reportadas = (
        RegistroSAP.objects
        .filter(estado=RegistroSAP.Estado.OK, unidad_serializada__isnull=False)
        .values("unidad_serializada_id")
    )
    return (
        Asignacion.objects
        .filter(estado=Asignacion.Estado.INSTALADA,
                unidad_serializada__isnull=False)
        .exclude(unidad_serializada_id__in=reportadas)
        .select_related("producto", "unidad_serializada", "operacion", "tecnico")
        .order_by("fecha_instalacion")
    )


def limite_gracia():
    return timezone.localdate() - timedelta(days=DIAS_GRACIA_SAP)

