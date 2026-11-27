import hashlib
import re
import unicodedata
from io import BytesIO

from django.db import transaction
from django.db.models import Q
from openpyxl import load_workbook

from almacen.models import Asignacion, Producto, UnidadSerializada
from operaciones.models import CargaInventario, RegistroInventario

from .config_sap import STATUS_ALMACEN, tipos_sin_ods

R = RegistroInventario.Resultado
T = RegistroInventario.TipoError
E5 = UnidadSerializada.Estado

MAX_FILAS = 100000

# Encabezado del IQ09 (normalizado) -> nombre interno
ALIAS = {
    "material": "codigo_producto",
    "numero_de_serie": "numero_serie",
    "status_sistema": "status_sap",
    "descripcion": "descripcion",
    "centro": "centro",
    "almacen": "almacen_sap",
    "tipo_de_stock": "tipo_stock",
    "orden": "orden_sap",
    "codigo_producto": "codigo_producto",
    "numero_serie": "numero_serie",
}

OBLIGATORIAS = {"codigo_producto": "Material", "numero_serie": "Número de serie"}

ACCIONES = {
    T.SERIE_VACIA: "Complete el número de serie en esa fila del Excel.",
    T.SERIE_DUPLICADA_ARCHIVO: "El mismo serial está repetido en el archivo; elimine el duplicado.",
    T.PRODUCTO_VACIO: "Complete la columna Material en esa fila.",
    T.PRODUCTO_NO_EXISTE: "Cree el material en Productos con ese código, o corrija el código en el archivo.",
    T.PRODUCTO_FUERA_CATEGORIA: "Es de otra categoría: cárguelo eligiendo esa categoría, o ignore la fila.",
    T.PRODUCTO_NO_SERIALIZADO: "Ese material se maneja por cantidad; no se compara por serie.",
    T.NO_EXISTE_EN_SISTEMA: "El medidor está en SAP pero no en CICB: regístrelo en Ingresos.",
    T.SERIE_OTRO_PRODUCTO: "La serie está asociada a otro material en CICB: corrija el producto de la unidad o el archivo.",
    T.CODIGO_UNIDAD_NO_COINCIDE: "El código interno no corresponde a esa serie: revise ambos datos.",
    T.ESTADO_NO_DISPONIBLE: "Revise el historial de la unidad en CICB.",
    T.STATUS_NO_ALMACEN: "SAP no lo tiene en almacén: verifique en SAP.",
    T.STATUS_NO_COINCIDE: "CICB y SAP no dicen lo mismo del estado del medidor: revise el historial de la unidad y actualice el que esté desactualizado.",
    T.ORDEN_NO_COINCIDE: "La orden en SAP es distinta a la orden de instalación en CICB: verifique cuál es la correcta.",
    T.FALTANTE_EN_ARCHIVO: "Está en CICB pero no viene en el IQ09: verifique que el serial esté bien digitado en CICB o si falta en SAP.",
}


class ErrorArchivo(Exception):
    """El archivo no se puede procesar (formato, columnas, vacío)."""


def _texto(valor, mayus=True):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    valor = str(valor).strip()
    return valor.upper() if mayus else valor


def _columna(nombre):
    s = unicodedata.normalize("NFKD", str(nombre or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")
    return ALIAS.get(s, s)


def _orden_norm(valor):
    return _texto(valor).lstrip("0")


def _lotes(items, n=500):
    for i in range(0, len(items), n):
        yield items[i:i + n]


def _leer_filas(contenido):
    try:
        wb = load_workbook(BytesIO(contenido), read_only=True, data_only=True)
    except Exception:
        raise ErrorArchivo("No se pudo leer el archivo. Verifique que sea un .xlsx válido.")

    ws = wb.active
    ws.reset_dimensions()   # los exports de SAP suelen traer mal las dimensiones
    filas_excel = ws.iter_rows(values_only=True)
    try:
        encabezado = next(filas_excel)
    except StopIteration:
        raise ErrorArchivo("El archivo está vacío.")

    cols = [_columna(c) for c in encabezado]
    faltan = [nombre for interno, nombre in OBLIGATORIAS.items() if interno not in cols]
    if faltan:
        raise ErrorArchivo("Faltan las columnas obligatorias: " + ", ".join(faltan) + ".")
    pos = {c: i for i, c in enumerate(cols) if c}

    def celda(fila, nombre, mayus=True):
        i = pos.get(nombre)
        return _texto(fila[i], mayus) if i is not None and i < len(fila) else ""

    filas = []
    for n, fila in enumerate(filas_excel, start=2):
        reg = {
            "fila": n,
            "serie": celda(fila, "numero_serie"),
            "cod_prod": celda(fila, "codigo_producto"),
            "descripcion": celda(fila, "descripcion", mayus=False),
            "status_sap": celda(fila, "status_sap"),
            "centro": celda(fila, "centro"),
            "almacen_sap": celda(fila, "almacen_sap"),
            "tipo_stock": celda(fila, "tipo_stock"),
            "orden_sap": celda(fila, "orden_sap"),
        }
        if not (reg["serie"] or reg["cod_prod"]):
            continue
        filas.append(reg)
        if len(filas) > MAX_FILAS:
            raise ErrorArchivo(f"El archivo supera el máximo de {MAX_FILAS} filas.")

    if not filas:
        raise ErrorArchivo("El archivo no tiene filas con datos.")
    return filas


def _evaluar(f, categoria, productos, unidad, inst, sin_ods_ids):
    """Devuelve (tipo_error, detalle) o (None, '') si coincide."""
    serie, cod_prod = f["serie"], f["cod_prod"]

    if not cod_prod:
        return T.PRODUCTO_VACIO, "La fila no trae código de material."

    producto = productos.get(cod_prod)
    if producto is None:
        return T.PRODUCTO_NO_EXISTE, f"El material {cod_prod} no existe en el sistema."
    if producto.categoria_id != categoria.pk:
        return T.PRODUCTO_FUERA_CATEGORIA, (
            f"El material {cod_prod} pertenece a la categoría {producto.categoria}, "
            f"no a {categoria}."
        )
    if not producto.es_serializado:
        return T.PRODUCTO_NO_SERIALIZADO, f"El material {cod_prod} se maneja por cantidad, no por serial."

    if unidad is None:
        return T.NO_EXISTE_EN_SISTEMA, f"La serie {serie} no está registrada en CICB."
    if unidad.producto_id != producto.pk:
        return T.SERIE_OTRO_PRODUCTO, (
            f"La serie {serie} está registrada en CICB como {unidad.producto.nombre_completo} "
            f"({unidad.producto.codigo_producto}), pero SAP la trae con el material {cod_prod}."
        )

    # ---- status de SAP contra estado de CICB ----
    status = f["status_sap"]
    if status:
        en_alma = status in STATUS_ALMACEN
        if unidad.estado in (E5.DISPONIBLE, E5.ASIGNADO) and not en_alma:
            return T.STATUS_NO_COINCIDE, (
                f"En CICB está «{unidad.get_estado_display()}» (en almacén o en manos del técnico), "
                f"pero en SAP el status es {status}."
            )
        if unidad.estado == E5.INSTALADO and en_alma:
            es_sin_ods = bool(
                inst and inst.operacion_id
                and inst.operacion.tipo_operacion_id in sin_ods_ids
            )
            if not es_sin_ods:
                return T.STATUS_NO_COINCIDE, (
                    f"Instalado en CICB, pero SAP todavía lo tiene en almacén ({status}): "
                    "falta reportar la instalación en SAP."
                )

    # ---- orden de SAP contra orden de CICB ----
    if (unidad.estado == E5.INSTALADO and inst and inst.operacion_id and f["orden_sap"]
            and _orden_norm(f["orden_sap"]) != _orden_norm(inst.operacion.numero_orden)):
        return T.ORDEN_NO_COINCIDE, (
            f"SAP tiene la orden {f['orden_sap']} y CICB lo instaló en la orden "
            f"{inst.operacion.numero_orden}."
        )

    return None, ""


@transaction.atomic
def procesar_carga(*, archivo, categoria, usuario, observaciones=""):
    contenido = archivo.read()
    archivo.seek(0)
    huella = hashlib.sha256(contenido).hexdigest()
    filas = _leer_filas(contenido)

    productos = {
        p.codigo_producto.upper(): p
        for p in Producto.objects.select_related("categoria")
    }

    # Universo: TODAS las unidades de la categoría, en cualquier estado
    universo = {
        u.pk: u for u in
        UnidadSerializada.objects.filter(producto__categoria=categoria).select_related("producto")
    }
    indice = {}

    def indexar(u):
        if u.numero_serie:
            indice[u.numero_serie.upper()] = u
        indice.setdefault(u.codigo.upper(), u)

    for u in universo.values():
        indexar(u)

    # Series que no están en la categoría: se buscan aparte para distinguir
    # "existe en otra categoría" de "no existe"
    sin_encontrar = sorted({f["serie"] for f in filas if f["serie"] and f["serie"] not in indice})
    for lote in _lotes(sin_encontrar):
        for u in (UnidadSerializada.objects
                  .filter(Q(numero_serie__in=lote) | Q(codigo__in=lote))
                  .select_related("producto")):
            indexar(u)

    # Instalación vigente de cada unidad (para validar orden y status)
    inst_por_unidad = {
        a.unidad_serializada_id: a
        for a in (Asignacion.objects
                  .filter(estado=Asignacion.Estado.INSTALADA,
                          unidad_serializada__producto__categoria=categoria)
                  .select_related("operacion__tipo_operacion"))
    }
    sin_ods_ids = {t.pk for t in tipos_sin_ods()}

    carga = CargaInventario.objects.create(
        categoria=categoria, archivo=archivo, nombre_original=archivo.name,
        hash_archivo=huella, usuario=usuario, observaciones=observaciones,
        total_filas=len(filas), disponibles_sistema=len(universo),
    )
    carga.archivo_repetido = (
        CargaInventario.objects.filter(hash_archivo=huella).exclude(pk=carga.pk).exists()
    )

    def orden_de(unidad):
        inst = inst_por_unidad.get(unidad.pk) if unidad else None
        return inst.operacion.numero_orden if inst and inst.operacion_id else ""

    registros, vistas, encontradas = [], {}, set()
    ok = error = 0

    for f in filas:
        serie = f["serie"]
        unidad = indice.get(serie) if serie else None
        if unidad is not None and unidad.pk in universo:
            encontradas.add(unidad.pk)
        inst = inst_por_unidad.get(unidad.pk) if unidad else None

        if not serie:
            tipo, detalle = T.SERIE_VACIA, "La fila no trae número de serie."
        elif serie in vistas:
            tipo, detalle = T.SERIE_DUPLICADA_ARCHIVO, f"Ya aparece en la fila {vistas[serie]}."
        else:
            vistas[serie] = f["fila"]
            tipo, detalle = _evaluar(f, categoria, productos, unidad, inst, sin_ods_ids)

        if tipo:
            error += 1
        else:
            ok += 1

        registros.append(RegistroInventario(
            carga=carga, fila=f["fila"], numero_serie=serie, codigo_producto=f["cod_prod"],
            descripcion=f["descripcion"][:200], status_sap=f["status_sap"][:20],
            centro=f["centro"][:10], almacen_sap=f["almacen_sap"][:10],
            tipo_stock=f["tipo_stock"][:10], orden_sap=f["orden_sap"][:30],
            orden_five=orden_de(unidad)[:30],
            resultado=R.ERROR if tipo else R.OK,
            tipo_error=tipo or "", detalle=detalle,
            producto=productos.get(f["cod_prod"]), unidad=unidad,
            estado_sistema=unidad.estado if unidad else "",
        ))

    # Faltantes: están en CICB (cualquier estado) y no vienen en el IQ09
    faltantes = [u for pk, u in universo.items() if pk not in encontradas]
    for u in faltantes:
        registros.append(RegistroInventario(
            carga=carga, numero_serie=u.numero_serie or "", codigo_unidad=u.codigo,
            codigo_producto=u.producto.codigo_producto,
            descripcion=u.producto.nombre_completo[:200],
            orden_five=orden_de(u)[:30],
            resultado=R.FALTANTE, tipo_error=T.FALTANTE_EN_ARCHIVO,
            detalle=f"Figura en CICB como «{u.get_estado_display()}» pero no viene en el IQ09.",
            producto=u.producto, unidad=u, estado_sistema=u.estado,
        ))

    RegistroInventario.objects.bulk_create(registros, batch_size=1000)

    carga.filas_ok, carga.filas_error, carga.faltantes_archivo = ok, error, len(faltantes)
    carga.save(update_fields=["filas_ok", "filas_error", "faltantes_archivo"])
    return carga