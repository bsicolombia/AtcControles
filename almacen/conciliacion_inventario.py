import hashlib
import re
import unicodedata
from io import BytesIO

from django.db import transaction
from openpyxl import load_workbook

from operaciones.models import (
    CargaInventario, Producto, RegistroInventario, UnidadSerializada,
)

R = RegistroInventario.Resultado
T = RegistroInventario.TipoError

MAX_FILAS = 20000

# Status de SAP que significan "está en el almacén"
STATUS_ALMACEN = {"ALMA"}

# Encabezado del Excel (normalizado) -> nombre interno
ALIAS = {
    "material": "codigo_producto",
    "numero_de_serie": "numero_serie",
    "status_sistema": "status_sap",
    "descripcion": "descripcion",
    "centro": "centro",
    "almacen": "almacen_sap",
    "tipo_de_stock": "tipo_stock",
    "orden": "orden_sap",
    # también acepta la plantilla simple
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
    T.NO_EXISTE_EN_SISTEMA: "El medidor no se ingresó al sistema: regístrelo en Ingresos.",
    T.SERIE_OTRO_PRODUCTO: "La serie está asociada a otro material en el sistema: corrija el producto de la unidad o el archivo.",
    T.CODIGO_UNIDAD_NO_COINCIDE: "El código interno no corresponde a esa serie: revise ambos datos.",
    T.ESTADO_NO_DISPONIBLE: "El archivo lo tiene en almacén pero el sistema lo muestra asignado, instalado o dañado: revise el historial de la unidad.",
    T.STATUS_NO_ALMACEN: "SAP no lo tiene en almacén (status distinto de ALMA): verifique en SAP.",
    T.FALTANTE_EN_ARCHIVO: "El sistema lo muestra disponible pero no vino en el archivo: verifique si se entregó, se instaló sin registrar, o si el archivo está incompleto.",
}

class ErrorArchivo(Exception):
    """El archivo no se puede procesar (formato, columnas, vacío)."""


def _texto(valor, mayus=True):
    """123456.0 -> '123456'; quita espacios; opcionalmente a mayúsculas."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    valor = str(valor).strip()
    return valor.upper() if mayus else valor


def _columna(nombre):
    """'Número de serie' -> 'numero_de_serie' -> alias interno."""
    s = unicodedata.normalize("NFKD", str(nombre or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")
    return ALIAS.get(s, s)


def _leer_filas(contenido):
    try:
        wb = load_workbook(BytesIO(contenido), read_only=True, data_only=True)
    except Exception:
        raise ErrorArchivo("No se pudo leer el archivo. Verifique que sea un .xlsx válido.")

    ws = wb.active
    ws.reset_dimensions()   # los exports de SAP suelen traer mal las dimensiones y openpyxl lee de menos
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
            continue  # fila vacía
        filas.append(reg)
        if len(filas) > MAX_FILAS:
            raise ErrorArchivo(f"El archivo supera el máximo de {MAX_FILAS} filas.")

    if not filas:
        raise ErrorArchivo("El archivo no tiene filas con datos.")
    return filas


def _evaluar(f, categoria, productos, unidad):
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
        return T.NO_EXISTE_EN_SISTEMA, f"La serie {serie} no está registrada en el sistema."
    if unidad.producto_id != producto.pk:
        return T.SERIE_OTRO_PRODUCTO, (
            f"La serie {serie} está registrada como {unidad.producto.nombre_completo} "
            f"({unidad.producto.codigo_producto}), pero el archivo la trae con el material {cod_prod}."
        )
    if unidad.estado != UnidadSerializada.Estado.DISPONIBLE:
        return T.ESTADO_NO_DISPONIBLE, (
            f"En el sistema la unidad figura «{unidad.get_estado_display()}», no disponible."
        )
    if f["status_sap"] and f["status_sap"] not in STATUS_ALMACEN:
        return T.STATUS_NO_ALMACEN, (
            f"En SAP el status es {f['status_sap']}; se esperaba "
            f"{' / '.join(sorted(STATUS_ALMACEN))}."
        )
    return None, ""


@transaction.atomic
def procesar_carga(*, archivo, categoria, usuario, observaciones=""):
    contenido = archivo.read()
    archivo.seek(0)
    huella = hashlib.sha256(contenido).hexdigest()
    filas = _leer_filas(contenido)

    # ---- datos del sistema (pocas consultas, sin N+1) ----
    productos = {
        p.codigo_producto.upper(): p
        for p in Producto.objects.select_related("categoria")
    }
    series = {f["serie"] for f in filas if f["serie"]}
    unidades = {}
    consulta = (
        UnidadSerializada.objects.filter(numero_serie__in=series)
        | UnidadSerializada.objects.filter(codigo__in=series)
    ).select_related("producto")
    for u in consulta:
        if u.numero_serie:
            unidades[u.numero_serie.upper()] = u
        unidades.setdefault(u.codigo.upper(), u)

    disponibles = {
        u.pk: u for u in UnidadSerializada.objects.filter(
            producto__categoria=categoria,
            estado=UnidadSerializada.Estado.DISPONIBLE,
        ).select_related("producto")
    }

    carga = CargaInventario.objects.create(
        categoria=categoria, archivo=archivo, nombre_original=archivo.name,
        hash_archivo=huella, usuario=usuario, observaciones=observaciones,
        total_filas=len(filas), disponibles_sistema=len(disponibles),
    )
    carga.archivo_repetido = (
        CargaInventario.objects.filter(hash_archivo=huella).exclude(pk=carga.pk).exists()
    )

    # ---- comparar fila por fila ----
    registros, vistas, encontradas = [], {}, set()
    ok = error = 0

    for f in filas:
        serie = f["serie"]
        unidad = unidades.get(serie) if serie else None
        if unidad:
            encontradas.add(unidad.pk)

        if not serie:
            tipo, detalle = T.SERIE_VACIA, "La fila no trae número de serie."
        elif serie in vistas:
            tipo, detalle = T.SERIE_DUPLICADA_ARCHIVO, f"Ya aparece en la fila {vistas[serie]}."
        else:
            vistas[serie] = f["fila"]
            tipo, detalle = _evaluar(f, categoria, productos, unidad)

        if tipo:
            error += 1
        else:
            ok += 1

        registros.append(RegistroInventario(
            carga=carga, fila=f["fila"], numero_serie=serie, codigo_producto=f["cod_prod"],
            descripcion=f["descripcion"][:200], status_sap=f["status_sap"][:20],
            centro=f["centro"][:10], almacen_sap=f["almacen_sap"][:10],
            tipo_stock=f["tipo_stock"][:10], orden_sap=f["orden_sap"][:30],
            resultado=R.ERROR if tipo else R.OK,
            tipo_error=tipo or "", detalle=detalle,
            producto=productos.get(f["cod_prod"]), unidad=unidad,
            estado_sistema=unidad.estado if unidad else "",
        ))

    # ---- faltantes: disponibles en el sistema que no vinieron en el archivo ----
    faltantes = [u for pk, u in disponibles.items() if pk not in encontradas]
    for u in faltantes:
        registros.append(RegistroInventario(
            carga=carga, numero_serie=u.numero_serie or "", codigo_unidad=u.codigo,
            codigo_producto=u.producto.codigo_producto,
            descripcion=u.producto.nombre_completo[:200],
            resultado=R.FALTANTE, tipo_error=T.FALTANTE_EN_ARCHIVO,
            detalle="Figura disponible en el almacén (sistema) pero no está en el archivo cargado.",
            producto=u.producto, unidad=u, estado_sistema=u.estado,
        ))

    RegistroInventario.objects.bulk_create(registros, batch_size=1000)

    carga.filas_ok, carga.filas_error, carga.faltantes_archivo = ok, error, len(faltantes)
    carga.save(update_fields=["filas_ok", "filas_error", "faltantes_archivo"])
    return carga