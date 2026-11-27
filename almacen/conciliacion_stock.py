import hashlib
import re
import unicodedata
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from io import BytesIO

from django.db import transaction
from django.db.models import Count
from openpyxl import load_workbook

from almacen.models import Asignacion, Inventario, Producto, UnidadSerializada
from operaciones.models import CargaStock, RegistroStock

from .config_sap import ErrorConfiguracion, tipos_sin_ods

R = RegistroStock.Resultado
T = RegistroStock.TipoError

MAX_FILAS = 100000

ALIAS = {
    "material": "codigo_producto",
    "descripcion_material": "descripcion",
    "libre_utilizacion": "libre",
    "inspecc_de_calidad": "calidad",
    "stock_bloqueado": "bloqueado",
    "stock_en_transito": "transito",
}
OBLIGATORIAS = {"codigo_producto": "Material", "libre": "Libre utilización"}
NUMERICAS = ("libre", "calidad", "bloqueado", "transito")

ACCIONES = {
    T.MATERIAL_NO_EXISTE: "El material está en SAP pero no en CICB: créelo en Productos con ese código (o ignore si no lo maneja CICB).",
    T.SAP_MAYOR: "SAP tiene más que CICB: pudo salir material en CICB (entrega, instalación con ODS) que SAP aún no descuenta, o falta registrar un ingreso en CICB.",
    T.SAP_MENOR: "SAP tiene menos que CICB: pudo salir material en SAP que CICB no registró, o hay un ingreso en CICB que SAP no tiene.",
    T.NO_EN_MB52: "CICB espera stock pero SAP no lo muestra: verifique el ingreso o si el MB52 está filtrado por almacén/centro.",
    T.CANTIDAD_INVALIDA: "Corrija la cantidad en esa fila del Excel.",
}


class ErrorArchivo(Exception):
    """El archivo no se puede procesar."""


def _texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip().upper()


def _columna(nombre):
    s = unicodedata.normalize("NFKD", str(nombre or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")
    return ALIAS.get(s, s)


def _decimal(valor):
    if valor is None or valor == "":
        return Decimal(0)
    if isinstance(valor, (int, float)):
        return Decimal(str(valor))
    s = str(valor).strip().replace(" ", "")
    if not s:
        return Decimal(0)
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    return Decimal(s)


def _leer(contenido):
    """Devuelve ({material: datos sumados}, [filas ilegibles])."""
    try:
        wb = load_workbook(BytesIO(contenido), read_only=True, data_only=True)
    except Exception:
        raise ErrorArchivo("No se pudo leer el archivo. Verifique que sea un .xlsx válido.")

    ws = wb.active
    ws.reset_dimensions()
    filas = ws.iter_rows(values_only=True)
    try:
        encabezado = next(filas)
    except StopIteration:
        raise ErrorArchivo("El archivo está vacío.")

    cols = [_columna(c) for c in encabezado]
    faltan = [nombre for interno, nombre in OBLIGATORIAS.items() if interno not in cols]
    if faltan:
        raise ErrorArchivo("Faltan las columnas obligatorias del MB52: " + ", ".join(faltan) + ".")
    pos = {c: i for i, c in enumerate(cols) if c}

    def crudo(fila, nombre):
        i = pos.get(nombre)
        return fila[i] if i is not None and i < len(fila) else None

    materiales, invalidas, n_filas = {}, [], 0
    for n, fila in enumerate(filas, start=2):
        cod = _texto(crudo(fila, "codigo_producto"))
        if not cod:
            continue                      # fila vacía o de totales
        n_filas += 1
        if n_filas > MAX_FILAS:
            raise ErrorArchivo(f"El archivo supera el máximo de {MAX_FILAS} filas.")
        try:
            valores = {c: _decimal(crudo(fila, c)) for c in NUMERICAS}
        except InvalidOperation:
            invalidas.append((n, cod))
            continue

        m = materiales.setdefault(cod, {
            "fila": n, "descripcion": str(crudo(fila, "descripcion") or "").strip(),
            **{c: Decimal(0) for c in NUMERICAS},
        })
        for c in NUMERICAS:
            m[c] += valores[c]

    if not materiales and not invalidas:
        raise ErrorArchivo("El archivo no tiene filas con material.")
    return materiales, invalidas


def _stock_five(sin_ods_ids):
    """
    Lo que CICB espera que SAP todavía vea en el almacén, por producto:
      disp    = disponible en almacén
      asig    = asignado a técnicos y sin instalar
      sin_ods = instalado en órdenes de los tipos «sin ods»
    """
    A = Asignacion.Estado
    datos = defaultdict(lambda: {"disp": 0, "asig": 0, "sin_ods": 0})

    for r in (UnidadSerializada.objects.filter(estado=UnidadSerializada.Estado.DISPONIBLE)
              .values("producto_id").annotate(n=Count("id"))):
        datos[r["producto_id"]]["disp"] += r["n"]

    for pid, cant in (Inventario.objects
                      .filter(producto__tipo_inventario=Producto.TipoInventario.CANTIDAD)
                      .values_list("producto_id", "cantidad_actual")):
        datos[pid]["disp"] += cant or 0

    for pid, unidad_id, cant, instalada in (
            Asignacion.objects.filter(estado=A.ASIGNADA)
            .values_list("producto_id", "unidad_serializada_id", "cantidad", "cantidad_instalada")):
        datos[pid]["asig"] += 1 if unidad_id else max((cant or 0) - (instalada or 0), 0)

    if sin_ods_ids:
        for pid, unidad_id, cant in (
                Asignacion.objects
                .filter(estado=A.INSTALADA, operacion__tipo_operacion_id__in=sin_ods_ids)
                .values_list("producto_id", "unidad_serializada_id", "cantidad")):
            datos[pid]["sin_ods"] += 1 if unidad_id else (cant or 0)

    return datos


@transaction.atomic
def procesar_stock(*, archivo, usuario, observaciones=""):
    contenido = archivo.read()
    archivo.seek(0)
    huella = hashlib.sha256(contenido).hexdigest()
    materiales, invalidas = _leer(contenido)

    try:
        tipos = tipos_sin_ods(estricto=True)
    except ErrorConfiguracion as exc:
        raise ErrorArchivo(str(exc))
    sin_ods_ids = {t.pk for t in tipos}

    productos = {p.codigo_producto.upper(): p for p in Producto.objects.all()}
    por_id = {p.pk: p for p in productos.values()}
    five = _stock_five(sin_ods_ids)
    cero = {"disp": 0, "asig": 0, "sin_ods": 0}

    carga = CargaStock.objects.create(
        archivo=archivo, nombre_original=archivo.name, hash_archivo=huella,
        usuario=usuario, observaciones=observaciones,
        tipos_sin_ods="; ".join(t.nombre for t in tipos)[:300],
    )
    carga.archivo_repetido = (
        CargaStock.objects.filter(hash_archivo=huella).exclude(pk=carga.pk).exists()
    )

    registros, ok = [], 0

    for cod, m in materiales.items():
        p = productos.get(cod)
        f = five.get(p.pk, cero) if p else cero
        total = f["disp"] + f["asig"] + f["sin_ods"]
        dif = m["libre"] - total

        if p is None:
            tipo = T.MATERIAL_NO_EXISTE
            detalle = f"El material {cod} no existe en CICB. SAP tiene {m['libre']} libre."
        elif dif == 0:
            tipo, detalle = "", ""
        else:
            tipo = T.SAP_MAYOR if dif > 0 else T.SAP_MENOR
            detalle = (
                f"SAP libre: {m['libre']}. CICB espera {total} "
                f"(disponible {f['disp']} + asignado {f['asig']} + instalado sin ODS {f['sin_ods']}). "
                f"Diferencia: {dif:+}."
            )

        if not tipo:
            ok += 1

        registros.append(RegistroStock(
            carga=carga, fila=m["fila"], codigo_producto=cod,
            descripcion=(m["descripcion"] or (p.nombre_completo if p else ""))[:200],
            producto=p,
            sap_libre=m["libre"], sap_calidad=m["calidad"],
            sap_bloqueado=m["bloqueado"], sap_transito=m["transito"],
            five_disponible=f["disp"], five_asignado=f["asig"],
            five_sin_ods=f["sin_ods"], five_total=total, diferencia=dif,
            resultado=R.ERROR if tipo else R.OK, tipo_error=tipo, detalle=detalle,
        ))

    # Filas con cantidad ilegible
    for n, cod in invalidas:
        registros.append(RegistroStock(
            carga=carga, fila=n, codigo_producto=cod, producto=productos.get(cod),
            resultado=R.ERROR, tipo_error=T.CANTIDAD_INVALIDA,
            detalle="No se pudo leer la cantidad de esta fila.",
        ))

    # CICB espera stock de materiales que no vinieron en el MB52
    for pid, f in five.items():
        total = f["disp"] + f["asig"] + f["sin_ods"]
        p = por_id.get(pid)
        if not p or total <= 0 or p.codigo_producto.upper() in materiales:
            continue
        registros.append(RegistroStock(
            carga=carga, codigo_producto=p.codigo_producto, descripcion=p.nombre_completo[:200],
            producto=p,
            five_disponible=f["disp"], five_asignado=f["asig"],
            five_sin_ods=f["sin_ods"], five_total=total, diferencia=-total,
            resultado=R.ERROR, tipo_error=T.NO_EN_MB52,
            detalle=(f"CICB espera {total} (disponible {f['disp']} + asignado {f['asig']} + "
                     f"instalado sin ODS {f['sin_ods']}) y el material no viene en el MB52."),
        ))

    RegistroStock.objects.bulk_create(registros, batch_size=1000)

    carga.total_materiales = len(registros)
    carga.coinciden = ok
    carga.con_diferencia = len(registros) - ok
    carga.save(update_fields=["total_materiales", "coinciden", "con_diferencia"])
    return carga