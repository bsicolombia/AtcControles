import logging
from decimal import Decimal, InvalidOperation

import openpyxl

from django.db import transaction
from django.db.models import Q

from operaciones.models import Operacion, CargaSAP, RegistroSAP
from .models import Producto, UnidadSerializada

logger = logging.getLogger(__name__)

COLUMNAS_OBLIGATORIAS = {
    "numero_orden",
    "codigo_producto",
}


class ErrorArchivo(Exception):
    """Error que impide procesar el archivo completo."""


def _txt(valor):
    if valor is None:
        return ""

    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)

    return str(valor).strip()


def _leer_excel(archivo):
    try:
        archivo.seek(0)
        wb = openpyxl.load_workbook(
            archivo,
            read_only=True,
            data_only=True,
        )
    except Exception as exc:
        raise ErrorArchivo(
            "No se pudo leer el archivo. Debe ser un Excel .xlsx válido."
        ) from exc

    try:
        ws = wb.active
        filas = ws.iter_rows(values_only=True)

        try:
            encabezado_original = next(filas)
        except StopIteration:
            raise ErrorArchivo("El archivo está vacío.")

        encabezado = [_txt(c).lower() for c in encabezado_original]

        if not any(encabezado):
            raise ErrorArchivo(
                "La primera fila debe contener los encabezados."
            )

        if len(encabezado) != len(set(encabezado)):
            raise ErrorArchivo(
                "El archivo contiene encabezados duplicados."
            )

        faltantes = COLUMNAS_OBLIGATORIAS - set(encabezado)
        if faltantes:
            raise ErrorArchivo(
                "Faltan columnas obligatorias: "
                + ", ".join(sorted(faltantes))
            )

        for numero_fila, fila in enumerate(filas, start=2):
            datos = {
                encabezado[i]: _txt(valor)
                for i, valor in enumerate(fila)
                if i < len(encabezado) and encabezado[i]
            }

            if any(datos.values()):
                yield numero_fila, datos

    finally:
        wb.close()


def _error(tipo, mensaje, refs):
    refs["tipo_error"] = tipo
    refs["diagnostico"] = mensaje
    return mensaje, refs


def _validar(datos):
    refs = {
        "cantidad": 0,
        "tipo_error": "",
        "diagnostico": "",
    }

    orden = datos.get("numero_orden", "")
    codigo = datos.get("codigo_producto", "")
    serie = datos.get("numero_serie", "")
    cantidad_original = datos.get("cantidad", "")

    if not orden:
        return _error(
            "ORDEN_VACIA",
            "Falta el número de orden en el archivo SAP.",
            refs,
        )

    if not codigo:
        return _error(
            "PRODUCTO_VACIO",
            "Falta el código del producto.",
            refs,
        )

    operacion = (
        Operacion.objects
        .select_related("tecnico")
        .filter(numero_orden=orden)
        .first()
    )

    if not operacion:
        return _error(
            "ORDEN_NO_EXISTE",
            f"La orden «{orden}» no existe en el sistema.",
            refs,
        )

    refs["operacion"] = operacion
    refs["tecnico"] = operacion.tecnico

    producto = (
        Producto.objects
        .filter(codigo_producto__iexact=codigo)
        .first()
    )

    if not producto:
        return _error(
            "PRODUCTO_NO_EXISTE",
            f"El producto «{codigo}» no existe en el sistema.",
            refs,
        )

    refs["producto"] = producto

    # Productos serializados, por ejemplo medidores.
    if producto.es_serializado:
        if not serie:
            return _error(
                "SERIAL_VACIO",
                f"El producto «{codigo}» es serializado, "
                "pero falta el número de serie.",
                refs,
            )

        # SAP está reportando una unidad para este serial.
        refs["cantidad"] = 1

        unidad = (
            UnidadSerializada.objects
            .filter(producto=producto)
            .filter(
                Q(numero_serie__iexact=serie)
                | Q(codigo__iexact=serie)
            )
            .first()
        )

        if not unidad:
            return _error(
                "SERIAL_NO_EXISTE",
                (
                    f"El serial «{serie}» no existe para el producto "
                    f"«{codigo}» en el sistema."
                ),
                refs,
            )

        refs["unidad_serializada"] = unidad

        if cantidad_original:
            try:
                cantidad = Decimal(cantidad_original)
            except (InvalidOperation, ValueError):
                return _error(
                    "CANTIDAD_INVALIDA",
                    f"La cantidad del serial «{serie}» no es válida.",
                    refs,
                )

            if not cantidad.is_finite() or cantidad != 1:
                return _error(
                    "CANTIDAD_SERIAL_INVALIDA",
                    f"El serial «{serie}» representa una unidad. "
                    "La cantidad SAP debe ser 1.",
                    refs,
                )

        if unidad.estado != UnidadSerializada.Estado.INSTALADO:
            return _error(
                "SERIAL_NO_INSTALADO_FIVE",
                f"SAP reporta el serial «{serie}» como instalado, "
                f"pero FIVE lo tiene en estado "
                f"«{unidad.get_estado_display()}».",
                refs,
            )

        return "", refs

    # Productos controlados por cantidad.
    if serie:
        return _error(
            "SERIAL_EN_PRODUCTO_CANTIDAD",
            f"El producto «{codigo}» no es serializado, "
            "pero el Excel contiene un número de serie.",
            refs,
        )

    if not cantidad_original:
        return _error(
            "CANTIDAD_VACIA",
            f"Falta la cantidad SAP del producto «{codigo}».",
            refs,
        )

    try:
        cantidad_decimal = Decimal(cantidad_original)
    except (InvalidOperation, ValueError):
        return _error(
            "CANTIDAD_INVALIDA",
            f"La cantidad «{cantidad_original}» no es numérica.",
            refs,
        )

    if not cantidad_decimal.is_finite():
        return _error(
            "CANTIDAD_INVALIDA",
            "La cantidad debe ser un número finito.",
            refs,
        )

    if cantidad_decimal <= 0:
        return _error(
            "CANTIDAD_NO_POSITIVA",
            "La cantidad SAP debe ser mayor que cero.",
            refs,
        )

    if cantidad_decimal != cantidad_decimal.to_integral_value():
        return _error(
            "CANTIDAD_DECIMAL",
            "La cantidad SAP debe ser un número entero.",
            refs,
        )

    refs["cantidad"] = int(cantidad_decimal)
    return "", refs


@transaction.atomic
def procesar_carga(archivo, usuario):
    filas = list(_leer_excel(archivo))

    if not filas:
        raise ErrorArchivo(
            "El archivo no contiene filas con información para procesar."
        )

    carga = CargaSAP.objects.create(
        archivo=archivo,
        usuario=usuario,
    )

    total = ok = errores = 0

    for numero_fila, datos in filas:
        total += 1

        orden = datos.get("numero_orden", "")
        codigo = datos.get("codigo_producto", "")
        serie = datos.get("numero_serie", "")
        cantidad_original = datos.get("cantidad", "")

        try:
            mensaje, refs = _validar(datos)
        except Exception:
            logger.exception(
                "Error validando fila %s de carga SAP %s",
                numero_fila,
                carga.pk,
            )
            refs = {
                "cantidad": 0,
                "tipo_error": "ERROR_INTERNO",
                "diagnostico": (
                    "No fue posible validar esta fila. "
                    "Revise los datos e inténtelo nuevamente."
                ),
            }
            mensaje = refs["diagnostico"]

        tipo_error = refs.get("tipo_error", "")
        diagnostico = refs.get("diagnostico", "")
        tiene_error = bool(mensaje)

        RegistroSAP.objects.create(
            carga=carga,
            fila=numero_fila,
            numero_orden=orden[:30],
            codigo_producto=codigo[:30],
            numero_serie=serie[:100],
            cantidad=refs.get("cantidad", 0),
            cantidad_sap_original=cantidad_original[:50],
            estado=(
                RegistroSAP.Estado.ERROR
                if tiene_error
                else RegistroSAP.Estado.OK
            ),
            error=mensaje,
            tipo_error=tipo_error[:30],
            diagnostico=diagnostico,
            operacion=refs.get("operacion"),
            tecnico=refs.get("tecnico"),
            producto=refs.get("producto"),
            unidad_serializada=refs.get("unidad_serializada"),
        )

        if tiene_error:
            errores += 1
        else:
            ok += 1

    carga.total_filas = total
    carga.filas_ok = ok
    carga.filas_error = errores
    carga.save(
        update_fields=["total_filas", "filas_ok", "filas_error"]
    )

    return carga