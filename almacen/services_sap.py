import logging
from decimal import Decimal, InvalidOperation

import openpyxl

from django.db import transaction
from django.db.models import Q

from operaciones.models import Operacion
from .models import (
    CargaSAP,
    Producto,
    RegistroSAP,
    UnidadSerializada,
)

logger = logging.getLogger(__name__)

COLUMNAS_OBLIGATORIAS = {
    "numero_orden",
    "codigo_producto",
}


class ErrorArchivo(Exception):
    """Errores que impiden procesar el archivo completo."""


def _txt(valor):
    """Normaliza valores de Excel sin perder ceros de texto."""
    if valor is None:
        return ""

    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)

    return str(valor).strip()


def _leer_excel(archivo):
    """
    Lee el Excel y valida que tenga las columnas requeridas.
    Devuelve tuplas con el número real de fila y sus datos.
    """
    try:
        archivo.seek(0)
        wb = openpyxl.load_workbook(
            archivo,
            read_only=True,
            data_only=True,
        )
    except Exception as exc:
        raise ErrorArchivo(
            "No se pudo leer el archivo. Suba un Excel .xlsx válido."
        ) from exc

    try:
        ws = wb.active
        filas = ws.iter_rows(values_only=True)

        try:
            encabezado_original = next(filas)
        except StopIteration:
            raise ErrorArchivo("El archivo está vacío.")

        encabezado = [
            _txt(c).lower()
            for c in encabezado_original
        ]

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


def _validar(datos):
    """
    Devuelve (mensaje_error, referencias).

    Un mensaje vacío significa que las validaciones realizadas
    fueron satisfactorias. No modifica inventarios ni instalaciones.
    """
    refs = {}

    numero_orden = datos.get("numero_orden", "")
    codigo = datos.get("codigo_producto", "")
    serie = datos.get("numero_serie", "")
    cantidad_excel = datos.get("cantidad", "")

    if not numero_orden:
        return "Falta el número de orden.", refs

    if not codigo:
        return "Falta el código del producto.", refs

    operacion = (
        Operacion.objects
        .select_related("tecnico")
        .filter(numero_orden=numero_orden)
        .first()
    )

    if not operacion:
        return (
            f"La orden {numero_orden} no existe en el sistema.",
            refs,
        )

    refs["operacion"] = operacion
    refs["tecnico"] = getattr(operacion, "tecnico", None)

    producto = (
        Producto.objects
        .filter(codigo_producto__iexact=codigo)
        .first()
    )

    if not producto:
        return (
            f"El producto «{codigo}» no existe en el sistema.",
            refs,
        )

    refs["producto"] = producto

    # ---------------------------------------------------------
    # MEDIDORES Y OTROS PRODUCTOS SERIALIZADOS
    # ---------------------------------------------------------
    if producto.es_serializado:
        if not serie:
            return (
                "El producto es serializado, pero falta el número de serie.",
                refs,
            )

        unidad = (
            UnidadSerializada.objects
            .filter(
                Q(numero_serie__iexact=serie)
                | Q(codigo__iexact=serie)
            )
            .first()
        )

        if not unidad:
            return (
                f"El serial «{serie}» no existe en el sistema.",
                refs,
            )

        refs["unidad_serializada"] = unidad
        refs["cantidad"] = 1

        # Primero se comprueba que el serial pertenezca al producto.
        if unidad.producto_id != producto.pk:
            return (
                f"El serial «{serie}» existe, pero no pertenece "
                f"al producto «{codigo}».",
                refs,
            )

        # Si el Excel incluye cantidad, un medidor representa una unidad.
        if cantidad_excel:
            try:
                cantidad = Decimal(cantidad_excel)
            except InvalidOperation:
                return (
                    f"La cantidad del serial «{serie}» no es válida.",
                    refs,
                )

            if cantidad != 1:
                return (
                    f"El serial «{serie}» representa una unidad; "
                    "la cantidad SAP debe ser 1.",
                    refs,
                )

        if unidad.estado != UnidadSerializada.Estado.INSTALADO:
            estado_actual = unidad.get_estado_display()

            return (
                f"SAP reporta el serial «{serie}» como instalado, "
                f"pero en el sistema figura como «{estado_actual}».",
                refs,
            )

        return "", refs

    # ---------------------------------------------------------
    # PRODUCTOS QUE SE CONTROLAN POR CANTIDAD
    # ---------------------------------------------------------
    if serie:
        return (
            f"El producto «{codigo}» no es serializado, "
            "pero el Excel contiene un número de serie.",
            refs,
        )

    if not cantidad_excel:
        return (
            f"Falta la cantidad SAP para el producto «{codigo}».",
            refs,
        )

    try:
        cantidad_decimal = Decimal(cantidad_excel)
    except InvalidOperation:
        return "La cantidad no es un número válido.", refs

    if not cantidad_decimal.is_finite():
        return "La cantidad debe ser un número finito.", refs

    if cantidad_decimal <= 0:
        return "La cantidad debe ser mayor que cero.", refs

    if cantidad_decimal != cantidad_decimal.to_integral_value():
        return "La cantidad debe ser un número entero.", refs

    refs["cantidad"] = int(cantidad_decimal)

    return "", refs


@transaction.atomic
def procesar_carga(archivo, usuario):
    """
    Guarda la carga y todos sus resultados.

    Las cargas anteriores se conservan como historial.
    Esta función no instala medidores ni modifica existencias.
    """
    carga = CargaSAP.objects.create(
        archivo=archivo,
        usuario=usuario,
    )

    total = 0
    ok = 0
    errores = 0

    try:
        filas = _leer_excel(carga.archivo.path)

        for numero_fila, datos in filas:
            total += 1

            try:
                mensaje, refs = _validar(datos)

            except Exception:
                logger.exception(
                    "Error validando la fila %s de la carga SAP %s",
                    numero_fila,
                    carga.pk,
                )
                mensaje = (
                    "No fue posible validar esta fila. "
                    "Revise los datos e inténtelo nuevamente."
                )
                refs = {}

            registro = RegistroSAP.objects.create(
                carga=carga,
                fila=numero_fila,
                numero_orden=datos.get("numero_orden", "")[:30],
                codigo_producto=datos.get("codigo_producto", "")[:30],
                numero_serie=datos.get("numero_serie", "")[:100],
                cantidad=refs.get("cantidad", 1),
                estado=(
                    RegistroSAP.Estado.ERROR
                    if mensaje
                    else RegistroSAP.Estado.OK
                ),
                error=mensaje,
                operacion=refs.get("operacion"),
                tecnico=refs.get("tecnico"),
                producto=refs.get("producto"),
                unidad_serializada=refs.get("unidad_serializada"),
            )

            if registro.estado == RegistroSAP.Estado.OK:
                ok += 1
            else:
                errores += 1

        carga.total_filas = total
        carga.filas_ok = ok
        carga.filas_error = errores
        carga.save(
            update_fields=[
                "total_filas",
                "filas_ok",
                "filas_error",
            ]
        )

        return carga

    except ErrorArchivo:
        # La transacción se revierte si el archivo completo es inválido.
        raise

    except Exception:
        logger.exception(
            "No se pudo completar la carga SAP %s",
            carga.pk,
        )
        raise