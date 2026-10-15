from django.db import transaction
from django.utils import timezone
from operaciones.models import Operacion
from .models import (
    Asignacion,
    Inventario,
    MovimientoInventario,
    Producto,
    UnidadSerializada,
)
Mov = MovimientoInventario.TipoMovimiento
EstadoU = UnidadSerializada.Estado
EstadoA = Asignacion.Estado
E = UnidadSerializada.Estado

# ============================================================
# ESTADOS PERMITIDOS PARA UNIDADES SERIALIZADAS
# ============================================================

TRANSICIONES = {
    E.DISPONIBLE: (
        E.DAÑADO,
        E.BAJA,
    ),

    E.DAÑADO: (
        E.DISPONIBLE,
        E.BAJA,
    ),

    E.INSTALADO: (
        E.DAÑADO,
    ),

    E.BAJA: (
        E.DISPONIBLE,
    ),
}


class ErrorInventario(Exception):
    """Error de regla de negocio apto para mostrar al usuario."""


# ============================================================
# INVENTARIO
# ============================================================

def _inventario(producto):
    """
    Obtiene y bloquea el inventario del producto
    dentro de una transacción.
    """
    inventario, _ = (
        Inventario.objects
        .select_for_update()
        .get_or_create(producto=producto)
    )
    return inventario


# ============================================================
# CONSECUTIVO DE UNIDADES SERIALIZADAS
# ============================================================

def _siguiente_numero(prefijo):
    """
    Obtiene el siguiente consecutivo para el código
    interno de una unidad serializada.

    Ejemplos:
        MED-001
        MED-002
        REG-001
        VAL-001
    """
    mayor = 0

    codigos = (
        UnidadSerializada.objects
        .filter(codigo__startswith=f"{prefijo}-")
        .values_list("codigo", flat=True)
    )

    for codigo in codigos:
        try:
            numero = int(codigo.rsplit("-", 1)[1])
            mayor = max(mayor, numero)
        except (ValueError, IndexError):
            continue

    return mayor + 1


# ============================================================
# CÓDIGO AUTOMÁTICO DE PRODUCTO
# ============================================================

def generar_codigo_producto(categoria):
    """
    Genera códigos de producto:

        MED-001
        REG-001
        VAL-001
        COD-001
    """

    base = (
        categoria.prefijo
        or categoria.nombre[:3]
    ).strip().upper()

    if not base:
        raise ErrorInventario(
            "La categoría debe tener un nombre "
            "o un prefijo válido."
        )

    numero = (
        Producto.objects
        .filter(
            codigo_producto__startswith=f"{base}-"
        )
        .count()
        + 1
    )

    while Producto.objects.filter(
        codigo_producto=f"{base}-{numero:03d}"
    ).exists():
        numero += 1

    return f"{base}-{numero:03d}"


# ============================================================
# REGISTRAR INGRESO
# ============================================================

@transaction.atomic
def registrar_ingreso(
    *,
    producto,
    cantidad,
    usuario,
    fecha_ingreso,
    seriales=None,
    observacion="",
    remision=None,  # <-- 1. Añadimos el parámetro remision aquí
):
    """
    Registra el ingreso de material al almacén.

    Productos serializados:
        - Si requieren número de serie físico:
          el usuario debe suministrar un serial por unidad.
        - Si NO requieren número de serie físico:
          Django genera únicamente el código interno.

    Productos por cantidad:
        - Solo se registra la cantidad.
    """

    if cantidad < 1:
        raise ErrorInventario(
            "La cantidad debe ser mayor que cero."
        )

    seriales = seriales or []

    inv = _inventario(producto)

    creadas = []

    # ========================================================
    # PRODUCTO SERIALIZADO
    # ========================================================

    if producto.es_serializado:

        prefijo = (
            producto.categoria.prefijo or ""
        ).strip().upper()

        if not prefijo:
            raise ErrorInventario(
                f"La categoría «{producto.categoria}» "
                "no tiene prefijo. Edítela antes de registrar unidades."
            )

        # ====================================================
        # PRODUCTO QUE REQUIERE SERIAL FÍSICO
        # EJEMPLO: MEDIDOR
        # ====================================================

        if producto.requiere_numero_serie:

            if len(seriales) != cantidad:
                raise ErrorInventario(
                    f"Se esperaban {cantidad} números de serie "
                    f"y se recibieron {len(seriales)}."
                )

            seriales_normalizados = [
                serie.strip().upper()
                for serie in seriales
            ]

            # No permitir duplicados dentro del mismo ingreso.
            if (
                len(seriales_normalizados)
                != len(set(seriales_normalizados))
            ):
                raise ErrorInventario(
                    "No puede repetir números de serie."
                )

            # Buscar seriales que ya existan.
            existentes = (
                UnidadSerializada.objects
                .filter(
                    numero_serie__in=seriales_normalizados
                )
                .values_list(
                    "numero_serie",
                    flat=True,
                )
            )

            existentes = {
                serie.upper()
                for serie in existentes
                if serie
            }

            repetidos = [
                serie
                for serie in seriales
                if serie.strip().upper() in existentes
            ]

            if repetidos:
                raise ErrorInventario(
                    "Ya existen estos números de serie: "
                    + ", ".join(repetidos)
                )

        # ====================================================
        # SERIALIZADO SIN SERIAL FÍSICO
        # EJEMPLO: REGULADOR / VALVULA
        # ====================================================

        else:

            if seriales:
                raise ErrorInventario(
                    "Este producto no requiere número de serie físico."
                )

        # ====================================================
        # GENERAR CÓDIGOS INTERNOS
        # ====================================================

        numero = _siguiente_numero(prefijo)

        for i in range(cantidad):

            numero_serie = None

            # Solo productos configurados con serial físico
            # reciben numero_serie.
            if producto.requiere_numero_serie:
                numero_serie = (
                    seriales[i]
                    .strip()
                    .upper()
                )

            unidad = UnidadSerializada.objects.create(
                producto=producto,
                codigo=f"{prefijo}-{numero + i:03d}",
                numero_serie=numero_serie,
                marca=producto.marca,
                referencia=producto.referencia,
                medida=producto.medida,
                fecha_ingreso=fecha_ingreso,
                estado=UnidadSerializada.Estado.DISPONIBLE,
                fecha_cambio_estado=timezone.now(),
                observaciones=observacion,
                remision=remision,
            )

            MovimientoInventario.objects.create(
                producto=producto,
                unidad_serializada=unidad,
                tipo_movimiento=Mov.ENTRADA,
                cantidad=1,
                usuario=usuario,
                observacion=observacion,
                remision=remision,  # <-- 3. Asociamos la remisión al movimiento de la unidad
            )

            creadas.append(unidad)

    # ========================================================
    # PRODUCTO POR CANTIDAD
    # ========================================================

    else:

        if seriales:
            raise ErrorInventario(
                "Este producto se maneja por cantidad "
                "y não utiliza números de serie."
            )

        MovimientoInventario.objects.create(
            producto=producto,
            tipo_movimiento=Mov.ENTRADA,
            cantidad=cantidad,
            usuario=usuario,
            observacion=observacion,
            remision=remision,      # <-- 4. Asociamos la remisión al movimiento de inventario general
        )

    # ========================================================
    # ACTUALIZAR INVENTARIO
    # ========================================================

    inv.cantidad_actual += cantidad
    inv.save()

    return creadas


# ============================================================
# ASIGNAR MATERIAL
# ============================================================

@transaction.atomic
def asignar(
    *,
    producto,
    tecnico,
    usuario,
    cantidad,
    unidad=None,
    fecha=None,
    observaciones="",
):
    """
    Entrega material a un técnico.

    Serializado:
        una unidad = cantidad 1.

    No serializado:
        se descuenta la cantidad solicitada
        del stock disponible.
    """

    if cantidad < 1:
        raise ErrorInventario(
            "La cantidad debe ser mayor que cero."
        )

    inv = _inventario(producto)

    if producto.es_serializado:

        if unidad is None:
            raise ErrorInventario(
                "Seleccione la unidad (código) que se entrega."
            )

        unidad = (
            UnidadSerializada.objects
            .select_for_update()
            .select_related("producto")
            .get(pk=unidad.pk)
        )

        if unidad.producto_id != producto.id:
            raise ErrorInventario(
                "La unidad seleccionada no pertenece al producto."
            )

        if unidad.estado != EstadoU.DISPONIBLE:
            raise ErrorInventario(
                f"La unidad {unidad.codigo} no está disponible "
                f"(estado: {unidad.get_estado_display()})."
            )

        cantidad = 1

        unidad.estado = EstadoU.ASIGNADO
        unidad.fecha_cambio_estado = timezone.now()

        unidad.save(
            update_fields=[
                "estado",
                "fecha_cambio_estado",
            ]
        )

    else:
        unidad = None

    if inv.cantidad_actual < cantidad:
        raise ErrorInventario(
            f"Stock insuficiente: hay "
            f"{inv.cantidad_actual} disponible(s) "
            f"y se piden {cantidad}."
        )

    asignacion = Asignacion.objects.create(
        tecnico=tecnico,
        producto=producto,
        unidad_serializada=unidad,
        cantidad=cantidad,
        fecha_asignacion=(
            fecha or timezone.localdate()
        ),
        observaciones=observaciones,
        usuario=usuario,
    )

    inv.cantidad_actual -= cantidad
    inv.cantidad_asignada += cantidad

    inv.save()

    MovimientoInventario.objects.create(
        producto=producto,
        unidad_serializada=unidad,
        asignacion=asignacion,
        tipo_movimiento=Mov.ASIGNACION,
        cantidad=cantidad,
        usuario=usuario,
        observacion=(
            f"Entregado a {tecnico}. "
            f"{observaciones}"
        ).strip(),
    )

    return asignacion


# ============================================================
# INSTALAR
# ============================================================

@transaction.atomic
def instalar(
    *,
    asignacion,
    usuario,
    fecha_instalacion,
    operacion,
    observaciones="",
    cantidad=None,
):
    """
    Registra la instalación de una asignación y la liga a una operación.

    Serializado:
        instala siempre la unidad completa.

    No serializado:
        puede instalarse toda la asignación o una parte.
    """

    a = (
        Asignacion.objects
        .select_for_update()
        .select_related(
            "producto",
            "unidad_serializada",
            "tecnico",
        )
        .get(pk=asignacion.pk)
    )

    if a.estado != EstadoA.ASIGNADA:
        raise ErrorInventario(
            "Esta asignación ya no está pendiente de instalar."
        )

    if a.unidad_serializada_id:
        cantidad = 1

    else:
        if cantidad is None:
            cantidad = a.cantidad

        if cantidad < 1 or cantidad > a.cantidad:
            raise ErrorInventario(
                f"La cantidad debe estar entre 1 y {a.cantidad}."
            )

        # Instalación parcial: se divide la asignación.
        if cantidad < a.cantidad:
            a.cantidad -= cantidad
            a.save(update_fields=["cantidad"])

            a = Asignacion.objects.create(
                tecnico=a.tecnico,
                producto=a.producto,
                cantidad=cantidad,
                estado=EstadoA.ASIGNADA,
                fecha_asignacion=a.fecha_asignacion,
                observaciones=a.observaciones,
                usuario=a.usuario,
            )

    a.estado = EstadoA.INSTALADA
    a.operacion = operacion
    a.fecha_instalacion = fecha_instalacion
    a.observaciones_instalacion = observaciones or None
    a.save()

    if a.unidad_serializada_id:
        unidad = (
            UnidadSerializada.objects
            .select_for_update()
            .get(pk=a.unidad_serializada_id)
        )

        if unidad.estado != EstadoU.ASIGNADO:
            raise ErrorInventario(
                f"La unidad {unidad.codigo} "
                "ya no está asignada al técnico."
            )

        unidad.estado = EstadoU.INSTALADO
        unidad.fecha_cambio_estado = timezone.now()

        unidad.save(
            update_fields=[
                "estado",
                "fecha_cambio_estado",
            ]
        )

    inv = _inventario(a.producto)

    inv.cantidad_asignada = max(
        inv.cantidad_asignada - cantidad,
        0,
    )
    inv.save(
        update_fields=[
            "cantidad_asignada",
            "fecha_actualizacion",
        ]
    )

    MovimientoInventario.objects.create(
        producto=a.producto,
        unidad_serializada=a.unidad_serializada,
        asignacion=a,
        tipo_movimiento=Mov.SALIDA,
        cantidad=cantidad,
        usuario=usuario,
        observacion=(
            f"{operacion.get_tipo_operacion_display()} · "
            f"orden {operacion.numero_orden} · "
            f"cuenta {operacion.numero_cuenta_contrato} · "
            f"técnico {a.tecnico}"
        ),
    )

    return a


@transaction.atomic
def obtener_o_crear_operacion(*, numero_orden, cuenta_contrato, tipo,
                              tecnico, usuario, fecha):
    numero_orden = numero_orden.strip().upper()
    cuenta = cuenta_contrato.strip()

    op, creada = (
        Operacion.objects.select_for_update()
        .get_or_create(
            numero_orden=numero_orden,
            defaults=dict(
                numero_cuenta_contrato=cuenta,
                tipo_operacion=tipo,
                tecnico=tecnico,
                estado=Operacion.Estado.EJECUTADA,
                origen=Operacion.Origen.MANUAL,
                fecha_operacion=fecha,
                fecha_ejecucion=fecha,
                usuario=usuario,
            ),
        )
    )

    if creada:
        return op

    # La orden ya existe (manual previa o, en etapa 2, cargada del Excel)
    if op.estado in (Operacion.Estado.CANCELADA, Operacion.Estado.ANULADA):
        raise ErrorInventario(f"La orden {numero_orden} está {op.get_estado_display().lower()}.")
    if op.numero_cuenta_contrato != cuenta:
        raise ErrorInventario(
            f"La orden {numero_orden} pertenece a la cuenta contrato "
            f"{op.numero_cuenta_contrato}, no a {cuenta}."
        )
    if op.tipo_operacion != tipo:
        raise ErrorInventario(
            f"La orden {numero_orden} es de tipo {op.get_tipo_operacion_display()}."
        )
    if op.tecnico_id and op.tecnico_id != tecnico.id:
        raise ErrorInventario(f"La orden {numero_orden} está asignada a {op.tecnico}.")

        op.tecnico = op.tecnico or tecnico

    # Solo avanza estados "previos"; no pisa EFECTIVA / NO_EJECUTADA.
    if op.estado in (
        Operacion.Estado.CREADA,
        Operacion.Estado.PENDIENTE,
        Operacion.Estado.ASIGNADA,
    ):
        op.estado = Operacion.Estado.EJECUTADA
        op.fecha_ejecucion = fecha

    op.save()
    return op

# ============================================================
# CAMBIAR ESTADO
# ============================================================

@transaction.atomic
def cambiar_estado(
    *,
    unidad,
    nuevo_estado,
    usuario,
    motivo,
):
    """
    Cambia el estado de una unidad serializada
    y registra el historial.
    """

    unidad = (
        UnidadSerializada.objects
        .select_for_update()
        .select_related("producto")
        .get(pk=unidad.pk)
    )

    anterior = unidad.estado

    valores_permitidos = {
        estado.value
        if hasattr(estado, "value")
        else estado
        for estado in TRANSICIONES.get(
            anterior,
            ()
        )
    }

    nuevo_valor = (
        nuevo_estado.value
        if hasattr(nuevo_estado, "value")
        else nuevo_estado
    )

    if nuevo_valor not in valores_permitidos:
        raise ErrorInventario(
            f"No se puede pasar de "
            f"«{unidad.get_estado_display()}» "
            "a ese estado."
        )

    etiqueta_anterior = (
        unidad.get_estado_display()
    )

    unidad.estado = nuevo_valor
    unidad.fecha_cambio_estado = timezone.now()

    unidad.save(
        update_fields=[
            "estado",
            "fecha_cambio_estado",
        ]
    )

    # El inventario actual representa
    # únicamente unidades disponibles.
    ajuste = 0

    if anterior == E.DISPONIBLE:
        ajuste = -1

    elif nuevo_valor == E.DISPONIBLE:
        ajuste = 1

    if ajuste:

        inv = _inventario(
            unidad.producto
        )

        inv.cantidad_actual = max(
            inv.cantidad_actual + ajuste,
            0,
        )

        inv.save(
            update_fields=[
                "cantidad_actual",
                "fecha_actualizacion",
            ]
        )

    MovimientoInventario.objects.create(
        producto=unidad.producto,
        unidad_serializada=unidad,
        tipo_movimiento=Mov.AJUSTE,
        cantidad=1,
        usuario=usuario,
        observacion=(
            f"Estado: {etiqueta_anterior} → "
            f"{unidad.get_estado_display()}. "
            f"{motivo}"
        ),
    )

    return unidad


# ============================================================
# DEVOLVER
# ============================================================

@transaction.atomic
def devolver(
    *,
    asignacion,
    usuario,
    observacion="",
):
    """
    Devuelve una asignación que todavía está
    en manos del técnico.
    """

    asignacion = (
        Asignacion.objects
        .select_for_update()
        .select_related(
            "producto",
            "unidad_serializada",
            "tecnico",
        )
        .get(pk=asignacion.pk)
    )

    if asignacion.estado != EstadoA.ASIGNADA:
        raise ErrorInventario(
            "Solo se pueden devolver asignaciones "
            "en estado «Asignada»."
        )

    inv = _inventario(
        asignacion.producto
    )

    if asignacion.unidad_serializada_id:

        unidad = (
            UnidadSerializada.objects
            .select_for_update()
            .get(pk=asignacion.unidad_serializada_id)
        )

        if unidad.estado != EstadoU.ASIGNADO:
            raise ErrorInventario(
                f"La unidad {unidad.codigo} "
                "no está actualmente asignada."
            )

        unidad.estado = EstadoU.DISPONIBLE
        unidad.fecha_cambio_estado = timezone.now()

        unidad.save(
            update_fields=[
                "estado",
                "fecha_cambio_estado",
            ]
        )

    asignacion.estado = EstadoA.DEVUELTA

    asignacion.save(
        update_fields=["estado"]
    )

    inv.cantidad_asignada = max(
        inv.cantidad_asignada - asignacion.cantidad,
        0,
    )

    inv.cantidad_actual += asignacion.cantidad

    inv.save()

    MovimientoInventario.objects.create(
        producto=asignacion.producto,
        unidad_serializada=(
            asignacion.unidad_serializada
        ),
        asignacion=asignacion,
        tipo_movimiento=Mov.DEVOLUCION,
        cantidad=asignacion.cantidad,
        usuario=usuario,
        observacion=(
            f"Devuelto por {asignacion.tecnico}. "
            f"{observacion}"
        ).strip(),
    )

    return asignacion