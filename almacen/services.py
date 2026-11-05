from django.db import transaction
from django.db.models import Q
from django.utils import timezone

from operaciones.models import Operacion

from .models import (
    Asignacion,
    Inconsistencia,
    Inventario,
    MovimientoInventario,
    Producto,
    UnidadSerializada,
    RemisionIngreso,
)
T = Inconsistencia.Tipo
REGULARIZAR_AUTOMATICO = False
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

@transaction.atomic
def obtener_o_crear_remision(
    *,
    numero,
    usuario,
    imagen=None,
    fecha_recibo_bsi=None,
    recogido_por="",
    recibido_almacen_por="",
):
    """
    Busca la remisión por número o la crea con sus datos de recepción.
    Si ya existía, solo completa lo que esté vacío; nunca pisa lo guardado.
    """
    remision, creada = RemisionIngreso.objects.get_or_create(
        numero_remision=numero,
        defaults={
            "comprobante_imagen": imagen,
            "fecha_recibo_bsi": fecha_recibo_bsi,
            "recogido_por": recogido_por or "",
            "recibido_almacen_por": recibido_almacen_por or "",
            "usuario": usuario,
        },
    )

    if not creada:
        cambios = []

        if imagen and not remision.comprobante_imagen:
            remision.comprobante_imagen = imagen
            cambios.append("comprobante_imagen")
        if fecha_recibo_bsi and not remision.fecha_recibo_bsi:
            remision.fecha_recibo_bsi = fecha_recibo_bsi
            cambios.append("fecha_recibo_bsi")
        if recogido_por and not remision.recogido_por:
            remision.recogido_por = recogido_por
            cambios.append("recogido_por")
        if recibido_almacen_por and not remision.recibido_almacen_por:
            remision.recibido_almacen_por = recibido_almacen_por
            cambios.append("recibido_almacen_por")

        if cambios:
            remision.save(update_fields=cambios)

    return remision

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
    remision=None,
    seriales=None,
    observacion="",
):
    """
    Registra el ingreso de material al almacén, asociado a una remisión.

    Productos serializados:
        - Si requieren número de serie físico:
          el usuario debe suministrar un serial por unidad.
        - Si NO requieren número de serie físico:
          Django genera únicamente el código interno.

    Productos por cantidad:
        - Solo se registra la cantidad.
    """

    if cantidad < 1:
        raise ErrorInventario("La cantidad debe ser mayor que cero.")

    seriales = seriales or []

    inv = _inventario(producto)

    creadas = []

    # ========================================================
    # PRODUCTO SERIALIZADO
    # ========================================================

    if producto.es_serializado:

        prefijo = (producto.categoria.prefijo or "").strip().upper()

        if not prefijo:
            raise ErrorInventario(
                f"La categoría «{producto.categoria}» "
                "no tiene prefijo. Edítela antes de registrar unidades."
            )

        # ----- Producto que requiere serial físico (ej. medidor) -----

        if producto.requiere_numero_serie:

            if len(seriales) != cantidad:
                raise ErrorInventario(
                    f"Se esperaban {cantidad} números de serie "
                    f"y se recibieron {len(seriales)}."
                )

            seriales_normalizados = [s.strip().upper() for s in seriales]

            # No permitir duplicados dentro del mismo ingreso.
            if len(seriales_normalizados) != len(set(seriales_normalizados)):
                raise ErrorInventario("No puede repetir números de serie.")

            # Buscar seriales que ya existan.
            existentes = {
                serie.upper()
                for serie in (
                    UnidadSerializada.objects
                    .filter(numero_serie__in=seriales_normalizados)
                    .values_list("numero_serie", flat=True)
                )
                if serie
            }

            repetidos = [s for s in seriales_normalizados if s in existentes]

            if repetidos:
                raise ErrorInventario(
                    "Ya existen estos números de serie: "
                    + ", ".join(repetidos)
                )

        # ----- Serializado sin serial físico (ej. regulador, válvula) -----

        else:
            seriales_normalizados = []

            if seriales:
                raise ErrorInventario(
                    "Este producto no requiere número de serie físico."
                )

        # ----- Generar códigos internos -----

        numero = _siguiente_numero(prefijo)

        for i in range(cantidad):

            numero_serie = (
                seriales_normalizados[i]
                if producto.requiere_numero_serie
                else None
            )

            unidad = UnidadSerializada.objects.create(
                producto=producto,
                remision=remision,
                codigo=f"{prefijo}-{numero + i:03d}",
                numero_serie=numero_serie,
                marca=producto.marca,
                referencia=producto.referencia,
                medida=producto.medida,
                fecha_ingreso=fecha_ingreso,
                estado=UnidadSerializada.Estado.DISPONIBLE,
                fecha_cambio_estado=timezone.now(),
                observaciones=observacion,
            )

            MovimientoInventario.objects.create(
                producto=producto,
                unidad_serializada=unidad,
                remision=remision,
                tipo_movimiento=Mov.ENTRADA,
                cantidad=1,
                usuario=usuario,
                observacion=observacion,
            )

            creadas.append(unidad)

    # ========================================================
    # PRODUCTO POR CANTIDAD
    # ========================================================

    else:

        if seriales:
            raise ErrorInventario(
                "Este producto se maneja por cantidad "
                "y no utiliza números de serie."
            )

        MovimientoInventario.objects.create(
            producto=producto,
            remision=remision,
            tipo_movimiento=Mov.ENTRADA,
            cantidad=cantidad,
            usuario=usuario,
            observacion=observacion,
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
        observacion=(observaciones or "").strip(),
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
    Instala una asignación y la vincula a una operación.

    Serializados:
        Siempre se instala una unidad completa.

    Por cantidad:
        Permite instalaciones parciales.
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

    # --------------------------------------------------------
    # UNIDADES SERIALIZADAS
    # --------------------------------------------------------

    if a.unidad_serializada_id:

        if cantidad not in (None, 1):
            raise ErrorInventario(
                "Una unidad serializada solo puede instalarse "
                "en cantidad de una."
            )

        cantidad = 1

        unidad = (
            UnidadSerializada.objects
            .select_for_update()
            .get(pk=a.unidad_serializada_id)
        )

        if unidad.estado != EstadoU.ASIGNADO:
            raise ErrorInventario(
                f"La unidad {unidad.codigo} no está asignada "
                "actualmente."
            )

        unidad.estado = EstadoU.INSTALADO
        unidad.fecha_cambio_estado = timezone.now()

        unidad.save(
            update_fields=[
                "estado",
                "fecha_cambio_estado",
            ]
        )

    # --------------------------------------------------------
    # PRODUCTOS POR CANTIDAD
    # --------------------------------------------------------

    else:

        if cantidad is None:
            cantidad = a.cantidad

        if not isinstance(cantidad, int) or isinstance(cantidad, bool):
            raise ErrorInventario(
                "La cantidad de instalación debe ser un entero."
            )

        if cantidad < 1 or cantidad > a.cantidad:
            raise ErrorInventario(
                f"La cantidad debe estar entre 1 y {a.cantidad}."
            )

        # Instalación parcial: separar la cantidad instalada.
        if cantidad < a.cantidad:

            a.cantidad -= cantidad

            a.save(
                update_fields=[
                    "cantidad",
                    "fecha_modificacion",
                ]
            )

            a = Asignacion.objects.create(
                tecnico=a.tecnico,
                producto=a.producto,
                cantidad=cantidad,
                cantidad_instalada=cantidad,
                estado=EstadoA.INSTALADA,
                fecha_asignacion=a.fecha_asignacion,
                observaciones=a.observaciones,
                usuario=a.usuario,
            )

    # --------------------------------------------------------
    # ACTUALIZAR ASIGNACIÓN
    # --------------------------------------------------------

    a.estado = EstadoA.INSTALADA
    a.cantidad_instalada = cantidad
    a.operacion = operacion
    a.fecha_instalacion = fecha_instalacion
    a.observaciones_instalacion = observaciones or None

    a.save(
        update_fields=[
            "estado",
            "cantidad_instalada",
            "operacion",
            "fecha_instalacion",
            "observaciones_instalacion",
            "fecha_modificacion",
        ]
    )

    # --------------------------------------------------------
    # ACTUALIZAR INVENTARIO ASIGNADO
    # --------------------------------------------------------

    inventario = _inventario(a.producto)

    if inventario.cantidad_asignada < cantidad:
        raise ErrorInventario(
            f"El inventario de {a.producto.nombre_completo} "
            "no tiene suficientes unidades registradas como asignadas "
            "para completar esta instalación."
        )

    inventario.cantidad_asignada -= cantidad

    inventario.save(
        update_fields=[
            "cantidad_asignada",
            "fecha_actualizacion",
        ]
    )

    # --------------------------------------------------------
    # REGISTRAR MOVIMIENTO
    # --------------------------------------------------------

    MovimientoInventario.objects.create(
        producto=a.producto,
        unidad_serializada=a.unidad_serializada,
        asignacion=a,
        tipo_movimiento=Mov.INSTALACION,
        cantidad=cantidad,
        usuario=usuario,
        observacion=(observaciones or "").strip(),
    )

    return a


def _sumar_observacion(actual, nueva):
    nueva = (nueva or "").strip()
    if not actual:
        return nueva
    if nueva in actual:
        return actual
    return f"{actual}\n[{timezone.localdate():%d/%m/%Y}] {nueva}"


@transaction.atomic
def obtener_o_crear_operacion(*, numero_orden, cuenta_contrato, tipo,
                              tecnico, usuario, fecha, estado, anomalia,
                              observaciones, archivo=None):
    numero_orden = numero_orden.strip().upper()
    cuenta = cuenta_contrato.strip()
    ejecutada = estado in Operacion.ESTADOS_QUE_INSTALAN

    op, creada = (
        Operacion.objects.select_for_update()
        .get_or_create(
            numero_orden=numero_orden,
            defaults=dict(
                numero_cuenta_contrato=cuenta,
                tipo_operacion=tipo,
                tecnico=tecnico,
                estado=estado,
                anomalia=anomalia,
                observaciones=(observaciones or "").strip(),
                archivo_soporte=archivo,
                origen=Operacion.Origen.MANUAL,
                fecha_operacion=fecha,
                fecha_ejecucion=fecha if ejecutada else None,
                usuario=usuario,
            ),
        )
    )

    if creada:
        return op

    # La orden ya existía (manual previa o cargada del Excel)
    if op.estado in (Operacion.Estado.CANCELADA, Operacion.Estado.ANULADA):
        raise ErrorInventario(
            f"La orden {numero_orden} está {op.get_estado_display().lower()}."
        )
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

    # Solo avanza desde estados "previos"; no pisa EFECTIVA / NO_EJECUTADA.
    if op.estado in (
        Operacion.Estado.CREADA,
        Operacion.Estado.PENDIENTE,
        Operacion.Estado.ASIGNADA,
    ):
        op.estado = estado
        op.fecha_ejecucion = fecha if ejecutada else None

    op.anomalia = anomalia
    op.observaciones = _sumar_observacion(op.observaciones, observaciones)
    if archivo:
        op.archivo_soporte = archivo   # el save() del model borra el anterior

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
# DEVOLVER UNIDAD SERIALIZADA
# ============================================================

@transaction.atomic
def devolver(
    *,
    asignacion,
    usuario,
    observacion,
    estado_retorno=EstadoU.DISPONIBLE,
):
    """
    Devuelve una unidad serializada al almacén.

    DISPONIBLE: aumenta el stock utilizable.
    DAÑADO: aumenta el inventario dañado.
    """

    observacion = (observacion or "").strip()

    if not observacion:
        raise ErrorInventario(
            "La observación es obligatoria para registrar la devolución."
        )

    if estado_retorno not in (
        EstadoU.DISPONIBLE,
        EstadoU.DAÑADO,
    ):
        raise ErrorInventario("Estado de devolución no válido.")

    a = (
        Asignacion.objects
        .select_for_update()
        .select_related("producto", "unidad_serializada", "tecnico")
        .get(pk=asignacion.pk)
    )

    if a.estado != EstadoA.ASIGNADA:
        raise ErrorInventario(
            "Solo se pueden devolver asignaciones pendientes."
        )

    if not a.unidad_serializada_id:
        raise ErrorInventario(
            "Esta devolución requiere una unidad serializada."
        )

    unidad = (
        UnidadSerializada.objects
        .select_for_update()
        .get(pk=a.unidad_serializada_id)
    )

    if unidad.estado != EstadoU.ASIGNADO:
        raise ErrorInventario(
            f"La unidad {unidad.codigo} no está actualmente asignada."
        )

    inv = _inventario(a.producto)

    if inv.cantidad_asignada < 1:
        raise ErrorInventario(
            f"El inventario de {a.producto.nombre_completo} "
            "no tiene unidades asignadas suficientes."
        )

    # Cambiar estado de la unidad.
    unidad.estado = estado_retorno
    unidad.fecha_cambio_estado = timezone.now()

    unidad.observaciones = _sumar_observacion(
        unidad.observaciones,
        f"Devolución: {observacion}"
    )

    unidad.save(
        update_fields=[
            "estado",
            "fecha_cambio_estado",
            "observaciones",
        ]
    )

    # Cerrar asignación.
    a.estado = EstadoA.DEVUELTA

    a.observaciones = _sumar_observacion(
        a.observaciones,
        f"Devolución: {observacion}"
    )

    a.save(
        update_fields=[
            "estado",
            "observaciones",
            "fecha_modificacion",
        ]
    )

    # Actualizar inventario.
    inv.cantidad_asignada -= 1

    if estado_retorno == EstadoU.DISPONIBLE:
        inv.cantidad_actual += 1
    else:
        inv.cantidad_danada += 1

    inv.save(
        update_fields=[
            "cantidad_asignada",
            "cantidad_actual",
            "cantidad_danada",
        ]
    )

    # Registrar movimiento.
    MovimientoInventario.objects.create(
        producto=a.producto,
        unidad_serializada=unidad,
        asignacion=a,
        tipo_movimiento=Mov.DEVOLUCION,
        cantidad=1,
        usuario=usuario,
        observacion=(
            f"Devuelto por {a.tecnico}. "
            f"Estado de recepción: {unidad.get_estado_display()}. "
            f"Observación: {observacion}"
        ),
    )

    return a
# ============================================================
# DEVOLVER MATERIAL POR CANTIDAD
# ============================================================

# ============================================================
# DEVOLVER MATERIAL POR CANTIDAD
# ============================================================

@transaction.atomic
def devolver_cantidad(
    *,
    tecnico,
    producto,
    cantidad,
    usuario,
    observacion,
    estado_retorno=EstadoU.DISPONIBLE,
    detalle_devoluciones=None,
):
    """
    Devuelve material no serializado.

    Registra devoluciones parciales y completas, conservando
    las cantidades instaladas y generando el detalle para el comprobante.
    """

    observacion = (observacion or "").strip()

    if not observacion:
        raise ErrorInventario(
            "La observación es obligatoria para registrar la devolución."
        )

    if estado_retorno not in (
        EstadoU.DISPONIBLE,
        EstadoU.DAÑADO,
    ):
        raise ErrorInventario("Estado de devolución no válido.")

    if isinstance(cantidad, bool) or not isinstance(cantidad, int):
        raise ErrorInventario("La cantidad debe ser un número entero.")

    if cantidad <= 0:
        raise ErrorInventario("La cantidad debe ser mayor que cero.")

    if producto.es_serializado:
        raise ErrorInventario(
            "Este producto debe devolverse seleccionando sus códigos individuales."
        )

    if detalle_devoluciones is None:
        detalle_devoluciones = []

    # Bloquear asignaciones pendientes.
    asigs = list(
        Asignacion.objects
        .select_for_update()
        .filter(
            tecnico=tecnico,
            producto=producto,
            estado=EstadoA.ASIGNADA,
            unidad_serializada__isnull=True,
        )
        .order_by("pk")
    )

    total_pendiente = sum(a.cantidad_pendiente for a in asigs)

    if cantidad > total_pendiente:
        raise ErrorInventario(
            f"{producto.nombre_completo}: el técnico tiene "
            f"{total_pendiente} unidad(es) pendientes, "
            f"pero intentas devolver {cantidad}."
        )

    # Bloquear inventario.
    inv = (
        Inventario.objects
        .select_for_update()
        .get(producto=producto)
    )

    if inv.cantidad_asignada < cantidad:
        raise ErrorInventario(
            f"El inventario de {producto.nombre_completo} "
            "no tiene suficientes unidades asignadas."
        )

    resto = cantidad

    for a in asigs:
        if resto <= 0:
            break

        pendiente = a.cantidad_pendiente

        if pendiente <= 0:
            continue

        tomar = min(resto, pendiente)
        nueva_cantidad = a.cantidad - tomar

        # Devolución parcial: separar la cantidad devuelta.
        if nueva_cantidad > 0:
            a.cantidad = nueva_cantidad

            if a.cantidad_instalada >= nueva_cantidad:
                a.estado = EstadoA.INSTALADA

            a.save(
                update_fields=[
                    "cantidad",
                    "estado",
                    "fecha_modificacion",
                ]
            )

            devuelta = Asignacion.objects.create(
                tecnico=a.tecnico,
                producto=a.producto,
                cantidad=tomar,
                cantidad_instalada=0,
                estado=EstadoA.DEVUELTA,
                fecha_asignacion=a.fecha_asignacion,
                observaciones=_sumar_observacion(
                    a.observaciones,
                    f"Devolución: {observacion}",
                ),
                usuario=a.usuario,
            )

        # Devolución completa: la asignación original queda devuelta.
        else:
            devuelta = a
            devuelta.estado = EstadoA.DEVUELTA
            devuelta.observaciones = _sumar_observacion(
                devuelta.observaciones,
                f"Devolución: {observacion}",
            )
            devuelta.save(
                update_fields=[
                    "estado",
                    "observaciones",
                    "fecha_modificacion",
                ]
            )

        # IMPORTANTE: registrar tanto las parciales como las completas.
        detalle_devoluciones.append({
            "asignacion": devuelta,
            "producto": producto,
            "cantidad": tomar,
        })

        estado_display = (
            "Disponible"
            if estado_retorno == EstadoU.DISPONIBLE
            else "Dañado"
        )

        MovimientoInventario.objects.create(
            producto=producto,
            asignacion=devuelta,
            tipo_movimiento=Mov.DEVOLUCION,
            cantidad=tomar,
            usuario=usuario,
            observacion=(
                f"Devuelto por {tecnico}. "
                f"Estado de recepción: {estado_display}. "
                f"Observación: {observacion}"
            ),
        )

        resto -= tomar

    if resto != 0:
        raise ErrorInventario(
            "No fue posible completar la devolución."
        )

    # Actualizar inventario una sola vez.
    inv.cantidad_asignada -= cantidad

    if estado_retorno == EstadoU.DISPONIBLE:
        inv.cantidad_actual += cantidad
    else:
        inv.cantidad_danada += cantidad

    inv.save(
        update_fields=[
            "cantidad_asignada",
            "cantidad_actual",
            "cantidad_danada",
        ]
    )

    return cantidad


# ============================================================
# INSTALACIÓN CON INCONSISTENCIAS
# ============================================================

# Estos tipos solo se reportan: no se instala ni se mueve inventario.
TIPOS_SOLO_REPORTE = {T.NO_REGISTRADO, T.YA_INSTALADO, T.NO_UTILIZABLE, T.SIN_STOCK}
def _resolver_novedad(novedad, usuario, nota):
    novedad.estado = Inconsistencia.Estado.RESUELTA
    novedad.resuelto_por = usuario
    novedad.fecha_resolucion = timezone.now()
    novedad.nota_resolucion = nota
    novedad.save()


def _revertir_serial(*, asignacion, usuario, motivo):
    """Devuelve la unidad al lugar donde estaba antes de instalarse."""
    a = (Asignacion.objects.select_for_update()
         .select_related("producto", "tecnico").get(pk=asignacion.pk))
    if a.estado != EstadoA.INSTALADA:
        raise ErrorInventario("La asignación ya no figura instalada.")

    unidad = UnidadSerializada.objects.select_for_update().get(pk=a.unidad_serializada_id)
    if unidad.estado != EstadoU.INSTALADO:
        raise ErrorInventario(f"La unidad {unidad.codigo} no figura instalada.")

    novedad = (Inconsistencia.objects.select_for_update()
               .filter(asignacion_regularizada=a, instalado=True)
               .order_by("-pk").first())
    inv = _inventario(a.producto)
    ahora = timezone.now()

    if novedad and novedad.estado_anterior == EstadoU.DISPONIBLE:
        # Estaba en almacén: vuelve al almacén
        unidad.estado = EstadoU.DISPONIBLE
        inv.cantidad_actual += 1
        a.estado = EstadoA.DEVUELTA            # conserva la operación: trazabilidad
        a.save(update_fields=["estado", "fecha_modificacion"])
        destino = "almacén (estaba disponible antes de instalarse)"

    elif novedad and novedad.tecnico_anterior_id:
        # Era de otro técnico: se le devuelve a él
        unidad.estado = EstadoU.ASIGNADO
        inv.cantidad_asignada += 1
        a.estado = EstadoA.DEVUELTA
        a.save(update_fields=["estado", "fecha_modificacion"])
        Asignacion.objects.create(
            tecnico=novedad.tecnico_anterior, producto=a.producto,
            unidad_serializada=unidad, cantidad=1, estado=EstadoA.ASIGNADA,
            fecha_asignacion=timezone.localdate(), usuario=usuario,
            observaciones=f"Restaurada al quitarla de la orden {a.operacion.numero_orden}.")
        destino = f"{novedad.tecnico_anterior} (quien la tenía antes)"

    else:
        # Estaba asignada al mismo técnico
        unidad.estado = EstadoU.ASIGNADO
        inv.cantidad_asignada += 1
        a.estado = EstadoA.ASIGNADA
        a.cantidad_instalada = 0          # <-- nueva
        a.operacion = None
        a.fecha_instalacion = None
        a.observaciones_instalacion = None
        a.save()
        destino = f"{a.tecnico} (pendiente de instalar)"

    unidad.fecha_cambio_estado = ahora
    unidad.save(update_fields=["estado", "fecha_cambio_estado"])
    inv.save()

    MovimientoInventario.objects.create(
        producto=a.producto, unidad_serializada=unidad, asignacion=a,
        tipo_movimiento=Mov.AJUSTE, cantidad=1, usuario=usuario,
        observacion=f"{motivo} Regresó a: {destino}.")

    if novedad:
        _resolver_novedad(novedad, usuario, f"Instalación revertida. Regresó a: {destino}.")


def _devolver_instalacion_a_almacen(*, asignacion, cantidad, usuario, motivo):
    """Revierte material por cantidad que se había sacado del stock al regularizar."""
    a = Asignacion.objects.select_for_update().get(pk=asignacion.pk)
    if a.estado != EstadoA.INSTALADA or cantidad < 1 or cantidad > a.cantidad:
        raise ErrorInventario("No se puede revertir esa cantidad.")

    if cantidad < a.cantidad:
        a.cantidad -= cantidad
        a.save(update_fields=["cantidad", "fecha_modificacion"])
        a = Asignacion.objects.create(
            tecnico=a.tecnico, producto=a.producto, cantidad=cantidad,
            estado=EstadoA.DEVUELTA, operacion=a.operacion,
            fecha_asignacion=a.fecha_asignacion,
            fecha_instalacion=a.fecha_instalacion, usuario=a.usuario)
        parcial = True
    else:
        a.estado = EstadoA.DEVUELTA
        a.save(update_fields=["estado", "fecha_modificacion"])
        parcial = False

    inv = _inventario(a.producto)
    inv.cantidad_actual += cantidad
    inv.save(update_fields=["cantidad_actual", "fecha_actualizacion"])

    MovimientoInventario.objects.create(
        producto=a.producto, asignacion=a, tipo_movimiento=Mov.AJUSTE,
        cantidad=cantidad, usuario=usuario,
        observacion=f"{motivo} {cantidad} regresaron al almacén (no estaban asignadas).")

    for n in Inconsistencia.objects.select_for_update().filter(
            asignacion_regularizada_id=asignacion.pk, instalado=True,
            estado=Inconsistencia.Estado.PENDIENTE):
        if parcial and n.cantidad > cantidad:
            n.cantidad -= cantidad
            n.detalle += f" [Se revirtieron {cantidad} al editar la orden.]"
            n.save()
        else:
            _resolver_novedad(n, usuario, "Instalación revertida: regresó al almacén.")


def _revertir_cantidad(*, asignacion, cantidad, usuario, motivo):
    """Material por cantidad que SÍ estaba asignado: vuelve al técnico."""
    a = Asignacion.objects.select_for_update().get(pk=asignacion.pk)
    if a.estado != EstadoA.INSTALADA or cantidad < 1 or cantidad > a.cantidad:
        raise ErrorInventario("No se puede revertir esa cantidad.")

    if cantidad < a.cantidad:
        a.cantidad -= cantidad
        a.cantidad_instalada = a.cantidad      # lo que sigue instalado
        a.save(update_fields=["cantidad", "cantidad_instalada", "fecha_modificacion"])
        a = Asignacion.objects.create(
            tecnico=a.tecnico, producto=a.producto, cantidad=cantidad,
            cantidad_instalada=0, estado=EstadoA.ASIGNADA,
            fecha_asignacion=a.fecha_asignacion,
            observaciones=a.observaciones, usuario=a.usuario)
    else:
        a.estado = EstadoA.ASIGNADA
        a.cantidad_instalada = 0
        a.operacion = None
        a.fecha_instalacion = None
        a.observaciones_instalacion = None
        a.save()

    inv = _inventario(a.producto)
    inv.cantidad_asignada += cantidad
    inv.save(update_fields=["cantidad_asignada", "fecha_actualizacion"])

    MovimientoInventario.objects.create(
        producto=a.producto, asignacion=a, tipo_movimiento=Mov.AJUSTE,
        cantidad=cantidad, usuario=usuario,
        observacion=f"{motivo} {cantidad} regresaron al técnico {a.tecnico}.")


@transaction.atomic
def sincronizar_materiales_operacion(*, operacion, payload, usuario):
    """Deja los materiales de la operación iguales al payload, aplicando solo diferencias."""
    tecnico = operacion.tecnico
    if not tecnico:
        raise ErrorInventario("La operación no tiene un técnico asignado.")

    fecha = operacion.fecha_ejecucion or operacion.fecha_operacion
    motivo = f"Edición de la orden {operacion.numero_orden}."

    # Las novedades que solo se reportaron se reevalúan: las que siguen
    # siendo válidas se vuelven a crear más abajo.
    Inconsistencia.objects.filter(
        operacion=operacion, estado=Inconsistencia.Estado.PENDIENTE, instalado=False,
    ).update(estado=Inconsistencia.Estado.RESUELTA, resuelto_por=usuario,
             fecha_resolucion=timezone.now(),
             nota_resolucion="Reevaluada al editar la orden.")

    actuales = list(Asignacion.objects.select_for_update()
                    .filter(operacion=operacion, estado=EstadoA.INSTALADA)
                    .order_by("-pk"))
    por_unidad = {a.unidad_serializada_id: a for a in actuales if a.unidad_serializada_id}
    por_producto = {}
    for a in actuales:
        if not a.unidad_serializada_id:
            por_producto.setdefault(a.producto_id, []).append(a)

    # ----- lo que se pide ahora -----
    deseado_cant, deseado_serial = {}, []
    for item in payload:
        try:
            pid = int(item.get("producto_id"))
        except (TypeError, ValueError):
            raise ErrorInventario("Hay un material sin producto válido.")
        producto = Producto.objects.filter(pk=pid).first()
        if not producto:
            raise ErrorInventario(f"No existe el producto con ID {pid}.")

        if producto.es_serializado:
            vistos = set()
            for s in item.get("seriales") or []:
                s = str(s).strip()
                if not s:
                    continue
                if s.lower() in vistos:
                    raise ErrorInventario(f"Serial repetido en {producto.nombre_completo}: {s}.")
                vistos.add(s.lower())
                deseado_serial.append((pid, s))
        else:
            cant = int(item.get("cantidad") or 0)
            if cant > 0:
                deseado_cant[pid] = deseado_cant.get(pid, 0) + cant

    novedades = []

    # ----- SERIALIZADOS -----
    conservar, por_instalar = set(), []
    for pid, serial in deseado_serial:
        unidad = (UnidadSerializada.objects
                  .filter(Q(numero_serie__iexact=serial) | Q(codigo__iexact=serial))
                  .first())
        if unidad and unidad.pk in por_unidad:
            conservar.add(unidad.pk)
        else:
            por_instalar.append((pid, serial))

    for uid, a in por_unidad.items():            # primero se quita lo que sobra
        if uid not in conservar:
            _revertir_serial(asignacion=a, usuario=usuario, motivo=motivo)

    for pid, serial in por_instalar:             # luego se instala lo nuevo
        n = instalar_serial(serial=serial, producto_id=pid, tecnico=tecnico,
                            usuario=usuario, operacion=operacion, fecha=fecha)
        if n:
            novedades.append(n)

    # ----- POR CANTIDAD -----
    reg_ids = set(Inconsistencia.objects.filter(
        operacion=operacion, instalado=True, asignacion_regularizada__isnull=False,
    ).values_list("asignacion_regularizada_id", flat=True))

    for pid in set(deseado_cant) | set(por_producto):
        asigs = sorted(por_producto.get(pid, []),
                       key=lambda a: (a.pk not in reg_ids, -a.pk))   # primero las regularizadas
        actual = sum(a.cantidad for a in asigs)
        deseado = deseado_cant.get(pid, 0)

        if deseado < actual:
            resto = actual - deseado
            for a in asigs:
                if resto <= 0:
                    break
                tomar = min(resto, a.cantidad)
                if a.pk in reg_ids:
                    _devolver_instalacion_a_almacen(asignacion=a, cantidad=tomar,
                                                    usuario=usuario, motivo=motivo)
                else:
                    _revertir_cantidad(asignacion=a, cantidad=tomar,
                                       usuario=usuario, motivo=motivo)
                resto -= tomar

        elif deseado > actual:
            novedades.extend(instalar_cantidad(
                producto=Producto.objects.get(pk=pid), cantidad=deseado - actual,
                tecnico=tecnico, usuario=usuario, operacion=operacion, fecha=fecha))

    return novedades

@transaction.atomic
def revertir_instalacion(
    *,
    asignacion,
    usuario,
):
    """
    Revierte una instalación y recupera el estado original
    de la unidad serializada.

    Si tenía una asignación anterior, la restaura.
    Si estaba disponible, vuelve a almacén.

    Si no hay registro del estado anterior, se considera
    asignada al técnico de la instalación.
    """

    asignacion = (
        Asignacion.objects
        .select_for_update()
        .select_related(
            "producto",
            "unidad_serializada",
            "tecnico",
            "operacion",
        )
        .get(pk=asignacion.pk)
    )

    if asignacion.estado != EstadoA.INSTALADA:
        raise ErrorInventario(
            "Solo se pueden revertir instalaciones."
        )

    if not asignacion.unidad_serializada_id:
        raise ErrorInventario(
            "Esta función está destinada a unidades serializadas. "
            "El material por cantidad requiere una reversión parcial "
            "o total específica."
        )

    unidad = (
        UnidadSerializada.objects
        .select_for_update()
        .get(pk=asignacion.unidad_serializada_id)
    )

    if unidad.estado != EstadoU.INSTALADO:
        raise ErrorInventario(
            f"La unidad {unidad.codigo} no figura instalada."
        )

    # Buscar la inconsistencia que originó la regularización.
    novedad = (
        Inconsistencia.objects
        .select_for_update()
        .filter(
            unidad_serializada=unidad,
            operacion=asignacion.operacion,
            instalado=True,
        )
        .order_by("-pk")
        .first()
    )

    # --------------------------------------------------------
    # DETERMINAR ESTADO ORIGINAL
    # --------------------------------------------------------

    estado_anterior = (
        novedad.estado_anterior
        if novedad and novedad.estado_anterior
        else EstadoU.ASIGNADO
    )

    asignacion_anterior = (
        novedad.asignacion_anterior
        if novedad and novedad.asignacion_anterior_id
        else None
    )

    inv = _inventario(asignacion.producto)

    # --------------------------------------------------------
    # EVITAR QUE LA ASIGNACIÓN DE INSTALACIÓN SIGA ACTIVA
    # --------------------------------------------------------

    asignacion.estado = EstadoA.DEVUELTA

    asignacion.save(
        update_fields=[
            "estado",
            "fecha_modificacion",
        ]
    )

    # --------------------------------------------------------
    # RESTAURAR ASIGNACIÓN ORIGINAL
    # --------------------------------------------------------

    if (
        estado_anterior == EstadoU.ASIGNADO
        and asignacion_anterior
    ):

        original = (
            Asignacion.objects
            .select_for_update()
            .get(pk=asignacion_anterior.pk)
        )

        if original.estado != EstadoA.DEVUELTA:
            raise ErrorInventario(
                "La asignación original ya no está devuelta. "
                "Revise el historial antes de revertir."
            )

        original.estado = EstadoA.ASIGNADA

        original.save(
            update_fields=[
                "estado",
                "fecha_modificacion",
            ]
        )

        unidad.estado = EstadoU.ASIGNADO

        inv.cantidad_asignada += 1

        detalle = (
            f"Instalación revertida. Unidad restaurada "
            f"como asignada a {original.tecnico}."
        )

    # --------------------------------------------------------
    # RESTAURAR DISPONIBILIDAD EN ALMACÉN
    # --------------------------------------------------------

    elif estado_anterior == EstadoU.DISPONIBLE:

        unidad.estado = EstadoU.DISPONIBLE

        inv.cantidad_actual += 1

        detalle = (
            "Instalación revertida. Unidad devuelta "
            "al inventario disponible del almacén."
        )

    # --------------------------------------------------------
    # ESTADO PREDETERMINADO: ASIGNADO AL TÉCNICO
    # --------------------------------------------------------

    else:

        unidad.estado = EstadoU.ASIGNADO

        # Si no hay asignación original, recuperar la asignación
        # de la instalación para el técnico que la realizó.
        asignacion.estado = EstadoA.ASIGNADA

        asignacion.save(
            update_fields=[
                "estado",
                "fecha_modificacion",
            ]
        )

        inv.cantidad_asignada += 1

        detalle = (
            f"Instalación revertida. Unidad nuevamente "
            f"asignada a {asignacion.tecnico}."
        )

    unidad.fecha_cambio_estado = timezone.now()

    unidad.save(
        update_fields=[
            "estado",
            "fecha_cambio_estado",
        ]
    )

    inv.save(
        update_fields=[
            "cantidad_actual",
            "cantidad_asignada",
            "fecha_actualizacion",
        ]
    )

    # --------------------------------------------------------
    # REGISTRAR MOVIMIENTO
    # --------------------------------------------------------

    MovimientoInventario.objects.create(
        producto=unidad.producto,
        unidad_serializada=unidad,
        asignacion=(
            asignacion_anterior
            if estado_anterior == EstadoU.ASIGNADO
            and asignacion_anterior
            else asignacion
        ),
        tipo_movimiento=Mov.AJUSTE,
        cantidad=1,
        usuario=usuario,
        observacion=detalle,
    )

    # Marcar inconsistencia como resuelta si corresponde.
    if novedad:
        novedad.estado = Inconsistencia.Estado.RESUELTA
        novedad.resuelto_por = usuario
        novedad.fecha_resolucion = timezone.now()
        novedad.nota_resolucion = detalle

        novedad.save(
            update_fields=[
                "estado",
                "resuelto_por",
                "fecha_resolucion",
                "nota_resolucion",
            ]
        )

    return asignacion
    
def _novedad(
    *, tipo, tecnico, usuario, operacion, detalle, producto=None, unidad=None,
    serial="", cantidad=1, instalado=False, asignacion_anterior=None,
    estado_anterior=None, tecnico_anterior=None, asignacion_regularizada=None,
):
    return Inconsistencia.objects.create(
        tipo=tipo, tecnico=tecnico, usuario=usuario, operacion=operacion,
        detalle=detalle, producto=producto, unidad_serializada=unidad,
        serial_reportado=serial, cantidad=cantidad, instalado=instalado,
        asignacion_anterior=asignacion_anterior, estado_anterior=estado_anterior,
        tecnico_anterior=tecnico_anterior,
        asignacion_regularizada=asignacion_regularizada,
    )

def clasificar_serial(*, serial, tecnico):
    """
    Devuelve (tipo, unidad, asignacion, detalle).
    tipo = None significa que está todo bien: lo tiene asignado ese técnico.
    """
    serial = (serial or "").strip()
    unidad = (
        UnidadSerializada.objects
        .select_related("producto")
        .filter(
            Q(numero_serie__iexact=serial)
            | Q(codigo__iexact=serial)
        )
        .first()
    )
    if not unidad:
        return (T.NO_REGISTRADO, None, None,
                "El serial no existe en el sistema: BSI no lo registró en un ingreso.")

    if unidad.estado == EstadoU.ASIGNADO:
        a = (
            Asignacion.objects.select_related("tecnico")
            .filter(unidad_serializada=unidad, estado=EstadoA.ASIGNADA)
            .first()
        )
        if a and a.tecnico_id == tecnico.id:
            return None, unidad, a, ""
        quien = a.tecnico if a else "otro técnico"
        return T.OTRO_TECNICO, unidad, a, f"Figura asignado a {quien}."

    if unidad.estado == EstadoU.DISPONIBLE:
        return (T.NO_ASIGNADO, unidad, None,
                "Figura disponible en almacén; nunca se asignó al técnico.")

    if unidad.estado == EstadoU.INSTALADO:
        prev = (
            Asignacion.objects.select_related("tecnico", "operacion")
            .filter(unidad_serializada=unidad, estado=EstadoA.INSTALADA)
            .order_by("-pk").first()
        )
        txt = "Ya figura instalado"
        if prev:
            txt += f" por {prev.tecnico}"
            if prev.operacion:
                txt += f" en la orden {prev.operacion.numero_orden}"
        return T.YA_INSTALADO, unidad, prev, txt + "."

    return (T.NO_UTILIZABLE, unidad, None,
            f"Figura en estado «{unidad.get_estado_display()}».")

@transaction.atomic
def instalar_serial(*, serial, producto_id, tecnico, usuario, operacion,
                    fecha, observaciones=""):
    tipo, unidad, asig, detalle = clasificar_serial(serial=serial, tecnico=tecnico)

    # Bloquea la unidad y reclasifica, por si cambió de estado
    if unidad:
        UnidadSerializada.objects.select_for_update().get(pk=unidad.pk)
        tipo, unidad, asig, detalle = clasificar_serial(serial=serial, tecnico=tecnico)

    if unidad and unidad.producto_id != int(producto_id):
        raise ErrorInventario(f"El serial {serial} no pertenece al producto seleccionado.")

    # Correcto: lo tiene asignado ese técnico
    if tipo is None:
        if not asig:
            raise ErrorInventario(f"No se encontró la asignación del serial {serial}.")
        instalar(asignacion=asig, usuario=usuario, fecha_instalacion=fecha,
                 operacion=operacion, observaciones=observaciones)
        return None

    producto = unidad.producto if unidad else Producto.objects.filter(pk=producto_id).first()
    if not producto:
        raise ErrorInventario(f"No existe el producto con ID {producto_id}.")

    serial_txt = (unidad.numero_serie or unidad.codigo) if unidad else serial
    regulariza = REGULARIZAR_AUTOMATICO and tipo in (T.NO_ASIGNADO, T.OTRO_TECNICO)

    # Solo reportar
    if not regulariza:
        return _novedad(
            tipo=tipo, tecnico=tecnico, usuario=usuario, operacion=operacion,
            detalle=detalle, producto=producto, unidad=unidad, serial=serial_txt,
            asignacion_anterior=asig,
            estado_anterior=unidad.estado if unidad else None,
            tecnico_anterior=asig.tecnico if asig else None,
        )

    # Regularizar: se guarda de dónde viene antes de moverla
    estado_prev = unidad.estado
    asig_prev = asig
    tecnico_prev = asig.tecnico if asig else None

    if tipo == T.OTRO_TECNICO:
        if not asig:
            raise ErrorInventario(
                f"La unidad {unidad.codigo} figura asignada pero no se encontró la asignación.")
        devolver(asignacion=asig, usuario=usuario,
                 observacion=f"Regularización: la instaló {tecnico}.")
        detalle += f" Se trasladó a {tecnico}, que fue quien la instaló."

    nueva = asignar(
        producto=unidad.producto, tecnico=tecnico, usuario=usuario, cantidad=1,
        unidad=unidad, fecha=fecha,
        observaciones="Regularización automática: instalada sin asignación previa.",
    )
    nueva = instalar(asignacion=nueva, usuario=usuario, fecha_instalacion=fecha,
                     operacion=operacion, observaciones=observaciones)

    return _novedad(
        tipo=tipo, tecnico=tecnico, usuario=usuario, operacion=operacion,
        detalle=detalle, producto=unidad.producto, unidad=unidad, serial=serial_txt,
        instalado=True, asignacion_anterior=asig_prev, estado_anterior=estado_prev,
        tecnico_anterior=tecnico_prev, asignacion_regularizada=nueva,
    )

@transaction.atomic
def instalar_cantidad(
    *,
    producto,
    cantidad,
    tecnico,
    usuario,
    operacion,
    fecha,
    observaciones="",
):
    """
    Instala únicamente cantidades previamente asignadas
    al técnico.

    Todo excedente queda como inconsistencia.
    Nunca consume inventario disponible automáticamente.
    """

    if cantidad < 1:
        raise ErrorInventario(
            "La cantidad que se va a instalar debe ser mayor que cero."
        )

    novedades = []
    resto = cantidad

    asignaciones = list(
        Asignacion.objects
        .select_for_update()
        .filter(
            tecnico=tecnico,
            producto=producto,
            estado=EstadoA.ASIGNADA,
            unidad_serializada__isnull=True,
        )
        .order_by("pk")
    )

    total_asignado = sum(
        a.cantidad_pendiente
        for a in asignaciones
    )

    # --------------------------------------------------------
    # INSTALAR SOLO LO QUE TIENE ASIGNADO EL TÉCNICO
    # --------------------------------------------------------

    for asignacion in asignaciones:

        if resto <= 0:
            break

        pendiente = asignacion.cantidad_pendiente

        if pendiente <= 0:
            continue

        tomar = min(resto, pendiente)

        instalar(
            asignacion=asignacion,
            usuario=usuario,
            fecha_instalacion=fecha,
            operacion=operacion,
            observaciones=observaciones,
            cantidad=tomar,
        )

        resto -= tomar

    # --------------------------------------------------------
    # REGISTRAR EXCEDENTE SIN CONSUMIR INVENTARIO
    # --------------------------------------------------------

    if resto > 0:

        tipo = (
            T.EXCEDE_CANTIDAD
            if total_asignado > 0
            else T.NO_ASIGNADO
        )

        novedades.append(
            _novedad(
                tipo=tipo,
                tecnico=tecnico,
                usuario=usuario,
                operacion=operacion,
                producto=producto,
                cantidad=resto,
                instalado=False,
                detalle=(
                    f"El técnico reportó {cantidad} "
                    f"unidad(es) instaladas de "
                    f"{producto.nombre_completo}. "
                    f"Tenía {total_asignado} unidad(es) "
                    f"pendientes de instalación. "
                    f"Se instalaron únicamente las asignadas. "
                    f"Las {resto} restantes quedan como "
                    f"inconsistencia y no se descontaron "
                    f"del inventario disponible."
                ),
            )
        )

    return novedades

@transaction.atomic
def reabrir_operacion(*, operacion, usuario):
    """
    Revierte todos los materiales de la orden a su estado anterior
    y elimina la orden para poder registrarla de nuevo.
    """
    from operaciones.models import RegistroSAP

    operacion = Operacion.objects.select_for_update().get(pk=operacion.pk)
    numero = operacion.numero_orden

    if RegistroSAP.objects.filter(operacion=operacion).exists():
        raise ErrorInventario(
            f"La orden {numero} ya tiene registros de SAP asociados y no se puede "
            "eliminar. Corrija la conciliación primero."
        )

    # 1) Revierte todo lo instalado (payload vacío = no queda nada)
    sincronizar_materiales_operacion(operacion=operacion, payload=[], usuario=usuario)

    # 2) Cierra las inconsistencias que queden y deja el número de orden anotado
    for n in Inconsistencia.objects.select_for_update().filter(operacion=operacion):
        n.detalle = f"[Orden {numero} reabierta y eliminada] {n.detalle}"
        if n.estado == Inconsistencia.Estado.PENDIENTE:
            n.estado = Inconsistencia.Estado.RESUELTA
            n.resuelto_por = usuario
            n.fecha_resolucion = timezone.now()
            n.nota_resolucion = f"Orden {numero} reabierta: materiales revertidos."
        n.operacion = None
        n.save()

    # 3) Las asignaciones devueltas conservaban la orden: se desligan
    Asignacion.objects.filter(operacion=operacion).update(operacion=None)

    # 4) Elimina el soporte y la orden
    if operacion.archivo_soporte:
        operacion.archivo_soporte.delete(save=False)
    operacion.delete()
    return numero