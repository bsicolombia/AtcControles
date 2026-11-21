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
    MovimientoIVT,
    IVT,
    CajaIVT,
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

# Cuántas unidades serializadas (medidor, regulador, válvula...) de una
# misma categoría se permiten por orden. Si Categoria.max_por_orden está
# vacío, se usa este valor.
MAX_SERIALIZADO_POR_ORDEN = 1


def _limite_por_orden(categoria):
    return categoria.max_por_orden or MAX_SERIALIZADO_POR_ORDEN


def _se_instalara(serial, tecnico, instaladas_en_orden):
    """
    True si ese serial realmente se va a instalar en la orden.
    (Un serial inexistente / ya instalado en otra orden solo se reporta
    como inconsistencia y no ocupa el cupo.)
    """
    unidad = (
        UnidadSerializada.objects
        .filter(Q(numero_serie__iexact=serial) | Q(codigo__iexact=serial))
        .first()
    )
    if not unidad:
        return False
    if unidad.pk in instaladas_en_orden:      # ya estaba instalada en ESTA orden
        return True
    tipo = clasificar_serial(serial=serial, tecnico=tecnico)[0]
    return tipo is None or (
        REGULARIZAR_AUTOMATICO and tipo in (T.NO_ASIGNADO, T.OTRO_TECNICO)
    )
def _validar_unico_por_orden(*, operacion, producto, excluir_asignacion=None):
    """Una orden no puede tener más unidades de una categoría que su límite."""
    if not producto.es_serializado:
        return

    categoria = producto.categoria
    limite = _limite_por_orden(categoria)          # <- ya no se salta si está vacío

    # Bloquea la orden: dos usuarios a la vez no se saltan la regla
    Operacion.objects.select_for_update().get(pk=operacion.pk)

    qs = (
        Asignacion.objects
        .select_related("unidad_serializada", "producto")
        .filter(
            operacion=operacion,
            estado=EstadoA.INSTALADA,
            unidad_serializada__isnull=False,
            producto__categoria_id=categoria.pk,
        )
    )
    if excluir_asignacion is not None:
        qs = qs.exclude(pk=excluir_asignacion.pk)

    if qs.count() >= limite:
        ya = qs.first().unidad_serializada
        raise ErrorInventario(
            f"Duplicado: la orden {operacion.numero_orden} ya tiene instalado "
            f"un/a {categoria.nombre.lower()} ({ya.numero_serie or ya.codigo}). "
            f"Solo se permite {limite} por orden. Si necesita cambiarlo, "
            f"quite el anterior primero."
        )

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
    Instala una asignacion y la vincula a una operacion.

    Serializados: siempre una unidad completa (max. por categoria segun
    Categoria.max_por_orden).
    Por cantidad: permite instalaciones parciales.
    """
    a = (
        Asignacion.objects
        .select_for_update()
        .select_related("producto", "producto__categoria",
                        "unidad_serializada", "tecnico")
        .get(pk=asignacion.pk)
    )

    if a.estado != EstadoA.ASIGNADA:
        raise ErrorInventario(
            "Esta asignacion ya no esta pendiente de instalar."
        )

    # ---------------- SERIALIZADOS ----------------
    if a.unidad_serializada_id:

        if cantidad not in (None, 1):
            raise ErrorInventario(
                "Una unidad serializada solo puede instalarse en cantidad de una."
            )
        cantidad = 1

        # Regla: uno por categoria en cada orden (antes de tocar nada)
        _validar_unico_por_orden(
            operacion=operacion, producto=a.producto, excluir_asignacion=a,
        )

        unidad = (
            UnidadSerializada.objects
            .select_for_update()
            .get(pk=a.unidad_serializada_id)
        )

        if unidad.estado != EstadoU.ASIGNADO:
            raise ErrorInventario(
                f"La unidad {unidad.codigo} no esta asignada actualmente."
            )

        unidad.estado = EstadoU.INSTALADO
        unidad.fecha_cambio_estado = timezone.now()
        unidad.save(update_fields=["estado", "fecha_cambio_estado"])

    # ---------------- POR CANTIDAD ----------------
    else:

        if cantidad is None:
            cantidad = a.cantidad

        if not isinstance(cantidad, int) or isinstance(cantidad, bool):
            raise ErrorInventario(
                "La cantidad de instalacion debe ser un entero."
            )

        if cantidad < 1 or cantidad > a.cantidad:
            raise ErrorInventario(
                f"La cantidad debe estar entre 1 y {a.cantidad}."
            )

        # Instalacion parcial: separar la cantidad instalada.
        if cantidad < a.cantidad:
            a.cantidad -= cantidad
            a.save(update_fields=["cantidad", "fecha_modificacion"])

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

    # ---------------- ACTUALIZAR ASIGNACION ----------------
    a.estado = EstadoA.INSTALADA
    a.cantidad_instalada = cantidad
    a.operacion = operacion
    a.fecha_instalacion = fecha_instalacion
    a.observaciones_instalacion = observaciones or None
    a.save(update_fields=[
        "estado", "cantidad_instalada", "operacion",
        "fecha_instalacion", "observaciones_instalacion",
        "fecha_modificacion",
    ])

    # ---------------- INVENTARIO ----------------
    inventario = _inventario(a.producto)

    if inventario.cantidad_asignada < cantidad:
        raise ErrorInventario(
            f"El inventario de {a.producto.nombre_completo} "
            "no tiene suficientes unidades registradas como asignadas "
            "para completar esta instalacion."
        )

    inventario.cantidad_asignada -= cantidad
    inventario.save(update_fields=["cantidad_asignada", "fecha_actualizacion"])

    # ---------------- MOVIMIENTO ----------------
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
def actualizar_operacion(*, operacion, numero_orden, cuenta_contrato, tipo,
                         estado, anomalia, observaciones, fecha, archivo=None):
    """Edita todos los datos de la orden. Las observaciones se reemplazan."""
    op = Operacion.objects.select_for_update().get(pk=operacion.pk)

    numero = (numero_orden or "").strip().upper()
    if not numero:
        raise ErrorInventario("El número de orden es obligatorio.")
    if Operacion.objects.filter(numero_orden=numero).exclude(pk=op.pk).exists():
        raise ErrorInventario(f"Ya existe otra orden con el número {numero}.")

    ejecutada = estado in Operacion.ESTADOS_QUE_INSTALAN

    op.numero_orden = numero
    op.numero_cuenta_contrato = (cuenta_contrato or "").strip()
    op.tipo_operacion = tipo
    op.estado = estado
    op.anomalia = anomalia
    op.observaciones = (observaciones or "").strip()
    op.fecha_operacion = fecha
    op.fecha_ejecucion = fecha if ejecutada else None
    if archivo:
        op.archivo_soporte = archivo   # el save() del modelo borra el anterior
    op.save()
    return op
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
        # La orden ya fue registrada: desde "instalar" no se puede volver a guardar.
    ya_tiene_material = (
        op.materiales.filter(estado=EstadoA.INSTALADA).exists()
        or op.ivts.filter(estado=EstadoI.INSTALADO).exists()
    )
    if ya_tiene_material or op.estado not in (
        Operacion.Estado.CREADA,
        Operacion.Estado.PENDIENTE,
        Operacion.Estado.ASIGNADA,
    ):
        raise ErrorInventario(
            f"La orden {numero_orden} ya fue registrada "
            f"(estado: {op.get_estado_display()}). "
            "Si necesita corregirla, solicítelo al administrador."
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
    """
    Deja los materiales de la operacion iguales al payload, aplicando solo
    diferencias.

    payload = [
        {"producto_id": 3, "seriales": ["ABC123"]},     # serializado
        {"producto_id": 8, "cantidad": 2},              # por cantidad
    ]

    Todas las validaciones del payload ocurren ANTES de modificar nada.
    """
    # Bloquea la orden durante toda la sincronizacion
    operacion = Operacion.objects.select_for_update().get(pk=operacion.pk)

    tecnico = operacion.tecnico
    if not tecnico:
        raise ErrorInventario("La operacion no tiene un tecnico asignado.")

    fecha = operacion.fecha_ejecucion or operacion.fecha_operacion
    motivo = f"Edicion de la orden {operacion.numero_orden}."

        # ========================================================
    # 1) LEER Y VALIDAR LO QUE SE PIDE (sin tocar la BD)
    # ========================================================
    deseado_cant = {}          # {producto_id: cantidad}
    deseado_serial = []        # [(producto_id, serial)]
    conteo_categoria = {}      # {categoria_id: unidades que SE INSTALAN}
    vistos = set()             # seriales ya usados en este formulario

    instaladas_en_orden = set(
        Asignacion.objects
        .filter(operacion=operacion, estado=EstadoA.INSTALADA,
                unidad_serializada__isnull=False)
        .values_list("unidad_serializada_id", flat=True)
    )

    for item in payload or []:
        try:
            pid = int(item.get("producto_id"))
        except (TypeError, ValueError):
            raise ErrorInventario("Hay un material sin producto valido.")

        producto = (
            Producto.objects.select_related("categoria").filter(pk=pid).first()
        )
        if not producto:
            raise ErrorInventario(f"No existe el producto con ID {pid}.")

        if producto.es_serializado:
            cat = producto.categoria
            limite = _limite_por_orden(cat)

            for s in item.get("seriales") or []:
                s = str(s).strip()
                if not s:
                    continue

                clave = s.lower()
                if clave in vistos:
                    raise ErrorInventario(
                        f"Duplicado: el serial {s} esta repetido en la orden."
                    )
                vistos.add(clave)
                deseado_serial.append((pid, s))

                if _se_instalara(s, tecnico, instaladas_en_orden):
                    n = conteo_categoria.get(cat.pk, 0) + 1
                    conteo_categoria[cat.pk] = n
                    if n > limite:
                        raise ErrorInventario(
                            f"Duplicado: la orden {operacion.numero_orden} solo "
                            f"permite {limite} {cat.nombre.lower()} y esta "
                            f"enviando mas ({s}). Deje solo uno."
                        )
        else:
            try:
                cant = int(item.get("cantidad") or 0)
            except (TypeError, ValueError):
                raise ErrorInventario(
                    f"Cantidad invalida en {producto.nombre_completo}."
                )
            if cant < 0:
                raise ErrorInventario(
                    f"La cantidad de {producto.nombre_completo} no puede ser negativa."
                )
            if cant > 0:
                deseado_cant[pid] = deseado_cant.get(pid, 0) + cant

    # ========================================================
    # 2) ESTADO ACTUAL
    # ========================================================

    # Novedades que solo se reportaron: se regeneran mas abajo
    Inconsistencia.objects.filter(
        operacion=operacion,
        estado=Inconsistencia.Estado.PENDIENTE,
        instalado=False,
    ).exclude(tipo__in=TIPOS_IVT).delete()

    actuales = list(
        Asignacion.objects.select_for_update()
        .filter(operacion=operacion, estado=EstadoA.INSTALADA)
        .order_by("-pk")
    )

    # Depurar filas duplicadas de una misma unidad
    por_unidad = {}
    for a in actuales:
        if not a.unidad_serializada_id:
            continue
        if a.unidad_serializada_id in por_unidad:
            a.estado = EstadoA.DEVUELTA
            a.observaciones = _sumar_observacion(
                a.observaciones, "Fila duplicada depurada al editar la orden.")
            a.save(update_fields=["estado", "observaciones", "fecha_modificacion"])
        else:
            por_unidad[a.unidad_serializada_id] = a
    actuales = [a for a in actuales if a.estado == EstadoA.INSTALADA]

    por_producto = {}
    for a in actuales:
        if not a.unidad_serializada_id:
            por_producto.setdefault(a.producto_id, []).append(a)

    novedades = []

    # ========================================================
    # 3) SERIALIZADOS
    # ========================================================
    conservar, por_instalar = set(), []
    for pid, serial in deseado_serial:
        unidad = (
            UnidadSerializada.objects
            .filter(Q(numero_serie__iexact=serial) | Q(codigo__iexact=serial))
            .first()
        )
        if unidad and unidad.pk in por_unidad:
            conservar.add(unidad.pk)
        else:
            por_instalar.append((pid, serial))

    # Primero se quita lo que sobra (asi se puede cambiar un medidor por otro)
    for uid, a in por_unidad.items():
        if uid not in conservar:
            _revertir_serial(asignacion=a, usuario=usuario, motivo=motivo)

    # Luego se instala lo nuevo (instalar() vuelve a validar el limite)
    for pid, serial in por_instalar:
        n = instalar_serial(
            serial=serial, producto_id=pid, tecnico=tecnico,
            usuario=usuario, operacion=operacion, fecha=fecha,
        )
        if n:
            novedades.append(n)

    # ========================================================
    # 4) POR CANTIDAD
    # ========================================================
    reg_ids = set(
        Inconsistencia.objects.filter(
            operacion=operacion, instalado=True,
            asignacion_regularizada__isnull=False,
        ).values_list("asignacion_regularizada_id", flat=True)
    )

    for pid in set(deseado_cant) | set(por_producto):
        asigs = sorted(
            por_producto.get(pid, []),
            key=lambda a: (a.pk not in reg_ids, -a.pk),  # primero las regularizadas
        )
        actual = sum(a.cantidad for a in asigs)
        deseado = deseado_cant.get(pid, 0)

        if deseado < actual:
            resto = actual - deseado
            for a in asigs:
                if resto <= 0:
                    break
                tomar = min(resto, a.cantidad)
                if a.pk in reg_ids:
                    _devolver_instalacion_a_almacen(
                        asignacion=a, cantidad=tomar,
                        usuario=usuario, motivo=motivo)
                else:
                    _revertir_cantidad(
                        asignacion=a, cantidad=tomar,
                        usuario=usuario, motivo=motivo)
                resto -= tomar

        elif deseado > actual:
            novedades.extend(instalar_cantidad(
                producto=Producto.objects.get(pk=pid),
                cantidad=deseado - actual,
                tecnico=tecnico, usuario=usuario,
                operacion=operacion, fecha=fecha,
            ))

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
    Revierte todos los materiales y el IVT de la orden a su estado anterior
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
 
    # 2) NUEVO: el IVT de la orden vuelve a ASIGNADO (mismo técnico)
    liberar_ivt_operacion(
        operacion=operacion,
        usuario=usuario,
        motivo="Orden reabierta",
    )
    if IVT.objects.filter(operacion=operacion, estado=EstadoI.INSTALADO).exists():
        raise ErrorInventario(
            f"No se pudo liberar el IVT de la orden {numero}; no se eliminó nada."
        )
 
    # 3) Cierra las inconsistencias que queden y deja el número de orden anotado
    for n in Inconsistencia.objects.select_for_update().filter(operacion=operacion):
        n.detalle = f"[Orden {numero} reabierta y eliminada] {n.detalle}"
        if n.estado == Inconsistencia.Estado.PENDIENTE:
            n.estado = Inconsistencia.Estado.RESUELTA
            n.resuelto_por = usuario
            n.fecha_resolucion = timezone.now()
            n.nota_resolucion = f"Orden {numero} reabierta: materiales revertidos."
        n.operacion = None
        n.save()
 
    # 4) Las asignaciones devueltas conservaban la orden: se desligan
    Asignacion.objects.filter(operacion=operacion).update(operacion=None)
 
    # 5) Elimina el soporte y la orden
    if operacion.archivo_soporte:
        operacion.archivo_soporte.delete(save=False)
    operacion.delete()
    return numero
# ============================================================
# IVT (INFORME DE VISITA TÉCNICA)
# ============================================================

EstadoI = IVT.Estado
MAX_IVT_POR_CAJA = 100_000

TRANSICIONES_IVT = {
    EstadoI.DISPONIBLE: (EstadoI.ANULADO, EstadoI.PERDIDO),
    EstadoI.ASIGNADO: (EstadoI.DISPONIBLE, EstadoI.ANULADO, EstadoI.PERDIDO),
    EstadoI.PERDIDO: (EstadoI.ASIGNADO, EstadoI.DISPONIBLE),   # apareció
}


def fmt_ivt(numero):
    return f"{int(numero):05d}"


def _mov_ivt(ivt, anterior, usuario, obs="", operacion=None):
    return MovimientoIVT(
        ivt=ivt, estado_anterior=anterior, estado_nuevo=ivt.estado,
        tecnico=ivt.tecnico, operacion=operacion or ivt.operacion,
        usuario=usuario, observacion=obs,
    )


# ---------------- ingreso por cajas ----------------

@transaction.atomic
def registrar_ingreso_ivt(*, cajas, remision, usuario, fecha_ingreso, observacion=""):
    """cajas = [(consecutivo_inicial, cantidad), ...]. Devuelve las CajaIVT creadas."""
    if not cajas:
        raise ErrorInventario("Agregue al menos una caja.")

    rangos = []
    for i, (ini, cant) in enumerate(cajas, 1):
        if ini < 1 or cant < 1:
            raise ErrorInventario(f"Caja {i}: inicial y cantidad deben ser mayores que cero.")
        if cant > MAX_IVT_POR_CAJA:
            raise ErrorInventario(f"Caja {i}: máximo {MAX_IVT_POR_CAJA} IVT por caja.")
        rangos.append((ini, ini + cant - 1))

    ordenados = sorted(rangos)
    for (_, f1), (i2, _) in zip(ordenados, ordenados[1:]):
        if i2 <= f1:
            raise ErrorInventario("Los consecutivos de las cajas se cruzan entre sí.")

    for ini, fin in rangos:
        ya = IVT.objects.filter(numero__range=(ini, fin)).order_by("numero").first()
        if ya:
            raise ErrorInventario(
                f"El rango {fmt_ivt(ini)}-{fmt_ivt(fin)} ya tiene IVT registrados "
                f"(por ejemplo el {ya.codigo})."
            )

    base = CajaIVT.objects.filter(remision=remision).count() if remision else 0
    creadas = []
    for pos, (ini, fin) in enumerate(rangos, 1):
        caja = CajaIVT.objects.create(
            remision=remision, numero_caja=base + pos,
            consecutivo_inicial=ini, consecutivo_final=fin, cantidad=fin - ini + 1,
            fecha_ingreso=fecha_ingreso, usuario=usuario, observacion=observacion,
        )
        IVT.objects.bulk_create(
            [IVT(numero=n, caja=caja) for n in range(ini, fin + 1)],
            batch_size=1000,
        )
        creadas.append(caja)
    return creadas


# ---------------- asignación por rango ----------------

@transaction.atomic
def asignar_ivt_rango(*, tecnico, inicial, final, usuario, fecha, observacion=""):
    if inicial < 1 or final < inicial:
        raise ErrorInventario("El consecutivo final debe ser mayor o igual al inicial.")

    total = final - inicial + 1
    rango = IVT.objects.filter(numero__range=(inicial, final))

    # Diagnóstico claro ANTES de modificar nada
    if rango.filter(estado=EstadoI.DISPONIBLE).count() != total:
        partes = []
        faltan = total - rango.count()
        if faltan:
            partes.append(f"{faltan} número(s) del rango no existen en el sistema")
        no_disp = list(
            rango.exclude(estado=EstadoI.DISPONIBLE)
            .order_by("numero").values_list("numero", "estado")[:5]
        )
        if no_disp:
            ejemplos = ", ".join(f"{fmt_ivt(n)} ({EstadoI(e).label})" for n, e in no_disp)
            partes.append(f"hay IVT que no están disponibles: {ejemplos}")
        raise ErrorInventario(
            f"No se puede asignar {fmt_ivt(inicial)}-{fmt_ivt(final)}: " + "; ".join(partes) + "."
        )

    ahora = timezone.now()
    # El filtro por estado hace la operación segura frente a concurrencia
    actualizados = rango.filter(estado=EstadoI.DISPONIBLE).update(
        estado=EstadoI.ASIGNADO, tecnico=tecnico,
        fecha_asignacion=fecha, fecha_cambio_estado=ahora,
    )
    if actualizados != total:
        raise ErrorInventario("Otro usuario modificó esos IVT al mismo tiempo. Intente de nuevo.")

    ids = list(rango.values_list("pk", flat=True))
    obs = f"Asignado a {tecnico} (rango {fmt_ivt(inicial)}-{fmt_ivt(final)}). {observacion}".strip()
    MovimientoIVT.objects.bulk_create(
        [MovimientoIVT(ivt_id=pk, estado_anterior=EstadoI.DISPONIBLE,
                       estado_nuevo=EstadoI.ASIGNADO, tecnico=tecnico,
                       usuario=usuario, observacion=obs) for pk in ids],
        batch_size=1000,
    )
    return total


# ---------------- anular / perdido / devolver (masivo) ----------------

@transaction.atomic
def cambiar_estado_ivt(*, ids, nuevo_estado, usuario, motivo):
    motivo = (motivo or "").strip()
    if not motivo:
        raise ErrorInventario("Escriba el motivo del cambio.")
    if nuevo_estado not in EstadoI.values:
        raise ErrorInventario("Acción no válida.")

    ivts = list(IVT.objects.select_for_update().filter(pk__in=ids))
    if not ivts:
        raise ErrorInventario("Seleccione al menos un IVT.")

    ahora = timezone.now()
    movs = []
    for ivt in ivts:
        if nuevo_estado not in TRANSICIONES_IVT.get(ivt.estado, ()):
            raise ErrorInventario(
                f"IVT {ivt.codigo}: no se puede pasar de «{ivt.get_estado_display()}» "
                f"a «{EstadoI(nuevo_estado).label}»."
            )
        if nuevo_estado == EstadoI.ASIGNADO and not ivt.tecnico_id:
            raise ErrorInventario(f"IVT {ivt.codigo}: no tiene técnico a quien reasignarlo.")

        anterior = ivt.estado
        ivt.estado = nuevo_estado
        if nuevo_estado == EstadoI.DISPONIBLE:
            ivt.tecnico = None
            ivt.fecha_asignacion = None
        ivt.fecha_cambio_estado = ahora
        ivt.observaciones = _sumar_observacion(
            ivt.observaciones, f"{EstadoI(nuevo_estado).label}: {motivo}")
        movs.append(_mov_ivt(ivt, anterior, usuario, motivo))

    IVT.objects.bulk_update(
        ivts, ["estado", "tecnico", "fecha_asignacion", "fecha_cambio_estado", "observaciones"],
        batch_size=1000)
    MovimientoIVT.objects.bulk_create(movs, batch_size=1000)
    return len(ivts)


@transaction.atomic
def marcar_ivt_fisico_perdido(*, ids, usuario, motivo):
    """IVT ya instalado cuyo papel físico se extravió: sigue ligado a su orden."""
    motivo = (motivo or "").strip()
    if not motivo:
        raise ErrorInventario("Escriba el motivo.")
    ivts = list(IVT.objects.select_for_update().filter(pk__in=ids))
    if not ivts:
        raise ErrorInventario("Seleccione al menos un IVT.")
    movs = []
    for ivt in ivts:
        if ivt.estado != EstadoI.INSTALADO:
            raise ErrorInventario(f"IVT {ivt.codigo}: solo un IVT instalado puede marcarse con físico perdido.")
        ivt.fisico_perdido = True
        ivt.observaciones = _sumar_observacion(ivt.observaciones, f"Físico extraviado: {motivo}")
        movs.append(_mov_ivt(ivt, ivt.estado, usuario, f"Físico extraviado: {motivo}"))
    IVT.objects.bulk_update(ivts, ["fisico_perdido", "observaciones"])
    MovimientoIVT.objects.bulk_create(movs)
    return len(ivts)


# ---------------- relación con la instalación ----------------

def validar_ivt_para_operacion(*, numero, tecnico, operacion=None, bloquear=False):
    try:
        numero = int(str(numero).strip())
    except (TypeError, ValueError):
        raise ErrorInventario("El IVT debe ser un número.")

    qs = IVT.objects.select_for_update() if bloquear else IVT.objects.select_related("tecnico")
    ivt = qs.filter(numero=numero).first()
    codigo = fmt_ivt(numero)

    if not ivt:
        raise ErrorInventario(f"El IVT {codigo} no existe. Verifique el consecutivo.")
    if ivt.estado == EstadoI.INSTALADO:
        if operacion is not None and ivt.operacion_id == operacion.pk:
            return ivt
        orden = ivt.operacion.numero_orden if ivt.operacion_id else "sin orden"
        raise ErrorInventario(f"El IVT {codigo} ya fue usado ({orden}).")
    if ivt.estado != EstadoI.ASIGNADO:
        raise ErrorInventario(
            f"El IVT {codigo} está «{ivt.get_estado_display()}»; "
            "solo se puede usar uno asignado al técnico.")
    if tecnico is not None and ivt.tecnico_id != tecnico.id:
        raise ErrorInventario(f"El IVT {codigo} está asignado a {ivt.tecnico}, no a {tecnico}.")
    return ivt


@transaction.atomic
def liberar_ivt_operacion(*, operacion, usuario, motivo):
    """Los IVT instalados en la orden vuelven a Asignado (mismo técnico)."""
    ahora = timezone.now()
    for ivt in IVT.objects.select_for_update().filter(
            operacion=operacion, estado=EstadoI.INSTALADO):
        ivt.estado = EstadoI.ASIGNADO
        ivt.operacion = None
        ivt.fecha_instalacion = None
        ivt.fisico_perdido = False
        ivt.fecha_cambio_estado = ahora
        ivt.save(update_fields=["estado", "operacion", "fecha_instalacion",
                                "fisico_perdido", "fecha_cambio_estado"])
        MovimientoIVT.objects.create(
            ivt=ivt, estado_anterior=EstadoI.INSTALADO, estado_nuevo=EstadoI.ASIGNADO,
            tecnico=ivt.tecnico, operacion=operacion, usuario=usuario,
            observacion=f"{motivo} (orden {operacion.numero_orden})")


@transaction.atomic
def vincular_ivt(*, numero, operacion, tecnico, usuario):
    ivt = validar_ivt_para_operacion(
        numero=numero, tecnico=tecnico, operacion=operacion, bloquear=True)
    if ivt.estado == EstadoI.INSTALADO:      # ya estaba en esta orden
        return ivt

    # Una orden = un IVT: si tenía otro, vuelve a Asignado
    liberar_ivt_operacion(
        operacion=operacion, usuario=usuario,
        motivo=f"Reemplazado por el IVT {ivt.codigo}")

    ivt.estado = EstadoI.INSTALADO
    ivt.operacion = operacion
    ivt.fecha_instalacion = operacion.fecha_ejecucion or operacion.fecha_operacion
    ivt.fecha_cambio_estado = timezone.now()
    ivt.save(update_fields=["estado", "operacion", "fecha_instalacion", "fecha_cambio_estado"])
    MovimientoIVT.objects.create(
        ivt=ivt, estado_anterior=EstadoI.ASIGNADO, estado_nuevo=EstadoI.INSTALADO,
        tecnico=ivt.tecnico, operacion=operacion, usuario=usuario,
        observacion=f"Instalado en la orden {operacion.numero_orden}")
    return ivt

def rangos_ivt(*, tecnico, estado=None):
    """Rangos consecutivos de IVT que tiene el técnico (por defecto, asignados)."""
    estado = estado or EstadoI.ASIGNADO
    numeros = (
        IVT.objects.filter(tecnico=tecnico, estado=estado)
        .order_by("numero").values_list("numero", flat=True)
    )
    rangos, ini, prev = [], None, None
    for n in numeros:
        if ini is None:
            ini = prev = n
        elif n == prev + 1:
            prev = n
        else:
            rangos.append((ini, prev))
            ini = prev = n
    if ini is not None:
        rangos.append((ini, prev))

    return [
        {
            "inicial": a, "final": b, "cantidad": b - a + 1,
            "texto": f"{fmt_ivt(a)} - {fmt_ivt(b)}" if a != b else fmt_ivt(a),
        }
        for a, b in rangos
    ]


@transaction.atomic
def devolver_ivt_rango(*, tecnico, inicial, final, usuario, estado_retorno, observacion):
    """
    Devuelve un rango de IVT que el técnico tiene ASIGNADO.
    estado_retorno: DISPONIBLE (vuelve al almacén), ANULADO o PERDIDO.
    Si UN solo número del rango no es del técnico o no está asignado, no se mueve nada.
    """
    observacion = (observacion or "").strip()
    if not observacion:
        raise ErrorInventario("La observación es obligatoria para devolver IVT.")
    if estado_retorno not in (EstadoI.DISPONIBLE, EstadoI.ANULADO, EstadoI.PERDIDO):
        raise ErrorInventario("Estado de devolución de IVT no válido.")
    if inicial < 1 or final < inicial:
        raise ErrorInventario("El consecutivo final debe ser mayor o igual al inicial.")

    total = final - inicial + 1
    ids = list(
        IVT.objects.select_for_update()
        .filter(numero__range=(inicial, final),
                estado=EstadoI.ASIGNADO, tecnico=tecnico)
        .values_list("pk", flat=True)
    )

    if len(ids) != total:
        partes = []
        existentes = IVT.objects.filter(numero__range=(inicial, final)).count()
        if existentes != total:
            partes.append(f"{total - existentes} número(s) no existen en el sistema")
        otros = (
            IVT.objects.filter(numero__range=(inicial, final))
            .exclude(estado=EstadoI.ASIGNADO, tecnico=tecnico)
            .select_related("tecnico").order_by("numero")[:5]
        )
        if otros:
            ejemplos = ", ".join(
                f"{o.codigo} ({o.get_estado_display()}"
                f"{', de ' + str(o.tecnico) if o.tecnico_id and o.tecnico_id != tecnico.id else ''})"
                for o in otros
            )
            partes.append(f"no están asignados a {tecnico}: {ejemplos}")
        raise ErrorInventario(
            f"No se puede devolver {fmt_ivt(inicial)}-{fmt_ivt(final)}: " + "; ".join(partes) + "."
        )

    cambiar_estado_ivt(
        ids=ids, nuevo_estado=estado_retorno, usuario=usuario, motivo=observacion,
    )
    return total
def _agrupar_rangos(numeros):
    rangos, ini, prev = [], None, None
    for n in numeros:
        if ini is None:
            ini = prev = n
        elif n == prev + 1:
            prev = n
        else:
            rangos.append((ini, prev)); ini = prev = n
    if ini is not None:
        rangos.append((ini, prev))
    return [{"inicial": a, "final": b, "cantidad": b - a + 1} for a, b in rangos]


def rangos_ivt_disponibles():
    """Rangos consecutivos de IVT que están en almacén."""
    numeros = (IVT.objects.filter(estado=EstadoI.DISPONIBLE)
               .order_by("numero").values_list("numero", flat=True))
    return _agrupar_rangos(numeros)

# ============================================================
# IVT: INCONSISTENCIAS (la orden se registra igual)
# ============================================================

TIPOS_IVT = {
    T.IVT_NO_REGISTRADO,
    T.IVT_NO_ASIGNADO,
    T.IVT_OTRO_TECNICO,
    T.IVT_YA_USADO,
    T.IVT_NO_UTILIZABLE,
}


def clasificar_ivt(*, numero, tecnico, operacion=None):
    """
    Devuelve (tipo, ivt, detalle).
    tipo = None significa que todo está bien: el IVT está asignado a ese
    técnico, o ya está instalado en esta misma orden.
    """
    try:
        numero = int(str(numero).strip())
    except (TypeError, ValueError):
        return T.IVT_NO_REGISTRADO, None, "El IVT debe ser un número."

    codigo = fmt_ivt(numero)
    ivt = (
        IVT.objects
        .select_related("tecnico", "operacion")
        .filter(numero=numero)
        .first()
    )

    if not ivt:
        return (
            T.IVT_NO_REGISTRADO, None,
            f"El IVT {codigo} no existe en el sistema: no se registró en "
            "ningún ingreso de cajas.",
        )

    if ivt.estado == EstadoI.INSTALADO:
        if operacion is not None and ivt.operacion_id == operacion.pk:
            return None, ivt, ""
        orden = ivt.operacion.numero_orden if ivt.operacion_id else "sin orden"
        return (
            T.IVT_YA_USADO, ivt,
            f"El IVT {codigo} ya fue usado en la orden {orden}.",
        )

    if ivt.estado == EstadoI.DISPONIBLE:
        return (
            T.IVT_NO_ASIGNADO, ivt,
            f"El IVT {codigo} figura en almacén; nunca se asignó al técnico.",
        )

    if ivt.estado in (EstadoI.ANULADO, EstadoI.PERDIDO):
        return (
            T.IVT_NO_UTILIZABLE, ivt,
            f"El IVT {codigo} figura como «{ivt.get_estado_display()}».",
        )

    # Estado ASIGNADO
    if tecnico is not None and ivt.tecnico_id != tecnico.id:
        return (
            T.IVT_OTRO_TECNICO, ivt,
            f"El IVT {codigo} está asignado a {ivt.tecnico}.",
        )

    return None, ivt, ""


@transaction.atomic
def vincular_ivt_o_reportar(*, numero, operacion, tecnico, usuario):
    """
    Si el IVT está bien, lo vincula a la orden.
    Si tiene algún problema NO bloquea la orden: queda como Inconsistencia
    y devuelve esa Inconsistencia.
    """
    if numero in (None, ""):
        return None

    # La orden se está guardando de nuevo: se regeneran las novedades de IVT
    Inconsistencia.objects.filter(
        operacion=operacion,
        tipo__in=TIPOS_IVT,
        estado=Inconsistencia.Estado.PENDIENTE,
        instalado=False,
    ).delete()

    tipo, ivt, detalle = clasificar_ivt(
        numero=numero, tecnico=tecnico, operacion=operacion,
    )

    if tipo is None:
        vincular_ivt(
            numero=numero, operacion=operacion, tecnico=tecnico, usuario=usuario,
        )
        return None

    return _novedad(
        tipo=tipo,
        tecnico=tecnico,
        usuario=usuario,
        operacion=operacion,
        detalle=detalle,
        serial=str(numero).strip().zfill(5),
        cantidad=1,
        instalado=False,
        estado_anterior=ivt.estado if ivt else None,
        tecnico_anterior=ivt.tecnico if ivt and ivt.tecnico_id else None,
    )