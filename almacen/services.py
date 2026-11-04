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
from .models import RemisionIngreso
from django.db.models import Q
from .models import (
    Asignacion, Inconsistencia, Inventario,
    MovimientoInventario, Producto, UnidadSerializada,
)
T = Inconsistencia.Tipo
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
        tipo_movimiento=Mov.INSTALACION,      # <-- antes Mov.SALIDA
        cantidad=cantidad,
        usuario=usuario,
        observacion=(observaciones or "").strip(),
    )

    return a


@transaction.atomic
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

    asignacion.save(update_fields=["estado", "fecha_modificacion"])

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
@transaction.atomic
def devolver_cantidad(*, tecnico, producto, cantidad, usuario, observacion=""):
    """Devuelve parte (o todo) de un material por cantidad que tiene el técnico."""
    asigs = list(
        Asignacion.objects.select_for_update()
        .filter(
            tecnico=tecnico,
            producto=producto,
            estado=EstadoA.ASIGNADA,
            unidad_serializada__isnull=True,
        )
        .order_by("pk")
    )

    total = sum(a.cantidad for a in asigs)
    if cantidad < 1 or cantidad > total:
        raise ErrorInventario(
            f"{producto.nombre_completo}: el técnico solo tiene {total} para devolver."
        )

    resto = cantidad
    for a in asigs:
        if resto <= 0:
            break

        tomar = min(resto, a.cantidad)

        if tomar < a.cantidad:
            # Devolución parcial: se divide la asignación
            a.cantidad -= tomar
            a.save(update_fields=["cantidad", "fecha_modificacion"])

            a = Asignacion.objects.create(
                tecnico=a.tecnico,
                producto=a.producto,
                cantidad=tomar,
                estado=EstadoA.ASIGNADA,
                fecha_asignacion=a.fecha_asignacion,
                observaciones=a.observaciones,
                usuario=a.usuario,
            )

        devolver(asignacion=a, usuario=usuario, observacion=observacion)
        resto -= tomar
# ============================================================
# INSTALACIÓN CON INCONSISTENCIAS
# ============================================================

# Estos tipos solo se reportan: no se instala ni se mueve inventario.
TIPOS_SOLO_REPORTE = {T.NO_REGISTRADO, T.YA_INSTALADO, T.NO_UTILIZABLE, T.SIN_STOCK}


def _novedad(*, tipo, tecnico, usuario, operacion, detalle, producto=None,
             unidad=None, serial="", cantidad=1, instalado=False):
    return Inconsistencia.objects.create(
        tipo=tipo, tecnico=tecnico, usuario=usuario, operacion=operacion,
        detalle=detalle, producto=producto, unidad_serializada=unidad,
        serial_reportado=serial, cantidad=cantidad, instalado=instalado,
    )


def clasificar_serial(*, serial, tecnico):
    """
    Devuelve (tipo, unidad, asignacion, detalle).
    tipo = None significa que está todo bien: lo tiene asignado ese técnico.
    """
    serial = (serial or "").strip()
    unidad = (
        UnidadSerializada.objects.select_related("producto")
        .filter(Q(numero_serie__iexact=serial) | Q(codigo__iexact=serial))
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
    """Instala una unidad por serial. Devuelve la Inconsistencia creada o None."""
    tipo, unidad, asig, detalle = clasificar_serial(serial=serial, tecnico=tecnico)

    if tipo is None:
        instalar(asignacion=asig, usuario=usuario, fecha_instalacion=fecha,
                 operacion=operacion, observaciones=observaciones)
        return None

    if tipo in TIPOS_SOLO_REPORTE:
        producto = unidad.producto if unidad else Producto.objects.filter(pk=producto_id).first()
        return _novedad(tipo=tipo, tecnico=tecnico, usuario=usuario, operacion=operacion,
                        detalle=detalle, producto=producto, unidad=unidad,
                        serial=serial, instalado=False)

    # NO_ASIGNADO u OTRO_TECNICO: se regulariza el inventario y se instala
    if tipo == T.OTRO_TECNICO:
        if not asig:
            raise ErrorInventario(
                f"La unidad {unidad.codigo} figura asignada pero no se encontró "
                "la asignación. Revísela en Unidades."
            )
        devolver(asignacion=asig, usuario=usuario,
                 observacion=f"Regularización: la instaló {tecnico}.")
        detalle += f" Se trasladó a {tecnico}, que fue quien la instaló."

    nueva = asignar(
        producto=unidad.producto, tecnico=tecnico, usuario=usuario, cantidad=1,
        unidad=unidad, fecha=fecha,
        observaciones="Regularización automática: instalada sin asignación previa.",
    )
    instalar(asignacion=nueva, usuario=usuario, fecha_instalacion=fecha,
             operacion=operacion, observaciones=observaciones)
    return _novedad(tipo=tipo, tecnico=tecnico, usuario=usuario, operacion=operacion,
                    detalle=detalle, producto=unidad.producto, unidad=unidad,
                    serial=unidad.numero_serie or unidad.codigo, instalado=True)


@transaction.atomic
def instalar_cantidad(*, producto, cantidad, tecnico, usuario, operacion,
                      fecha, observaciones=""):
    """Instala material por cantidad. Devuelve la lista de inconsistencias creadas."""
    novedades = []
    asigs = list(
        Asignacion.objects.select_for_update()
        .filter(tecnico=tecnico, producto=producto, estado=EstadoA.ASIGNADA,
                unidad_serializada__isnull=True)
        .order_by("pk")
    )

    resto = cantidad
    for a in asigs:
        if resto <= 0:
            break
        pend = a.cantidad_pendiente
        if pend <= 0:
            continue
        tomar = min(resto, pend)
        instalar(asignacion=a, usuario=usuario, fecha_instalacion=fecha,
                 operacion=operacion, observaciones=observaciones, cantidad=tomar)
        resto -= tomar

    if resto <= 0:
        return novedades

    tipo = T.EXCEDE_CANTIDAD if resto < cantidad else T.NO_ASIGNADO
    inv = _inventario(producto)
    cubrir = min(resto, inv.cantidad_actual)

    if cubrir > 0:
        nueva = asignar(
            producto=producto, tecnico=tecnico, usuario=usuario, cantidad=cubrir,
            fecha=fecha,
            observaciones="Regularización automática: instalado sin asignación previa.",
        )
        instalar(asignacion=nueva, usuario=usuario, fecha_instalacion=fecha,
                 operacion=operacion, observaciones=observaciones, cantidad=cubrir)
        novedades.append(_novedad(
            tipo=tipo, tecnico=tecnico, usuario=usuario, operacion=operacion,
            producto=producto, cantidad=cubrir, instalado=True,
            detalle=f"{cubrir} unidad(es) no estaban asignadas al técnico; "
                    "se tomaron del stock del almacén.",
        ))

    faltan = resto - cubrir
    if faltan > 0:
        novedades.append(_novedad(
            tipo=T.SIN_STOCK, tecnico=tecnico, usuario=usuario, operacion=operacion,
            producto=producto, cantidad=faltan, instalado=False,
            detalle=f"Reportó {faltan} unidad(es) que no están asignadas y no hay "
                    f"existencias en el sistema (stock: {inv.cantidad_actual}). "
                    "Probable material que BSI no ingresó.",
        ))
    return novedades