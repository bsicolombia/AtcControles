from django.conf import settings
from django.db import models
from django.utils import timezone
import os

def ruta_remision(instance, filename):
    # Extrae la extensión del archivo original (ej: .jpg, .png, .pdf)
    ext = filename.split('.')[-1]
    # Limpia el número de remisión por seguridad para que no tenga caracteres raros
    nombre_limpio = instance.numero_remision.strip().upper().replace(" ", "_")
    # Retorna la ruta: media/remisiones_ingreso/NUMERO_REMISION.ext
    return os.path.join("remisiones_ingreso", f"{nombre_limpio}.{ext}")


class RemisionIngreso(models.Model):
    numero_remision = models.CharField(
        "Número de orden / remisión",
        max_length=30,
        unique=True
    )
    
    comprobante_imagen = models.ImageField(
        "Foto del comprobante",
        upload_to=ruta_remision,
        blank=True,
        null=True
    )
    
    fecha_registro = models.DateTimeField(auto_now_add=True)
    
    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="remisiones_ingreso"
    )

    class Meta:
        verbose_name = "Remisión de ingreso"
        verbose_name_plural = "Remisiones de ingreso"
        ordering = ["-fecha_registro"]

    def __str__(self):
        return self.numero_remision
    

# ============================================================
# CATEGORÍA
# ============================================================

class Categoria(models.Model):
    nombre = models.CharField(max_length=100, unique=True)

    prefijo = models.CharField(
        "Prefijo de código",
        max_length=5,
        blank=True,
        help_text="Solo para categorías serializadas. Ej: DET, REG, VAL."
    )

    estado = models.BooleanField(default=True)

    fecha_creacion = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Categoría"
        verbose_name_plural = "Categorías"
        ordering = ["nombre"]

    def __str__(self):
        return self.nombre


# ============================================================
# PRODUCTO
# ============================================================

class Producto(models.Model):

    class TipoInventario(models.TextChoices):
        SERIALIZADO = "SERIALIZADO", "Serializado (una unidad = un código)"
        CANTIDAD = "CANTIDAD", "Por cantidad"

    class UnidadMedida(models.TextChoices):
        UNIDAD = "UNIDAD", "Unidad"
        METRO = "METRO", "Metro"
        KILOGRAMO = "KILOGRAMO", "Kilogramo"
        LITRO = "LITRO", "Litro"
        TARRO = "TARRO", "Tarro"
        ROLLO = "ROLLO", "Rollo"
        CAJA = "CAJA", "Caja"
        PAQUETE = "PAQUETE", "Paquete"

    categoria = models.ForeignKey(
        Categoria,
        on_delete=models.PROTECT,
        related_name="productos"
    )

    codigo_producto = models.CharField(
        max_length=30,
        unique=True
    )

    nombre = models.CharField(
        max_length=150
    )

    marca = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    referencia = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    medida = models.CharField(
        max_length=50,
        blank=True,
        null=True
    )

    unidad_medida = models.CharField(
        max_length=20,
        choices=UnidadMedida.choices,
        default=UnidadMedida.UNIDAD
    )

    color = models.CharField(
        max_length=50,
        blank=True,
        null=True
    )

    presentacion = models.CharField(
        "Presentación (tarro)",
        max_length=50,
        blank=True,
        null=True
    )

    tipo_inventario = models.CharField(
        max_length=20,
        choices=TipoInventario.choices
    )

    requiere_numero_serie = models.BooleanField(
        "Requiere número de serie físico",
        default=False,
        help_text=(
            "Actívelo cuando el producto tenga un número de serie "
            "físico que deba registrarse."
        )
    )

    estado = models.BooleanField(default=True)

    fecha_creacion = models.DateTimeField(
        auto_now_add=True
    )

    class Meta:
        verbose_name = "Producto"
        verbose_name_plural = "Productos"
        ordering = ["nombre", "medida"]

    def __str__(self):
        return f"{self.codigo_producto} - {self.nombre}"

    @property
    def es_serializado(self):
        return self.tipo_inventario == self.TipoInventario.SERIALIZADO

    @property
    def nombre_completo(self):
        extra = " ".join(
            x for x in (
                self.medida,
                self.color,
                self.presentacion
            )
            if x
        )
        return f"{self.nombre} {extra}".strip()


# ============================================================
# UNIDAD SERIALIZADA
# ============================================================

class UnidadSerializada(models.Model):

    class Estado(models.TextChoices):
        DISPONIBLE = "DISPONIBLE", "Disponible"
        ASIGNADO = "ASIGNADO", "Asignado"
        INSTALADO = "INSTALADO", "Instalado"
        DAÑADO = "DANADO", "Dañado"
        BAJA = "BAJA", "Baja"

    remision = models.ForeignKey(
        RemisionIngreso, 
        on_delete=models.PROTECT, 
        blank=True, 
        null=True, 
        related_name="unidades"
    )
    
    producto = models.ForeignKey(
        Producto,
        on_delete=models.PROTECT,
        related_name="unidades_serializadas"
    )

    codigo = models.CharField(
        max_length=50,
        unique=True
    )

    numero_serie = models.CharField(
        "Número de serie",
        max_length=100,
        blank=True,
        null=True
    )

    marca = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    referencia = models.CharField(
        max_length=100,
        blank=True,
        null=True
    )

    medida = models.CharField(
        max_length=50,
        blank=True,
        null=True
    )

    fecha_ingreso = models.DateField()

    estado = models.CharField(
        max_length=20,
        choices=Estado.choices,
        default=Estado.DISPONIBLE
    )
    
    fecha_cambio_estado = models.DateTimeField(
        null=True,
        blank=True,
        editable=False,
    )

    observaciones = models.TextField(
        blank=True,
        null=True
    )

    class Meta:
        verbose_name = "Unidad serializada"
        verbose_name_plural = "Unidades serializadas"
        ordering = ["-fecha_ingreso", "-codigo"]

    def __str__(self):
        return self.codigo
    
    @property
    def etiqueta(self):
        """Texto para mostrar en selectores: serial primero, código como apoyo."""
        base = self.numero_serie or self.codigo
        if self.numero_serie:
            base = f"{self.numero_serie} · {self.codigo}"
        if self.medida:
            base += f" · {self.medida}"
        return base


# ============================================================
# INVENTARIO
# ============================================================

class Inventario(models.Model):

    producto = models.OneToOneField(
        Producto,
        on_delete=models.PROTECT,
        related_name="inventario"
    )

    cantidad_actual = models.PositiveIntegerField(
        default=0
    )

    cantidad_asignada = models.PositiveIntegerField(
        default=0
    )

    inventario_almacen = models.PositiveIntegerField(
        "Inventario de almacén",
        default=0
    )

    cantidad_minima = models.PositiveIntegerField(
        default=0
    )

    fecha_actualizacion = models.DateTimeField(
        auto_now=True
    )

    class Meta:
        verbose_name = "Inventario"
        verbose_name_plural = "Inventario"
        ordering = ["producto__nombre", "producto__medida"]

    def __str__(self):
        return f"{self.producto.nombre} - Stock: {self.cantidad_actual}"

    @property
    def total(self):
        return self.cantidad_actual + self.cantidad_asignada

    @property
    def bajo_stock(self):
        return (
            self.cantidad_minima > 0
            and self.cantidad_actual <= self.cantidad_minima
        )
        

    @property
    def total_asignado(self):
        """Total de material que se encuentra actualmente en estado ASIGNADA"""
        return sum(
            a.cantidad for a in self.producto.asignaciones.all() 
            if a.estado == Asignacion.Estado.ASIGNADA
        )

    @property
    def instalado_five(self):
        """Total de material instalado según Five (excluyendo devueltas)"""
        total = 0
        for a in self.producto.asignaciones.all():
            if a.estado == Asignacion.Estado.DEVUELTA:
                continue  # Ignoramos las devueltas
            
            if a.estado == Asignacion.Estado.INSTALADA:
                total += a.cantidad or 0
            else:
                total += a.cantidad_instalada or 0
        return total

    @property
    def en_poder_tecnicos(self):
        """Material en poder de técnicos: Lo asignado activo menos lo ya instalado de esas asignaciones"""
        total = 0
        for a in self.producto.asignaciones.all():
            if a.estado == Asignacion.Estado.ASIGNADA:
                # Cantidad asignada menos lo que ya hayan instalado de esta misma orden
                pendiente = max((a.cantidad or 0) - (a.cantidad_instalada or 0), 0)
                total += pendiente
        return total

# ============================================================
# ASIGNACIÓN
# ============================================================

class Asignacion(models.Model):

    class Estado(models.TextChoices):
        ASIGNADA = "ASIGNADA", "Asignada"
        INSTALADA = "INSTALADA", "Instalada"
        DEVUELTA = "DEVUELTA", "Devuelta"
        
    operacion = models.ForeignKey(
        "operaciones.Operacion",
        on_delete=models.PROTECT,
        related_name="materiales",
        null=True, blank=True,
    )

    tecnico = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="asignaciones",
        verbose_name="Técnico",
    )

    producto = models.ForeignKey(
        Producto,
        on_delete=models.PROTECT,
        related_name="asignaciones"
    )

    unidad_serializada = models.ForeignKey(
        UnidadSerializada,
        on_delete=models.PROTECT,
        related_name="asignaciones",
        blank=True,
        null=True
    )

    cantidad = models.PositiveIntegerField(
        default=1
    )

    cantidad_instalada = models.PositiveIntegerField(
        "Cantidad instalada",
        default=0
    )

    estado = models.CharField(
        max_length=20,
        choices=Estado.choices,
        default=Estado.ASIGNADA
    )

    fecha_asignacion = models.DateField(
        default=timezone.localdate
    )
    
    fecha_instalacion = models.DateField(null=True, blank=True)
    observaciones_instalacion = models.TextField(null=True, blank=True)

    observaciones = models.TextField(
        blank=True,
        null=True
    )

    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="asignaciones_registradas"
    )

    fecha_modificacion = models.DateTimeField(
        auto_now=True
    )

    class Meta:
        verbose_name = "Asignación"
        verbose_name_plural = "Asignaciones"
        ordering = ["-fecha_asignacion", "-id"]

    def __str__(self):
        return f"{self.producto.nombre} -> {self.tecnico}"

    @property
    def cantidad_pendiente(self):
        if self.unidad_serializada_id:
            return 0 if self.estado == self.Estado.INSTALADA else 1

        return max(
            self.cantidad - self.cantidad_instalada,
            0
        )


# ============================================================
# MOVIMIENTO DE INVENTARIO
# ============================================================

class MovimientoInventario(models.Model):

    class TipoMovimiento(models.TextChoices):
        ENTRADA = "ENTRADA", "Entrada"
        SALIDA = "SALIDA", "Salida"
        AJUSTE = "AJUSTE", "Ajuste"
        DEVOLUCION = "DEVOLUCION", "Devolución"
        ASIGNACION = "ASIGNACION", "Asignación"
        DESASIGNACION = "DESASIGNACION", "Desasignación"
        INSTALACION = "INSTALACION", "Instalación"
        
    remision = models.ForeignKey(
        RemisionIngreso,
        on_delete=models.PROTECT,
        related_name="movimientos",
        blank=True,
        null=True
    )

    producto = models.ForeignKey(
        Producto,
        on_delete=models.PROTECT,
        related_name="movimientos"
    )

    unidad_serializada = models.ForeignKey(
        UnidadSerializada,
        on_delete=models.PROTECT,
        related_name="movimientos",
        blank=True,
        null=True
    )

    asignacion = models.ForeignKey(
        Asignacion,
        on_delete=models.PROTECT,
        related_name="movimientos",
        blank=True,
        null=True
    )

    tipo_movimiento = models.CharField(
        max_length=20,
        choices=TipoMovimiento.choices
    )

    cantidad = models.PositiveIntegerField()

    fecha = models.DateTimeField(
        auto_now_add=True
    )

    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="movimientos"
    )

    observacion = models.TextField(
        blank=True,
        null=True
    )

    class Meta:
        verbose_name = "Movimiento de inventario"
        verbose_name_plural = "Movimientos de inventario"
        ordering = ["-fecha", "-id"]

    def __str__(self):
        return f"{self.producto.nombre} - {self.tipo_movimiento} ({self.cantidad})"