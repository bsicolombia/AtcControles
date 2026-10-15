from django.db import models
from django.conf import settings
from django.utils import timezone
from almacen.models import UnidadSerializada, Producto

class Operacion(models.Model):
    class Origen(models.TextChoices):
        MANUAL = "MANUAL", "Registro manual"
        CARGA_MASIVA = "CARGA_MASIVA", "Carga masiva"

    class TipoOperacion(models.TextChoices):
        INSTALACION = "INSTALACION", "Instalación"
        RETIRO = "RETIRO", "Retiro"
        MANTENIMIENTO = "MANTENIMIENTO", "Mantenimiento"

    class Estado(models.TextChoices):
        CREADA = "CREADA", "Creada"
        PENDIENTE = "PENDIENTE", "Pendiente"
        ASIGNADA = "ASIGNADA", "Asignada"
        EJECUTADA = "EJECUTADA", "Ejecutada"
        NO_EJECUTADA = "NO_EJECUTADA", "No ejecutada"
        EFECTIVA = "EFECTIVA", "Efectiva"
        CANCELADA ="CANCELADA", "Cancelada"
        ANULADA = "ANULADA", "Anulada"

    numero_orden = models.CharField(max_length=10, unique=True)
    
    numero_cuenta_contrato = models.CharField(max_length=10, db_index=True)

    tipo_operacion = models.CharField(
        max_length=30,
        choices=TipoOperacion.choices
    )

    tecnico = models.ForeignKey(
        "usuarios.Tecnicos",
        on_delete=models.PROTECT,
        related_name="operaciones",
        null=True,
        blank=True
    )

    estado = models.CharField(
        max_length=20,
        choices=Estado.choices,
        default=Estado.PENDIENTE
    )

    fecha_operacion = models.DateField(
        default=timezone.localdate
    )

    observaciones = models.TextField(
        blank=True
    )

    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="operaciones_creadas"
    )
    
    origen = models.CharField(
        max_length=20, choices=Origen.choices, default=Origen.MANUAL
    )
    
    fecha_ejecucion = models.DateField(null=True, blank=True)
    
    # Columnas extra del Excel/PowerBI sin migrar cada vez
    datos_origen = models.JSONField(default=dict, blank=True)

    fecha_creacion = models.DateTimeField(
        auto_now_add=True
    )

    fecha_modificacion = models.DateTimeField(
        auto_now=True
    )
    
class CargaSAP(models.Model):
    archivo = models.FileField(upload_to="cargas_sap/")
    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name="cargas_sap"
    )
    fecha = models.DateTimeField(auto_now_add=True)
    total_filas = models.PositiveIntegerField(default=0)
    filas_ok = models.PositiveIntegerField(default=0)
    filas_error = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ["-fecha"]


class RegistroSAP(models.Model):
    class Estado(models.TextChoices):
        OK = "OK", "Conciliado"
        ERROR = "ERROR", "Con error"

    carga = models.ForeignKey(CargaSAP, on_delete=models.CASCADE, related_name="registros")
    fila = models.PositiveIntegerField()

    # Lo que vino en el Excel (se guarda tal cual, aunque sea inválido)
    numero_orden = models.CharField(max_length=30, blank=True)
    codigo_producto = models.CharField(max_length=30, blank=True)
    numero_serie = models.CharField(max_length=100, blank=True)
    cantidad = models.PositiveIntegerField(default=1)

    estado = models.CharField(max_length=10, choices=Estado.choices)
    error = models.TextField(blank=True)

    # Referencias resueltas (solo si hubo coincidencia)
    operacion = models.ForeignKey("operaciones.Operacion", null=True, blank=True,
                                  on_delete=models.PROTECT, related_name="registros_sap")
    tecnico = models.ForeignKey("usuarios.Tecnicos", null=True, blank=True,
                                on_delete=models.PROTECT, related_name="registros_sap")
    producto = models.ForeignKey(Producto, null=True, blank=True,
                                 on_delete=models.PROTECT, related_name="registros_sap")
    unidad_serializada = models.ForeignKey(UnidadSerializada, null=True, blank=True,
                                           on_delete=models.PROTECT, related_name="registros_sap")

    class Meta:
        ordering = ["fila"]