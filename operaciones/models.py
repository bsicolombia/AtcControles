from django.db import models
from django.conf import settings
from django.utils import timezone

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