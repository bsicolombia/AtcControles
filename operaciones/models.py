import os

from django.conf import settings
from django.core.exceptions import ValidationError
from django.core.validators import FileExtensionValidator
from django.db import models
from django.utils import timezone

# =====================================================================
# Constantes
# =====================================================================
SIN_ANOMALIA = "00"
EXTENSIONES_SOPORTE = ["pdf", "jpg", "jpeg", "png"]
LIMITE_SOPORTE_MB = 5


# =====================================================================
# Utilidades de archivos
# =====================================================================
def ruta_soporte_operacion(instance, filename):
    """Guarda el archivo como operaciones/soportes/NUMERO_ORDEN.ext"""
    ext = os.path.splitext(filename)[1].lower()
    nombre = instance.numero_orden.strip().upper().replace(" ", "_")
    return os.path.join("operaciones", "soportes", f"{nombre}{ext}")


def validar_tamano_archivo(archivo):
    if archivo.size > LIMITE_SOPORTE_MB * 1024 * 1024:
        raise ValidationError(f"El archivo no puede superar {LIMITE_SOPORTE_MB} MB.")


# =====================================================================
# Anomalía
# =====================================================================
class Anomalia(models.Model):
    codigo = models.CharField("Código", max_length=5, unique=True)
    descripcion = models.CharField("Descripción", max_length=150)
    activa = models.BooleanField(default=True)

    class Meta:
        verbose_name = "Anomalía"
        verbose_name_plural = "Anomalías"
        ordering = ["codigo"]

    def __str__(self):
        return f"{self.codigo} - {self.descripcion}"

    @property
    def es_sin_anomalia(self):
        return self.codigo == SIN_ANOMALIA


# =====================================================================
# Operación (orden de trabajo)
# =====================================================================
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
        CANCELADA = "CANCELADA", "Cancelada"
        ANULADA = "ANULADA", "Anulada"

    # Estados en los que se instala material
    ESTADOS_QUE_INSTALAN = {Estado.EJECUTADA, Estado.EFECTIVA}
    # Estados en los que la orden NO se hizo (exigen una anomalía real)
    ESTADOS_NEGATIVOS = {Estado.NO_EJECUTADA, Estado.CANCELADA, Estado.ANULADA}

    # --- Identificación ---
    numero_orden = models.CharField(max_length=10, unique=True)
    numero_cuenta_contrato = models.CharField(max_length=10, db_index=True)
    tipo_operacion = models.CharField(max_length=30, choices=TipoOperacion.choices)

    # --- Responsables ---
    tecnico = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="operaciones_tecnico",
    )
    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="operaciones_creadas",
    )

    # --- Estado y anomalía ---
    estado = models.CharField(
        max_length=20, choices=Estado.choices, default=Estado.PENDIENTE
    )
    # null=True: las operaciones antiguas y las de carga masiva no la traen
    anomalia = models.ForeignKey(
        Anomalia,
        on_delete=models.PROTECT,
        related_name="operaciones",
        null=True,
        blank=True,
    )

    # --- Fechas ---
    fecha_operacion = models.DateField(default=timezone.localdate)
    fecha_ejecucion = models.DateField(null=True, blank=True)

    # --- Soporte y notas ---
    archivo_soporte = models.FileField(
        "Archivo de soporte (PDF o imagen)",
        upload_to=ruta_soporte_operacion,
        blank=True,
        null=True,
        validators=[
            FileExtensionValidator(EXTENSIONES_SOPORTE),
            validar_tamano_archivo,
        ],
    )
    observaciones = models.TextField(blank=True)

    # --- Origen ---
    origen = models.CharField(
        max_length=20, choices=Origen.choices, default=Origen.MANUAL
    )
    # Columnas extra del Excel/PowerBI sin migrar cada vez
    datos_origen = models.JSONField(default=dict, blank=True)

    # --- Auditoría ---
    fecha_creacion = models.DateTimeField(auto_now_add=True)
    fecha_modificacion = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Operación"
        verbose_name_plural = "Operaciones"
        ordering = ["-fecha_operacion", "-id"]
        indexes = [
            # Aceleran el resumen de ejecución por técnico / año / mes / día
            models.Index(fields=["fecha_operacion"], name="op_fecha_idx"),
            models.Index(fields=["tecnico", "fecha_operacion"], name="op_tec_fecha_idx"),
        ]

    def __str__(self):
        return f"{self.numero_orden} - {self.get_tipo_operacion_display()}"

    def clean(self):
        super().clean()
        errores = {}
        if self.estado in self.ESTADOS_QUE_INSTALAN and not self.fecha_ejecucion:
            errores["fecha_ejecucion"] = "Indique la fecha de ejecución."
        if (
            self.estado in self.ESTADOS_NEGATIVOS
            and self.anomalia_id
            and self.anomalia.es_sin_anomalia
        ):
            errores["anomalia"] = "Indique la anomalía que explica este estado."
        if errores:
            raise ValidationError(errores)

    def save(self, *args, **kwargs):
        # Si suben un archivo nuevo, borra el anterior: sin huérfanos ni
        # nombres con sufijos aleatorios.
        if self.pk and self.archivo_soporte and not self.archivo_soporte._committed:
            anterior = Operacion.objects.filter(pk=self.pk).first()
            if anterior and anterior.archivo_soporte:
                anterior.archivo_soporte.delete(save=False)
        super().save(*args, **kwargs)


# =====================================================================
# Carga masiva SAP
# =====================================================================
class CargaSAP(models.Model):
    archivo = models.FileField(upload_to="cargas_sap/")
    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="cargas_sap",
    )
    fecha = models.DateTimeField(auto_now_add=True)

    total_filas = models.PositiveIntegerField(default=0)
    filas_ok = models.PositiveIntegerField(default=0)
    filas_error = models.PositiveIntegerField(default=0)
    filas_ignoradas = models.PositiveIntegerField(default=0)

    hash_archivo = models.CharField(max_length=64, blank=True, db_index=True)

    class Meta:
        ordering = ["-fecha"]

    def __str__(self):
        return f"Carga SAP #{self.pk} - {self.fecha}"


# =====================================================================
# Registro de cada fila de SAP
# =====================================================================
class RegistroSAP(models.Model):

    class Estado(models.TextChoices):
        OK = "OK", "Conciliado"
        ERROR = "ERROR", "Con error"
        IGNORADO = "IGNORADO", "Ya conciliado antes"

    RESUELTO_POR = [
        ("FIVE", "Se corrigió en Five"),
        ("SAP", "Se corrigió en SAP"),
        ("AMBOS", "Se corrigió en ambos"),
    ]

    # --- Origen de la fila ---
    carga = models.ForeignKey(
        CargaSAP, on_delete=models.CASCADE, related_name="registros"
    )
    fila = models.PositiveIntegerField()

    # --- Datos tal como vienen de SAP ---
    numero_orden = models.CharField(max_length=30, blank=True, db_index=True)
    cuenta_contrato = models.CharField(max_length=30, blank=True, default="")
    codigo_producto = models.CharField(max_length=30, blank=True)
    numero_serie = models.CharField(max_length=100, blank=True, db_index=True)
    codigo_unidad = models.CharField(
        max_length=50, blank=True, default="", db_index=True
    )
    cantidad = models.PositiveIntegerField(default=1)
    cantidad_sap_original = models.CharField(max_length=50, blank=True, default="")

    # --- Resultado de la conciliación ---
    estado = models.CharField(max_length=10, choices=Estado.choices)
    error = models.TextField(blank=True)
    tipo_error = models.CharField(max_length=30, blank=True, default="")
    diagnostico = models.TextField(blank=True, default="")
    comparacion = models.JSONField(default=dict, blank=True)

    # --- Relaciones con Five ---
    operacion = models.ForeignKey(
        "operaciones.Operacion",
        null=True, blank=True,
        on_delete=models.PROTECT,
        related_name="registros_sap",
    )
    tecnico = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True, blank=True,
        on_delete=models.SET_NULL,
        related_name="registros_sap_tecnico",
    )
    producto = models.ForeignKey(
        "almacen.Producto",
        null=True, blank=True,
        on_delete=models.PROTECT,
        related_name="registros_sap",
    )
    unidad_serializada = models.ForeignKey(
        "almacen.UnidadSerializada",
        null=True, blank=True,
        on_delete=models.PROTECT,
        related_name="registros_sap",
    )

    # --- Seguimiento de repeticiones y resolución ---
    clave_unica = models.CharField(max_length=40, blank=True, default="", db_index=True)
    registro_original = models.ForeignKey(
        "self",
        null=True, blank=True,
        on_delete=models.SET_NULL,
        related_name="repeticiones",
    )
    # último estado de ese caso, aún sin resolver
    vigente = models.BooleanField(default=False, db_index=True)
    resuelto = models.BooleanField(default=False, db_index=True)
    resuelto_en = models.ForeignKey(
        CargaSAP,
        null=True, blank=True,
        on_delete=models.SET_NULL,
        related_name="errores_resueltos",
    )
    resuelto_por = models.CharField(
        max_length=10, blank=True, default="", choices=RESUELTO_POR
    )

    class Meta:
        ordering = ["fila"]
        indexes = [models.Index(fields=["clave_unica", "estado"])]

    def __str__(self):
        return f"Fila {self.fila} - {self.numero_orden} - {self.estado}"