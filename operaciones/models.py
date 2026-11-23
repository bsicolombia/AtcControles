import os

from django.conf import settings
from django.core.exceptions import ValidationError
from django.core.validators import FileExtensionValidator
from django.db import models
from django.utils import timezone

# =====================================================================
# Constantes
# =====================================================================
SIN_ANOMALIA = "0000"
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

class Anomalia(models.Model):
    codigo = models.CharField("Código", max_length=5, unique=True)
    descripcion = models.CharField("Descripción", max_length=150)

    # Clasificación según las tablas operativas
    aplica_ejecutada = models.BooleanField(
        "Disponible en ejecutada",
        default=False,
    )

    aplica_no_ejecutada = models.BooleanField(
        "Disponible en no ejecutada",
        default=False,
    )

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
    @classmethod
    def para_estado(cls, estado):
        """Anomalías que se pueden elegir según el estado de la orden."""
        qs = cls.objects.filter(activa=True)
        if estado in Operacion.ESTADOS_EJECUTADOS:
            return qs.filter(aplica_ejecutada=True)
        if estado in Operacion.ESTADOS_NEGATIVOS:
            return qs.filter(aplica_no_ejecutada=True)
        return qs.none()

class TipoOperacion(models.Model):
    """
    Una fila por cada "Descripción ODS" de la tabla de SAP.

    El código ODS (ZRCM, ZSCO...) NO es único: ZSCO lo comparten
    «WM-Reparación CM / Sustit MED» y «WM-Mtto y Eme Servihogar».
    Lo único es el nombre.
    """

    class PerfilEstados(models.TextChoices):
        MATERIAL = "MATERIAL", "Ejecutada / No ejecutada / Ejecutada suspendida"
        FUGAS = "FUGAS", "Ejecutada sin fugas / Ejecutada con fuga / No ejecutada"

    codigo_ods = models.CharField(
        "Código ODS", max_length=10, blank=True, db_index=True,
        help_text="Abreviación de SAP: ZRCM, ZSCO, etc. Puede repetirse.",
    )
    nombre = models.CharField("Descripción ODS", max_length=120, unique=True)

    perfil_estados = models.CharField(
        "Estados permitidos", max_length=10,
        choices=PerfilEstados.choices, default=PerfilEstados.MATERIAL,
    )
    instala_material = models.BooleanField(
        "Permite instalar material", default=True,
        help_text="Si se desmarca, la orden solo lleva los datos del encabezado.",
    )
    requiere_ivt = models.BooleanField(
        "Pide número de IVT", default=True,
    )

    prefijo_sin_ods = models.CharField(
        "Prefijo si no hay ODS", max_length=5, blank=True,
        help_text="Ej. RSO. Si se llena, este tipo permite registrar sin número de "
                "orden y el sistema genera RSO0001, RSO0002…",
    )

    activo = models.BooleanField(default=True)
    orden = models.PositiveSmallIntegerField(default=0)

    class Meta:
        verbose_name = "Tipo de operación"
        verbose_name_plural = "Tipos de operación"
        ordering = ["orden", "nombre"]

    def __str__(self):
        return self.etiqueta

    @property
    def etiqueta(self):
        return f"{self.codigo_ods} · {self.nombre}" if self.codigo_ods else self.nombre

    # ---- reglas que antes estaban repartidas por todo el código ----
    def estados_validos(self):
        E = Operacion.Estado
        if self.perfil_estados == self.PerfilEstados.FUGAS:
            return [E.EJECUTADA_SIN_FUGAS, E.EJECUTADA_CON_FUGA, E.NO_EJECUTADA]
        return [E.EJECUTADA, E.NO_EJECUTADA, E.EJECUTADA_SUSPENDIDA]

    def instala(self, estado):
        """¿Esta combinación tipo + estado instala material?"""
        return self.instala_material and estado in Operacion.ESTADOS_QUE_INSTALAN

    def pide_ivt(self, estado):
        """¿Esta combinación tipo + estado exige IVT?"""
        return self.requiere_ivt and estado in Operacion.ESTADOS_QUE_REQUIEREN_IVT

def contexto_tipos():
    return {
        "config_tipos": {
            str(t.pk): {
                "instala": t.instala_material,
                "ivt": t.requiere_ivt,
                "sin_ods": t.prefijo_sin_ods,
                "estados": [[e.value, e.label] for e in t.estados_validos()],
            }
            for t in TipoOperacion.objects.filter(activo=True)
        },
        "estados_ejecutados": sorted(e.value for e in Operacion.ESTADOS_EJECUTADOS),
        "estados_ivt": sorted(e.value for e in Operacion.ESTADOS_QUE_REQUIEREN_IVT),
    }
# ---------------------------------------------------------------------
# OPCIONAL · fase 2: "Act Mtto" y "Producto o servicio" de la tabla de SAP
# ---------------------------------------------------------------------
class ActividadMtto(models.Model):
    tipo = models.ForeignKey(
        TipoOperacion, on_delete=models.CASCADE, related_name="actividades"
    )
    codigo = models.CharField("Act Mtto", max_length=5)          # F10, J10, J20...
    descripcion = models.CharField("Producto o servicio", max_length=150)
    activa = models.BooleanField(default=True)

    class Meta:
        ordering = ["tipo", "codigo", "descripcion"]
        # J10 y J30 se repiten con distinto producto/servicio
        unique_together = [("tipo", "codigo", "descripcion")]

    def __str__(self):
        return f"{self.codigo} · {self.descripcion}"
    # Si la usa: en Operacion agregue
    #   actividad = models.ForeignKey(ActividadMtto, null=True, blank=True,
    #                                 on_delete=models.PROTECT, related_name="operaciones")

# =====================================================================
# Operación (orden de trabajo)
# =====================================================================
class Operacion(models.Model):

    class Origen(models.TextChoices):
        MANUAL = "MANUAL", "Registro manual"
        CARGA_MASIVA = "CARGA_MASIVA", "Carga masiva"

    class Estado(models.TextChoices):
        CREADA = "CREADA", "Creada"
        PENDIENTE = "PENDIENTE", "Pendiente"
        ASIGNADA = "ASIGNADA", "Asignada"
        EJECUTADA = "EJECUTADA", "Ejecutada"
        NO_EJECUTADA = "NO_EJECUTADA", "No ejecutada"
        EJECUTADA_SUSPENDIDA = "EJECUTADA_SUSPENDIDA", "Ejecutada suspendida"
        EJECUTADA_SIN_FUGAS = "EJECUTADA_SIN_FUGAS", "Ejecutada sin fugas"  
        EJECUTADA_CON_FUGA = "EJECUTADA_CON_FUGA", "Ejecutada con fuga"     
        CANCELADA = "CANCELADA", "Cancelada"
        ANULADA = "ANULADA", "Anulada"

    ESTADOS_EJECUTADOS = {
        Estado.EJECUTADA,
        Estado.EJECUTADA_SUSPENDIDA,
        Estado.EJECUTADA_SIN_FUGAS,
        Estado.EJECUTADA_CON_FUGA,
    }

    # Estados en los que se instala material
    # operaciones/models.py, dentro de Operacion
    ESTADOS_QUE_INSTALAN = {
        Estado.EJECUTADA,
        Estado.EJECUTADA_SUSPENDIDA,
        Estado.EJECUTADA_SIN_FUGAS,
        Estado.EJECUTADA_CON_FUGA,
    }

    # Estados que requieren IVT y soporte de visita
    ESTADOS_QUE_REQUIEREN_IVT = {
        Estado.EJECUTADA,
        Estado.NO_EJECUTADA,
        Estado.EJECUTADA_SUSPENDIDA,
        Estado.EJECUTADA_SIN_FUGAS,
        Estado.EJECUTADA_CON_FUGA,
    }

    # Estados en los que la orden NO se hizo
    ESTADOS_NEGATIVOS = {
        Estado.NO_EJECUTADA,
    }
    # --- Identificación ---
    numero_orden = models.CharField(max_length=10, unique=True)
    numero_cuenta_contrato = models.CharField(max_length=10, db_index=True)
    tipo_operacion = models.ForeignKey(
        "operaciones.TipoOperacion",
        on_delete=models.PROTECT,
        related_name="operaciones",
    )

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
        return f"{self.numero_orden} - {self.tipo_operacion.nombre}"

    @property
    def es_sin_ods(self):
        p = (self.tipo_operacion.prefijo_sin_ods or "")
        return bool(p) and self.numero_orden.startswith(p) and self.numero_orden[len(p):].isdigit()

    def clean(self):
        super().clean()
        errores = {}

        if self.estado in self.ESTADOS_EJECUTADOS and not self.fecha_ejecucion:
            errores["fecha_ejecucion"] = "Indique la fecha de ejecución."

        if (self.estado in self.ESTADOS_NEGATIVOS
                and self.anomalia_id and self.anomalia.es_sin_anomalia):
            errores["anomalia"] = "Indique la anomalía que explica este estado."

        if (self.tipo_operacion_id
                and self.estado in (self.ESTADOS_EJECUTADOS | self.ESTADOS_NEGATIVOS)):
            validos = {e.value for e in self.tipo_operacion.estados_validos()}
            if self.estado not in validos:
                errores["estado"] = (
                    f"«{self.tipo_operacion.nombre}» no admite el estado "
                    f"{self.get_estado_display()}."
                )

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