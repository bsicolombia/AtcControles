import os
from almacen.models import UnidadSerializada, Categoria,Producto
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

    tipo_con_ods = models.ForeignKey(
        "self", null=True, blank=True,
        on_delete=models.SET_NULL,
        related_name="variantes_sin_ods",
        verbose_name="Tipo cuando tenga ODS",
        help_text="Solo para tipos «sin ODS»: al asignar la ODS real, "
                "la orden pasa a este tipo.",
    )

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

    @property
    def tipo_actividades(self):
        """Tipo cuyas actividades aplican: las de su versión con ODS, si es «sin ODS»."""
        return self.tipo_con_ods or self
    
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
        """El IVT es obligatorio en toda operación."""
        return True

def contexto_tipos():
    return {
        "config_tipos": {
            str(t.pk): {
                "instala": t.instala_material,
                "ivt": t.requiere_ivt,
                "sin_ods": t.prefijo_sin_ods,
                "act_tipo": str(t.tipo_con_ods_id or t.pk),
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
        MANUAL = "MANUAL", "Digitada en el sistema"
        CARGA_MASIVA = "CARGA_MASIVA", "Carga masiva"
        APP_TECNICO = "APP_TECNICO", "Reportada por el técnico (app)"

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

        # --- Actividad de mantenimiento (depende del tipo de operación) ---
    actividad = models.ForeignKey(
        "operaciones.ActividadMtto",
        on_delete=models.PROTECT,
        related_name="operaciones",
        null=True,
        blank=True,
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

        if (self.actividad_id and self.tipo_operacion_id
                and self.actividad.tipo_id != self.tipo_operacion.tipo_actividades.pk):
            errores["actividad"] = "La actividad no corresponde al tipo de operación."

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
        ("CICB", "Se corrigió en CICB"),
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

    # --- Relaciones con CICB ---
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

class RegistroActividad(models.Model):
    class Accion(models.TextChoices):
        ORDEN_CREADA = "ORDEN_CREADA", "Orden registrada"
        ORDEN_EDITADA = "ORDEN_EDITADA", "Orden editada"
        ORDEN_REACTIVADA = "ORDEN_REACTIVADA", "Orden reactivada (eliminada)"
        ODS_ASIGNADA = "ODS_ASIGNADA", "ODS real asignada"
        ENTREGA = "ENTREGA", "Entrega de material"
        DEVOLUCION = "DEVOLUCION", "Devolución de material"
        INGRESO = "INGRESO", "Ingreso a almacén"
        IVT = "IVT", "Movimiento de IVT"
        ESTADO_UNIDAD = "ESTADO_UNIDAD", "Cambio de estado de unidad"
        INCONSISTENCIA = "INCONSISTENCIA", "Inconsistencia resuelta"
        CARGA_SAP = "CARGA_SAP", "Carga SAP"
        PASSWORD_FALLIDA = "PASSWORD_FALLIDA", "Contraseña incorrecta"
        CARGA_INVENTARIO = "CARGA_INVENTARIO", "Carga de inventario físico"

    ACCIONES_ORDEN = ("ORDEN_CREADA", "ORDEN_EDITADA", "ORDEN_REACTIVADA", "ODS_ASIGNADA")

    fecha = models.DateTimeField(auto_now_add=True, db_index=True)
    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True,
        on_delete=models.SET_NULL, related_name="actividades",
    )
    usuario_nombre = models.CharField(max_length=150, blank=True)   # copia por si borran el usuario
    accion = models.CharField(max_length=30, choices=Accion.choices, db_index=True)
    referencia = models.CharField(max_length=60, blank=True, db_index=True)  # nº de orden, código de acta...
    tecnico = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True,
        on_delete=models.SET_NULL, related_name="actividades_como_tecnico",
    )
    detalle = models.TextField(blank=True)
    datos = models.JSONField(default=dict, blank=True)    # antes/después, motivo, etc.
    ip = models.CharField(max_length=45, blank=True)

    class Meta:
        ordering = ["-fecha", "-id"]
        verbose_name = "Registro de actividad"
        verbose_name_plural = "Registros de actividad"

    def __str__(self):
        return f"{self.fecha:%d/%m/%Y %H:%M} · {self.usuario_nombre} · {self.get_accion_display()}"

class CargaInventario(models.Model):
    categoria = models.ForeignKey(
        Categoria, on_delete=models.PROTECT, related_name="cargas_inventario"
    )
    archivo = models.FileField(upload_to="cargas_inventario/%Y/%m/")
    nombre_original = models.CharField(max_length=255)
    hash_archivo = models.CharField(max_length=64, blank=True, db_index=True)
    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL, on_delete=models.PROTECT,
        related_name="cargas_inventario",
    )
    fecha = models.DateTimeField(auto_now_add=True, db_index=True)
    observaciones = models.TextField(blank=True)

    total_filas = models.PositiveIntegerField(default=0)
    filas_ok = models.PositiveIntegerField(default=0)
    filas_error = models.PositiveIntegerField(default=0)
    faltantes_archivo = models.PositiveIntegerField(default=0)
    disponibles_sistema = models.PositiveIntegerField(
        default=0, help_text="Unidades disponibles en el sistema al momento de la carga."
    )

    class Meta:
        ordering = ["-fecha", "-id"]
        verbose_name = "Carga de inventario"
        verbose_name_plural = "Cargas de inventario"

    def __str__(self):
        return f"Carga inventario #{self.pk} · {self.categoria} · {self.fecha:%d/%m/%Y %H:%M}"

    @property
    def coincide_todo(self):
        return self.filas_error == 0 and self.faltantes_archivo == 0


class RegistroInventario(models.Model):

    class Resultado(models.TextChoices):
        OK = "OK", "Coincide"
        ERROR = "ERROR", "Con error"
        FALTANTE = "FALTANTE", "Falta en el archivo"

    class TipoError(models.TextChoices):
        SERIE_VACIA = "SERIE_VACIA", "Fila sin número de serie"
        SERIE_DUPLICADA_ARCHIVO = "SERIE_DUPLICADA_ARCHIVO", "Serie repetida en el archivo"
        PRODUCTO_VACIO = "PRODUCTO_VACIO", "Fila sin código de material"
        PRODUCTO_NO_EXISTE = "PRODUCTO_NO_EXISTE", "El código de material no existe"
        PRODUCTO_FUERA_CATEGORIA = "PRODUCTO_FUERA_CATEGORIA", "El material es de otra categoría"
        PRODUCTO_NO_SERIALIZADO = "PRODUCTO_NO_SERIALIZADO", "El material no se maneja por serial"
        NO_EXISTE_EN_SISTEMA = "NO_EXISTE_EN_SISTEMA", "La serie no existe en el sistema"
        SERIE_OTRO_PRODUCTO = "SERIE_OTRO_PRODUCTO", "La serie pertenece a otro material"
        CODIGO_UNIDAD_NO_COINCIDE = "CODIGO_UNIDAD_NO_COINCIDE", "El código de unidad no corresponde a la serie"
        ESTADO_NO_DISPONIBLE = "ESTADO_NO_DISPONIBLE", "La unidad no figura disponible en el sistema"
        FALTANTE_EN_ARCHIVO = "FALTANTE_EN_ARCHIVO", "Disponible en el sistema pero no está en el archivo"
        STATUS_NO_ALMACEN = "STATUS_NO_ALMACEN", "En SAP no figura en almacén (status distinto de ALMA)"
        STATUS_NO_COINCIDE = "STATUS_NO_COINCIDE", "El status de SAP no coincide con el estado en CICB"
        ORDEN_NO_COINCIDE = "ORDEN_NO_COINCIDE", "La orden de SAP no coincide con la de CICB"

    carga = models.ForeignKey(
        CargaInventario, on_delete=models.CASCADE, related_name="registros"
    )
    fila = models.PositiveIntegerField(null=True, blank=True)  # null en faltantes

    # Tal como viene en el archivo
    numero_serie = models.CharField(max_length=100, blank=True, db_index=True)
    codigo_unidad = models.CharField(max_length=50, blank=True)
    codigo_producto = models.CharField(max_length=30, blank=True)
    orden_five = models.CharField("Orden en CICB", max_length=30, blank=True)

    resultado = models.CharField(max_length=10, choices=Resultado.choices, db_index=True)
    tipo_error = models.CharField(max_length=30, choices=TipoError.choices, blank=True, db_index=True)
    detalle = models.TextField(blank=True)

    # Relaciones y foto del estado en el sistema
    producto = models.ForeignKey(Producto, null=True, blank=True, on_delete=models.PROTECT,
                                 related_name="registros_inventario")
    unidad = models.ForeignKey(UnidadSerializada, null=True, blank=True, on_delete=models.PROTECT,
                               related_name="registros_inventario")
    estado_sistema = models.CharField(max_length=20, blank=True)
    descripcion = models.CharField(max_length=200, blank=True)
    status_sap = models.CharField("Status en SAP", max_length=20, blank=True)
    centro = models.CharField(max_length=10, blank=True)
    almacen_sap = models.CharField("Almacén SAP", max_length=10, blank=True)
    tipo_stock = models.CharField(max_length=10, blank=True)
    orden_sap = models.CharField("Orden en SAP", max_length=30, blank=True)

    class Meta:
        ordering = ["resultado", "fila", "id"]
        indexes = [models.Index(fields=["carga", "resultado"])]

    def __str__(self):
        return f"{self.numero_serie or '—'} · {self.get_resultado_display()}"

class CargaStock(models.Model):
    """Una carga del reporte MB52 de SAP (stock por material)."""
    archivo = models.FileField(upload_to="cargas_stock/%Y/%m/")
    nombre_original = models.CharField(max_length=255)
    hash_archivo = models.CharField(max_length=64, blank=True, db_index=True)
    usuario = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT,
                                related_name="cargas_stock")
    fecha = models.DateTimeField(auto_now_add=True, db_index=True)
    observaciones = models.TextField(blank=True)

    total_materiales = models.PositiveIntegerField(default=0)
    coinciden = models.PositiveIntegerField(default=0)
    con_diferencia = models.PositiveIntegerField(default=0)
    tipos_sin_ods = models.CharField(max_length=300, blank=True)

    class Meta:
        ordering = ["-fecha", "-id"]
        verbose_name = "Carga de stock MB52"
        verbose_name_plural = "Cargas de stock MB52"

    def __str__(self):
        return f"MB52 #{self.pk} · {self.fecha:%d/%m/%Y %H:%M}"

    @property
    def coincide_todo(self):
        return self.con_diferencia == 0


class RegistroStock(models.Model):

    class Resultado(models.TextChoices):
        OK = "OK", "Coincide"
        ERROR = "ERROR", "Con diferencia"

    class TipoError(models.TextChoices):
        MATERIAL_NO_EXISTE = "MATERIAL_NO_EXISTE", "El material no existe en el sistema"
        SAP_MAYOR = "SAP_MAYOR", "SAP tiene más de lo que CICB espera"
        SAP_MENOR = "SAP_MENOR", "SAP tiene menos de lo que CICB espera"
        NO_EN_MB52 = "NO_EN_MB52", "CICB espera stock y el material no viene en el MB52"
        CANTIDAD_INVALIDA = "CANTIDAD_INVALIDA", "Cantidad ilegible en el archivo"

    carga = models.ForeignKey(CargaStock, on_delete=models.CASCADE, related_name="registros")
    fila = models.PositiveIntegerField(null=True, blank=True)

    codigo_producto = models.CharField(max_length=30, db_index=True)
    descripcion = models.CharField(max_length=200, blank=True)
    producto = models.ForeignKey("almacen.Producto", null=True, blank=True,
                                 on_delete=models.PROTECT, related_name="registros_stock")

    # Lo que dice SAP (MB52)
    sap_libre = models.DecimalField(max_digits=14, decimal_places=3, default=0)
    sap_calidad = models.DecimalField(max_digits=14, decimal_places=3, default=0)
    sap_bloqueado = models.DecimalField(max_digits=14, decimal_places=3, default=0)
    sap_transito = models.DecimalField(max_digits=14, decimal_places=3, default=0)

    # Lo que espera CICB
    five_disponible = models.DecimalField(max_digits=14, decimal_places=3, default=0)
    five_asignado = models.DecimalField(max_digits=14, decimal_places=3, default=0)
    five_sin_ods = models.DecimalField(max_digits=14, decimal_places=3, default=0)
    five_total = models.DecimalField(max_digits=14, decimal_places=3, default=0)

    diferencia = models.DecimalField(max_digits=14, decimal_places=3, default=0)  # SAP - CICB

    resultado = models.CharField(max_length=10, choices=Resultado.choices, db_index=True)
    tipo_error = models.CharField(max_length=30, choices=TipoError.choices, blank=True, db_index=True)
    detalle = models.TextField(blank=True)

    class Meta:
        ordering = ["codigo_producto"]
        indexes = [models.Index(fields=["carga", "resultado"])]