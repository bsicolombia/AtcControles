from django.db import models
from django.conf import settings
# Create your models here.
# ============================================================
# EVIDENCIAS DE INSTALACIÓN
# ============================================================

class EvidenciaInstalacion(models.Model):
    asignacion = models.ForeignKey(
        "almacen.Asignacion",
        on_delete=models.PROTECT,
        related_name="evidencias",
    )

    imagen = models.ImageField(
        upload_to="evidencias/instalaciones/%Y/%m/",
    )

    descripcion = models.CharField(
        max_length=250,
        blank=True,
    )

    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="evidencias_instalacion",
    )

    fecha_subida = models.DateTimeField(
        auto_now_add=True,
    )

    class Meta:
        verbose_name = "Evidencia de instalación"
        verbose_name_plural = "Evidencias de instalación"
        ordering = ["-fecha_subida"]

    def __str__(self):
        return f"Evidencia de asignación #{self.asignacion_id}"