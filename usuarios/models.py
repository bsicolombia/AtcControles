from django.db import models


class Tecnicos(models.Model):
    supervisor = models.CharField(max_length=150)
    tecnico_cedula = models.CharField(max_length=10, unique=True)
    tecnico_apellido_nombres = models.CharField(max_length=150)
    activo = models.BooleanField(default=True)

    class Meta:
        verbose_name = "Técnico"
        verbose_name_plural = "Técnicos"
        ordering = ["tecnico_apellido_nombres"]

    def __str__(self):
        return self.tecnico_apellido_nombres
