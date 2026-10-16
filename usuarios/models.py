from django.conf import settings
from django.db import models


class PerfilTecnico(models.Model):
    usuario = models.OneToOneField(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="perfil_tecnico",
    )
    cedula = models.CharField(max_length=20, unique=True)
    supervisor = models.CharField(max_length=150, blank=True)

    def __str__(self):
        return f"{self.usuario.get_full_name()} - CC {self.cedula}"