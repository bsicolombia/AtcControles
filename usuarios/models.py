
from django.db import models
from django.contrib.auth.models import AbstractUser
from django.core.exceptions import ValidationError


class User(AbstractUser):

    ROLES = [
        ('ADMIN', 'Administrador'),
        ('SUPERVISOR', 'Supervisor'),
        ('TECNICO', 'Tecnico'),
        ('COORDINADOR', 'Coordinador'),
        ('ALMACENISTA', 'Almacenista'),
        ('DIGITADOR', 'Digitador'),
    ]

    cedula = models.CharField(
        max_length=20,
        unique=True
    )

    rol = models.CharField(
        max_length=20,
        choices=ROLES
    )

    supervisor = models.ForeignKey(
        'self',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='tecnicos'
    )

    def clean(self):
        super().clean()

        if self.rol == 'TECNICO' and not self.supervisor:
            raise ValidationError({
                'supervisor': 'Un técnico debe tener un supervisor.'
            })

        if self.rol != 'TECNICO' and self.supervisor:
            raise ValidationError({
                'supervisor': 'Solo los técnicos pueden tener supervisor.'
            })

        if self.supervisor and self.supervisor.rol != 'SUPERVISOR':
            raise ValidationError({
                'supervisor': 'Debes seleccionar un usuario con rol Supervisor.'
            })

    def __str__(self):
        return f"{self.cedula} - {self.username}"
