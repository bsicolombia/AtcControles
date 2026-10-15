from django.contrib import admin

from .models import Tecnicos


@admin.register(Tecnicos)
class TecnicosAdmin(admin.ModelAdmin):
    list_display = ("tecnico_apellido_nombres", "tecnico_cedula", "supervisor", "activo")
    list_filter = ("activo", "supervisor")
    search_fields = ("tecnico_apellido_nombres", "tecnico_cedula")
