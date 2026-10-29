from django.contrib import admin

# Register your models here.
from django.contrib import admin
from .models import Anomalia


@admin.register(Anomalia)
class AnomaliaAdmin(admin.ModelAdmin):
    list_display = ("codigo", "descripcion", "activa")
    list_filter = ("activa",)
    search_fields = ("codigo", "descripcion")
    list_editable = ("activa",)
    ordering = ("codigo",)
