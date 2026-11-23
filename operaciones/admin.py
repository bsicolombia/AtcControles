from django.contrib import admin

# Register your models here.
from django.contrib import admin
from .models import Anomalia, TipoOperacion


@admin.register(Anomalia)
class AnomaliaAdmin(admin.ModelAdmin):
    list_display = ("codigo", "descripcion", "activa")
    list_filter = ("activa",)
    search_fields = ("codigo", "descripcion")
    list_editable = ("activa",)
    ordering = ("codigo",)

@admin.register(TipoOperacion)  
class TipoOperacionAdmin(admin.ModelAdmin):
    list_display = (
        "codigo_ods", "nombre", "perfil_estados",
        "instala_material", "requiere_ivt", "activo", "orden",
    )
    list_editable = ("instala_material", "requiere_ivt", "activo", "orden")
    list_filter = ("codigo_ods", "perfil_estados", "activo")
    search_fields = ("codigo_ods", "nombre")
    ordering = ("orden", "nombre")