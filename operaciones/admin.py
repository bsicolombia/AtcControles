from django.contrib import admin

from .models import ActividadMtto, Anomalia, Operacion, TipoOperacion


@admin.register(Anomalia)
class AnomaliaAdmin(admin.ModelAdmin):
    list_display = ("codigo", "descripcion", "activa")
    list_filter = ("activa",)
    search_fields = ("codigo", "descripcion")
    list_editable = ("activa",)
    ordering = ("codigo",)


class ActividadMttoInline(admin.TabularInline):
    model = ActividadMtto
    extra = 1
    fields = ("codigo", "descripcion", "activa")
    ordering = ("codigo", "descripcion")


@admin.register(TipoOperacion)
class TipoOperacionAdmin(admin.ModelAdmin):
    list_display = (
        "codigo_ods", "nombre", "perfil_estados",
        "instala_material", "requiere_ivt", "n_actividades", "activo", "orden",
    )
    list_editable = ("instala_material", "requiere_ivt", "activo", "orden")
    list_filter = ("codigo_ods", "perfil_estados", "activo")
    search_fields = ("codigo_ods", "nombre")
    ordering = ("orden", "nombre")
    inlines = [ActividadMttoInline]

    @admin.display(description="Actividades")
    def n_actividades(self, obj):
        return obj.actividades.count()


@admin.register(ActividadMtto)
class ActividadMttoAdmin(admin.ModelAdmin):
    list_display = ("codigo", "descripcion", "tipo", "activa")
    list_editable = ("activa",)
    list_filter = ("tipo", "activa")
    search_fields = ("codigo", "descripcion", "tipo__nombre")
    list_select_related = ("tipo",)
    ordering = ("tipo__orden", "codigo", "descripcion")