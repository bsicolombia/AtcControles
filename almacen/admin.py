from django.contrib import admin

from .models import Asignacion, Categoria, Inventario, MovimientoInventario, Producto, UnidadSerializada


@admin.register(Categoria)
class CategoriaAdmin(admin.ModelAdmin):
    list_display = ("nombre", "prefijo", "estado")
    search_fields = ("nombre",)


from django.contrib import admin
from .models import Producto

@admin.register(Producto)
class ProductoAdmin(admin.ModelAdmin):
    list_display = (
        "codigo_producto",
        "nombre",
        "tipo_inventario",
        "requiere_numero_serie",
        "estado",
    )

    fields = (
        "categoria",
        "nombre",
        "codigo_producto",
        "tipo_inventario",
        "requiere_numero_serie",
        "marca",
        "medida",
        "referencia",
        "unidad_medida",
        "color",
        "presentacion",
        "estado",
    )


@admin.register(UnidadSerializada)
class UnidadAdmin(admin.ModelAdmin):
    list_display = ("codigo", "producto", "numero_serie", "marca", "estado", "fecha_ingreso")
    list_filter = ("estado", "producto")
    search_fields = ("codigo", "numero_serie")


@admin.register(Inventario)
class InventarioAdmin(admin.ModelAdmin):
    list_display = ("producto", "cantidad_actual", "cantidad_asignada", "cantidad_minima")


@admin.register(Asignacion)
class AsignacionAdmin(admin.ModelAdmin):
    list_display = ("producto", "unidad_serializada", "tecnico", "estado", "fecha_asignacion", "operacion")
    list_filter = ("estado",)


@admin.register(MovimientoInventario)
class MovimientoAdmin(admin.ModelAdmin):
    list_display = ("fecha", "tipo_movimiento", "producto", "unidad_serializada", "cantidad", "usuario")
    list_filter = ("tipo_movimiento",)
