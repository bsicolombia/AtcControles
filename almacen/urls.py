from django.urls import path

from . import views
from .views import DashboardView
app_name = "almacen"

urlpatterns = [
    
    path("inicio/", DashboardView.as_view(), name="inicio"),
    path("categorias/", views.CategoriaListView.as_view(), name="categorias"),
    path("categorias/nueva/", views.CategoriaCreateView.as_view(), name="categoria_crear"),
    path("categorias/<int:pk>/editar/", views.CategoriaUpdateView.as_view(), name="categoria_editar"),

    path("productos/", views.ProductoListView.as_view(), name="productos"),
    path("productos/nuevo/", views.ProductoCreateView.as_view(), name="producto_crear"),
    path("productos/<int:pk>/editar/", views.ProductoUpdateView.as_view(), name="producto_editar"),

    
    path("inventario/<int:pk>/minimo/", views.MinimoUpdateView.as_view(), name="minimo_editar"),
    path("inventario/<int:pk>/editar-almacen/", views.InventarioAlmacenUpdateView.as_view(), name="inventario_almacen_editar"),
    path("ingresos/nuevo/", views.IngresoView.as_view(), name="ingreso_crear"),

    path("unidades/", views.UnidadListView.as_view(), name="unidades"),
    path("unidades/<int:pk>/", views.UnidadDetailView.as_view(), name="unidad_detalle"),
    path("unidades/<int:pk>/editar/", views.UnidadUpdateView.as_view(), name="unidad_editar"),
    path("unidades/<int:pk>/estado/", views.UnidadEstadoView.as_view(), name="unidad_estado"),

    path("asignaciones/", views.AsignacionListView.as_view(), name="asignaciones"),
    path("asignaciones/nueva/", views.AsignacionCreateView.as_view(), name="asignacion_crear"),
    path("asignaciones/<int:pk>/instalar/", views.AsignacionInstalarView.as_view(), name="asignacion_instalar"),
    path("asignaciones/<int:pk>/devolver/", views.AsignacionDevolverView.as_view(), name="asignacion_devolver"),
    path("tecnicos/<int:pk>/asignaciones/", views.TecnicoAsignacionesView.as_view(), name="tecnico_asignaciones"),
    path("api/unidades-disponibles/", views.UnidadesDisponiblesJson.as_view(), name="api_unidades"),

    path("movimientos/", views.MovimientoListView.as_view(), name="movimientos"),
]