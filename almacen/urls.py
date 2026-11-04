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
    path("asignaciones/nueva/", views.EntregaCreateView.as_view(), name="asignacion_crear"),
    path("asignaciones/<int:pk>/instalar/", views.AsignacionInstalarView.as_view(), name="asignacion_instalar"),
    path("tecnicos/<int:pk>/devolver/", views.TecnicoDevolverView.as_view(), name="tecnico_devolver"),
    path("tecnicos/<int:pk>/asignaciones/", views.TecnicoAsignacionesView.as_view(), name="tecnico_asignaciones"),
    path("api/unidades-disponibles/", views.UnidadesDisponiblesJson.as_view(), name="api_unidades"),
    path("movimientos/", views.MovimientoListView.as_view(), name="movimientos"),
    path(
    "operaciones/<int:pk>/editar/",
    views.editar_instalacion,
    name="editar_instalacion",
),

path("tecnicos/<int:pk>/asignaciones/", views.TecnicoAsignacionesView.as_view(), name="tecnico_asignaciones"),
path("tecnicos/<int:pk>/devolver/", views.TecnicoDevolverView.as_view(), name="tecnico_devolver"),
    path("inventario/", views.InventarioListView.as_view(), name="inventario"), 
    path("entregas/", views.EntregaListView.as_view(), name="entregas"),
    path("entregas/nueva/", views.EntregaCreateView.as_view(), name="entrega_nueva"),
    path("entregas/<int:pk>/", views.EntregaDetailView.as_view(), name="entrega_detalle"),
    path("entregas/<int:pk>/pdf/", views.EntregaPdfView.as_view(), name="entrega_pdf"),
    path("entregas/<int:pk>/excel/", views.EntregaExcelView.as_view(), name="entrega_excel"),
    path("api/validar-serial/", views.ValidarSerialJson.as_view(), name="api_validar_serial"),
    path("api/validar-serial-instalacion/", views.ValidarSerialInstalacionJson.as_view(), name="api_validar_serial_instalacion"),
    path("inconsistencias/", views.InconsistenciaListView.as_view(), name="inconsistencias"),
    path("inconsistencias/<int:pk>/resolver/", views.InconsistenciaResolverView.as_view(), name="inconsistencia_resolver"),

]