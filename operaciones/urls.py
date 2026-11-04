
from django.urls import path
from . import views

app_name = "operaciones"

urlpatterns = [
    path("carga_sap/",views.CargaSAPView.as_view(),name="carga_sap",),
    path("carga_sap_detalle/<int:pk>/",views.CargaSAPDetalleView.as_view(),name="carga_sap_detalle",),
    path("carga-sap/plantilla/",views.descargar_plantilla_sap,name="descargar_plantilla_sap",),
    path("cargas-sap/",views.HistorialCargaSAPView.as_view(),name="carga_sap_historial",),
    path("conciliacion-sap/",views.ConciliacionSAPView.as_view(),name="conciliacion_sap",),
    path("consulta-orden/",views.ConsultaOrdenView.as_view(),name="consulta_orden",),
    path("errores/", views.ErroresActualesView.as_view(), name="errores_actuales"),
    path("soporte/<int:pk>/", views.SoporteOperacionView.as_view(), name="soporte"),
    path("operacion/<int:pk>/editar/",views.OperacionUpdateView.as_view(),name="editar_instalacion",),
]
