from django.urls import path

from . import views

app_name = "tecnicos"

urlpatterns = [
    path("", views.TecnicoListView.as_view(), name="lista"),
    path("nuevo/", views.TecnicoCreateView.as_view(), name="crear"),
    path("<int:pk>/editar/", views.TecnicoUpdateView.as_view(), name="editar"),
    path("carga-masiva/",views.CargaMasivaTecnicosView.as_view(),name="carga_masiva",),
    path("plantilla/",views.DescargarPlantillaTecnicosView.as_view(),name="plantilla",),
]
