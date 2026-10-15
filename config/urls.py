from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.contrib.auth import views as auth_views
from django.urls import include, path
from almacen.views import InventarioListView
from almacen.views import DashboardView

admin.site.site_header = "ACT CONTROLES"
admin.site.site_title = "ACT CONTROLES"
admin.site.index_title = "Administración del sistema"

urlpatterns = [
    path("admin/", admin.site.urls),
    path(
        "login/",
        auth_views.LoginView.as_view(
            template_name="registration/login.html",
            redirect_authenticated_user=True,
        ),
        name="login",
    ),
    path("logout/", auth_views.LogoutView.as_view(), name="logout"),
    path("", InventarioListView.as_view(), name="dashboard"),
    path("almacen/", include("almacen.urls")),
    path("tecnicos/", include("usuarios.urls")),
]

# Añadir esto para que Django sirva las imágenes en desarrollo
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

handler403 = "usuarios.views.error_403"