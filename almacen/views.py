from collections import OrderedDict
from .models import RemisionIngreso
from django.contrib import messages
from django.db import transaction
from django.db.models import F, Q
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect
from django.urls import reverse, reverse_lazy
from django.utils import timezone
from django.utils.http import url_has_allowed_host_and_scheme
from django.http import HttpResponseRedirect
from django.views import View
from django.views.generic import (
    CreateView,
    DetailView,
    FormView,
    ListView,
    TemplateView,
    UpdateView,
)
from django.db.models import Prefetch
from usuarios.mixins import RolRequeridoMixin
from django.contrib.auth import get_user_model
from usuarios.models import PerfilTecnico

User = get_user_model()
from usuarios.roles import ADMINISTRADOR
from urllib.parse import urlencode
from . import services
from . import services_sap
from .forms import (
    AsignacionForm,
    CambioEstadoForm,
    CategoriaForm,
    IngresoForm,
    InstalacionForm,
    InstalacionMasivaForm,
    InventarioAlmacenForm,
    MinimoForm,
    ProductoForm,
    UnidadForm,
)
from .models import (
    Asignacion,
    Categoria,
    Inventario,
    MovimientoInventario,
    Producto,
    UnidadSerializada,
)
from operaciones.models import CargaSAP
from .forms import CargaSAPForm
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from io import BytesIO
from operaciones.models import CargaSAP, RegistroSAP
from django.db.models import Sum


def tecnicos_queryset():
    return User.objects.filter(
        groups__name="Tecnico",
        perfil_tecnico__isnull=False,
    ).select_related(
        "perfil_tecnico"
    ).distinct().order_by(
        "first_name",
        "last_name",
        "username",
    )
# ============================================================
# UTILIDADES
# ============================================================
class BuscarMixin:
    """Búsqueda ?q=; código exacto y otros campos parciales."""

    paginate_by = 20
    campos_busqueda = ()

    def filtrar(self, qs):
        q = self.request.GET.get("q", "").strip()

        if q and self.campos_busqueda:
            cond = Q()

            for campo in self.campos_busqueda:
                if campo == "codigo":
                    # Coincidencia exacta para el código
                    cond |= Q(codigo__iexact=q)
                else:
                    # Búsqueda parcial para los demás campos
                    cond |= Q(**{f"{campo}__icontains": q})

            qs = qs.filter(cond)

        return qs

    def get_queryset(self):
        return self.filtrar(super().get_queryset())

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        params = self.request.GET.copy()
        params.pop("page", None)
        ctx["qs"] = params.urlencode()
        return ctx
class FormPageMixin:
    """Usa el template genérico form.html (salvo que la vista defina otro)."""

    template_name = "form.html"
    cancel_url = None
    titulo = ""
    boton = "Guardar"
    subtitulo = ""

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx.update(
            titulo=self.titulo,
            boton=self.boton,
            subtitulo=self.subtitulo,
            cancel_url=self.cancel_url or reverse("dashboard"),
        )
        return ctx


# ============================================================
# DASHBOARD
# ============================================================

class DashboardView(RolRequeridoMixin, TemplateView):
    template_name = "dashboard.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        E = UnidadSerializada.Estado

        conteo = {
            estado: UnidadSerializada.objects.filter(estado=estado).count()
            for estado in E.values
        }
        total = sum(conteo.values())

        ctx["unidades"] = [
            {
                "estado": estado,
                "nombre": E(estado).label,
                "n": conteo[estado],
                "pct": round(conteo[estado] * 100 / total, 1) if total else 0,
            }
            for estado in (E.DISPONIBLE, E.ASIGNADO, E.INSTALADO, E.DAÑADO, E.BAJA)
        ]
        ctx["total_unidades"] = total

        ctx["bajo_stock"] = (
            Inventario.objects
            .filter(
                cantidad_minima__gt=0,
                cantidad_actual__lte=F("cantidad_minima"),
                producto__estado=True,
            )
            .select_related("producto")[:8]
        )

        productos = (
            Producto.objects
            .filter(estado=True)
            .select_related("inventario")
            .prefetch_related("unidades_serializadas")
            .order_by("nombre")
        )

        productos_stock = []
        for producto in productos:
            if producto.es_serializado:
                cantidad_disponible = (
                    producto.unidades_serializadas
                    .filter(estado=E.DISPONIBLE)
                    .count()
                )
            else:
                cantidad_disponible = (
                    producto.inventario.cantidad_actual
                    if hasattr(producto, "inventario")
                    else 0
                )
            productos_stock.append(
                {"producto": producto, "cantidad_disponible": cantidad_disponible}
            )
        ctx["productos_stock"] = productos_stock

        ctx["materiales"] = (
            Inventario.objects
            .filter(
                producto__tipo_inventario=Producto.TipoInventario.CANTIDAD,
                producto__estado=True,
            )
            .select_related("producto")[:8]
        )

        ctx["pendientes"] = (
            Asignacion.objects.filter(estado=Asignacion.Estado.ASIGNADA).count()
        )
        ctx["productos_activos"] = Producto.objects.filter(estado=True).count()

        ctx["ultimos"] = (
            MovimientoInventario.objects
            .select_related("producto", "unidad_serializada", "usuario")[:8]
        )

        return ctx


# ============================================================
# CATEGORÍAS
# ============================================================

class CategoriaListView(RolRequeridoMixin, BuscarMixin, ListView):
    model = Categoria
    template_name = "almacen/categoria_list.html"
    campos_busqueda = ("nombre", "prefijo")


class CategoriaFormMixin(FormPageMixin, RolRequeridoMixin):
    roles_permitidos = (ADMINISTRADOR,)

    model = Categoria
    form_class = CategoriaForm
    success_url = reverse_lazy("almacen:categorias")
    cancel_url = reverse_lazy("almacen:categorias")

    def form_valid(self, form):
        messages.success(self.request, "Categoría guardada.")
        return super().form_valid(form)


class CategoriaCreateView(CategoriaFormMixin, CreateView):
    titulo = "Nuevo material"
    boton = "Guardar material"


class CategoriaUpdateView(CategoriaFormMixin, UpdateView):
    titulo = "Editar categoría"
    boton = "Guardar cambios"


# ============================================================
# PRODUCTOS
# ============================================================

class ProductoListView(RolRequeridoMixin, BuscarMixin, ListView):
    model = Producto
    template_name = "almacen/producto_list.html"
    campos_busqueda = (
        "nombre",
        "codigo_producto",
        "marca",
        "referencia",
        "medida",
        "categoria__nombre",
    )

    def get_queryset(self):
        qs = super().get_queryset().select_related("categoria", "inventario")

        cat = self.request.GET.get("categoria")
        if cat:
            qs = qs.filter(categoria_id=cat)

        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["categorias"] = Categoria.objects.filter(estado=True)
        ctx["categoria_sel"] = self.request.GET.get("categoria", "")
        return ctx


class ProductoFormMixin(FormPageMixin, RolRequeridoMixin):
    model = Producto
    form_class = ProductoForm
    success_url = reverse_lazy("almacen:productos")
    cancel_url = reverse_lazy("almacen:productos")

    def form_valid(self, form):
        resp = super().form_valid(form)
        Inventario.objects.get_or_create(producto=self.object)
        messages.success(self.request, "Producto guardado.")
        return resp


class ProductoCreateView(ProductoFormMixin, CreateView):
    titulo = "Nuevo material"
    boton = "Guardar material"
    subtitulo = (
        "Un material por cada tipo y medida: "
        "«Codo 1/2» y «Codo 1/4» son material distintos."
    )


class ProductoUpdateView(ProductoFormMixin, UpdateView):
    titulo = "Editar material"
    boton = "Guardar cambios"


# ============================================================
# INVENTARIO
# ============================================================

class InventarioListView(RolRequeridoMixin, BuscarMixin, ListView):
    model = Inventario
    template_name = "almacen/inventario_list.html"
    campos_busqueda = (
        "producto__nombre",
        "producto__codigo_producto",
        "producto__medida",
        "producto__categoria__nombre",
    )

    def get_queryset(self):
        qs = (
            super().get_queryset()
            .select_related("producto", "producto__categoria")
        )

        if self.request.GET.get("bajo"):
            qs = qs.filter(
                cantidad_minima__gt=0,
                cantidad_actual__lte=F("cantidad_minima"),
            )

        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        tecnico_sel = self.request.GET.get("tecnico", "").strip()

        ctx["tecnicos_lista"] = tecnicos_queryset()
        ctx["tecnico_sel"] = tecnico_sel

        if tecnico_sel and tecnico_sel.isdigit():
            ctx.update(_obtener_contexto_asignaciones(self.request))
        else:
            # Inventario general: agregamos el dato SAP a cada fila
            sap = _sap_total_por_producto()
            filas = list(ctx["object_list"])
            for i in filas:
                i.instalado_sap = sap.get(i.producto_id, 0)
            ctx["object_list"] = filas

        return ctx


class InventarioAlmacenUpdateView(RolRequeridoMixin, UpdateView):
    """Edición manual del inventario de almacén (antes era una función sin rol)."""

    model = Inventario
    form_class = InventarioAlmacenForm
    template_name = "almacen/inventario_almacen_editar.html"
    context_object_name = "inventario"
    success_url = reverse_lazy("almacen:inventario")

    def get_queryset(self):
        return super().get_queryset().select_related("producto")

    def form_valid(self, form):
        messages.success(
            self.request,
            f"Inventario de almacén de {form.instance.producto.nombre_completo} "
            "actualizado correctamente.",
        )
        return super().form_valid(form)


class MinimoUpdateView(FormPageMixin, RolRequeridoMixin, UpdateView):
    model = Inventario
    form_class = MinimoForm
    success_url = reverse_lazy("almacen:inventario")
    cancel_url = reverse_lazy("almacen:inventario")

    titulo = "Stock mínimo"
    boton = "Guardar mínimo"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["subtitulo"] = self.object.producto.nombre_completo
        return ctx


# ============================================================
# INGRESOS
# ============================================================

# ============================================================
# INGRESOS
# ============================================================

class IngresoView(FormPageMixin, RolRequeridoMixin, FormView):
    form_class = IngresoForm
    template_name = "almacen/ingreso_form.html"
    success_url = reverse_lazy("almacen:unidades")
    cancel_url = reverse_lazy("almacen:inventario")

    titulo = "Registrar ingreso"
    boton = "Registrar ingreso"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        # Pasamos todas las remisiones existentes para alimentar el buscador/select
        ctx["remisiones_existentes"] = RemisionIngreso.objects.all().order_by("-fecha_registro")[:50]
        return ctx

    def post(self, request, *args, **kwargs):
        form = self.form_class(request.POST, request.FILES)
        if form.is_valid():
            return self.form_valid(form)
        return self.render_to_response(self.get_context_data(form=form))

    def form_valid(self, form):
        d = form.cleaned_data
        producto = d["producto"]
        num_orden = d["numero_orden"]
        imagen = d.get("comprobante_imagen")

        try:
            with transaction.atomic():
                # Buscar o crear la remisión
                remision, creada = RemisionIngreso.objects.get_or_create(
                    numero_remision=num_orden,
                    defaults={
                        "comprobante_imagen": imagen,
                        "usuario": self.request.user,
                    }
                )

                # Si ya existía pero subieron foto nueva, la guardamos
                if not creada and imagen and not remision.comprobante_imagen:
                    remision.comprobante_imagen = imagen
                    remision.save()

                # Registrar el ingreso asociando la remisión
                unidades = services.registrar_ingreso(
                    producto=producto,
                    cantidad=d["cantidad"],
                    usuario=self.request.user,
                    fecha_ingreso=timezone.localdate(),
                    seriales=d.get("seriales", []),
                    observacion=d["observacion"],
                    remision=remision
                )

        except services.ErrorInventario as e:
            form.add_error(None, str(e))
            return self.render_to_response(self.get_context_data(form=form))

        messages.success(
            self.request,
            f"Ingreso registrado con éxito asociado a la remisión {num_orden}."
        )
        return super().form_valid(form)

# ============================================================
# UNIDADES SERIALIZADAS
# ============================================================

class UnidadListView(RolRequeridoMixin, BuscarMixin, ListView):
    model = UnidadSerializada
    template_name = "almacen/unidad_list.html"

    campos_busqueda = (
        "codigo",
        "numero_serie",
        "marca",
        "referencia",
        "producto__nombre",
        "producto__codigo_producto",
        "remision__numero_remision",
    )

    def get_queryset(self):
        tecnico_id = self.request.GET.get("tecnico", "").strip()

        asignaciones_qs = (
            Asignacion.objects
            .select_related(
                "tecnico",
                "operacion",
                "producto",
            )
            .order_by("-fecha_modificacion", "-id")
        )

        qs = (
            super()
            .get_queryset()
            .select_related("producto", "remision")
        )

        # Filtrar unidades por el técnico seleccionado.
        if tecnico_id and tecnico_id.isdigit():
            qs = qs.filter(
                (
                    Q(
                        estado=UnidadSerializada.Estado.ASIGNADO,
                        asignaciones__tecnico_id=int(tecnico_id),
                        asignaciones__estado=Asignacion.Estado.ASIGNADA,
                    )
                    | Q(
                        estado=UnidadSerializada.Estado.INSTALADO,
                        asignaciones__tecnico_id=int(tecnico_id),
                        asignaciones__estado=Asignacion.Estado.INSTALADA,
                    )
                )
            ).distinct()

            asignaciones_qs = asignaciones_qs.filter(
                tecnico_id=int(tecnico_id)
            )

        qs = qs.prefetch_related(
            Prefetch(
                "asignaciones",
                queryset=asignaciones_qs,
                to_attr="asignaciones_lista",
            )
        )

        estado = self.request.GET.get("estado", "").strip()
        if estado:
            qs = qs.filter(estado=estado)

        producto = self.request.GET.get("producto", "").strip()
        if producto and producto.isdigit():
            qs = qs.filter(producto_id=int(producto))

        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        # Total de resultados filtrados.
        ctx["total_unidades_filtradas"] = self.object_list.count()

        ctx["estados"] = UnidadSerializada.Estado.choices
        ctx["estado_sel"] = self.request.GET.get("estado", "")
        ctx["producto_sel"] = self.request.GET.get("producto", "")
        ctx["tecnico_sel"] = self.request.GET.get("tecnico", "")

        # Usuarios que pertenecen al grupo Técnico.
        ctx["tecnicos_lista"] = (
            User.objects
            .filter(
                groups__name="Tecnico",
                perfil_tecnico__isnull=False,
            )
            .select_related("perfil_tecnico")
            .distinct()
            .order_by("first_name", "last_name", "username")
        )

        ctx["productos"] = (
            Producto.objects
            .filter(
                tipo_inventario=Producto.TipoInventario.SERIALIZADO,
                estado=True,
            )
            .order_by("nombre")
        )

        ctx["estados_editables"] = list(services.TRANSICIONES)

        ctx["hay_instaladas"] = self.object_list.filter(
            estado=UnidadSerializada.Estado.INSTALADO
        ).exists()

        for unidad in ctx["object_list"]:
            unidad.asignacion_relevante = None

            if unidad.estado == UnidadSerializada.Estado.ASIGNADO:
                estado_asignacion = Asignacion.Estado.ASIGNADA

            elif unidad.estado == UnidadSerializada.Estado.INSTALADO:
                estado_asignacion = Asignacion.Estado.INSTALADA

            else:
                continue

            unidad.asignacion_relevante = next(
                (
                    asignacion
                    for asignacion in getattr(
                        unidad,
                        "asignaciones_lista",
                        [],
                    )
                    if asignacion.estado == estado_asignacion
                ),
                None,
            )

        return ctx



class UnidadDetailView(RolRequeridoMixin, DetailView):
    model = UnidadSerializada
    template_name = "almacen/unidad_detail.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["asignaciones"] = self.object.asignaciones.select_related("tecnico")
        ctx["movimientos"] = self.object.movimientos.select_related("usuario")
        return ctx


class UnidadEstadoView(FormPageMixin, RolRequeridoMixin, FormView):
    form_class = CambioEstadoForm
    cancel_url = reverse_lazy("almacen:unidades")
    success_url = reverse_lazy("almacen:unidades")

    titulo = "Cambiar estado"
    boton = "Guardar estado"

    def dispatch(self, request, *args, **kwargs):
        self.unidad = get_object_or_404(
            UnidadSerializada.objects.select_related("producto"),
            pk=kwargs["pk"],
        )

        if self.unidad.estado not in services.TRANSICIONES:
            messages.error(
                request,
                f"La unidad {self.unidad.codigo} "
                f"está en «{self.unidad.get_estado_display()}» "
                "y su estado no se cambia desde aquí.",
            )
            return redirect("almacen:unidades")

        return super().dispatch(request, *args, **kwargs)

    def get_form_kwargs(self):
        kwargs = super().get_form_kwargs()
        kwargs["unidad"] = self.unidad
        return kwargs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["subtitulo"] = (
            f"{self.unidad.codigo} · "
            f"estado actual: {self.unidad.get_estado_display()}"
        )
        return ctx

    def form_valid(self, form):
        try:
            services.cambiar_estado(
                unidad=self.unidad,
                nuevo_estado=form.cleaned_data["estado"],
                usuario=self.request.user,
                motivo=form.cleaned_data["motivo"],
            )
        except services.ErrorInventario as e:
            form.add_error(None, str(e))
            return self.form_invalid(form)

        messages.success(
            self.request, f"Estado de {self.unidad.codigo} actualizado."
        )
        return super().form_valid(form)


class UnidadUpdateView(FormPageMixin, RolRequeridoMixin, UpdateView):
    model = UnidadSerializada
    form_class = UnidadForm

    titulo = "Editar unidad"
    boton = "Guardar cambios"

    def get_success_url(self):
        return reverse("almacen:unidad_detalle", args=[self.object.pk])

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["cancel_url"] = self.get_success_url()
        ctx["subtitulo"] = self.object.codigo
        return ctx


# ============================================================
# ASIGNACIONES · CONSOLIDADO
# ============================================================
def _sap_total_por_producto():
    """Total reportado por SAP en la última carga, agrupado por producto."""
    ultima_carga = CargaSAP.objects.order_by("-fecha", "-pk").first()

    if not ultima_carga:
        return {}

    registros = (
        RegistroSAP.objects
        .filter(carga=ultima_carga, producto__isnull=False)
        .values("producto_id")
        .annotate(total=Sum("cantidad"))
    )

    return {r["producto_id"]: r["total"] or 0 for r in registros}
def _obtener_datos_sap():
    """
    Obtiene las cantidades reportadas en la última carga SAP.
    Incluye filas OK y ERROR si tienen técnico y producto identificados.
    No modifica el inventario.
    """
    ultima_carga = CargaSAP.objects.order_by("-fecha", "-pk").first()

    if not ultima_carga:
        return {}

    registros = (
        RegistroSAP.objects
        .filter(
            carga=ultima_carga,
            tecnico__isnull=False,
            producto__isnull=False,
        )
        .values("tecnico_id", "producto_id")
        .annotate(total=Sum("cantidad"))
    )

    return {
        (r["tecnico_id"], r["producto_id"]): r["total"] or 0
        for r in registros
    }

def _datos_vacios():
    return {
        "asignados": 0,
        "five": 0,
        "sap": 0,
        "debe_tener": 0,
    }
def _obtener_contexto_asignaciones(request):
    """Arma el consolidado técnico × producto."""

    q = request.GET.get("q", "").strip()
    estado_sel = request.GET.get("estado", "").strip()
    tecnico_sel = request.GET.get("tecnico", "").strip()

    qs = (
        Asignacion.objects
        .select_related(
            "tecnico",
            "producto",
            "producto__categoria",
            "unidad_serializada",
            "operacion",
        )
        .order_by("-fecha_asignacion", "-id")
    )

    # Filtro por técnico seleccionado
    if tecnico_sel.isdigit():
        qs = qs.filter(tecnico_id=int(tecnico_sel))

    # Filtro por estado
    if estado_sel:
        qs = qs.filter(estado=estado_sel)

    # Búsqueda por cédula, nombre, producto y demás campos
    # Búsqueda por cédula, nombre, producto, código o serial
    if q:
        filtros = (
            Q(tecnico__username__icontains=q)
            | Q(tecnico__first_name__icontains=q)
            | Q(tecnico__last_name__icontains=q)
            | Q(tecnico__perfil_tecnico__cedula__icontains=q)
            | Q(tecnico__perfil_tecnico__supervisor__icontains=q)
            | Q(producto__nombre__icontains=q)
            | Q(producto__codigo_producto__icontains=q)
            | Q(producto__marca__icontains=q)
            | Q(producto__referencia__icontains=q)
            | Q(producto__medida__icontains=q)
            | Q(producto__color__icontains=q)
            | Q(producto__presentacion__icontains=q)
            | Q(operacion__numero_orden__icontains=q)
            | Q(operacion__numero_cuenta_contrato__icontains=q)
            | Q(unidad_serializada__codigo__icontains=q)
            | Q(unidad_serializada__numero_serie__icontains=q)
        )

        # Comprobar si la búsqueda corresponde al serial de un medidor.
        unidades_encontradas = UnidadSerializada.objects.filter(
            Q(codigo__iexact=q) | Q(numero_serie__iexact=q)
        )

        unidad = unidades_encontradas.first()

        if unidad:
            # Mostrar únicamente la asignación vigente del medidor.
            qs = qs.filter(
                unidad_serializada=unidad,
                estado__in=[
                    Asignacion.Estado.ASIGNADA,
                    Asignacion.Estado.INSTALADA,
                ],
            )
        else:
            qs = qs.filter(filtros)
            
    productos_dict = OrderedDict()
    tecnicos_dict = OrderedDict()
    datos_sap = _obtener_datos_sap()

    for a in qs:
        productos_dict.setdefault(a.producto_id, a.producto)

        bloque = tecnicos_dict.setdefault(
            a.tecnico_id,
            {"tecnico": a.tecnico, "productos": {}},
        )

        datos = bloque["productos"].setdefault(
            a.producto_id,
            _datos_vacios(),
        )

        # Una asignación devuelta ya no está en poder del técnico.
        if a.estado == Asignacion.Estado.DEVUELTA:
            continue

        datos["asignados"] += a.cantidad or 0

        # Cantidades instaladas
        if a.estado == Asignacion.Estado.INSTALADA:
            datos["five"] += a.cantidad or 0
        else:
            datos["five"] += a.cantidad_instalada or 0

    productos = list(productos_dict.values())
    tecnicos = []

    base_unidades = reverse("almacen:unidades")
    base_inventario = reverse("dashboard")

    for bloque in tecnicos_dict.values():
        filas = {
            "ASIGNADOS": [],
            "FIVE": [],
            "SAP": [],
            "DEBE_TENER": [],
        }

        productos_detalle = []

        for p in productos:
            d = bloque["productos"].get(p.pk, _datos_vacios())

            asignados = d["asignados"]
            five = d["five"]

            sap = datos_sap.get(
                (bloque["tecnico"].pk, p.pk),
                0,
            )

            debe_tener = max(asignados - sap, 0)

            coincide = sap == five

            # Datos para cada celda de la tabla.
            # Las URL conservan el técnico y el producto.
            def crear_celda(cantidad, tipo):
                if p.es_serializado:
                    estado = (
                        UnidadSerializada.Estado.INSTALADO
                        if tipo == "FIVE"
                        else UnidadSerializada.Estado.ASIGNADO
                    )

                    parametros = urlencode({
                        "tecnico": bloque["tecnico"].pk,
                        "producto": p.pk,
                        "estado": estado,
                    })

                    url = f"{base_unidades}?{parametros}"

                else:
                    # Para materiales por cantidad, mostrar
                    # el resumen filtrado por técnico y producto.
                    parametros = urlencode({
                        "tecnico": bloque["tecnico"].pk,
                        "q": p.codigo_producto or p.nombre,
                    })

                    url = f"{base_inventario}?{parametros}"

                return {
                    "cantidad": cantidad,
                    "url": url,
                }

            filas["ASIGNADOS"].append(
                crear_celda(asignados, "ASIGNADOS")
            )
            filas["FIVE"].append(
                crear_celda(five, "FIVE")
            )
            celda_sap = crear_celda(sap, "SAP")
            celda_sap["coincide"] = coincide
            celda_sap["estado_conciliacion"] = (
                "ok" if coincide else "error"
            )
            filas["SAP"].append(celda_sap)
            filas["DEBE_TENER"].append(
                crear_celda(debe_tener, "DEBE_TENER")
            )

            productos_detalle.append({
                "producto": p,
                "asignados": asignados,
                "five": five,
                "sap": sap,
                "debe_tener": debe_tener,
            })

        tecnicos.append({
            "tecnico": bloque["tecnico"],
            "filas": filas,
            "productos_detalle": productos_detalle,
        })

    return {
        "tecnicos": tecnicos,
        "productos": productos,
        "estados": Asignacion.Estado.choices,
        "estado_sel": estado_sel,
        "q": q,
        "tecnico_sel": tecnico_sel,
    }
    
class AsignacionListView(RolRequeridoMixin, TemplateView):
    template_name = "almacen/asignacion_list.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        ctx.update(_obtener_contexto_asignaciones(self.request))

        ctx["tecnicos_lista"] = (
            User.objects
            .filter(
                groups__name="Tecnico",
                perfil_tecnico__isnull=False,
            )
            .select_related("perfil_tecnico")
            .distinct()
            .order_by("first_name", "last_name", "username")
        )

        return ctx

# ============================================================
# ASIGNACIONES · CREAR
# ============================================================

class AsignacionCreateView(FormPageMixin, RolRequeridoMixin, FormView):
    form_class = AsignacionForm
    template_name = "almacen/asignacion_form.html"

    success_url = reverse_lazy("almacen:asignaciones")
    cancel_url = reverse_lazy("almacen:asignaciones")

    titulo = "Nueva asignación"
    boton = "Entregar al técnico"
    subtitulo = (
        "Entrega de material o de una unidad serializada a un técnico."
    )

    def form_valid(self, form):
        d = form.cleaned_data

        try:
            services.asignar(
                producto=d["producto"],
                tecnico=d["tecnico"],
                usuario=self.request.user,
                cantidad=d["cantidad"],
                unidad=d.get("unidad"),
                fecha=d["fecha_asignacion"],
                observaciones=d["observaciones"],
            )
        except services.ErrorInventario as e:
            form.add_error(None, str(e))
            return self.form_invalid(form)

        messages.success(self.request, "Asignación registrada.")
        return super().form_valid(form)


# ============================================================
# INSTALAR / DEVOLVER · PANTALLA DEL TÉCNICO (varios materiales)
# ============================================================


class TecnicoAsignacionesView(RolRequeridoMixin, TemplateView):
    """
    Un técnico, un formulario: se escribe la operación una sola vez
    y las cantidades o unidades instaladas de cada material.
    """

    template_name = "almacen/asignacion_instalar.html"

    def _cargar(self):
        self.tecnico = get_object_or_404(
            User.objects.filter(
                groups__name="Tecnico",
                perfil_tecnico__isnull=False,
            ).select_related("perfil_tecnico").distinct(),
            pk=self.kwargs["pk"],
        )

        pendientes = (
            Asignacion.objects
            .filter(
                tecnico=self.tecnico,
                estado=Asignacion.Estado.ASIGNADA,
            )
            .select_related(
                "producto",
                "unidad_serializada",
                "operacion",
            )
            .order_by(
                "producto__nombre",
                "producto__medida",
                "id",
            )
        )

        self.pendientes = [
            a for a in pendientes
            if a.cantidad_pendiente > 0
        ]

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        self._cargar()

        # Conservar los valores del formulario si hubo errores.
        if self.request.method == "POST":
            for a in self.pendientes:
                a.valor = self.request.POST.get(
                    f"cant_{a.pk}", ""
                )
                a.marcado = bool(
                    self.request.POST.get(f"inst_{a.pk}")
                )

        ctx["tecnico"] = self.tecnico
        ctx["pendientes"] = self.pendientes
        ctx["form"] = (
            kwargs.get("form") or InstalacionMasivaForm()
        )

        return ctx

    def post(self, request, *args, **kwargs):
        self._cargar()
        form = InstalacionMasivaForm(request.POST)

        items = []
        errores = []

        for a in self.pendientes:
            nombre = a.producto.nombre_completo

            # Unidad serializada: se marca como instalada.
            if a.unidad_serializada_id:
                if request.POST.get(f"inst_{a.pk}"):
                    items.append((a, None))
                continue

            raw = request.POST.get(
                f"cant_{a.pk}", ""
            ).strip()

            if not raw:
                continue

            try:
                cant = int(raw)
            except ValueError:
                errores.append(
                    f"{nombre}: la cantidad no es válida."
                )
                continue

            if cant <= 0:
                continue

            if cant > a.cantidad_pendiente:
                errores.append(
                    f"{nombre}: solo tiene "
                    f"{a.cantidad_pendiente} pendiente(s)."
                )
                continue

            items.append((a, cant))

        if not items and not errores:
            errores.append(
                "Indique al menos un material para instalar."
            )

        if errores or not form.is_valid():
            for error in errores:
                messages.error(request, error)

            return self.render_to_response(
                self.get_context_data(form=form)
            )

        d = form.cleaned_data

        try:
            with transaction.atomic():
                op = services.obtener_o_crear_operacion(
                    numero_orden=d["numero_orden"],
                    cuenta_contrato=d["numero_cuenta_contrato"],
                    tipo=d["tipo_operacion"],
                    tecnico=self.tecnico,
                    usuario=request.user,
                    fecha=d["fecha_instalacion"],
                )

                for asignacion, cantidad in items:
                    services.instalar(
                        asignacion=asignacion,
                        usuario=request.user,
                        cantidad=cantidad,
                        fecha_instalacion=d["fecha_instalacion"],
                        operacion=op,
                        observaciones=d["observaciones"],
                    )

        except services.ErrorInventario as e:
            messages.error(request, str(e))
            return self.render_to_response(
                self.get_context_data(form=form)
            )

        messages.success(
            request,
            f"Instalación registrada en la orden "
            f"{op.numero_orden} ({len(items)} material(es)).",
        )

        return redirect(
            "almacen:tecnico_asignaciones",
            pk=self.tecnico.pk,
        )



# ============================================================
# INSTALAR · UNA SOLA ASIGNACIÓN
# ============================================================

class AsignacionInstalarView(FormPageMixin, RolRequeridoMixin, FormView):
    form_class = InstalacionForm
    template_name = "almacen/asignacion_instalar.html"

    titulo = "Registrar instalación"
    boton = "Registrar instalación"

    def dispatch(self, request, *args, **kwargs):
        self.asignacion = get_object_or_404(
            Asignacion.objects.select_related(
                "tecnico", "producto", "unidad_serializada"
            ),
            pk=kwargs["pk"],
        )

        if self.asignacion.estado != Asignacion.Estado.ASIGNADA:
            messages.error(
                request,
                "Esta asignación ya no está pendiente de instalar.",
            )
            return redirect("almacen:asignaciones")

        return super().dispatch(request, *args, **kwargs)

    def get_success_url(self):
        return reverse(
            "almacen:tecnico_asignaciones", args=[self.asignacion.tecnico_id]
        )

    def get_form_kwargs(self):
        kwargs = super().get_form_kwargs()
        kwargs["asignacion"] = self.asignacion
        return kwargs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        a = self.asignacion

        detalle = (
            a.unidad_serializada.codigo
            if a.unidad_serializada
            else f"{a.cantidad_pendiente} × {a.producto.nombre_completo}"
        )

        ctx["asignacion"] = a
        ctx["subtitulo"] = f"{detalle} · asignado a {a.tecnico}"
        ctx["cancel_url"] = self.get_success_url()
        return ctx

    def form_valid(self, form):
        d = form.cleaned_data
        cantidad = d.get("cantidad")

        try:
            with transaction.atomic():
                op = services.obtener_o_crear_operacion(
                    numero_orden=d["numero_orden"],
                    cuenta_contrato=d["numero_cuenta_contrato"],
                    tipo=d["tipo_operacion"],
                    tecnico=self.asignacion.tecnico,
                    usuario=self.request.user,
                    fecha=d["fecha_instalacion"],
                )
                services.instalar(
                    asignacion=self.asignacion,
                    usuario=self.request.user,
                    cantidad=cantidad,
                    fecha_instalacion=d["fecha_instalacion"],
                    operacion=op,
                    observaciones=d["observaciones"],
                )
        except services.ErrorInventario as e:
            messages.error(self.request, str(e))
            return redirect("almacen:asignaciones")

        if cantidad:
            self.asignacion.refresh_from_db()
            restan = self.asignacion.cantidad_pendiente
            messages.success(
                self.request,
                f"Instaladas {cantidad} × "
                f"{self.asignacion.producto.nombre_completo}. "
                f"Al técnico le quedan {restan} por instalar.",
            )
        else:
            messages.success(self.request, "Instalación registrada.")

        return super().form_valid(form)


# ============================================================
# DEVOLVER ASIGNACIÓN
# ============================================================

class AsignacionDevolverView(RolRequeridoMixin, View):

    def post(self, request, pk):
        asignacion = get_object_or_404(Asignacion, pk=pk)

        try:
            services.devolver(asignacion=asignacion, usuario=request.user)
            messages.success(
                request,
                "Devolución registrada: el material volvió al almacén.",
            )
        except services.ErrorInventario as e:
            messages.error(request, str(e))

        siguiente = request.POST.get("next", "")
        if siguiente and url_has_allowed_host_and_scheme(
            siguiente, allowed_hosts={request.get_host()}
        ):
            return redirect(siguiente)

        return redirect("almacen:asignaciones")


# ============================================================
# UNIDADES DISPONIBLES · JSON
# ============================================================
class UnidadesDisponiblesJson(RolRequeridoMixin, View):

    def get(self, request):
        producto = request.GET.get("producto")
        data = {"serializado": False, "stock": 0, "unidades": []}

        if producto:
            p = get_object_or_404(Producto, pk=producto)
            data["serializado"] = p.es_serializado

            if p.es_serializado:
                unidades = (
                    UnidadSerializada.objects
                    .filter(producto=p, estado=UnidadSerializada.Estado.DISPONIBLE)
                    .order_by("numero_serie", "codigo")
                )
                data["unidades"] = [
                    {"id": u.id, "texto": u.etiqueta} for u in unidades
                ]
                data["stock"] = len(data["unidades"])
            else:
                inv = Inventario.objects.filter(producto=p).first()
                data["stock"] = inv.cantidad_actual if inv else 0

        return JsonResponse(data)


# ============================================================
# MOVIMIENTOS / KARDEX
# ============================================================

class MovimientoListView(RolRequeridoMixin, BuscarMixin, ListView):
    model = MovimientoInventario
    template_name = "almacen/movimiento_list.html"
    campos_busqueda = (
        "producto__nombre",
        "producto__codigo_producto",
        "unidad_serializada__codigo",
        "usuario__username",
        "observacion",
    )

    def get_queryset(self):
        qs = super().get_queryset().select_related(
            "producto", "unidad_serializada", "usuario"
        )

        tipo = self.request.GET.get("tipo")
        if tipo:
            qs = qs.filter(tipo_movimiento=tipo)

        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["tipos"] = MovimientoInventario.TipoMovimiento.choices
        ctx["tipo_sel"] = self.request.GET.get("tipo", "")
        return ctx
    
class HistorialCargaSAPView(RolRequeridoMixin, ListView):
    model = CargaSAP
    template_name = "almacen/carga_sap_historial.html"
    context_object_name = "cargas"
    paginate_by = 25

    def get_queryset(self):
        qs = CargaSAP.objects.select_related("usuario").order_by(
            "-fecha", "-pk"
        )

        desde = self.request.GET.get("desde", "").strip()
        hasta = self.request.GET.get("hasta", "").strip()
        estado = self.request.GET.get("estado", "").strip()
        q = self.request.GET.get("q", "").strip()

        if desde:
            qs = qs.filter(fecha__date__gte=desde)

        if hasta:
            qs = qs.filter(fecha__date__lte=hasta)

        if estado == "OK":
            qs = qs.filter(filas_error=0, filas_ok__gt=0)
        elif estado == "ERROR":
            qs = qs.filter(filas_error__gt=0)

        if q:
            qs = qs.filter(
                Q(archivo__icontains=q)
                | Q(usuario__username__icontains=q)
            )

        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        ctx["desde"] = self.request.GET.get("desde", "")
        ctx["hasta"] = self.request.GET.get("hasta", "")
        ctx["estado_sel"] = self.request.GET.get("estado", "")
        ctx["q"] = self.request.GET.get("q", "")

        return ctx

class CargaSAPView(FormPageMixin, RolRequeridoMixin, FormView):
    form_class = CargaSAPForm
    template_name = "almacen/carga_sap_form.html"
    titulo = "Carga masiva SAP"
    boton = "Cargar y validar"
    cancel_url = reverse_lazy("almacen:asignaciones")

    def form_valid(self, form):
        try:
            carga = services_sap.procesar_carga(
                form.cleaned_data["archivo"],
                self.request.user,
            )

        except services_sap.ErrorArchivo as exc:
            form.add_error("archivo", str(exc))
            return self.form_invalid(form)

        except Exception:
            # Registra el detalle técnico en los logs del servidor.
            import logging
            logging.getLogger(__name__).exception(
                "Error procesando archivo de SAP"
            )
            form.add_error(
                "archivo",
                "No fue posible procesar el archivo. "
                "Verifique su contenido e inténtelo nuevamente.",
            )
            return self.form_invalid(form)

        messages.success(
            self.request,
            f"Carga procesada: {carga.filas_ok} conciliados y "
            f"{carga.filas_error} con errores.",
        )

        return HttpResponseRedirect(
            reverse(
                "almacen:carga_sap_detalle",
                kwargs={"pk": carga.pk},
            )
        )


class CargaSAPDetalleView(RolRequeridoMixin, DetailView):
    model = CargaSAP
    template_name = "almacen/carga_sap_detalle.html"
    context_object_name = "carga"

    def get_queryset(self):
        return CargaSAP.objects.select_related("usuario")

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        registros = (
            self.object.registros
            .select_related(
                "operacion",
                "tecnico",
                "producto",
                "unidad_serializada",
            )
            .order_by("fila")
        )

        estado = self.request.GET.get("estado", "").strip()
        tipo_error = self.request.GET.get("tipo_error", "").strip()
        q = self.request.GET.get("q", "").strip()

        if estado in RegistroSAP.Estado.values:
            registros = registros.filter(estado=estado)

        if tipo_error:
            registros = registros.filter(tipo_error=tipo_error)

        if q:
            registros = registros.filter(
                Q(numero_orden__icontains=q)
                | Q(codigo_producto__icontains=q)
                | Q(numero_serie__icontains=q)
                | Q(tecnico__tecnico_apellido_nombres__icontains=q)
                | Q(error__icontains=q)
                | Q(diagnostico__icontains=q)
            )

        ctx["registros"] = registros
        ctx["estado_sel"] = estado
        ctx["tipo_error_sel"] = tipo_error
        ctx["q"] = q
        ctx["tipos_error"] = (
            self.object.registros
            .exclude(tipo_error="")
            .values_list("tipo_error", flat=True)
            .distinct()
            .order_by("tipo_error")
        )

        return ctx

@login_required
def descargar_plantilla_sap(request):
    """
    Genera y descarga una plantilla Excel para la carga SAP.
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "Carga SAP"

    # Columnas que utiliza services_sap.py
    columnas = [
        "numero_orden",
        "codigo_producto",
        "numero_serie",
        "cantidad",
    ]

    ws.append(columnas)

    # Encabezados
    for celda in ws[1]:
        celda.font = Font(
            bold=True,
            color="FFFFFF",
        )
        celda.fill = PatternFill(
            fill_type="solid",
            fgColor="17365D",
        )
        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    # Ejemplos para orientar al usuario
    ws.append([
        "100001",
        "MED001",
        "SERIAL001",
        1,
    ])

    ws.append([
        "100002",
        "MAT001",
        "",
        5,
    ])

    # Anchos de columnas
    anchos = [22, 24, 30, 15]

    for indice, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[
            get_column_letter(indice)
        ].width = ancho

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:D3"

    # Mantener los códigos y seriales como texto
    for fila in ws.iter_rows(min_row=2, max_col=3):
        for celda in fila:
            celda.number_format = "@"

    # Crear archivo en memoria
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    response = HttpResponse(
        buffer.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    response["Content-Disposition"] = (
        'attachment; filename="plantilla_carga_sap.xlsx"'
    )

    return response