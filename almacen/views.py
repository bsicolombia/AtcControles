from collections import OrderedDict
from .models import RemisionIngreso
from django.db.models import Case, When, IntegerField, Sum, F, Q
from django.db.models.functions import Coalesce
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
from usuarios.roles import DIGITADOR, tiene_rol
from django.db.models import Prefetch
from usuarios.mixins import RolRequeridoMixin
from django.contrib.auth import get_user_model
from usuarios.roles import ROLES_ALMACEN, DIGITADOR, COORDINADOR,ROLES_DIGITADOR
User = get_user_model()
from usuarios.roles import ADMINISTRADOR
from urllib.parse import urlencode
from . import services
from operaciones import services_sap
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
    return (
        User.objects
        .filter(rol="TECNICO")
        .select_related("supervisor")
        .order_by("first_name", "last_name", "username")
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
# FUNCIONES AUXILIARES: ASIGNADO, FIVE Y SAP
# ============================================================

def _datos_vacios():
    return {
        "entregado": 0,   # todo lo entregado y no devuelto
        "asignados": 0,   # entregado y no devuelto
        "five": 0,
        "sap": 0,
        "debe_tener": 0,
    }

def _asignado_total_por_producto():
    """Total entregado y no devuelto (mismo criterio que la tabla por técnico)."""
    registros = (
        Asignacion.objects
        .exclude(estado=Asignacion.Estado.DEVUELTA)
        .values("producto_id")
        .annotate(total=Sum(Coalesce(F("cantidad"), 0)))
    )
    return {r["producto_id"]: r["total"] or 0 for r in registros}

def _five_total_por_tecnico_producto():
    """
    Total instalado según FIVE, agrupado por técnico y producto.

    - Asignación INSTALADA: cuenta cantidad completa.
    - Otras asignaciones no devueltas: cuenta cantidad_instalada.
    """
    registros = (
        Asignacion.objects
        .exclude(estado=Asignacion.Estado.DEVUELTA)
        .values("tecnico_id", "producto_id")
        .annotate(
            total=Sum(
                Case(
                    When(
                        estado=Asignacion.Estado.INSTALADA,
                        then=Coalesce(F("cantidad"), 0),
                    ),
                    default=Coalesce(F("cantidad_instalada"), 0),
                    output_field=IntegerField(),
                )
            )
        )
    )

    return {
        (r["tecnico_id"], r["producto_id"]): r["total"] or 0
        for r in registros
    }
def _five_total_por_producto():
    """
    Total instalado según FIVE, agrupado por producto.
    - Asignación INSTALADA -> cuenta su cantidad completa
    - Otras (no devueltas) -> cuenta cantidad_instalada
    """
    registros = (
        Asignacion.objects
        .exclude(estado=Asignacion.Estado.DEVUELTA)
        .values("producto_id")
        .annotate(
            total=Sum(
                Case(
                    When(
                        estado=Asignacion.Estado.INSTALADA,
                        then=Coalesce(F("cantidad"), 0),
                    ),
                    default=Coalesce(F("cantidad_instalada"), 0),
                    output_field=IntegerField(),
                )
            )
        )
    )

    return {r["producto_id"]: r["total"] or 0 for r in registros}


def _sap_total_por_producto():
    """Total reportado por SAP en la última carga, agrupado por producto."""
    ultima_carga = CargaSAP.objects.order_by("-fecha", "-pk").first()

    if not ultima_carga:
        return {}

    registros = (
        RegistroSAP.objects
        .filter(
            carga=ultima_carga,
            producto__isnull=False,
            tecnico__isnull=False,   # <-- agregar
        )
        .values("producto_id")
        .annotate(total=Sum("cantidad"))
    )

    return {r["producto_id"]: r["total"] or 0 for r in registros}


def _obtener_datos_sap():
    """
    Total reportado por SAP en la última carga, agrupado por
    (técnico, producto). Incluye filas OK y ERROR si tienen
    técnico y producto identificados.
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
# ============================================================
# DASHBOARD
# ============================================================

class DashboardView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
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
    roles_permitidos = ROLES_ALMACEN
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
    roles_permitidos = ROLES_ALMACEN
    titulo = "Nuevo material"
    boton = "Guardar material"


class CategoriaUpdateView(CategoriaFormMixin, UpdateView):
    roles_permitidos = ROLES_ALMACEN
    titulo = "Editar categoría"
    boton = "Guardar cambios"


# ============================================================
# PRODUCTOS
# ============================================================

class ProductoListView(RolRequeridoMixin, BuscarMixin, ListView):
    roles_permitidos = ROLES_ALMACEN
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
    roles_permitidos = ROLES_ALMACEN
    titulo = "Nuevo material"
    boton = "Guardar material"
    subtitulo = (
        "Un material por cada tipo y medida: "
        "«Codo 1/2» y «Codo 1/4» son material distintos."
    )


class ProductoUpdateView(ProductoFormMixin, UpdateView):
    roles_permitidos = ROLES_ALMACEN
    titulo = "Editar material"
    boton = "Guardar cambios"


# ============================================================
# INVENTARIO
# ============================================================

class InventarioListView(RolRequeridoMixin, BuscarMixin, ListView):
    roles_permitidos = ROLES_ALMACEN

    model = Inventario
    template_name = "almacen/inventario_list.html"
    context_object_name = "inventarios"

    campos_busqueda = (
        "producto__nombre",
        "producto__codigo_producto",
        "producto__medida",
        "producto__categoria__nombre",
    )

    def dispatch(self, request, *args, **kwargs):
        usuario = request.user

        if (
            usuario.is_authenticated
            and not usuario.is_superuser
            and not tiene_rol(usuario, ADMINISTRADOR)
            and tiene_rol(usuario, DIGITADOR)
        ):
            return redirect("almacen:asignaciones")

        return super().dispatch(request, *args, **kwargs)

    def get_queryset(self):
        qs = (
            super()
            .get_queryset()
            .select_related("producto", "producto__categoria")
        )

        if self.request.GET.get("bajo"):
            qs = qs.filter(
                cantidad_minima__gt=0,
                cantidad_actual__lte=F("cantidad_minima"),
            )

        comparacion = self.request.GET.get("comparacion", "").strip()

        if comparacion in ("coinciden", "diferentes"):
            five = _five_total_por_producto()
            sap = _sap_total_por_producto()

            inventarios = list(qs)

            if comparacion == "coinciden":
                inventarios = [
                    inv for inv in inventarios
                    if five.get(inv.producto_id, 0)
                    == sap.get(inv.producto_id, 0)
                ]
            else:
                inventarios = [
                    inv for inv in inventarios
                    if five.get(inv.producto_id, 0)
                    != sap.get(inv.producto_id, 0)
                ]

            return inventarios

        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["comparacion_sel"] = self.request.GET.get(
            "comparacion", ""
        ).strip()

        tecnico_sel = self.request.GET.get("tecnico", "").strip()

        ctx["tecnicos_lista"] = tecnicos_queryset()
        ctx["tecnico_sel"] = tecnico_sel

        if tecnico_sel.isdigit():
            ctx.update(_obtener_contexto_asignaciones(self.request))
        else:
            five = _five_total_por_producto()
            sap = _sap_total_por_producto()
            asignado = _asignado_total_por_producto()

            for inventario in ctx["inventarios"]:
                producto_id = inventario.producto_id

                inventario.five = five.get(producto_id, 0)
                inventario.sap = sap.get(producto_id, 0)
                inventario.asignado_estado = asignado.get(
                    producto_id, 0
                )

        return ctx


class InventarioAlmacenUpdateView(RolRequeridoMixin, UpdateView):
    roles_permitidos = ROLES_ALMACEN
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
    roles_permitidos = ROLES_ALMACEN
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
    roles_permitidos = ROLES_ALMACEN
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
    roles_permitidos = ROLES_ALMACEN
    model = UnidadSerializada
    template_name = "almacen/unidad_list.html"
   
    campos_busqueda = (
        # Unidad serializada
        "codigo",
        "numero_serie",
        "marca",
        "referencia",
        "medida",
        "estado",

        # Producto: usar campos reales del modelo
        "producto__nombre",
        "producto__codigo_producto",
        "producto__marca",
        "producto__referencia",
        "producto__medida",
        "producto__color",
        "producto__presentacion",
        "producto__categoria__nombre",

        # Remisión
        "remision__numero_remision",

        # Técnico
        "asignaciones__tecnico__first_name",
        "asignaciones__tecnico__last_name",
        "asignaciones__tecnico__username",
        "asignaciones__tecnico__cedula",

        # Operación asociada a la asignación
        "asignaciones__operacion__numero_orden",
        "asignaciones__operacion__numero_cuenta_contrato",
    )
    
    
    def aplicar_busqueda_general(self, qs):
        termino = self.request.GET.get("q", "").strip()

        if not termino:
            return qs

        filtros = Q()

        for campo in self.campos_busqueda:
            filtros |= Q(**{f"{campo}__icontains": termino})

        return qs.filter(filtros).distinct()



    def get_queryset(self):
        tecnico_id = self.request.GET.get("tecnico", "").strip()
        estado = self.request.GET.get("estado", "").strip()
        producto = self.request.GET.get("producto", "").strip()
        comparacion = self.request.GET.get("comparacion", "").strip()

        asignaciones_qs = (
            Asignacion.objects
            .select_related("tecnico", "operacion", "producto")
            .order_by("-fecha_modificacion", "-id")
        )

        qs = (
            super()
            .get_queryset()
            .select_related("producto", "remision")
        )

        if tecnico_id.isdigit():
            tecnico_pk = int(tecnico_id)

            qs = qs.filter(
                (
                    Q(
                        estado=UnidadSerializada.Estado.ASIGNADO,
                        asignaciones__tecnico_id=tecnico_pk,
                        asignaciones__estado=Asignacion.Estado.ASIGNADA,
                    )
                    | Q(
                        estado=UnidadSerializada.Estado.INSTALADO,
                        asignaciones__tecnico_id=tecnico_pk,
                        asignaciones__estado=Asignacion.Estado.INSTALADA,
                    )
                )
            ).distinct()

            asignaciones_qs = asignaciones_qs.filter(
                tecnico_id=tecnico_pk
            )

        if estado:
            qs = qs.filter(estado=estado)

        if producto.isdigit():
            qs = qs.filter(producto_id=int(producto))

        qs = qs.prefetch_related(
            Prefetch(
                "asignaciones",
                queryset=asignaciones_qs,
                to_attr="asignaciones_lista",
            )
        )

        # Comparación por técnico y producto.
        if comparacion in ("coinciden", "diferentes"):
            five = _five_total_por_tecnico_producto()
            sap = _obtener_datos_sap()

            # Se obtienen las unidades que coinciden con los
            # estados y asignaciones que corresponden a cada técnico.
            unidades = list(qs)

            ids_validos = []

            for unidad in unidades:
                if unidad.estado == UnidadSerializada.Estado.ASIGNADO:
                    estado_asignacion = Asignacion.Estado.ASIGNADA
                elif unidad.estado == UnidadSerializada.Estado.INSTALADO:
                    estado_asignacion = Asignacion.Estado.INSTALADA
                else:
                    continue

                asignacion = next(
                    (
                        a for a in getattr(
                            unidad, "asignaciones_lista", []
                        )
                        if a.estado == estado_asignacion
                    ),
                    None,
                )

                if asignacion is None:
                    continue

                clave = (asignacion.tecnico_id, unidad.producto_id)

                cantidad_five = five.get(clave, 0)
                cantidad_sap = sap.get(clave, 0)

                coincide = cantidad_five == cantidad_sap

                if (
                    (comparacion == "coinciden" and coincide)
                    or
                    (comparacion == "diferentes" and not coincide)
                ):
                    ids_validos.append(unidad.pk)

            qs = qs.filter(pk__in=ids_validos)

        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        ctx["total_unidades_filtradas"] = self.object_list.count()

        ctx["estados"] = UnidadSerializada.Estado.choices
        ctx["estado_sel"] = self.request.GET.get("estado", "")
        ctx["producto_sel"] = self.request.GET.get("producto", "")
        ctx["tecnico_sel"] = self.request.GET.get("tecnico", "")

        ctx["tecnicos_lista"] = tecnicos_queryset()

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

        # Resolver la asignación relevante de cada unidad.
        for unidad in ctx["object_list"]:
            unidad.asignacion_relevante = None

            if unidad.estado == UnidadSerializada.Estado.ASIGNADO:
                estado_asignacion = Asignacion.Estado.ASIGNADA
            elif unidad.estado == UnidadSerializada.Estado.INSTALADO:
                estado_asignacion = Asignacion.Estado.INSTALADA
            else:
                estado_asignacion = None

            if estado_asignacion:
                unidad.asignacion_relevante = next(
                    (
                        a for a in getattr(
                            unidad, "asignaciones_lista", []
                        )
                        if a.estado == estado_asignacion
                    ),
                    None,
                )

        # Valores de FIVE y SAP agrupados por técnico y producto.
        five = _five_total_por_tecnico_producto()
        sap = _obtener_datos_sap()

        for unidad in ctx["object_list"]:
            asignacion = unidad.asignacion_relevante

            if asignacion:
                clave = (
                    asignacion.tecnico_id,
                    unidad.producto_id,
                )

                unidad.five = five.get(clave, 0)
                unidad.sap = sap.get(clave, 0)
            else:
                unidad.five = 0
                unidad.sap = 0

            unidad.coincide_sap_five = (
                unidad.five == unidad.sap
            )

        ctx["comparacion_sel"] = self.request.GET.get(
            "comparacion", ""
        ).strip()

        return ctx



class UnidadDetailView(RolRequeridoMixin, DetailView):
    roles_permitidos = ROLES_ALMACEN
    model = UnidadSerializada
    template_name = "almacen/unidad_detail.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["asignaciones"] = self.object.asignaciones.select_related("tecnico")
        ctx["movimientos"] = self.object.movimientos.select_related("usuario")
        return ctx


class UnidadEstadoView(FormPageMixin, RolRequeridoMixin, FormView):
    roles_permitidos = ROLES_ALMACEN
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
    roles_permitidos = ROLES_ALMACEN
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
# CONSOLIDADO DE ASIGNACIONES, FIVE Y SAP
# ============================================================

def _obtener_contexto_asignaciones(request):
    """
    Construye el consolidado por técnico y producto.

    ASIGNADOS:
        Cantidad que se ha entregado al técnico y no se ha devuelto.

    FIVE:
        Cantidad instalada registrada en el sistema.

    SAP:
        Cantidad reportada en la última carga SAP.

    DEBE_TENER:
        Cantidad asignada menos la cantidad reportada por SAP,
        sin permitir resultados negativos.
    """

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

    # --------------------------------------------------------
    # FILTRO POR TÉCNICO
    # --------------------------------------------------------

    if tecnico_sel.isdigit():
        qs = qs.filter(
            tecnico_id=int(tecnico_sel)
        )

    # --------------------------------------------------------
    # FILTRO POR ESTADO DE ASIGNACIÓN
    # --------------------------------------------------------

    if estado_sel:
        qs = qs.filter(
            estado=estado_sel
        )

    # --------------------------------------------------------
    # BÚSQUEDA GENERAL
    # --------------------------------------------------------

    if q:
        filtros = (
            Q(tecnico__cedula__icontains=q)
            | Q(tecnico__first_name__icontains=q)
            | Q(tecnico__last_name__icontains=q)
            | Q(tecnico__username__icontains=q)

            # Supervisor
            | Q(tecnico__supervisor__first_name__icontains=q)
            | Q(tecnico__supervisor__last_name__icontains=q)
            | Q(tecnico__supervisor__username__icontains=q)

            # Producto
            | Q(producto__nombre__icontains=q)
            | Q(producto__codigo_producto__icontains=q)
            | Q(producto__marca__icontains=q)
            | Q(producto__referencia__icontains=q)
            | Q(producto__medida__icontains=q)
            | Q(producto__color__icontains=q)
            | Q(producto__presentacion__icontains=q)

            # Operación
            | Q(operacion__numero_orden__icontains=q)
            | Q(operacion__numero_cuenta_contrato__icontains=q)

            # Unidad serializada
            | Q(unidad_serializada__codigo__icontains=q)
            | Q(unidad_serializada__numero_serie__icontains=q)
        )

        unidad = (
            UnidadSerializada.objects
            .filter(
                Q(codigo__iexact=q)
                | Q(numero_serie__iexact=q)
            )
            .first()
        )

        if unidad:
            qs = qs.filter(
                unidad_serializada=unidad,
                estado__in=[
                    Asignacion.Estado.ASIGNADA,
                    Asignacion.Estado.INSTALADA,
                ],
            )
        else:
            qs = qs.filter(filtros)

    # --------------------------------------------------------
    # AGRUPAR ASIGNACIONES POR TÉCNICO Y PRODUCTO
    # --------------------------------------------------------

    productos_dict = OrderedDict()
    tecnicos_dict = OrderedDict()

    datos_sap = _obtener_datos_sap()

    for asignacion in qs:
        producto = asignacion.producto
        tecnico = asignacion.tecnico

        productos_dict.setdefault(
            producto.pk,
            producto,
        )

        bloque = tecnicos_dict.setdefault(
            tecnico.pk,
            {
                "tecnico": tecnico,
                "productos": {},
            },
        )

        datos = bloque["productos"].setdefault(
            producto.pk,
            _datos_vacios(),
        )

        # Las asignaciones devueltas no cuentan
        # como materiales en poder del técnico.
        if asignacion.estado == Asignacion.Estado.DEVUELTA:
            continue

        # ----------------------------------------------------
        # ASIGNADOS
        # ----------------------------------------------------

        datos["asignados"] += asignacion.cantidad or 0

        # ----------------------------------------------------
        # FIVE: CANTIDAD INSTALADA
        # ----------------------------------------------------

        if asignacion.estado == Asignacion.Estado.INSTALADA:
            datos["five"] += asignacion.cantidad or 0
        else:
            datos["five"] += asignacion.cantidad_instalada or 0

    productos = list(productos_dict.values())
    tecnicos = []

    base_unidades = reverse("almacen:unidades")
    base_inventario = reverse("dashboard")

    # --------------------------------------------------------
    # CONSTRUIR LAS FILAS DEL CONSOLIDADO
    # --------------------------------------------------------

    for bloque in tecnicos_dict.values():

        filas = {
            "ASIGNADOS": [],
            "FIVE": [],
            "SAP": [],
            "DEBE_TENER": [],
        }

        productos_detalle = []

        for producto in productos:

            datos = bloque["productos"].get(
                producto.pk,
                _datos_vacios(),
            )

            asignados = datos["asignados"]
            five = datos["five"]

            # Cantidad reportada por SAP para este técnico
            # y este producto.
            sap = datos_sap.get(
                (
                    bloque["tecnico"].pk,
                    producto.pk,
                ),
                0,
            )

            debe_tener = max(
                asignados - sap,
                0,
            )

            coincide = sap == five

            # ------------------------------------------------
            # ENLACE PARA CADA CELDA
            # ------------------------------------------------

            def crear_celda(cantidad, tipo):
                if producto.es_serializado:

                    estado = (
                        UnidadSerializada.Estado.INSTALADO
                        if tipo == "FIVE"
                        else UnidadSerializada.Estado.ASIGNADO
                    )

                    parametros = urlencode({
                        "tecnico": bloque["tecnico"].pk,
                        "producto": producto.pk,
                        "estado": estado,
                    })

                    url = f"{base_unidades}?{parametros}"

                else:
                    parametros = urlencode({
                        "tecnico": bloque["tecnico"].pk,
                        "q": (
                            producto.codigo_producto
                            or producto.nombre
                        ),
                    })

                    url = f"{base_inventario}?{parametros}"

                return {
                    "cantidad": cantidad,
                    "url": url,
                }

            # ASIGNADOS
            filas["ASIGNADOS"].append(
                crear_celda(
                    asignados,
                    "ASIGNADOS",
                )
            )

            # FIVE
            filas["FIVE"].append(
                crear_celda(
                    five,
                    "FIVE",
                )
            )

            # SAP
            celda_sap = crear_celda(
                sap,
                "SAP",
            )

            celda_sap["coincide"] = coincide

            celda_sap["estado_conciliacion"] = (
                "ok" if coincide else "error"
            )

            filas["SAP"].append(celda_sap)

            # DEBE TENER
            filas["DEBE_TENER"].append(
                crear_celda(
                    debe_tener,
                    "DEBE_TENER",
                )
            )

            # DETALLE DEL PRODUCTO
            productos_detalle.append({
                "producto": producto,
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
    roles_permitidos = ROLES_DIGITADOR
    template_name = "almacen/asignacion_list.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        ctx.update(_obtener_contexto_asignaciones(self.request))

        ctx["tecnicos_lista"] = tecnicos_queryset()
        return ctx

# ============================================================
# ASIGNACIONES · CREAR
# ============================================================

class AsignacionCreateView(FormPageMixin, RolRequeridoMixin, FormView):
    roles_permitidos = ROLES_ALMACEN
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
    roles_permitidos = ROLES_DIGITADOR
    """
    Un técnico, un formulario: se escribe la operación una sola vez
    y las cantidades o unidades instaladas de cada material.
    """

    template_name = "almacen/asignacion_instalar.html"

    def _cargar(self):
        self.tecnico = get_object_or_404(
            User.objects.select_related("supervisor").filter(
                rol="TECNICO"
            ),
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
    roles_permitidos = ROLES_DIGITADOR
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
    roles_permitidos = ROLES_ALMACEN

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
    roles_permitidos = ROLES_ALMACEN

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
    roles_permitidos = ROLES_ALMACEN
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
    
