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
from django.db.models import F, Q, Exists, OuterRef
from django.views import View
from django.views.generic import (
    CreateView,
    DetailView,
    FormView,
    ListView,
    TemplateView,
    UpdateView,
)

from .models import (Asignacion, Categoria, Inconsistencia, Inventario,
                     MovimientoInventario, Producto, UnidadSerializada)
from django.db.models import Count
from usuarios.roles import DIGITADOR, tiene_rol
from django.db.models import Prefetch
from usuarios.mixins import RolRequeridoMixin
from django.contrib.auth import get_user_model
from usuarios.roles import ROLES_ALMACEN, DIGITADOR, COORDINADOR,ROLES_DIGITADOR
User = get_user_model()
from usuarios.roles import ADMINISTRADOR
from urllib.parse import urlencode
from . import services
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
    
import json
from django.contrib.staticfiles import finders
from django.http import HttpResponse
from django.views.generic import DetailView
from .forms import EntregaForm
from .models import EntregaMaterial, EntregaItem

from .models import (
    Asignacion,
    Categoria,
    Inventario,
    MovimientoInventario,
    Producto,
    UnidadSerializada,
)
from operaciones.models import CargaSAP, Operacion
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
    
def _unidades_conciliadas_sap():
    """IDs de unidades que SAP ya reportó y que coinciden con Five (en cualquier carga)."""
    return set(
        RegistroSAP.objects
        .filter(estado=RegistroSAP.Estado.OK, unidad_serializada__isnull=False)
        .values_list("unidad_serializada_id", flat=True)
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
    Totales SAP por (técnico, producto).
    Conserva las conciliaciones OK de cargas anteriores,
    aunque la última carga haya ignorado registros repetidos.
    """
    from django.db.models import Q, Sum
    from operaciones.models import RegistroSAP

    E = RegistroSAP.Estado

    registros = (
        RegistroSAP.objects
        .filter(
            Q(estado=E.OK)
            | Q(estado=E.ERROR, vigente=True)
        )
        .filter(
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
        
        ctx["inconsistencias_pendientes"] = Inconsistencia.objects.filter(
            estado=Inconsistencia.Estado.PENDIENTE).count()

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

from django.db.models import F, Q, Exists, OuterRef


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

        q = self.request.GET.get("q", "").strip()

        if q:
            movimientos_remision = MovimientoInventario.objects.filter(
                producto_id=OuterRef("producto_id"),
                remision__numero_remision__icontains=q,
                tipo_movimiento=MovimientoInventario.TipoMovimiento.ENTRADA,
            )

            unidades_remision = UnidadSerializada.objects.filter(
                producto_id=OuterRef("producto_id"),
                remision__numero_remision__icontains=q,
            )

            qs = qs.annotate(
                tiene_remision_movimiento=Exists(
                    movimientos_remision
                ),
                tiene_remision_unidad=Exists(
                    unidades_remision
                ),
            ).filter(
                Q(producto__nombre__icontains=q)
                | Q(producto__codigo_producto__icontains=q)
                | Q(producto__medida__icontains=q)
                | Q(producto__categoria__nombre__icontains=q)
                | Q(tiene_remision_movimiento=True)
                | Q(tiene_remision_unidad=True)
            ).distinct()

        if self.request.GET.get("bajo"):
            qs = qs.filter(
                cantidad_minima__gt=0,
                cantidad_actual__lte=F("cantidad_minima"),
            )

        comparacion = self.request.GET.get(
            "comparacion", ""
        ).strip()

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

        tecnico_sel = self.request.GET.get(
            "tecnico", ""
        ).strip()

        ctx["tecnicos_lista"] = tecnicos_queryset()
        ctx["tecnico_sel"] = tecnico_sel

        if tecnico_sel.isdigit():
            ctx.update(
                _obtener_contexto_asignaciones(self.request)
            )
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

        try:
            with transaction.atomic():
                remision = services.obtener_o_crear_remision(
                    numero=num_orden,
                    usuario=self.request.user,
                    imagen=d.get("comprobante_imagen"),
                    fecha_recibo_bsi=d.get("fecha_recibo_bsi"),
                    recogido_por=d.get("recogido_por", ""),
                    recibido_almacen_por=d.get("recibido_almacen_por", ""),
                )

                unidades = services.registrar_ingreso(
                    producto=producto,
                    cantidad=d["cantidad"],
                    usuario=self.request.user,
                    fecha_ingreso=timezone.localdate(),
                    remision=remision,
                    seriales=d.get("seriales", []),
                    observacion=d["observacion"],
                )
        except services.ErrorInventario as e:
            form.add_error(None, str(e))
            return self.form_invalid(form)

        if unidades:
            mensaje = (
                f"Ingresaron {len(unidades)} unidad(es) de "
                f"{producto.nombre_completo} con la remisión {num_orden}."
            )
            if not producto.requiere_numero_serie:
                mensaje += " Los códigos internos fueron generados automáticamente."

            self.success_url = reverse("almacen:unidades")
        else:
            mensaje = (
                f"Ingresaron {d['cantidad']} {producto.nombre_completo} "
                f"con la remisión {num_orden}."
            )
            self.success_url = reverse("almacen:inventario")

        messages.success(self.request, mensaje)
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
            super().get_queryset()
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

        if comparacion in ("coinciden", "diferentes", "pendientes"):
            sap_ok = _unidades_conciliadas_sap()
            ids_validos = []

            for pk, est in qs.values_list("pk", "estado"):
                five = 1 if est == UnidadSerializada.Estado.INSTALADO else 0
                sap = 1 if pk in sap_ok else 0

                if five and sap:
                    resultado = "coinciden"
                elif five and not sap:
                    resultado = "pendientes"      # instalado en Five, SAP aún no lo reporta
                else:
                    resultado = "diferentes"      # SAP dice algo que Five no respalda

                if resultado == comparacion:
                    ids_validos.append(pk)

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

        # FIVE y SAP por unidad (0 o 1)
        sap_ok = _unidades_conciliadas_sap()

        for unidad in ctx["object_list"]:
            unidad.five = 1 if unidad.estado == UnidadSerializada.Estado.INSTALADO else 0
            unidad.sap = 1 if unidad.pk in sap_ok else 0

            if unidad.five and unidad.sap:
                unidad.estado_sap = "ok"
            elif unidad.five or unidad.sap:
                unidad.estado_sap = "error"
            else:
                unidad.estado_sap = "neutro"      # disponible, asignado, dañado...

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
# ASIGNACIONES · EDITAR
# ============================================================

class AsignacionUpdateView(
    FormPageMixin,
    RolRequeridoMixin,
    UpdateView
):
    roles_permitidos = ROLES_ALMACEN

    model = Asignacion
    form_class = AsignacionForm
    template_name = "almacen/asignacion_form.html"

    titulo = "Editar asignación"
    boton = "Guardar cambios"

    def get_queryset(self):
        return Asignacion.objects.select_related(
            "tecnico",
            "producto",
            "unidad_serializada",
        )

    def get_success_url(self):
        return reverse_lazy("almacen:asignaciones")

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        ctx["subtitulo"] = (
            f"Modificar asignación de {self.object.tecnico}"
        )

        ctx["cancel_url"] = self.get_success_url()

        return ctx

    def form_valid(self, form):
        messages.success(
            self.request,
            "Asignación actualizada correctamente."
        )
        return super().form_valid(form)
# ============================================================
# INSTALAR / DEVOLVER · PANTALLA DEL TÉCNICO (varios materiales)
# ============================================================


import json  # arriba del archivo

class TecnicoAsignacionesView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_DIGITADOR
    template_name = "almacen/asignacion_instalar.html"

    def _cargar(self):
        self.tecnico = get_object_or_404(
            User.objects.select_related("supervisor").filter(rol="TECNICO"),
            pk=self.kwargs["pk"],
        )
        pendientes = (
            Asignacion.objects
            .filter(tecnico=self.tecnico, estado=Asignacion.Estado.ASIGNADA)
            .select_related("producto", "unidad_serializada", "operacion")
            .order_by("producto__nombre", "producto__medida", "id")
        )
        self.pendientes = [a for a in pendientes if a.cantidad_pendiente > 0]

    def _datos_pendientes(self):
        """Todos los materiales por cantidad (con lo pendiente) y los seriales asignados."""
        pend = {}
        seriales = []
        for a in self.pendientes:
            if a.unidad_serializada_id:
                u = a.unidad_serializada
                seriales.append({
                    "id": u.pk,
                    "serial": u.numero_serie or u.codigo,
                    "codigo": u.codigo,
                    "producto": a.producto_id,
                    "nombre": a.producto.nombre_completo,
                })
            else:
                pend[a.producto_id] = pend.get(a.producto_id, 0) + a.cantidad_pendiente

        cantidad = [
            {"id": p.pk, "nombre": p.nombre_completo, "pendiente": pend.get(p.pk, 0)}
            for p in Producto.objects.filter(
                estado=True, tipo_inventario=Producto.TipoInventario.CANTIDAD
            ).order_by("nombre", "medida")
        ]
        return cantidad, seriales

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        self._cargar()
        pend_cantidad, pend_seriales = self._datos_pendientes()

        try:
            payload_previo = json.loads(self.request.POST.get("payload", "[]"))
            if not isinstance(payload_previo, list):
                payload_previo = []
        except ValueError:
            payload_previo = []

        ctx["tecnico"] = self.tecnico
        ctx["pend_cantidad"] = pend_cantidad
        ctx["pend_seriales"] = pend_seriales
        ctx["productos_serializados"] = Producto.objects.filter(
            estado=True, tipo_inventario=Producto.TipoInventario.SERIALIZADO
        ).order_by("nombre", "medida")
        ctx["tipos_inconsistencia"] = dict(Inconsistencia.Tipo.choices)
        ctx["payload_previo"] = payload_previo
        ctx["estados_instalan"] = sorted(e.value for e in Operacion.ESTADOS_QUE_INSTALAN)
        ctx["form"] = kwargs.get("form") or InstalacionMasivaForm()
        return ctx

    def post(self, request, *args, **kwargs):
        self._cargar()
        form = InstalacionMasivaForm(request.POST, request.FILES)

        try:
            lineas = json.loads(request.POST.get("payload", "[]"))
            if not isinstance(lineas, list):
                lineas = []
        except ValueError:
            lineas = []

        errores, instala = [], False
        form_ok = form.is_valid()
        if form_ok:
            instala = form.cleaned_data["estado"] in Operacion.ESTADOS_QUE_INSTALAN
            if instala and not lineas:
                errores.append("Agregue al menos un material o serial para instalar.")
            elif not instala and lineas:
                errores.append("Con este estado no se instala material: deje la lista vacía.")

        if errores or not form_ok:
            for e in errores:
                messages.error(request, e)
            return self.render_to_response(self.get_context_data(form=form))

        d = form.cleaned_data
        novedades, piezas = [], 0
        try:
            with transaction.atomic():
                op = services.obtener_o_crear_operacion(
                    numero_orden=d["numero_orden"],
                    cuenta_contrato=d["numero_cuenta_contrato"],
                    tipo=d["tipo_operacion"],
                    tecnico=self.tecnico,
                    usuario=request.user,
                    fecha=d["fecha_instalacion"],
                    estado=d["estado"],
                    anomalia=d["anomalia"],
                    observaciones=d["observaciones"],
                    archivo=d.get("archivo_soporte"),
                )

                if instala:
                    vistos = set()
                    for l in lineas:
                        if l.get("serializado"):
                            for u in l.get("unidades", []):
                                serial = str(u.get("serial", "")).strip()
                                if not serial or serial.lower() in vistos:
                                    continue
                                vistos.add(serial.lower())
                                piezas += 1
                                n = services.instalar_serial(
                                    serial=serial, producto_id=l.get("producto"),
                                    tecnico=self.tecnico, usuario=request.user,
                                    operacion=op, fecha=d["fecha_instalacion"],
                                    observaciones=d["observaciones"],
                                )
                                if n:
                                    novedades.append(n)
                        else:
                            try:
                                pid, cant = int(l.get("producto")), int(l.get("cantidad", 0))
                            except (TypeError, ValueError):
                                raise services.ErrorInventario("Hay un material inválido en la lista.")
                            producto = Producto.objects.filter(pk=pid, estado=True).first()
                            if not producto or producto.es_serializado:
                                raise services.ErrorInventario("Hay un material inválido en la lista.")
                            if cant <= 0:
                                continue
                            piezas += cant
                            novedades += services.instalar_cantidad(
                                producto=producto, cantidad=cant, tecnico=self.tecnico,
                                usuario=request.user, operacion=op,
                                fecha=d["fecha_instalacion"],
                                observaciones=d["observaciones"],
                            )
        except services.ErrorInventario as e:
            messages.error(request, str(e))
            return self.render_to_response(self.get_context_data(form=form))

        if instala:
            messages.success(
                request,
                f"Orden {op.numero_orden} registrada ({piezas} pieza(s) reportadas).",
            )
        else:
            messages.success(
                request,
                f"Orden {op.numero_orden} registrada como {op.get_estado_display().lower()}.",
            )
        if novedades:
            messages.warning(
                request,
                f"⚠ Se registraron {len(novedades)} inconsistencia(s) para revisar.",
            )
        return redirect("almacen:tecnico_asignaciones", pk=self.tecnico.pk)
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
                    tecnico=self.tecnico,
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

import json

from django.shortcuts import get_object_or_404, render, redirect
from django.contrib import messages
from django.contrib.auth.decorators import login_required

from operaciones.models import Operacion
from almacen.models import Asignacion


@login_required
def editar_instalacion(request, pk):

    operacion = get_object_or_404(
        Operacion,
        pk=pk,
        tipo_operacion=Operacion.TipoOperacion.INSTALACION,
    )

    asignaciones = (
        Asignacion.objects
        .filter(operacion=operacion)
        .select_related("producto", "unidad_serializada")
    )

    payload_previo = []

    for a in asignaciones:

        if a.unidad_serializada_id:
            unidad = a.unidad_serializada

            payload_previo.append({
                "producto": a.producto_id,
                "nombre": a.producto.nombre_completo,
                "serializado": True,
                "cantidad": 0,
                "unidades": [{
                    "id": unidad.id,
                    "serial": unidad.numero_serie or unidad.codigo,
                }],
            })

        else:
            payload_previo.append({
                "producto": a.producto_id,
                "nombre": a.producto.nombre_completo,
                "serializado": False,
                "cantidad": a.cantidad_instalada,
                "unidades": [],
            })

    # Aquí debes reutilizar la lógica que ya tienes
    # para construir pend_cantidad, pend_seriales,
    # estados_instalan y tipos_inconsistencia.

    if request.method == "POST":

        # Aquí se debe procesar el payload y actualizar
        # las asignaciones y movimientos existentes.
        pass

    context = {
        "operacion": operacion,
        "tecnico": operacion.tecnico,
        "payload_previo": payload_previo,
        "modo_edicion": True,
    }

    return render(
        request,
        "operaciones/registrar_operacion.html",
        context,
    )
# ============================================================
# DEVOLVER ASIGNACIÓN
# ============================================================

class TecnicoDevolverView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "almacen/tecnico_devolver.html"

    def _cargar(self):
        self.tecnico = get_object_or_404(User, pk=self.kwargs["pk"], rol="TECNICO")
        pend = (
            Asignacion.objects
            .filter(tecnico=self.tecnico, estado=Asignacion.Estado.ASIGNADA)
            .select_related("producto", "unidad_serializada")
            .order_by("producto__nombre", "producto__medida", "id")
        )
        self.grupos = OrderedDict()   # por cantidad
        self.seriales = []            # por serial
        for a in pend:
            if a.cantidad_pendiente <= 0:
                continue
            if a.unidad_serializada_id:
                self.seriales.append(a)
            else:
                g = self.grupos.setdefault(
                    a.producto_id, {"producto": a.producto, "pendiente": 0})
                g["pendiente"] += a.cantidad

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        self._cargar()
        ctx.update(tecnico=self.tecnico, grupos=list(self.grupos.values()),
                   seriales=self.seriales)
        return ctx

    def post(self, request, *args, **kwargs):
        self._cargar()
        obs = request.POST.get("observacion", "").strip()
        errores, por_serial, por_cantidad = [], [], []

        for a in self.seriales:
            if request.POST.get(f"ser_{a.pk}"):
                por_serial.append(a)

        for pid, g in self.grupos.items():
            raw = request.POST.get(f"cant_{pid}", "").strip()
            if not raw:
                continue
            try:
                n = int(raw)
            except ValueError:
                errores.append(f"{g['producto'].nombre_completo}: cantidad inválida.")
                continue
            if n <= 0:
                continue
            if n > g["pendiente"]:
                errores.append(
                    f"{g['producto'].nombre_completo}: solo tiene {g['pendiente']}.")
                continue
            por_cantidad.append((g["producto"], n))

        if not errores and not por_serial and not por_cantidad:
            errores.append("Seleccione al menos un material para devolver.")

        if errores:
            for e in errores:
                messages.error(request, e)
            return self.render_to_response(self.get_context_data())

        try:
            with transaction.atomic():
                for a in por_serial:
                    services.devolver(asignacion=a, usuario=request.user, observacion=obs)
                for producto, n in por_cantidad:
                    services.devolver_cantidad(
                        tecnico=self.tecnico, producto=producto, cantidad=n,
                        usuario=request.user, observacion=obs)
        except services.ErrorInventario as e:
            messages.error(request, str(e))
            return self.render_to_response(self.get_context_data())

        messages.success(
            request,
            f"Devolución registrada: {len(por_serial) + sum(n for _, n in por_cantidad)} "
            "pieza(s) volvieron al almacén.")
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

# ---------- helpers ----------
def _nombre(u):
    return (u.get_full_name() or u.username) if u else ""

def _cedula(u):
    return getattr(u, "cedula", "") or ""


# ---------- validar serial (JSON) ----------
class ValidarSerialJson(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN

    def get(self, request):
        producto = request.GET.get("producto", "").strip()
        serial = request.GET.get("serial", "").strip()

        if not serial:
            return JsonResponse({"ok": False, "mensaje": "Escriba el serial."})

        u = (
            UnidadSerializada.objects
            .select_related("producto")
            .filter(Q(numero_serie__iexact=serial) | Q(codigo__iexact=serial))
            .first()
        )
        if not u:
            return JsonResponse({"ok": False, "mensaje": "El serial no existe en el sistema."})

        if producto.isdigit() and u.producto_id != int(producto):
            return JsonResponse({
                "ok": False,
                "mensaje": f"Ese serial pertenece a otro material: {u.producto.nombre_completo}.",
            })

        if u.estado != UnidadSerializada.Estado.DISPONIBLE:
            return JsonResponse({
                "ok": False,
                "mensaje": f"No está disponible (estado: {u.get_estado_display()}).",
            })

        return JsonResponse({
            "ok": True,
            "id": u.pk,
            "serial": u.numero_serie or u.codigo,
            "mensaje": "Serial correcto y disponible.",
        })


# ---------- crear entrega ----------
class EntregaCreateView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "almacen/entrega_form.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["form"] = kwargs.get("form") or EntregaForm()
        ctx["productos"] = Producto.objects.filter(estado=True).order_by("nombre", "medida")
        ctx["payload_previo"] = kwargs.get("payload", [])
        return ctx

    def _resolver(self, items):
        """Valida el JSON del navegador contra la BD (no se confía en el cliente)."""
        resueltos, errores = [], []
        for it in items:
            try:
                p = Producto.objects.get(pk=int(it["producto"]), estado=True)
            except (Producto.DoesNotExist, KeyError, ValueError, TypeError):
                errores.append("Hay un material inválido en la lista.")
                continue

            if p.es_serializado:
                ids = [int(u["id"]) for u in it.get("unidades", [])]
                unidades = list(UnidadSerializada.objects.filter(pk__in=ids, producto=p))
                if not ids or len(unidades) != len(set(ids)):
                    errores.append(f"{p.nombre_completo}: seriales inválidos.")
                    continue
                resueltos.append((p, len(unidades), unidades))
            else:
                try:
                    cant = int(it.get("cantidad", 0))
                except (ValueError, TypeError):
                    cant = 0
                if cant <= 0:
                    errores.append(f"{p.nombre_completo}: cantidad inválida.")
                    continue
                resueltos.append((p, cant, []))
        return resueltos, errores

    def post(self, request, *args, **kwargs):
        form = EntregaForm(request.POST)
        try:
            items = json.loads(request.POST.get("payload", "[]"))
        except ValueError:
            items = []

        def volver():
            return self.render_to_response(self.get_context_data(form=form, payload=items))

        if not form.is_valid():
            return volver()
        if not items:
            messages.error(request, "Agregue al menos un material.")
            return volver()

        resueltos, errores = self._resolver(items)
        if errores:
            for e in errores:
                messages.error(request, e)
            return volver()

        d = form.cleaned_data
        try:
            with transaction.atomic():
                entrega = EntregaMaterial.objects.create(
                    tecnico=d["tecnico"], entregado_por=request.user,
                    fecha=d["fecha"], observaciones=d["observaciones"],
                )
                for producto, cantidad, unidades in resueltos:
                    if producto.es_serializado:
                        for u in unidades:
                            services.asignar(
                                producto=producto, tecnico=d["tecnico"], usuario=request.user,
                                cantidad=1, unidad=u, fecha=d["fecha"],
                                observaciones=d["observaciones"],
                            )
                            EntregaItem.objects.create(
                                entrega=entrega, producto=producto, unidad=u,
                                serial=u.numero_serie or u.codigo, cantidad=1,
                            )
                    else:
                        services.asignar(
                            producto=producto, tecnico=d["tecnico"], usuario=request.user,
                            cantidad=cantidad, unidad=None, fecha=d["fecha"],
                            observaciones=d["observaciones"],
                        )
                        EntregaItem.objects.create(
                            entrega=entrega, producto=producto, cantidad=cantidad,
                        )
        except services.ErrorInventario as e:
            messages.error(request, str(e))
            return volver()

        messages.success(request, f"Entrega {entrega.codigo} registrada.")
        return redirect("almacen:entrega_detalle", pk=entrega.pk)


class EntregaDetailView(RolRequeridoMixin, DetailView):
    roles_permitidos = ROLES_ALMACEN
    model = EntregaMaterial
    template_name = "almacen/entrega_detail.html"

    def get_queryset(self):
        return (super().get_queryset()
                .select_related("tecnico", "entregado_por")
                .prefetch_related("items__producto"))


# ---------- PDF de una entrega ----------
class EntregaPdfView(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN

    def get(self, request, pk):
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.platypus import (Image, Paragraph, SimpleDocTemplate,
                                        Spacer, Table, TableStyle)

        e = get_object_or_404(
            EntregaMaterial.objects.select_related("tecnico", "entregado_por")
            .prefetch_related("items__producto"), pk=pk)

        buf = BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, title=e.codigo,
                                leftMargin=1.6*cm, rightMargin=1.6*cm,
                                topMargin=1.4*cm, bottomMargin=1.4*cm)
        ss = getSampleStyleSheet()
        normal = ParagraphStyle("n", parent=ss["Normal"], fontSize=9, leading=11)
        small = ParagraphStyle("s", parent=normal, fontSize=8, textColor=colors.grey)
        titulo = ParagraphStyle("t", parent=ss["Title"], fontSize=15, leading=18)
        AZUL = colors.HexColor("#1f3a5f")

        # Encabezado: logo | título | código
        ruta_logo = finders.find("almacen/img/logo.png")
        logo = Image(ruta_logo, width=3.2*cm, height=1.6*cm, kind="proportional") if ruta_logo else ""
        cab = Table(
            [[logo,
              Paragraph("ACTA DE ENTREGA DE MATERIAL", titulo),
              Paragraph(f"<b>N.º</b><br/>{e.codigo}<br/><b>Fecha</b><br/>{e.fecha:%d/%m/%Y}", normal)]],
            colWidths=[3.8*cm, 9.2*cm, 4.4*cm])
        cab.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOX", (2, 0), (2, 0), 0.8, AZUL),
            ("LINEBELOW", (0, 0), (-1, 0), 1.2, AZUL),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))

        # Partes
        partes = Table([
            [Paragraph("<b>Técnico (recibe)</b>", normal), Paragraph("<b>Entrega</b>", normal)],
            [Paragraph(f"{_nombre(e.tecnico)}<br/>C.C. {_cedula(e.tecnico)}", normal),
             Paragraph(f"{_nombre(e.entregado_por)}<br/>C.C. {_cedula(e.entregado_por)}", normal)],
        ], colWidths=[8.7*cm, 8.7*cm])
        partes.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8eef6")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))

        # Materiales
        filas = [["#", "Código", "Material", "Serial / código unidad", "Cant."]]
        for i, it in enumerate(e.items.all(), 1):
            filas.append([i, it.producto.codigo_producto or "", Paragraph(it.producto.nombre_completo, normal),
                          it.serial or "—", it.cantidad])
        filas.append(["", "", "", "TOTAL", e.total_cantidad])
        tabla = Table(filas, colWidths=[1*cm, 2.8*cm, 7.2*cm, 4.4*cm, 2*cm], repeatRows=1)
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), AZUL), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ALIGN", (-1, 0), (-1, -1), "CENTER"), ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("FONTNAME", (-2, -1), (-1, -1), "Helvetica-Bold"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#f5f7fa")]),
        ]))

        # Firmas
        firmas = Table([
            ["", "", ""],
            [Paragraph(f"<b>{_nombre(e.entregado_por)}</b><br/>C.C. {_cedula(e.entregado_por)}<br/>Entrega (Almacén)", normal),
             "",
             Paragraph(f"<b>{_nombre(e.tecnico)}</b><br/>C.C. {_cedula(e.tecnico)}<br/>Recibe (Técnico)", normal)],
        ], colWidths=[7.4*cm, 2.6*cm, 7.4*cm], rowHeights=[2.2*cm, None])
        firmas.setStyle(TableStyle([
            ("LINEABOVE", (0, 1), (0, 1), 0.8, colors.black),
            ("LINEABOVE", (2, 1), (2, 1), 0.8, colors.black),
            ("VALIGN", (0, 1), (-1, 1), "TOP"),
        ]))

        story = [cab, Spacer(1, 10), partes, Spacer(1, 10), tabla]
        if e.observaciones:
            story += [Spacer(1, 8), Paragraph(f"<b>Observaciones:</b> {e.observaciones}", normal)]
        story += [Spacer(1, 10),
                  Paragraph("El técnico declara recibir los materiales relacionados y se hace "
                            "responsable de su custodia y uso.", small),
                  Spacer(1, 20), firmas, Spacer(1, 16),
                  Paragraph(f"Documento generado el {timezone.localtime():%d/%m/%Y %H:%M} · {e.codigo}", small)]
        doc.build(story)

        resp = HttpResponse(buf.getvalue(), content_type="application/pdf")
        resp["Content-Disposition"] = f'inline; filename="{e.codigo}.pdf"'
        return resp


# ---------- Excel de una entrega ----------
class EntregaExcelView(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN

    def get(self, request, pk):
        e = get_object_or_404(
            EntregaMaterial.objects.select_related("tecnico", "entregado_por")
            .prefetch_related("items__producto"), pk=pk)

        wb = Workbook()
        ws = wb.active
        ws.title = "Entrega"
        negrita = Font(bold=True)
        head = PatternFill("solid", fgColor="1F3A5F")

        ws["A1"] = "ACTA DE ENTREGA DE MATERIAL"; ws["A1"].font = Font(bold=True, size=14)
        datos = [
            ("N.º", e.codigo), ("Fecha", e.fecha.strftime("%d/%m/%Y")),
            ("Técnico", _nombre(e.tecnico)), ("C.C. técnico", _cedula(e.tecnico)),
            ("Entregado por", _nombre(e.entregado_por)), ("C.C. entrega", _cedula(e.entregado_por)),
            ("Observaciones", e.observaciones),
        ]
        for i, (k, v) in enumerate(datos, start=3):
            ws.cell(i, 1, k).font = negrita
            ws.cell(i, 2, v)

        fila = len(datos) + 4
        for c, t in enumerate(["#", "Código", "Material", "Serial / código unidad", "Cantidad"], 1):
            cel = ws.cell(fila, c, t)
            cel.font = Font(bold=True, color="FFFFFF"); cel.fill = head
            cel.alignment = Alignment(horizontal="center")
        for i, it in enumerate(e.items.all(), 1):
            fila += 1
            for c, v in enumerate([i, it.producto.codigo_producto or "", it.producto.nombre_completo,
                                   it.serial or "", it.cantidad], 1):
                ws.cell(fila, c, v)
        fila += 1
        ws.cell(fila, 4, "TOTAL").font = negrita
        ws.cell(fila, 5, e.total_cantidad).font = negrita

        fila += 3
        ws.cell(fila, 1, "Firma entrega: ____________________"); ws.cell(fila, 4, "Firma técnico: ____________________")
        for col, w in zip("ABCDE", (8, 16, 38, 26, 12)):
            ws.column_dimensions[col].width = w

        buf = BytesIO(); wb.save(buf)
        resp = HttpResponse(buf.getvalue(),
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        resp["Content-Disposition"] = f'attachment; filename="{e.codigo}.xlsx"'
        return resp


# ---------- Informe general (matriz como tu Excel) ----------
def _matriz_entregas(request):
    qs = (EntregaMaterial.objects
          .select_related("tecnico", "entregado_por")
          .prefetch_related("items__producto"))

    tecnico = request.GET.get("tecnico", "").strip()
    desde = request.GET.get("desde", "").strip()
    hasta = request.GET.get("hasta", "").strip()
    q = request.GET.get("q", "").strip()

    if tecnico.isdigit():
        qs = qs.filter(tecnico_id=int(tecnico))
    if desde:
        qs = qs.filter(fecha__gte=desde)
    if hasta:
        qs = qs.filter(fecha__lte=hasta)
    if q:
        qs = qs.filter(
            Q(codigo__icontains=q) | Q(items__serial__icontains=q)
            | Q(items__producto__nombre__icontains=q)
            | Q(tecnico__first_name__icontains=q) | Q(tecnico__last_name__icontains=q)
            | Q(tecnico__cedula__icontains=q)
        ).distinct()

    columnas, filas = {}, []
    for e in qs.order_by("-fecha", "-id"):
        cants = {}
        for it in e.items.all():
            columnas.setdefault(it.producto_id, it.producto)
            cants[it.producto_id] = cants.get(it.producto_id, 0) + it.cantidad
        filas.append({"entrega": e, "cants": cants, "total": sum(cants.values())})

    # Solo columnas de materiales que realmente se entregaron
    cols = sorted(columnas.values(), key=lambda p: (p.nombre, p.medida or ""))
    for f in filas:
        f["celdas"] = [f["cants"].get(p.pk, 0) for p in cols]
    totales = [sum(f["cants"].get(p.pk, 0) for f in filas) for p in cols]

    return {
        "filas": filas, "cols": cols, "totales": totales,
        "total_general": sum(totales),
        "tecnico_sel": tecnico, "desde": desde, "hasta": hasta, "q": q,
    }


class EntregaListView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "almacen/entrega_list.html"

    def get(self, request, *args, **kwargs):
        if request.GET.get("export") == "xlsx":
            return self._excel(request)
        return super().get(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx.update(_matriz_entregas(self.request))
        ctx["tecnicos_lista"] = tecnicos_queryset()
        return ctx

    def _excel(self, request):
        m = _matriz_entregas(request)
        wb = Workbook(); ws = wb.active; ws.title = "Informe entregas"
        head = PatternFill("solid", fgColor="1F3A5F")
        enc = ["Código", "Fecha", "Técnico", "C.C.", "Entregado por"] + \
              [p.nombre_completo for p in m["cols"]] + ["Total"]
        for c, t in enumerate(enc, 1):
            cel = ws.cell(1, c, t)
            cel.font = Font(bold=True, color="FFFFFF"); cel.fill = head
            cel.alignment = Alignment(horizontal="center", wrap_text=True)
        for r, f in enumerate(m["filas"], 2):
            e = f["entrega"]
            vals = [e.codigo, e.fecha.strftime("%d/%m/%Y"), _nombre(e.tecnico),
                    _cedula(e.tecnico), _nombre(e.entregado_por)] + \
                   [v or None for v in f["celdas"]] + [f["total"]]
            for c, v in enumerate(vals, 1):
                ws.cell(r, c, v)
        r = len(m["filas"]) + 2
        ws.cell(r, 5, "TOTALES").font = Font(bold=True)
        for c, v in enumerate(m["totales"] + [m["total_general"]], 6):
            ws.cell(r, c, v).font = Font(bold=True)
        ws.freeze_panes = "F2"
        for c in range(1, len(enc) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 14 if c > 5 else 22
        buf = BytesIO(); wb.save(buf)
        resp = HttpResponse(buf.getvalue(),
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        resp["Content-Disposition"] = 'attachment; filename="informe_entregas.xlsx"'
        return resp
    
class ValidarSerialInstalacionJson(RolRequeridoMixin, View):
    """Dice si un serial lo tiene el técnico o qué inconsistencia sería."""
    roles_permitidos = ROLES_DIGITADOR

    def get(self, request):
        tecnico = get_object_or_404(User, pk=request.GET.get("tecnico"), rol="TECNICO")
        serial = request.GET.get("serial", "").strip()
        if not serial:
            return JsonResponse({"tipo": "VACIO", "mensaje": "Escriba el serial."})

        tipo, unidad, _a, detalle = services.clasificar_serial(serial=serial, tecnico=tecnico)
        return JsonResponse({
            "tipo": tipo or "OK",
            "mensaje": detalle or "Lo tiene asignado.",
            "instala": tipo not in services.TIPOS_SOLO_REPORTE,
            "id": unidad.pk if unidad else None,
            "serial": (unidad.numero_serie or unidad.codigo) if unidad else serial,
            "producto": unidad.producto_id if unidad else None,
            "nombre": unidad.producto.nombre_completo if unidad else "",
        })


class InconsistenciaListView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "almacen/inconsistencia_list.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        tipo, estado = g.get("tipo", "").strip(), g.get("estado", "").strip()
        tecnico, q = g.get("tecnico", "").strip(), g.get("q", "").strip()

        qs = Inconsistencia.objects.select_related(
            "tecnico", "producto", "unidad_serializada", "operacion"
        )
        if tipo:
            qs = qs.filter(tipo=tipo)
        if estado:
            qs = qs.filter(estado=estado)
        if tecnico.isdigit():
            qs = qs.filter(tecnico_id=int(tecnico))
        if q:
            qs = qs.filter(
                Q(serial_reportado__icontains=q)
                | Q(producto__nombre__icontains=q)
                | Q(operacion__numero_orden__icontains=q)
                | Q(tecnico__first_name__icontains=q)
                | Q(tecnico__last_name__icontains=q)
            )

        conteo = dict(
            Inconsistencia.objects.filter(estado=Inconsistencia.Estado.PENDIENTE)
            .values_list("tipo").annotate(n=Count("id"))
        )
        ctx.update(
            items=qs[:300],
            resumen=[(v, l, conteo.get(v, 0)) for v, l in Inconsistencia.Tipo.choices],
            total_pendientes=sum(conteo.values()),
            tipos=Inconsistencia.Tipo.choices,
            estados=Inconsistencia.Estado.choices,
            tipo_sel=tipo, estado_sel=estado, tecnico_sel=tecnico, q=q,
            tecnicos_lista=tecnicos_queryset(),
        )
        return ctx


class InconsistenciaResolverView(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN

    def post(self, request, pk):
        i = get_object_or_404(Inconsistencia, pk=pk)
        if i.estado == Inconsistencia.Estado.PENDIENTE:
            i.estado = Inconsistencia.Estado.RESUELTA
            i.resuelto_por = request.user
            i.fecha_resolucion = timezone.now()
            i.nota_resolucion = request.POST.get("nota", "").strip()
            i.save()
            messages.success(request, "Inconsistencia marcada como resuelta.")
        return redirect(request.POST.get("next") or "almacen:inconsistencias")