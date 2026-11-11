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
from django.utils.dateparse import parse_date
from .models import (Asignacion, Categoria, Inconsistencia, Inventario,
                     MovimientoInventario, Producto, UnidadSerializada,IVT)
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
    AsignarIVTForm,
    IngresoIVTForm,
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
from django.db.models import Min, Max
ETIQUETAS_ERRORES = {
    "NO_ASIGNADO": "INSTALADO NO ASIGNADO A ESTE TECNICO",
    "OTRO_TECNICO": "INSTALADO ASIGNADO A OTRO TECNICO",
    "NO_REGISTRADO": "SERIAL INEXISTENTE",
    "YA_INSTALADO": "YA INSTALADO",
    "NO_UTILIZABLE": "DAÑADO / BAJA",
    "EXCEDE_CANTIDAD": "EXCEDE CANTIDAD ASIGNADA",
    "SIN_STOCK": "SIN STOCK",
}

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

        bajo = self.request.GET.get("bajo") == "1"

        if bajo:
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

        # ==================================================
        # CONTADORES DEL INVENTARIO CON FILTROS ACTIVOS
        # ==================================================

        qs = self.get_queryset()

        # Si get_queryset devuelve una lista por comparación,
        # también podemos contar sus elementos.
        ctx["total_productos"] = len(qs)

        ctx["total_existencias"] = sum(
            i.cantidad_actual or 0 for i in qs
        )
        ctx["total_danado"] = sum(
            i.cantidad_danada or 0 for i in qs
        )

        ctx["total_recibido"] = sum(
            i.total or 0 for i in qs
        )

        ctx["total_minimo"] = sum(
            i.cantidad_minima or 0 for i in qs
        )

        ctx["total_stock_bajo"] = sum(
            1 for i in qs
            if i.cantidad_minima > 0
            and i.cantidad_actual <= i.cantidad_minima
        )

        ctx["total_five"] = sum(
            five.get(i.producto_id, 0) for i in qs
        ) if not tecnico_sel.isdigit() else 0

        ctx["total_sap"] = sum(
            sap.get(i.producto_id, 0) for i in qs
        ) if not tecnico_sel.isdigit() else 0

        ctx["total_asignado"] = sum(
            asignado.get(i.producto_id, 0) for i in qs
        ) if not tecnico_sel.isdigit() else 0

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

class RemisionInfoJson(RolRequeridoMixin, View):
    """Dice si un número de remisión ya está registrado."""
    roles_permitidos = ROLES_ALMACEN

    def get(self, request):
        numero = request.GET.get("numero", "").strip().upper()
        if not numero:
            return JsonResponse({"vacio": True})

        r = RemisionIngreso.objects.filter(numero_remision=numero).first()
        if not r:
            return JsonResponse({"vacio": False, "existe": False})

        return JsonResponse({
            "vacio": False,
            "existe": True,
            "fecha": r.fecha_recibo_bsi.strftime("%d/%m/%Y") if r.fecha_recibo_bsi else "",
            "recogido_por": r.recogido_por,
            "recibido_almacen_por": r.recibido_almacen_por,
        })
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
        ctx["remisiones_existentes"] = RemisionIngreso.objects.all().order_by("-fecha_registro")[:50]
        ctx["siguiente_ivt"] = (IVT.objects.aggregate(m=Max("numero"))["m"] or 0) + 1
        return ctx

    def post(self, request, *args, **kwargs):
        if request.POST.get("tipo_ingreso") == "IVT":
            return self._post_ivt(request)

        form = self.form_class(request.POST, request.FILES)
        if form.is_valid():
            return self.form_valid(form)
        return self.render_to_response(self.get_context_data(form=form))

    def _post_ivt(self, request):
        p = request.POST
        numero = p.get("numero_orden", "").strip().upper()

        def volver():
            form = self.form_class(initial={"numero_orden": numero})
            return self.render_to_response(
                self.get_context_data(form=form, tipo_ingreso="IVT")
            )

        if not numero:
            messages.error(request, "Escriba el número de orden / remisión.")
            return volver()

        cajas = []
        try:
            for ini, cant in zip(p.getlist("ivt_inicial"), p.getlist("ivt_cantidad")):
                if not ini.strip() and not cant.strip():
                    continue
                cajas.append((int(ini), int(cant)))
        except ValueError:
            messages.error(request, "Hay una caja con consecutivo o cantidad inválidos.")
            return volver()

        if not cajas:
            messages.error(request, "Agregue al menos una caja de IVT.")
            return volver()

        try:
            with transaction.atomic():
                remision = services.obtener_o_crear_remision(
                    numero=numero,
                    usuario=request.user,
                    imagen=request.FILES.get("comprobante_imagen"),
                    fecha_recibo_bsi=parse_date(p.get("fecha_recibo_bsi") or ""),
                    recogido_por=p.get("recogido_por", "").strip(),
                    recibido_almacen_por=p.get("recibido_almacen_por", "").strip(),
                )
                creadas = services.registrar_ingreso_ivt(
                    cajas=cajas, remision=remision, usuario=request.user,
                    fecha_ingreso=timezone.localdate(),
                    observacion=p.get("observacion", "").strip(),
                )
        except services.ErrorInventario as e:
            messages.error(request, str(e))
            return volver()

        total = sum(c.cantidad for c in creadas)
        messages.success(
            request,
            f"Ingresaron {total} IVT en {len(creadas)} caja(s) con la remisión {numero}.",
        )
        return redirect(f"{reverse('almacen:ivt_lista')}?q={numero}")

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
    Consolidado por técnico y material.

    ASIGNADOS:    material entregado y no devuelto.
    FIVE:         cantidad instalada según FIVE.
    SAP:          cantidad instalada reportada por SAP.
    DEBE_TENER:   ASIGNADOS - mayor valor entre FIVE y SAP.
    DIFERENCIAS:  diferencia absoluta entre FIVE y SAP.
    ERRORES:      una fila por tipo de Inconsistencia pendiente,
                  con su cantidad en la columna del material.
    IVT:          consolidado independiente por técnico.
    """

    q = request.GET.get("q", "").strip()
    estado_sel = request.GET.get("estado", "").strip()
    tecnico_sel = request.GET.get("tecnico", "").strip()
    ivt_estado_sel = request.GET.get("ivt_estado", "").strip()

    # =========================================================
    # ASIGNACIONES
    # =========================================================

    qs = (
        Asignacion.objects
        .select_related(
            "tecnico",
            "tecnico__supervisor",
            "producto",
            "producto__categoria",
            "unidad_serializada",
            "operacion",
        )
        .order_by("-fecha_asignacion", "-id")
    )

    if tecnico_sel.isdigit():
        qs = qs.filter(tecnico_id=int(tecnico_sel))

    if estado_sel:
        qs = qs.filter(estado=estado_sel)

    if q:
        filtros = (
            Q(tecnico__cedula__icontains=q)
            | Q(tecnico__first_name__icontains=q)
            | Q(tecnico__last_name__icontains=q)
            | Q(tecnico__username__icontains=q)
            | Q(tecnico__supervisor__first_name__icontains=q)
            | Q(tecnico__supervisor__last_name__icontains=q)
            | Q(tecnico__supervisor__username__icontains=q)
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

        unidad = (
            UnidadSerializada.objects
            .filter(Q(codigo__iexact=q) | Q(numero_serie__iexact=q))
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

    # =========================================================
    # IVT (consulta independiente)
    # =========================================================

    ivt_queryset = (
        IVT.objects
        .select_related("tecnico", "operacion", "caja", "caja__remision")
        .filter(tecnico__isnull=False)
    )

    if tecnico_sel.isdigit():
        ivt_queryset = ivt_queryset.filter(tecnico_id=int(tecnico_sel))

    if ivt_estado_sel in IVT.Estado.values:
        ivt_queryset = ivt_queryset.filter(estado=ivt_estado_sel)

    if q:
        filtros_ivt = (
            Q(operacion__numero_orden__icontains=q)
            | Q(caja__remision__numero_remision__icontains=q)
        )
        if q.isdigit():
            filtros_ivt |= Q(numero=int(q))

        ivt_queryset = ivt_queryset.filter(filtros_ivt)

    # =========================================================
    # INCONSISTENCIAS PENDIENTES (técnico, material, tipo)
    # =========================================================

    inconsistencias_qs = Inconsistencia.objects.filter(
        tecnico__isnull=False,
        estado=Inconsistencia.Estado.PENDIENTE,
    )

    if tecnico_sel.isdigit():
        inconsistencias_qs = inconsistencias_qs.filter(
            tecnico_id=int(tecnico_sel)
        )

    if q:
        inconsistencias_qs = inconsistencias_qs.filter(
            Q(tecnico__cedula__icontains=q)
            | Q(tecnico__first_name__icontains=q)
            | Q(tecnico__last_name__icontains=q)
            | Q(tecnico__username__icontains=q)
            | Q(serial_reportado__icontains=q)
            | Q(producto__nombre__icontains=q)
            | Q(producto__codigo_producto__icontains=q)
            | Q(operacion__numero_orden__icontains=q)
        )

    errores_agrupados = list(
        inconsistencias_qs
        .values("tecnico_id", "producto_id", "tipo")
        .annotate(n=Count("id"))
    )

    errores_dict = {}    # (tecnico, producto, tipo) -> n
    errores_total = {}   # (tecnico, tipo) -> n  (incluye los sin material)

    for e in errores_agrupados:
        clave_total = (e["tecnico_id"], e["tipo"])
        errores_total[clave_total] = errores_total.get(clave_total, 0) + e["n"]

        if e["producto_id"]:
            errores_dict[(e["tecnico_id"], e["producto_id"], e["tipo"])] = e["n"]

    # =========================================================
    # AGRUPAR TÉCNICOS Y PRODUCTOS
    # =========================================================

    productos_dict = OrderedDict()
    tecnicos_dict = OrderedDict()

    datos_sap = _obtener_datos_sap()

    for asignacion in qs:
        producto = asignacion.producto
        tecnico = asignacion.tecnico

        productos_dict.setdefault(producto.pk, producto)

        bloque = tecnicos_dict.setdefault(
            tecnico.pk,
            {"tecnico": tecnico, "productos": {}},
        )

        datos = bloque["productos"].setdefault(producto.pk, _datos_vacios())

        if asignacion.estado == Asignacion.Estado.DEVUELTA:
            continue

        datos["asignados"] += asignacion.cantidad or 0

        if asignacion.estado == Asignacion.Estado.INSTALADA:
            datos["five"] += asignacion.cantidad or 0
        else:
            datos["five"] += asignacion.cantidad_instalada or 0

    # =========================================================
    # TÉCNICOS QUE SOLO TIENEN IVT O SOLO ERRORES
    # =========================================================

    ids_tecnicos_extra = set(
        ivt_queryset.values_list("tecnico_id", flat=True).distinct()
    )
    ids_tecnicos_extra |= {e["tecnico_id"] for e in errores_agrupados}
    ids_tecnicos_extra -= set(tecnicos_dict.keys())

    if ids_tecnicos_extra:
        for t in (
            User.objects
            .filter(pk__in=ids_tecnicos_extra)
            .select_related("supervisor")
        ):
            tecnicos_dict[t.pk] = {"tecnico": t, "productos": {}}

    # =========================================================
    # MATERIALES QUE SOLO TIENEN ERRORES
    # =========================================================

    ids_prod_extra = {
        e["producto_id"] for e in errores_agrupados if e["producto_id"]
    } - set(productos_dict.keys())

    if ids_prod_extra:
        for p in Producto.objects.filter(pk__in=ids_prod_extra):
            productos_dict[p.pk] = p

    # =========================================================
    # CONSOLIDADO IVT POR TÉCNICO
    # =========================================================

    ivt_por_tecnico = list(
        ivt_queryset
        .values(
            "tecnico_id",
            "tecnico__first_name",
            "tecnico__last_name",
            "tecnico__username",
        )
        .annotate(
            total=Count("id"),
            disponibles=Count("id", filter=Q(estado=IVT.Estado.DISPONIBLE)),
            asignados=Count("id", filter=Q(estado=IVT.Estado.ASIGNADO)),
            instalados=Count("id", filter=Q(estado=IVT.Estado.INSTALADO)),
            anulados=Count("id", filter=Q(estado=IVT.Estado.ANULADO)),
            perdidos=Count("id", filter=Q(estado=IVT.Estado.PERDIDO)),
            minimo_asignado=Min("numero", filter=Q(estado=IVT.Estado.ASIGNADO)),
            maximo_asignado=Max("numero", filter=Q(estado=IVT.Estado.ASIGNADO)),
            minimo_instalado=Min("numero", filter=Q(estado=IVT.Estado.INSTALADO)),
            maximo_instalado=Max("numero", filter=Q(estado=IVT.Estado.INSTALADO)),
        )
        .order_by("tecnico__first_name", "tecnico__last_name")
    )

    ivt_datos = {fila["tecnico_id"]: fila for fila in ivt_por_tecnico}

    ivt_vacio = {
        "total": 0, "disponibles": 0, "asignados": 0,
        "instalados": 0, "anulados": 0, "perdidos": 0,
        "minimo_asignado": None, "maximo_asignado": None,
        "minimo_instalado": None, "maximo_instalado": None,
    }

    # =========================================================
    # ORDENAR (técnicos por nombre, materiales por nombre/medida)
    # =========================================================

    productos = sorted(
        productos_dict.values(),
        key=lambda p: (p.nombre or "", p.medida or ""),
    )

    bloques = sorted(
        tecnicos_dict.values(),
        key=lambda b: (
            (b["tecnico"].first_name or "").lower(),
            (b["tecnico"].last_name or "").lower(),
            b["tecnico"].username,
        ),
    )

    # =========================================================
    # CONSTRUIR FILAS POR TÉCNICO
    # =========================================================

    base_unidades = reverse("almacen:unidades")
    base_inventario = reverse("almacen:inventario")
    base_inconsistencias = reverse("almacen:inconsistencias")

    def url_celda(tecnico_pk, producto, tipo):
        if producto.es_serializado:
            estado = (
                UnidadSerializada.Estado.INSTALADO
                if tipo == "FIVE"
                else UnidadSerializada.Estado.ASIGNADO
            )
            return f"{base_unidades}?" + urlencode({
                "tecnico": tecnico_pk,
                "producto": producto.pk,
                "estado": estado,
            })

        return f"{base_inventario}?" + urlencode({
            "tecnico": tecnico_pk,
            "q": producto.codigo_producto or producto.nombre,
        })

    tecnicos = []

    for bloque in bloques:
        tecnico = bloque["tecnico"]

        filas = {
            "ASIGNADOS": [],
            "DEBE_TENER": [],
            "FIVE": [],
            "SAP": [],
            "DIFERENCIAS": [],
        }

        productos_detalle = []

        for producto in productos:
            datos = bloque["productos"].get(producto.pk, _datos_vacios())

            asignados = datos["asignados"]
            five = datos["five"]
            sap = datos_sap.get((tecnico.pk, producto.pk), 0)

            mayor_instalado = max(five, sap)
            debe_tener = max(asignados - mayor_instalado, 0)
            diferencia = abs(five - sap)
            coincide = sap == five

            def celda(cantidad, tipo):
                return {
                    "cantidad": cantidad,
                    "url": url_celda(tecnico.pk, producto, tipo),
                }

            filas["ASIGNADOS"].append(celda(asignados, "ASIGNADOS"))
            filas["DEBE_TENER"].append(celda(debe_tener, "DEBE_TENER"))
            filas["FIVE"].append(celda(five, "FIVE"))

            celda_sap = celda(sap, "SAP")
            celda_sap["coincide"] = coincide
            celda_sap["estado_conciliacion"] = "ok" if coincide else "error"
            filas["SAP"].append(celda_sap)

            filas["DIFERENCIAS"].append(celda(diferencia, "DIFERENCIAS"))

            productos_detalle.append({
                "producto": producto,
                "asignados": asignados,
                "five": five,
                "sap": sap,
                "mayor_instalado": mayor_instalado,
                "diferencia": diferencia,
                "debe_tener": debe_tener,
            })

        # -----------------------------------------------------
        # FILAS DE ERRORES: una por tipo de inconsistencia
        # -----------------------------------------------------

        filas_errores = []

        for tipo, etiqueta in Inconsistencia.Tipo.choices:
            celdas_error = []

            for producto in productos:
                n = errores_dict.get((tecnico.pk, producto.pk, tipo), 0)

                url = ""
                if n:
                    url = f"{base_inconsistencias}?" + urlencode({
                        "tipo": tipo,
                        "estado": "PENDIENTE",
                        "tecnico": tecnico.pk,
                        "producto": producto.pk,
                    })

                celdas_error.append({"cantidad": n, "url": url})

            total_tipo = errores_total.get((tecnico.pk, tipo), 0)

            filas_errores.append({
                "tipo": tipo,
                "etiqueta": etiqueta,
                "corto": ETIQUETAS_ERRORES.get(tipo, tipo),
                "celdas": celdas_error,
                "total": total_tipo,
                "url_total": (
                    f"{base_inconsistencias}?" + urlencode({
                        "tipo": tipo,
                        "estado": "PENDIENTE",
                        "tecnico": tecnico.pk,
                    })
                ) if total_tipo else "",
            })

        tecnicos.append({
            "tecnico": tecnico,
            "filas": filas,
            "errores": filas_errores,
            "productos_detalle": productos_detalle,
            "ivt": ivt_datos.get(tecnico.pk, ivt_vacio),
        })

    # =========================================================
    # RESUMEN GENERAL IVT (sin filtros de pantalla)
    # =========================================================

    ivt_total = IVT.objects.count()
    ivt_disponibles = IVT.objects.filter(estado=IVT.Estado.DISPONIBLE).count()

    # =========================================================
    # CONTEXTO
    # =========================================================

    return {
        "tecnicos": tecnicos,
        "productos": productos,

        "estados": Asignacion.Estado.choices,
        "ivt_estados": IVT.Estado.choices,

        "estado_sel": estado_sel,
        "ivt_estado_sel": ivt_estado_sel,

        "q": q,
        "tecnico_sel": tecnico_sel,

        "tecnicos_lista": User.objects.filter(is_active=True).order_by(
            "first_name", "last_name"
        ),

        "ivt_por_tecnico": ivt_por_tecnico,
        "ivt_disponibles": ivt_disponibles,
        "ivt_total": ivt_total,
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

        payload_previo = []
        try:
            crudo = json.loads(self.request.POST.get("payload", "[]"))
            nombres = {
                p.pk: p.nombre_completo
                for p in Producto.objects.filter(
                    pk__in=[
                        int(l.get("producto_id"))
                        for l in crudo
                        if str(l.get("producto_id", "")).isdigit()
                    ]
                )
            }
            for l in crudo if isinstance(crudo, list) else []:
                pid = int(l.get("producto_id"))
                seriales = l.get("seriales") or []
                payload_previo.append({
                    "producto": pid,
                    "nombre": nombres.get(pid, "Material"),
                    "serializado": bool(seriales),
                    "cantidad": 0 if seriales else int(l.get("cantidad") or 0),
                    "unidades": [
                        {"id": None, "serial": s} for s in seriales
                    ],
                })
        except (ValueError, TypeError):
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
        
        form = InstalacionMasivaForm(request.POST, request.FILES, tecnico=self.tecnico)

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
                    services.vincular_ivt(
                        numero=d["numero_ivt"], operacion=op,
                        tecnico=self.tecnico, usuario=request.user)

                if instala:
                    novedades = services.sincronizar_materiales_operacion(
                        operacion=op,
                        payload=lineas,
                        usuario=request.user,
                    )
                    piezas = sum(
                        int(l.get("cantidad") or 0) for l in lineas
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

    # ==========================================================
    # CARGAR ASIGNACIONES PENDIENTES DEL TÉCNICO
    # ==========================================================

    def _cargar(self):

        self.tecnico = get_object_or_404(
            User,
            pk=self.kwargs["pk"],
            rol="TECNICO"
        )

        pend = (
            Asignacion.objects
            .filter(
                tecnico=self.tecnico,
                estado=Asignacion.Estado.ASIGNADA
            )
            .select_related("producto", "unidad_serializada")
            .order_by("producto__nombre", "producto__medida", "id")
        )

        self.grupos = OrderedDict()
        self.seriales = []

        for a in pend:

            if a.cantidad_pendiente <= 0:
                continue

            if a.unidad_serializada_id:

                # Incluye serial físico y códigos internos.
                self.seriales.append(a)

            else:

                g = self.grupos.setdefault(
                    a.producto_id,
                    {
                        "producto": a.producto,
                        "pendiente": 0,
                    }
                )

                g["pendiente"] += a.cantidad_pendiente
                
        self.ivt_rangos = services.rangos_ivt(tecnico=self.tecnico)

    # ==========================================================
    # CONSTRUIR CONTEXTO
    # ==========================================================

    def get_context_data(self, **kwargs):

        ctx = super().get_context_data(**kwargs)

        errores_celdas = kwargs.get("errores_celdas", {})
        valores_post = kwargs.get("valores_post", {})

        self._cargar()

        grupos = list(self.grupos.values())

        # ------------------------------------------------------
        # MATERIALES POR CANTIDAD
        # ------------------------------------------------------

        for g in grupos:

            pid = g["producto"].pk

            g["error_disponible"] = errores_celdas.get(
                f"disp_{pid}", []
            )

            g["error_danado"] = errores_celdas.get(
                f"dan_{pid}", []
            )

            g["valor_disponible"] = valores_post.get(
                f"disp_{pid}", "0"
            )

            g["valor_danado"] = valores_post.get(
                f"dan_{pid}", "0"
            )

        # ------------------------------------------------------
        # UNIDADES SERIALIZADAS
        # ------------------------------------------------------

        for a in self.seriales:

            a.error_estado = errores_celdas.get(
                f"estado_ser_{a.pk}", []
            )

            a.valor_estado = valores_post.get(
                f"estado_ser_{a.pk}", "DISPONIBLE"
            )

            a.seleccionado = (
                valores_post.get(f"ser_{a.pk}") == "1"
            )

        # ------------------------------------------------------
        # CONTEXTO GENERAL
        # ------------------------------------------------------

        ctx.update(
            tecnico=self.tecnico,
            grupos=grupos,
            seriales=self.seriales,

            error_observacion=errores_celdas.get(
                "observacion", []
            ),

            valor_observacion=valores_post.get(
                "observacion", ""
            ),

            valores_post=valores_post,
            ivt_rangos=self.ivt_rangos,
            ivt_filas=kwargs.get("ivt_filas", []),
            error_ivt=errores_celdas.get("ivt", []),
        )

        return ctx

    # ==========================================================
    # PROCESAR DEVOLUCIÓN
    # ==========================================================

    def post(self, request, *args, **kwargs):

        self._cargar()

        # IMPORTANTE:
        # Los grupos deben definirse dentro de POST.
        grupos = list(self.grupos.values())

        obs = request.POST.get(
            "observacion", ""
        ).strip()

        errores = []
        errores_celdas = {}

        por_serial = []
        por_cantidad = []

        valores_post = request.POST.dict()

        # ------------------------------------------------------
        # FUNCIÓN PARA REGISTRAR ERRORES POR CAMPO
        # ------------------------------------------------------

        def error_celda(campo, mensaje):

            errores_celdas.setdefault(
                campo, []
            ).append(mensaje)

        # ======================================================
        # 1. VALIDAR OBSERVACIÓN
        # ======================================================

        if not obs:

            error_celda(
                "observacion",
                "La observación es obligatoria."
            )

        elif len(obs) < 5:

            error_celda(
                "observacion",
                "La observación debe contener al menos 5 caracteres."
            )

        # ======================================================
        # 2. VALIDAR UNIDADES CON CÓDIGO INDIVIDUAL
        # ======================================================

        for a in self.seriales:

            if request.POST.get(f"ser_{a.pk}") != "1":
                continue

            campo_estado = f"estado_ser_{a.pk}"

            estado = request.POST.get(
                campo_estado, ""
            )

            if estado not in (
                UnidadSerializada.Estado.DISPONIBLE,
                UnidadSerializada.Estado.DAÑADO,
            ):

                codigo = (
                    a.unidad_serializada.codigo
                    if a.unidad_serializada
                    else f"Asignación {a.pk}"
                )

                error_celda(
                    campo_estado,
                    f"Selecciona un estado válido para {codigo}."
                )

                continue

            por_serial.append((a, estado))

        # ======================================================
        # 3. VALIDAR MATERIALES POR CANTIDAD
        # ======================================================

        materiales_sin_cantidad = []

        for g in grupos:

            producto = g["producto"]
            pendiente = int(g["pendiente"] or 0)

            nombre = getattr(
                producto,
                "nombre_completo",
                str(producto)
            )

            campo_disponible = f"disp_{producto.pk}"
            campo_danado = f"dan_{producto.pk}"

            # --------------------------------------------------
            # LEER VALORES DEL FORMULARIO
            # --------------------------------------------------

            raw_disponible = request.POST.get(
                campo_disponible, "0"
            ).strip()

            raw_danado = request.POST.get(
                campo_danado, "0"
            ).strip()

            # --------------------------------------------------
            # VALIDAR DISPONIBLE
            # --------------------------------------------------

            try:

                disponible = int(raw_disponible or 0)

            except (TypeError, ValueError):

                error_celda(
                    campo_disponible,
                    f'{nombre}: "{raw_disponible}" no es una cantidad válida. '
                    "Ingresa un número entero, por ejemplo: 0, 1, 2 o 3."
                )

                continue

            # --------------------------------------------------
            # VALIDAR DAÑADO
            # --------------------------------------------------

            try:

                danado = int(raw_danado or 0)

            except (TypeError, ValueError):

                error_celda(
                    campo_danado,
                    f'{nombre}: "{raw_danado}" no es una cantidad válida. '
                    "Ingresa un número entero, por ejemplo: 0, 1, 2 o 3."
                )

                continue

            # --------------------------------------------------
            # VALIDAR NÚMEROS NEGATIVOS
            # --------------------------------------------------

            cantidad_negativa = False

            if disponible < 0:

                error_celda(
                    campo_disponible,
                    f"{nombre}: la cantidad disponible no puede ser negativa."
                )

                cantidad_negativa = True

            if danado < 0:

                error_celda(
                    campo_danado,
                    f"{nombre}: la cantidad dañada no puede ser negativa."
                )

                cantidad_negativa = True

            if cantidad_negativa:
                continue

            # --------------------------------------------------
            # VALIDAR CANTIDAD MÁXIMA PENDIENTE
            # --------------------------------------------------

            total_producto = disponible + danado

            if total_producto > pendiente:

                mensaje = (
                    f"{nombre}: ingresaste {total_producto} unidades "
                    f"({disponible} disponibles y {danado} dañadas), "
                    f"pero solamente tienes {pendiente} pendientes. "
                    f"Debes ingresar una cantidad igual o inferior a "
                    f"{pendiente}."
                )

                # Mostrar el error en los campos que tienen cantidad.
                if disponible > 0:
                    error_celda(
                        campo_disponible,
                        mensaje
                    )

                if danado > 0:
                    error_celda(
                        campo_danado,
                        mensaje
                    )

                continue

            # --------------------------------------------------
            # IDENTIFICAR MATERIALES SIN CANTIDAD
            # --------------------------------------------------

            if total_producto == 0:

                materiales_sin_cantidad.append(
                    f"{nombre} (pendiente: {pendiente})"
                )

                continue

            # --------------------------------------------------
            # AGREGAR CANTIDADES VÁLIDAS
            # --------------------------------------------------

            if disponible > 0:

                por_cantidad.append((
                    producto,
                    disponible,
                    UnidadSerializada.Estado.DISPONIBLE,
                ))

            if danado > 0:

                por_cantidad.append((
                    producto,
                    danado,
                    UnidadSerializada.Estado.DAÑADO,
                ))
    
        # ======================================================
        # 3b. VALIDAR RANGOS DE IVT
        # ======================================================

        filas_ivt = []   # lo escrito, para repintar si hay error
        por_ivt = []     # (inicial, final, estado) válidos

        for ini_raw, fin_raw, est in zip(
            request.POST.getlist("ivt_ini"),
            request.POST.getlist("ivt_fin"),
            request.POST.getlist("ivt_estado"),
        ):
            ini_raw, fin_raw = ini_raw.strip(), fin_raw.strip()
            if not ini_raw and not fin_raw:
                continue

            filas_ivt.append({"ini": ini_raw, "fin": fin_raw, "estado": est})

            try:
                ini = int(ini_raw)
                fin = int(fin_raw or ini_raw)   # solo "desde" = un único IVT
            except ValueError:
                error_celda("ivt", f'IVT "{ini_raw}" - "{fin_raw}": no es un rango válido.')
                continue

            if ini < 1 or fin < ini:
                error_celda("ivt", f"IVT {ini_raw}-{fin_raw}: el final debe ser mayor o igual al inicial.")
                continue

            if est not in ("DISPONIBLE", "ANULADO", "PERDIDO"):
                error_celda("ivt", f"IVT {ini:05d}-{fin:05d}: seleccione cómo regresa.")
                continue

            por_ivt.append((ini, fin, est))

        ordenados = sorted(por_ivt)
        for (_, f1, _), (i2, _, _) in zip(ordenados, ordenados[1:]):
            if i2 <= f1:
                error_celda("ivt", "Hay rangos de IVT que se cruzan entre sí.")
                break

        # ======================================================
        # 4. CALCULAR TOTAL SOLICITADO
        # ======================================================

        total_solicitado = (
            len(por_serial)
            + sum(
                cantidad
                for producto, cantidad, estado in por_cantidad
            )
        )

        # ======================================================
        # 5. VALIDAR QUE SE HAYA SELECCIONADO ALGÚN MATERIAL
        # ======================================================

        if total_solicitado == 0 and not errores_celdas:

            if materiales_sin_cantidad:

                detalle = "; ".join(
                    materiales_sin_cantidad
                )

                mensaje = (
                    "No has indicado cantidades para devolver. "
                    f"Los materiales pendientes son: {detalle}. "
                    "Ingresa una cantidad mayor que cero en Disponible "
                    "o Dañado, o selecciona una unidad con código individual."
                )

                errores.append(mensaje)

                # También señalar los campos que debe completar.
                for g in grupos:

                    pid = g["producto"].pk

                    error_celda(
                        f"disp_{pid}",
                        "Ingresa una cantidad mayor que cero en Disponible "
                        "o Dañado para devolver este material."
                    )

            else:

                errores.append(
                    "No has seleccionado ninguna unidad para devolver. "
                    "Selecciona una unidad con código individual "
                    "o ingresa una cantidad mayor que cero."
                )

        # ======================================================
        # 6. IMPEDIR DEVOLUCIONES SI EXISTEN ERRORES
        # ======================================================

        if errores_celdas or errores:

            for mensaje in errores:
                messages.error(request, mensaje)

            if errores_celdas:

                messages.error(
                    request,
                    "No se registró ninguna devolución. "
                    "Revisa los campos señalados y corrige los errores."
                )

            return self.render_to_response(
                self.get_context_data(
                    errores_celdas=errores_celdas,
                    valores_post=valores_post,
                    ivt_filas=filas_ivt,
                )
            )

        
        # ======================================================
        # 7. REGISTRAR DEVOLUCIÓN Y GENERAR COMPROBANTE
        # ======================================================

        try:
            with transaction.atomic():

                # Crear el documento de devolución.
                documento = EntregaMaterial.objects.create(
                    tipo=EntregaMaterial.Tipo.DEVOLUCION,
                    tecnico=self.tecnico,
                    entregado_por=request.user,
                    fecha=timezone.localdate(),
                    observaciones=obs,
                )

                # ----------------------------------------------
                # UNIDADES CON CÓDIGO INDIVIDUAL
                # ----------------------------------------------

                for a, estado in por_serial:

                    devuelta = services.devolver(
                        asignacion=a,
                        usuario=request.user,
                        observacion=obs,
                        estado_retorno=estado,
                    )

                    unidad = devuelta.unidad_serializada

                    EntregaItem.objects.create(
                        entrega=documento,
                        producto=devuelta.producto,
                        unidad=unidad,
                        asignacion=devuelta,
                        serial=(
                            getattr(unidad, "numero_serie", "") or unidad.codigo
                        ),
                        cantidad=1,
                        estado_retorno=estado,
                        observacion=obs,
                    )

                # ----------------------------------------------
                # MATERIALES POR CANTIDAD
                # ----------------------------------------------

                for producto, cantidad, estado in por_cantidad:

                    detalles = []

                    services.devolver_cantidad(
                        tecnico=self.tecnico,
                        producto=producto,
                        cantidad=cantidad,
                        usuario=request.user,
                        observacion=obs,
                        estado_retorno=estado,
                        detalle_devoluciones=detalles,
                    )

                    # Cada bloque devuelto queda registrado
                    # con su asignación correspondiente.
                    for detalle in detalles:

                        EntregaItem.objects.create(
                            entrega=documento,
                            producto=detalle["producto"],
                            asignacion=detalle["asignacion"],
                            cantidad=detalle["cantidad"],
                            estado_retorno=estado,
                            observacion=obs,
                        )
                        
                                # ----------------------------------------------
                # IVT POR RANGO
                # ----------------------------------------------

                for ini, fin, estado in por_ivt:

                    total = services.devolver_ivt_rango(
                        tecnico=self.tecnico,
                        inicial=ini,
                        final=fin,
                        usuario=request.user,
                        estado_retorno=estado,
                        observacion=obs,
                    )

                    EntregaItem.objects.create(
                        entrega=documento,
                        producto=None,
                        ivt_inicial=ini,
                        ivt_final=fin,
                        serial=f"{ini:05d}-{fin:05d}",
                        cantidad=total,
                        estado_retorno=estado,
                        observacion=obs,
                    )

        except services.ErrorInventario as e:

            messages.error(
                request,
                f"No se pudo registrar la devolución: {e}"
            )

            return self.render_to_response(
                self.get_context_data(
                    errores_celdas=errores_celdas,
                    valores_post=valores_post,
                )
            )


        # ======================================================
        # 8. CONFIRMACIÓN
        # ======================================================

        total_devuelto = total_solicitado

        messages.success(
            request,
            f"Devolución registrada correctamente: "
            f"{total_devuelto} pieza(s)."
        )

        return redirect("almacen:entrega_detalle", documento.pk)

   
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
            "producto",
            "unidad_serializada",
            "usuario",
        )

        tipo = self.request.GET.get("tipo", "").strip()
        fecha = self.request.GET.get("fecha", "").strip()
        responsable = self.request.GET.get("responsable", "").strip()

        if tipo:
            qs = qs.filter(tipo_movimiento=tipo)

        if fecha:
            qs = qs.filter(fecha__date=fecha)

        if responsable.isdigit():
            qs = qs.filter(usuario_id=int(responsable))

        return qs.order_by("-fecha", "-pk")

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        ctx["tipos"] = MovimientoInventario.TipoMovimiento.choices
        ctx["tipo_sel"] = self.request.GET.get("tipo", "")
        ctx["fecha_sel"] = self.request.GET.get("fecha", "")

        ctx["responsables"] = (
            User.objects.filter(
                movimientos__isnull=False
            )
            .distinct()
            .order_by("first_name", "username")
        )

        ctx["responsable_sel"] = self.request.GET.get("responsable", "")

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
        ctx["payload_ivt_previo"] = kwargs.get("payload_ivt", [])

        disp = IVT.objects.filter(estado=IVT.Estado.DISPONIBLE)
        ctx["ivt_disponibles"] = disp.count()
        ctx["ivt_primero"] = disp.aggregate(m=Min("numero"))["m"] or ""
        return ctx
    
    def _resolver_ivt(self, rangos):
        resueltos, errores = [], []
        for r in rangos:
            try:
                ini, fin = int(r["inicial"]), int(r["final"])
            except (KeyError, ValueError, TypeError):
                errores.append("Hay un rango de IVT inválido en la lista.")
                continue
            if ini < 1 or fin < ini:
                errores.append(f"IVT {ini}-{fin}: el final debe ser mayor o igual al inicial.")
                continue
            resueltos.append((ini, fin))

        ordenados = sorted(resueltos)
        for (_, f1), (i2, _) in zip(ordenados, ordenados[1:]):
            if i2 <= f1:
                errores.append("Hay rangos de IVT que se cruzan entre sí.")
                break
        return resueltos, errores

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
        try:
            ivt_items = json.loads(request.POST.get("payload_ivt", "[]"))
            if not isinstance(ivt_items, list):
                ivt_items = []
        except ValueError:
            ivt_items = []

        def volver():
            return self.render_to_response(
                self.get_context_data(form=form, payload=items, payload_ivt=ivt_items)
            )

        if not form.is_valid():
            return volver()
        if not items and not ivt_items:
            messages.error(request, "Agregue al menos un material o un rango de IVT.")
            return volver()

        resueltos, errores = self._resolver(items)
        ivt_resueltos, errores_ivt = self._resolver_ivt(ivt_items)
        errores += errores_ivt
        if errores:
            for e in errores:
                messages.error(request, e)
            return volver()

        d = form.cleaned_data
        try:
            with transaction.atomic():
                entrega = EntregaMaterial.objects.create(
                    tipo=EntregaMaterial.Tipo.ENTREGA,
                    tecnico=d["tecnico"],
                    entregado_por=request.user,
                    fecha=d["fecha"],
                    observaciones=d["observaciones"],
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

                # ---- IVT: cada rango es una línea más del documento ----
                for ini, fin in ivt_resueltos:
                    total = services.asignar_ivt_rango(
                        tecnico=d["tecnico"], inicial=ini, final=fin,
                        usuario=request.user, fecha=d["fecha"],
                        observacion=d["observaciones"],
                    )
                    EntregaItem.objects.create(
                        entrega=entrega, producto=None,
                        ivt_inicial=ini, ivt_final=fin,
                        serial=f"{ini:05d}-{fin:05d}", cantidad=total,
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
        return (
            super().get_queryset()
            .select_related("tecnico", "entregado_por")
            .prefetch_related(
                "items__producto",
                "items__unidad",
                "items__asignacion",
            )
        )


class EntregaPdfView(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN

    def get(self, request, pk):
        from xml.sax.saxutils import escape

        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.platypus import (
            Image, Paragraph, SimpleDocTemplate,
            Spacer, Table, TableStyle,
        )

        e = get_object_or_404(
            EntregaMaterial.objects
            .select_related("tecnico", "entregado_por")
            .prefetch_related("items__producto", "items__unidad"),
            pk=pk,
        )

        es_devolucion = e.tipo == EntregaMaterial.Tipo.DEVOLUCION

        titulo_texto = (
            "ACTA DE DEVOLUCIÓN DE MATERIAL"
            if es_devolucion
            else "ACTA DE ENTREGA DE MATERIAL"
        )

        nombre_tecnico = escape(str(_nombre(e.tecnico)))
        cedula_tecnico = escape(str(_cedula(e.tecnico)))
        nombre_almacen = escape(str(_nombre(e.entregado_por)))
        cedula_almacen = escape(str(_cedula(e.entregado_por)))

        buf = BytesIO()

        doc = SimpleDocTemplate(
            buf,
            pagesize=A4,
            title=f"{titulo_texto} - {e.codigo}",
            leftMargin=1.3 * cm,
            rightMargin=1.3 * cm,
            topMargin=1.3 * cm,
            bottomMargin=1.3 * cm,
        )

        ss = getSampleStyleSheet()

        normal = ParagraphStyle(
            "NormalDocumento",
            parent=ss["Normal"],
            fontSize=8,
            leading=10,
        )

        small = ParagraphStyle(
            "SmallDocumento",
            parent=normal,
            fontSize=7,
            textColor=colors.grey,
        )

        titulo = ParagraphStyle(
            "TituloDocumento",
            parent=ss["Title"],
            fontSize=13,
            leading=16,
            alignment=1,
        )

        azul = colors.HexColor("#1f3a5f")

        # Encabezado.
        ruta_logo = finders.find("almacen/img/logo.png")
        logo = (
            Image(
                ruta_logo,
                width=3.0 * cm,
                height=1.5 * cm,
                kind="proportional",
            )
            if ruta_logo
            else ""
        )

        cabecera = Table(
            [[
                logo,
                Paragraph(titulo_texto, titulo),
                Paragraph(
                    f"<b>N.º</b><br/>{escape(str(e.codigo))}"
                    f"<br/><b>Fecha</b><br/>{e.fecha:%d/%m/%Y}",
                    normal,
                ),
            ]],
            colWidths=[3.2 * cm, 9.5 * cm, 4.5 * cm],
        )

        cabecera.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOX", (2, 0), (2, 0), 0.8, azul),
            ("LINEBELOW", (0, 0), (-1, 0), 1.2, azul),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))

        # Responsables.
        if es_devolucion:
            encabezado_izq = "Técnico (devuelve)"
            encabezado_der = "Recibió en almacén"
            firma_izq = "Devuelve (Técnico)"
            firma_der = "Recibe (Almacén)"

            parte_izq = (
                f"{nombre_tecnico}<br/>C.C. {cedula_tecnico}"
            )
            parte_der = (
                f"{nombre_almacen}<br/>C.C. {cedula_almacen}"
            )
        else:
            encabezado_izq = "Técnico (recibe)"
            encabezado_der = "Entrega (Almacén)"
            firma_izq = "Entrega (Almacén)"
            firma_der = "Recibe (Técnico)"

            parte_izq = (
                f"{nombre_tecnico}<br/>C.C. {cedula_tecnico}"
            )
            parte_der = (
                f"{nombre_almacen}<br/>C.C. {cedula_almacen}"
            )

        partes = Table(
            [
                [
                    Paragraph(f"<b>{encabezado_izq}</b>", normal),
                    Paragraph(f"<b>{encabezado_der}</b>", normal),
                ],
                [
                    Paragraph(parte_izq, normal),
                    Paragraph(parte_der, normal),
                ],
            ],
            colWidths=[8.6 * cm, 8.6 * cm],
        )

        partes.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8eef6")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))

        # Materiales.
        if es_devolucion:
            filas = [[
                "#", "Código", "Material", "Serial / unidad",
                "Cant.", "Estado",
            ]]
        else:
            filas = [[
                "#", "Código", "Material", "Serial / unidad", "Cant.",
            ]]

        for i, it in enumerate(e.items.all(), 1):
            fila = [
                str(i),
                escape(str(it.producto.codigo_producto or "")),
                Paragraph(
                    escape(str(it.producto.nombre_completo)),
                    normal,
                ),
                escape(str(it.serial or "—")),
                str(it.cantidad),
            ]

            if es_devolucion:
                fila.append(
                    escape(str(it.get_estado_retorno_display() or "—"))
                )

            filas.append(fila)

        if es_devolucion:
            filas.append(["", "", "", "TOTAL", str(e.total_cantidad), ""])
            anchos = [
                0.7 * cm, 2.2 * cm, 5.0 * cm,
                3.8 * cm, 1.2 * cm, 3.0 * cm,
            ]
        else:
            filas.append(["", "", "", "TOTAL", str(e.total_cantidad)])
            anchos = [
                0.8 * cm, 2.7 * cm, 6.0 * cm,
                4.5 * cm, 1.8 * cm,
            ]

        tabla = Table(filas, colWidths=anchos, repeatRows=1)

        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), azul),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (-1, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            (
                "ROWBACKGROUNDS",
                (0, 1), (-1, -2),
                [colors.white, colors.HexColor("#f5f7fa")],
            ),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))

        # Firmas.
        if es_devolucion:
            firma_1 = Paragraph(
                f"<b>{nombre_tecnico}</b><br/>"
                f"C.C. {cedula_tecnico}<br/>{firma_izq}",
                normal,
            )
            firma_2 = Paragraph(
                f"<b>{nombre_almacen}</b><br/>"
                f"C.C. {cedula_almacen}<br/>{firma_der}",
                normal,
            )
        else:
            firma_1 = Paragraph(
                f"<b>{nombre_almacen}</b><br/>"
                f"C.C. {cedula_almacen}<br/>{firma_izq}",
                normal,
            )
            firma_2 = Paragraph(
                f"<b>{nombre_tecnico}</b><br/>"
                f"C.C. {cedula_tecnico}<br/>{firma_der}",
                normal,
            )

        firmas = Table(
            [["", ""], [firma_1, firma_2]],
            colWidths=[8.6 * cm, 8.6 * cm],
            rowHeights=[1.8 * cm, None],
        )

        firmas.setStyle(TableStyle([
            ("LINEABOVE", (0, 1), (0, 1), 0.8, colors.black),
            ("LINEABOVE", (1, 1), (1, 1), 0.8, colors.black),
            ("VALIGN", (0, 1), (-1, 1), "TOP"),
        ]))

        story = [
            cabecera,
            Spacer(1, 10),
            partes,
            Spacer(1, 10),
            tabla,
        ]

        if e.observaciones:
            story.extend([
                Spacer(1, 8),
                Paragraph(
                    f"<b>Observaciones:</b> "
                    f"{escape(str(e.observaciones))}",
                    normal,
                ),
            ])

        if es_devolucion:
            texto_declaracion = (
                "El almacén deja constancia de la recepción de los "
                "materiales relacionados, según el estado indicado."
            )
        else:
            texto_declaracion = (
                "El técnico declara recibir los materiales relacionados "
                "y se hace responsable de su custodia y uso."
            )

        story.extend([
            Spacer(1, 12),
            Paragraph(texto_declaracion, small),
            Spacer(1, 25),
            firmas,
            Spacer(1, 15),
            Paragraph(
                f"Documento generado el "
                f"{timezone.localtime():%d/%m/%Y %H:%M} · {e.codigo}",
                small,
            ),
        ])

        doc.build(story)

        response = HttpResponse(
            buf.getvalue(),
            content_type="application/pdf",
        )
        response["Content-Disposition"] = (
            f'inline; filename="{e.codigo}.pdf"'
        )
        return response



class EntregaExcelView(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN

    def get(self, request, pk):
        e = get_object_or_404(
            EntregaMaterial.objects
            .select_related("tecnico", "entregado_por")
            .prefetch_related("items__producto"),
            pk=pk,
        )

        es_devolucion = e.tipo == EntregaMaterial.Tipo.DEVOLUCION

        titulo = (
            "ACTA DE DEVOLUCIÓN DE MATERIAL"
            if es_devolucion
            else "ACTA DE ENTREGA DE MATERIAL"
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "Devolucion" if es_devolucion else "Entrega"

        negrita = Font(bold=True)
        encabezado = PatternFill("solid", fgColor="1F3A5F")

        ws["A1"] = titulo
        ws["A1"].font = Font(bold=True, size=14)

        responsable = (
            "Recibido por almacén"
            if es_devolucion
            else "Entregado por"
        )

        datos = [
            ("Documento", e.codigo),
            ("Fecha", e.fecha.strftime("%d/%m/%Y")),
            ("Técnico", _nombre(e.tecnico)),
            ("C.C. técnico", _cedula(e.tecnico)),
            (responsable, _nombre(e.entregado_por)),
            ("C.C. responsable", _cedula(e.entregado_por)),
            ("Observaciones", e.observaciones or ""),
        ]

        for i, (clave, valor) in enumerate(datos, start=3):
            ws.cell(i, 1, clave).font = negrita
            ws.cell(i, 2, valor)

        fila = len(datos) + 4

        columnas = [
            "#",
            "Código",
            "Material",
            "Serial / código unidad",
            "Cantidad",
        ]

        if es_devolucion:
            columnas.extend(["Estado de recepción", "Observación"])

        for c, texto in enumerate(columnas, 1):
            celda = ws.cell(fila, c, texto)
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = encabezado
            celda.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        for i, it in enumerate(e.items.all(), 1):
            fila += 1

            valores = [
                i,
                it.producto.codigo_producto or "",
                it.producto.nombre_completo,
                it.serial or "",
                it.cantidad,
            ]

            if es_devolucion:
                valores.extend([
                    it.get_estado_retorno_display() or "",
                    it.observacion or "",
                ])

            for c, valor in enumerate(valores, 1):
                ws.cell(fila, c, valor)

        fila += 1
        ws.cell(fila, 4, "TOTAL").font = negrita
        ws.cell(fila, 5, e.total_cantidad).font = negrita

        fila += 3

        if es_devolucion:
            ws.cell(
                fila, 1,
                "Firma técnico que devuelve: __________________________",
            )
            ws.cell(
                fila, 5,
                "Firma almacén que recibe: __________________________",
            )
        else:
            ws.cell(
                fila, 1,
                "Firma almacén que entrega: __________________________",
            )
            ws.cell(
                fila, 5,
                "Firma técnico que recibe: __________________________",
            )

        anchos = {
            "A": 8,
            "B": 16,
            "C": 38,
            "D": 26,
            "E": 12,
            "F": 22,
            "G": 35,
        }

        for columna, ancho in anchos.items():
            ws.column_dimensions[columna].width = ancho

        ws.freeze_panes = "A" + str(len(datos) + 5)
        ws.sheet_view.showGridLines = False

        buf = BytesIO()
        wb.save(buf)

        response = HttpResponse(
            buf.getvalue(),
            content_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

        response["Content-Disposition"] = (
            f'attachment; filename="{e.codigo}.xlsx"'
        )

        return response



def _matriz_entregas(request):
    qs = (
        EntregaMaterial.objects
        .select_related("tecnico", "entregado_por")
        .prefetch_related("items__producto")
    )

    tecnico = request.GET.get("tecnico", "").strip()
    desde = request.GET.get("desde", "").strip()
    hasta = request.GET.get("hasta", "").strip()
    q = request.GET.get("q", "").strip()
    tipo = request.GET.get("tipo", "").strip()

    if tipo in (
        EntregaMaterial.Tipo.ENTREGA,
        EntregaMaterial.Tipo.DEVOLUCION,
    ):
        qs = qs.filter(tipo=tipo)

    if tecnico.isdigit():
        qs = qs.filter(tecnico_id=int(tecnico))

    if desde:
        qs = qs.filter(fecha__gte=desde)

    if hasta:
        qs = qs.filter(fecha__lte=hasta)

    if q:
        qs = qs.filter(
            Q(codigo__icontains=q)
            | Q(items__serial__icontains=q)
            | Q(items__producto__nombre__icontains=q)
            | Q(items__producto__codigo_producto__icontains=q)
            | Q(tecnico__first_name__icontains=q)
            | Q(tecnico__last_name__icontains=q)
            | Q(tecnico__username__icontains=q)
            | Q(tecnico__cedula__icontains=q)
        ).distinct()

    columnas = {}
    filas = []

    for documento in qs.order_by("-fecha", "-id"):
        cantidades = {}

        for item in documento.items.all():
            columnas.setdefault(item.producto_id, item.producto)

            cantidades[item.producto_id] = (
                cantidades.get(item.producto_id, 0) + item.cantidad
            )

        filas.append({
            "entrega": documento,
            "cants": cantidades,
            "total": sum(cantidades.values()),
        })

    cols = sorted(
        columnas.values(),
        key=lambda p: (p.nombre, p.medida or ""),
    )

    for fila in filas:
        fila["celdas"] = [
            fila["cants"].get(producto.pk, 0)
            for producto in cols
        ]

    totales = [
        sum(fila["cants"].get(producto.pk, 0) for fila in filas)
        for producto in cols
    ]

    return {
        "filas": filas,
        "cols": cols,
        "totales": totales,
        "total_general": sum(totales),
        "tecnico_sel": tecnico,
        "desde": desde,
        "hasta": hasta,
        "q": q,
        "tipo_sel": tipo,
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

        wb = Workbook()
        ws = wb.active
        ws.title = "Entregas y devoluciones"

        encabezado = PatternFill("solid", fgColor="1F3A5F")
        negrita = Font(bold=True)

        columnas = (
            ["Código", "Tipo", "Fecha", "Técnico", "C.C.", "Responsable almacén"]
            + [p.nombre_completo for p in m["cols"]]
            + ["Total"]
        )

        for c, texto in enumerate(columnas, 1):
            celda = ws.cell(1, c, texto)
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = encabezado
            celda.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        for r, fila in enumerate(m["filas"], 2):
            e = fila["entrega"]

            tipo_texto = (
                "Devolución"
                if e.tipo == EntregaMaterial.Tipo.DEVOLUCION
                else "Entrega"
            )

            valores = [
                e.codigo,
                tipo_texto,
                e.fecha.strftime("%d/%m/%Y"),
                _nombre(e.tecnico),
                _cedula(e.tecnico),
                _nombre(e.entregado_por),
            ]

            valores += [v or None for v in fila["celdas"]]
            valores.append(fila["total"])

            for c, valor in enumerate(valores, 1):
                ws.cell(r, c, valor)

        fila_total = len(m["filas"]) + 2

        ws.cell(fila_total, 6, "TOTALES").font = negrita

        for c, valor in enumerate(
            m["totales"] + [m["total_general"]],
            7,
        ):
            ws.cell(fila_total, c, valor).font = negrita

        ws.freeze_panes = "G2"
        ws.sheet_view.showGridLines = False

        for c in range(1, len(columnas) + 1):
            ws.column_dimensions[get_column_letter(c)].width = (
                14 if c > 6 else 22
            )

        buf = BytesIO()
        wb.save(buf)

        response = HttpResponse(
            buf.getvalue(),
            content_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

        response["Content-Disposition"] = (
            'attachment; filename="informe_entregas_devoluciones.xlsx"'
        )

        return response


class InconsistenciaListView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "almacen/inconsistencia_list.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        tipo, estado = g.get("tipo", "").strip(), g.get("estado", "").strip()
        tecnico, q = g.get("tecnico", "").strip(), g.get("q", "").strip()
        producto = g.get("producto", "").strip()

        qs = Inconsistencia.objects.select_related(
            "tecnico", "producto", "unidad_serializada", "operacion"
        )
        if tipo:
            qs = qs.filter(tipo=tipo)
        if estado:
            qs = qs.filter(estado=estado)
        if tecnico.isdigit():
            qs = qs.filter(tecnico_id=int(tecnico))
        if producto.isdigit():                             # NUEVO
           qs = qs.filter(producto_id=int(producto))
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
            items=qs[:500],
            resumen=[(v, l, conteo.get(v, 0)) for v, l in Inconsistencia.Tipo.choices],
            total_pendientes=sum(conteo.values()),
            tipos=Inconsistencia.Tipo.choices,
            estados=Inconsistencia.Estado.choices,
            tipo_sel=tipo, estado_sel=estado, tecnico_sel=tecnico, q=q,
            producto_sel=producto,                         # NUEVO
            productos_lista=Producto.objects.filter(estado=True).order_by("nombre", "medida"),  # NUEVO
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
    
class IngresoIVTView(RolRequeridoMixin, FormView):
    roles_permitidos = ROLES_ALMACEN
    form_class = IngresoIVTForm
    template_name = "almacen/ingreso_ivt_form.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        ctx["remisiones_existentes"] = RemisionIngreso.objects.order_by("-fecha_registro")[:50]
        ctx["siguiente_ivt"] = (IVT.objects.aggregate(m=Max("numero"))["m"] or 0) + 1
        return ctx

    def form_valid(self, form):
        d = form.cleaned_data
        try:
            with transaction.atomic():
                remision = services.obtener_o_crear_remision(
                    numero=d["numero_orden"], usuario=self.request.user,
                    imagen=d.get("comprobante_imagen"),
                    fecha_recibo_bsi=d.get("fecha_recibo_bsi"),
                    recogido_por=d.get("recogido_por", ""),
                    recibido_almacen_por=d.get("recibido_almacen_por", ""),
                )
                cajas = services.registrar_ingreso_ivt(
                    cajas=d["cajas"], remision=remision, usuario=self.request.user,
                    fecha_ingreso=timezone.localdate(), observacion=d["observacion"],
                )
        except services.ErrorInventario as e:
            form.add_error(None, str(e))
            return self.form_invalid(form)

        total = sum(c.cantidad for c in cajas)
        messages.success(
            self.request,
            f"Ingresaron {total} IVT en {len(cajas)} caja(s) con la remisión {d['numero_orden']}.")
        return redirect(f"{reverse('almacen:ivt_lista')}?q={d['numero_orden']}")


class AsignarIVTView(RolRequeridoMixin, FormView):
    roles_permitidos = ROLES_ALMACEN
    form_class = AsignarIVTForm
    template_name = "almacen/ivt_asignar_form.html"
    success_url = reverse_lazy("almacen:ivt_lista")

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        disp = IVT.objects.filter(estado=IVT.Estado.DISPONIBLE)
        ctx["disponibles"] = disp.count()
        ctx["primer_disponible"] = disp.aggregate(m=Min("numero"))["m"]
        return ctx

    def form_valid(self, form):
        d = form.cleaned_data
        try:
            n = services.asignar_ivt_rango(
                tecnico=d["tecnico"], inicial=d["consecutivo_inicial"],
                final=d["consecutivo_final"], usuario=self.request.user,
                fecha=d["fecha"], observacion=d["observaciones"])
        except services.ErrorInventario as e:
            form.add_error(None, str(e))
            return self.form_invalid(form)
        messages.success(self.request, f"{n} IVT asignados a {d['tecnico']}.")
        return super().form_valid(form)


class IVTListView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "almacen/ivt_list.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET

        estado = g.get("estado", "").strip()
        tecnico = g.get("tecnico", "").strip()
        desde = g.get("desde", "").strip()
        hasta = g.get("hasta", "").strip()
        q = g.get("q", "").strip()
        fisico = g.get("fisico", "").strip()

        # =====================================================
        # TODOS LOS IVT: SIN PAGINACIÓN
        # =====================================================
        qs = (
            IVT.objects
            .select_related(
                "tecnico",
                "operacion",
                "caja__remision"
            )
            .order_by("numero")
        )

        # =====================================================
        # FILTRO POR ESTADO
        # Si no selecciona estado, muestra TODOS.
        # Incluye DISPONIBLE, ASIGNADO, INSTALADO,
        # ANULADO y PERDIDO.
        # =====================================================
        if estado in IVT.Estado.values:
            qs = qs.filter(estado=estado)

        # =====================================================
        # FILTRO POR TÉCNICO
        # =====================================================
        if tecnico.isdigit():
            qs = qs.filter(tecnico_id=int(tecnico))

        # =====================================================
        # FILTRO POR RANGO IVT
        # =====================================================
        if desde.isdigit():
            qs = qs.filter(numero__gte=int(desde))

        if hasta.isdigit():
            qs = qs.filter(numero__lte=int(hasta))

        # =====================================================
        # FILTRO FÍSICO PERDIDO
        # =====================================================
        if fisico == "1":
            qs = qs.filter(fisico_perdido=True)

        # =====================================================
        # BÚSQUEDA POR ORDEN O REMISIÓN
        # =====================================================
        if q:
            qs = qs.filter(
                Q(operacion__numero_orden__icontains=q) |
                Q(caja__remision__numero_remision__icontains=q)
            )

        # =====================================================
        # CONTADORES POR ESTADO (SIN FILTROS)
        # =====================================================
        conteo = dict(
            IVT.objects
            .values_list("estado")
            .annotate(n=Count("id"))
        )

        resumen = [
            (v, l, conteo.get(v, 0))
            for v, l in IVT.Estado.choices
        ]

        # =====================================================
        # RESUMEN IVT POR TÉCNICO
        # Cuenta todos los estados de cada técnico
        # =====================================================
        por_tecnico = (
            IVT.objects
            .filter(tecnico__isnull=False)
            .values(
                "tecnico_id",
                "tecnico__first_name",
                "tecnico__last_name",
                "tecnico__username",
            )
            .annotate(
                asignados=Count("id", filter=Q(estado=IVT.Estado.ASIGNADO)),
                instalados=Count("id", filter=Q(estado=IVT.Estado.INSTALADO)),
                perdidos=Count("id", filter=Q(estado=IVT.Estado.PERDIDO)),
                anulados=Count("id", filter=Q(estado=IVT.Estado.ANULADO)),
                total=Count("id"),
            )
            .order_by("tecnico__first_name", "tecnico__username")
        )

        # =====================================================
        # QUERYSTRING PARA CONSERVAR FILTROS
        # =====================================================
        params = g.copy()
        params.pop("page", None)

        # =====================================================
        # CONTEXTO
        # page_obj AHORA ES EL QUERYSET COMPLETO
        # NO UTILIZAMOS PAGINATOR
        # =====================================================
        ctx.update({
            "page_obj": qs,

            "resumen": resumen,
            "por_tecnico": por_tecnico,

            "estados": IVT.Estado.choices,
            "tecnicos_lista": tecnicos_queryset(),

            "estado_sel": estado,
            "tecnico_sel": tecnico,
            "desde": desde,
            "hasta": hasta,
            "q": q,
            "fisico_sel": fisico,

            "querystring": params.urlencode(),
        })

        return ctx



class IVTCambiarEstadoView(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN

    def post(self, request):
        ids = [int(i) for i in request.POST.getlist("ids") if i.isdigit()]
        accion = request.POST.get("accion", "")
        motivo = request.POST.get("motivo", "")
        try:
            if accion == "FISICO_PERDIDO":
                n = services.marcar_ivt_fisico_perdido(ids=ids, usuario=request.user, motivo=motivo)
            else:
                n = services.cambiar_estado_ivt(
                    ids=ids, nuevo_estado=accion, usuario=request.user, motivo=motivo)
            messages.success(request, f"{n} IVT actualizados.")
        except services.ErrorInventario as e:
            messages.error(request, str(e))

        siguiente = request.POST.get("next", "")
        if siguiente and url_has_allowed_host_and_scheme(
                siguiente, allowed_hosts={request.get_host()}):
            return redirect(siguiente)
        return redirect("almacen:ivt_lista")


class ValidarIVTJson(RolRequeridoMixin, View):
    roles_permitidos = ROLES_DIGITADOR

    def get(self, request):
        tecnico = get_object_or_404(User, pk=request.GET.get("tecnico"), rol="TECNICO")
        op = Operacion.objects.filter(pk=request.GET.get("operacion") or 0).first()
        try:
            ivt = services.validar_ivt_para_operacion(
                numero=request.GET.get("numero", ""), tecnico=tecnico, operacion=op)
        except services.ErrorInventario as e:
            return JsonResponse({"ok": False, "mensaje": "✗ " + str(e)})
        return JsonResponse({"ok": True, "mensaje": f"✓ IVT {ivt.codigo} asignado a este técnico."})