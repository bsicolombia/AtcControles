from .context_processors import invalidar_burbujas_conciliacion, invalidar_burbuja_sin_ods
from .sin_ods import ordenes_sin_ods_qs, con_cantidades
from django.db.models.functions import Upper, Trim
from . import conciliacion_stock
from .forms import CargaStockForm
from operaciones.models import CargaStock, RegistroStock
from django.db.models.functions import Abs
import logging
from . import conciliacion_inventario
from .forms import CargaInventarioForm
from operaciones.models import CargaInventario, RegistroInventario
from operaciones.actividad import registrar
from operaciones.models import RegistroActividad
from .tokens import nuevo_token, token_valido, consumir_token
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
from .tokens import nuevo_token, token_para_render, consumir_token, devolver_token
from operaciones.models import CargaSAP, Operacion, contexto_tipos
from django.views.generic import (
    CreateView,
    DetailView,
    FormView,
    ListView,
    TemplateView,
    UpdateView,
)
from .config_sap import CODIGOS_EXCLUIDOS_VANTI
from django.views.decorators.cache import never_cache
from django.utils.decorators import method_decorator
from django.shortcuts import  render
from django.utils.dateparse import parse_date
from .models import (Asignacion, Categoria, Inconsistencia, Inventario,
                     MovimientoInventario, Producto, UnidadSerializada,IVT)
from django.urls import reverse
from django.utils.http import url_has_allowed_host_and_scheme
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
from .etiquetas import ETIQUETAS_ERRORES

from django.db.models.functions import Upper

def tecnicos_queryset():
    return (
        User.objects
        .filter(groups__name="TECNICO")
        .select_related("supervisor")
        .order_by("first_name", "last_name", "username")
        .distinct()
    )
    
def _unidades_conciliadas_sap():
    """IDs de unidades que SAP ya reportó y que coinciden con CICB (en cualquier carga)."""
    return set(
        RegistroSAP.objects
        .filter(estado=RegistroSAP.Estado.OK, unidad_serializada__isnull=False)
        .values_list("unidad_serializada_id", flat=True)
    )

def _fijar_pruebas_minimas(operacion, activo):
    datos = dict(operacion.datos_origen or {})
    if activo:
        datos["pruebas_minimas"] = True
    else:
        datos.pop("pruebas_minimas", None)
    if datos != (operacion.datos_origen or {}):
        operacion.datos_origen = datos
        operacion.save(update_fields=["datos_origen", "fecha_modificacion"])

def _fijar_actividad(operacion, actividad):
    """Guarda la actividad de mantenimiento elegida en el formulario."""
    nuevo = actividad.pk if actividad else None
    if operacion.actividad_id != nuevo:
        operacion.actividad = actividad
        operacion.save(update_fields=["actividad", "fecha_modificacion"])
# ============================================================
# UTILIDADES
# ============================================================
class BuscarMixin:
    """Búsqueda ?q=; código exacto y otros campos parciales."""

    paginate_by = None
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
# FUNCIONES AUXILIARES: ASIGNADO, CICB Y SAP
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
    Total instalado según CICB, agrupado por técnico y producto.

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
    Total instalado según CICB, agrupado por producto.
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

        # Total de materiales según los filtros aplicados
        ctx["total_materiales"] = self.get_queryset().count()

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
        # Obtener el inventario sin aplicar primero el filtro de BuscarMixin
        qs = (
            ListView.get_queryset(self)
            .select_related("producto", "producto__categoria")
        )

        q = self.request.GET.get("q", "").strip()

        if q:
            # Buscar por número de remisión en movimientos
            movimientos_remision = MovimientoInventario.objects.filter(
                producto_id=OuterRef("producto_id"),
                remision__numero_remision__icontains=q,
            )

            # Buscar por número de remisión en unidades serializadas
            unidades_remision = UnidadSerializada.objects.filter(
                producto_id=OuterRef("producto_id"),
                remision__numero_remision__icontains=q,
            )

            qs = (
                qs.annotate(
                    tiene_remision_movimiento=Exists(movimientos_remision),
                    tiene_remision_unidad=Exists(unidades_remision),
                )
                .filter(
                    Q(producto__nombre__icontains=q)
                    | Q(producto__codigo_producto__icontains=q)
                    | Q(producto__medida__icontains=q)
                    | Q(producto__categoria__nombre__icontains=q)
                    | Q(tiene_remision_movimiento=True)
                    | Q(tiene_remision_unidad=True)
                )
                .distinct()
            )

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

        request = self.request

        ctx["comparacion_sel"] = request.GET.get(
            "comparacion", ""
        ).strip()

        tecnico_sel = request.GET.get("tecnico", "").strip()

        ctx["tecnicos_lista"] = tecnicos_queryset()
        ctx["tecnico_sel"] = tecnico_sel

        if tecnico_sel.isdigit():

            ctx.update(
                _obtener_contexto_asignaciones(request)
            )

        else:

            five = _five_total_por_producto()
            sap = _sap_total_por_producto()
            asignado = _asignado_total_por_producto()

            # Obtener los inventarios que ya vienen filtrados
            inventarios = list(ctx["inventarios"])

            # Agregar datos de comparación a cada inventario
            for inventario in inventarios:

                producto_id = inventario.producto_id

                inventario.five = five.get(producto_id, 0)
                inventario.sap = sap.get(producto_id, 0)

                inventario.asignado_estado = asignado.get(
                    producto_id, 0
                )

                # NUEVO: Disponible CICB + En poder de técnicos
                inventario.disponible_vanti = (
                    (inventario.cantidad_actual or 0)
                    + (inventario.en_poder_tecnicos or 0)
                )

            CODIGOS_EXCLUIDOS_VANTI = {
                "N01", "N02", "N03", "N04", "N05",
                "N06", "N07", "N08", "N09", "N10",
                "N11", "N12", "N13", "N14", "N15",
                "N16", "U01", "U02", "UHH1",
                "C1", "C2", "CC1", "CC2",
                "B1", "AC", "FA", "FM",
            }

            inventarios_vanti = [
                inv for inv in inventarios
                if (
                    inv.producto.codigo_producto or ""
                ).strip().upper() not in CODIGOS_EXCLUIDOS_VANTI
            ]

            # ---------- TODOS LOS MATERIALES (como antes) ----------
            ctx["total_productos"] = len(inventarios)
            ctx["total_recibido"] = sum(inv.total or 0 for inv in inventarios)
            ctx["total_existencias"] = sum(inv.cantidad_actual or 0 for inv in inventarios)
            ctx["total_danado"] = sum(inv.cantidad_danada or 0 for inv in inventarios)
            ctx["total_asignado"] = sum(inv.asignado_estado or 0 for inv in inventarios)

            # ---------- SOLO VANTI (excluye los códigos de la lista) ----------
            ctx["total_productos_vanti"] = len(inventarios_vanti)
            ctx["total_disponible_vanti"] = sum(
                inv.disponible_vanti for inv in inventarios_vanti
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
@method_decorator(never_cache, name="dispatch")
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

        registrar(
            request,
            RegistroActividad.Accion.INGRESO,
            numero,
            detalle="Ingreso de IVT",
            datos={
                "cajas": len(creadas),
                "total_ivt": total,
                "remision": numero,
            },
        )

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

        registrar(
            self.request, RegistroActividad.Accion.INGRESO, num_orden,
            detalle=d["observacion"][:300],
            datos={
                "producto": producto.nombre_completo,
                "cantidad": d["cantidad"],
                "remision": num_orden,
            },
        )

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
        "producto__codigo_producto",
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
        categoria = self.request.GET.get("categoria", "").strip()

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

        # FILTRO POR CATEGORÍA
        if categoria.isdigit():
            qs = qs.filter(producto__categoria_id=int(categoria))

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
                    resultado = "pendientes"      # instalado en CICB, SAP aún no lo reporta
                else:
                    resultado = "diferentes"      # SAP dice algo que CICB no respalda

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
            .order_by("codigo_producto", "nombre")
        )

        ctx["categorias"] = (
            Categoria.objects
            .filter(
                productos__tipo_inventario=Producto.TipoInventario.SERIALIZADO,
                productos__estado=True,
            )
            .distinct()
            .order_by("nombre")
        )
        ctx["categoria_sel"] = self.request.GET.get("categoria", "").strip()

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

        # CICB y SAP por unidad (0 o 1)
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
        ctx["asignaciones"] = self.object.asignaciones.select_related("tecnico", "operacion")
        ctx["movimientos"] = self.object.movimientos.select_related("usuario")
        ctx["registros_sap"] = (
            RegistroSAP.objects
            .filter(unidad_serializada=self.object)
            .select_related("carga")
            .order_by("-pk")
        )
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

        registrar(
            self.request,
            RegistroActividad.Accion.ESTADO_UNIDAD,
            self.unidad.codigo,
            detalle=form.cleaned_data["motivo"][:300],
            datos={
                "estado_anterior": self.unidad.get_estado_display(),
                "nuevo_estado": form.cleaned_data["estado"],
            },
        )

        messages.success(
            self.request,
            f"Estado de {self.unidad.codigo} actualizado."
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
# CONSOLIDADO DE ASIGNACIONES, CICB Y SAP
# ============================================================
def _obtener_contexto_asignaciones(request):
    """
    Consolidado por técnico y material.

    ASIGNADOS:    material entregado y no devuelto.
    CICB:         cantidad instalada según CICB.
    SAP:          cantidad instalada reportada por SAP.
    DEBE_TENER:   ASIGNADOS - mayor valor entre CICB y SAP.
    DIFERENCIAS:  diferencia absoluta entre CICB y SAP.
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
        (p for p in productos_dict.values() if p is not None),
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
                if tipo == "CICB"
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
            "CICB": [],
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
            filas["CICB"].append(celda(five, "CICB"))

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
@method_decorator(never_cache, name="dispatch")
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

        registrar(
            self.request,
            RegistroActividad.Accion.ENTREGA,
            d["producto"].nombre_completo,
            tecnico=d["tecnico"],
            detalle=d["observaciones"][:300],
            datos={
                "cantidad": d["cantidad"],
                "unidad": (
                    d["unidad"].numero_serie or d["unidad"].codigo
                    if d.get("unidad") else None
                ),
            },
        )

        messages.success(self.request, "Asignación registrada.")
        return super().form_valid(form)
class ValidarOrdenJson(RolRequeridoMixin, View):
    """Dice si una orden ya está registrada (se consulta al salir del campo)."""
    roles_permitidos = ROLES_DIGITADOR

    def get(self, request):
        numero = (request.GET.get("numero") or "").strip().upper()
        if not numero:
            return JsonResponse({"vacio": True})

        op = (Operacion.objects
              .select_related("tecnico", "tipo_operacion")
              .filter(numero_orden=numero).first())
        if not op:
            return JsonResponse({"existe": False, "bloquea": False,
                                 "mensaje": ""})

        previos = (Operacion.Estado.CREADA, Operacion.Estado.PENDIENTE,
                   Operacion.Estado.ASIGNADA)
        registrada = (
            op.estado not in previos
            or op.materiales.filter(estado=Asignacion.Estado.INSTALADA).exists()
            or op.ivts.filter(estado=IVT.Estado.INSTALADO).exists()
        )
        if registrada:
            return JsonResponse({
                "existe": True, "bloquea": True,
                "mensaje": (f"⚠ La orden {numero} ya está registrada "
                            f"({op.get_estado_display()}, técnico {op.tecnico}). "
                            "Si necesita corregirla, solicítelo al administrador."),
            })
        return JsonResponse({
            "existe": True, "bloquea": False,
            "mensaje": f"La orden {numero} existe pero aún está pendiente: se puede registrar.",
        })
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
def _datos_pendientes_tecnico(tecnico, operacion=None):
    """
    Devuelve (cantidad, seriales) para el formulario de instalación.

    cantidad: TODOS los materiales por cantidad activos, con lo que el
              técnico tiene pendiente (0 si no tiene nada asignado).
              Si instala más de lo pendiente, el servicio lo registra
              como inconsistencia.
    seriales: unidades serializadas que el técnico tiene asignadas.

    Si se pasa `operacion` (modo edición), lo ya instalado en esa orden
    cuenta como disponible para que no se marque como excedente.
    """
    pendientes = (
        Asignacion.objects
        .filter(tecnico=tecnico, estado=Asignacion.Estado.ASIGNADA)
        .select_related("producto", "unidad_serializada")
        .order_by("producto__nombre", "producto__medida", "id")
    )

    pend = {}
    seriales = []

    for a in pendientes:
        if a.cantidad_pendiente <= 0:
            continue
        if a.unidad_serializada_id:
            u = a.unidad_serializada
            seriales.append({
                "id": u.pk,
                "serial": u.numero_serie or u.codigo,
                "codigo": u.codigo,
                "producto": a.producto_id,
                "categoria": a.producto.categoria_id,
                "nombre": a.producto.nombre_completo,
            })
        else:
            pend[a.producto_id] = pend.get(a.producto_id, 0) + a.cantidad_pendiente

    if operacion is not None:
        instaladas = (
            Asignacion.objects
            .filter(operacion=operacion, estado=Asignacion.Estado.INSTALADA)
            .select_related("producto", "unidad_serializada")
        )
        for a in instaladas:
            if a.unidad_serializada_id:
                u = a.unidad_serializada
                seriales.append({
                    "id": u.pk,
                    "serial": u.numero_serie or u.codigo,
                    "codigo": u.codigo,
                    "producto": a.producto_id,
                    "categoria": a.producto.categoria_id,
                    "nombre": a.producto.nombre_completo,
                })
            else:
                pend[a.producto_id] = pend.get(a.producto_id, 0) + a.cantidad

    cantidad = [
        {
            "id": p.pk,
            "nombre": p.nombre_completo,
            "pendiente": pend.get(p.pk, 0),
            "asignado": pend.get(p.pk, 0) > 0,
        }
        for p in Producto.objects.filter(
            estado=True,
            tipo_inventario=Producto.TipoInventario.CANTIDAD,
        ).order_by("nombre", "medida")
    ]
    return cantidad, seriales

@method_decorator(never_cache, name="dispatch")
class TecnicoAsignacionesView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_DIGITADOR
    template_name = "almacen/asignacion_instalar.html"

    def _cargar(self):
        self.tecnico = get_object_or_404(
            User.objects.select_related("supervisor").filter(rol="TECNICO"),
            pk=self.kwargs["pk"],
        )

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        self._cargar()

        pend_cantidad, pend_seriales = _datos_pendientes_tecnico(self.tecnico)

        # Si el POST falló, se repinta lo que el usuario había agregado
        payload_previo = _previo_desde_post(self.request) or []

        cancelar_url = (
            self.request.GET.get("next")
            or self.request.POST.get("next")
            or self.request.META.get("HTTP_REFERER")
        )
        if not cancelar_url or not url_has_allowed_host_and_scheme(
            cancelar_url,
            allowed_hosts={self.request.get_host()},
            require_https=self.request.is_secure(),
        ):
            cancelar_url = reverse(
                "almacen:tecnico_asignaciones", kwargs={"pk": self.tecnico.pk},
            )

        ctx["cancelar_url"] = cancelar_url
        ctx["tecnico"] = self.tecnico
        ctx["pend_cantidad"] = pend_cantidad
        ctx["pend_seriales"] = pend_seriales
        ctx["productos_serializados"] = Producto.objects.filter(
            estado=True, tipo_inventario=Producto.TipoInventario.SERIALIZADO
        ).select_related("categoria").order_by("nombre", "medida")
        ctx["tipos_inconsistencia"] = dict(Inconsistencia.Tipo.choices)
        ctx["payload_previo"] = payload_previo
        ctx["estados_instalan"] = sorted(e.value for e in Operacion.ESTADOS_QUE_INSTALAN)
        ctx["form"] = kwargs.get("form") or InstalacionMasivaForm(tecnico=self.tecnico)
        ctx["form_token"] = token_para_render(self.request)
        ctx.update(contexto_tipos())                    # config_tipos, estados_ejecutados, estados_ivt
        return ctx

    def post(self, request, *args, **kwargs):
        self._cargar()

        # 1) Validar TODO sin tocar el token
        form = InstalacionMasivaForm(request.POST, request.FILES, tecnico=self.tecnico)
        if not form.is_valid():
            messages.error(request, "Revise los campos señalados.")
            return self.render_to_response(self.get_context_data(form=form))

        d = form.cleaned_data
        tipo, estado = d["tipo_operacion"], d["estado"]
        instala = tipo.instala(estado)
        pide_ivt = tipo.pide_ivt(estado)
        pruebas_min = bool(d.get("pruebas_minimas"))   

        try:
            lineas = _leer_payload(request)
            if instala and not lineas and not pruebas_min:     # <-- cambia
                raise services.ErrorInventario(
                    "Agregue al menos un material o serial para instalar, "
                    "o marque «Pruebas mínimas».")
            if not instala and lineas:
                raise services.ErrorInventario(
                    "Este tipo de operación / estado no permite instalar material.")
        except services.ErrorInventario as e:
            messages.error(request, str(e))
            return self.render_to_response(self.get_context_data(form=form))

        # 2) Todo correcto: se reclama el token (bloquea el doble envío)
        if not consumir_token(request):
            messages.warning(
                request, "Este formulario ya fue procesado o expiró. No se registró nada nuevo.")
            return redirect("almacen:tecnico_asignaciones", pk=self.tecnico.pk)

        # 3) Guardar. Si falla por cualquier motivo, el token se devuelve
        novedades, piezas = [], 0
        try:
            with transaction.atomic():
                numero = d["numero_orden"] or services.generar_orden_sin_ods(tipo)
                op = services.obtener_o_crear_operacion(
                    numero_orden=numero,
                    cuenta_contrato=d["numero_cuenta_contrato"],
                    tipo=tipo, tecnico=self.tecnico, usuario=request.user,
                    fecha=d["fecha_instalacion"], estado=estado,
                    anomalia=d["anomalia"], observaciones=d["observaciones"],
                    archivo=d.get("archivo_soporte"),
                )
                _fijar_actividad(op, d.get("actividad"))
                _fijar_pruebas_minimas(op, pruebas_min)
                if pide_ivt:
                    n_ivt = services.vincular_ivt_o_reportar(
                        numero=d.get("numero_ivt"), operacion=op,
                        tecnico=self.tecnico, usuario=request.user,
                    )
                    if n_ivt:
                        novedades.append(n_ivt)
                if instala and lineas:
                    novedades += services.sincronizar_materiales_operacion(
                        operacion=op, payload=lineas, usuario=request.user,
                    )
                    piezas = sum(
                        len(l["seriales"] or []) or int(l["cantidad"] or 0) for l in lineas
                    )
        except services.ErrorInventario as e:
            devolver_token(request)
            messages.error(request, str(e))
            return self.render_to_response(self.get_context_data(form=form))
        except Exception:
            devolver_token(request)      # error inesperado: no pierde el formulario
            raise
        invalidar_burbuja_sin_ods()

        registrar(
            request, RegistroActividad.Accion.ORDEN_CREADA, op.numero_orden,
            detalle=d["observaciones"][:300],
            tecnico=self.tecnico,
            datos={
                "tipo": tipo.nombre,
                "estado": estado,
                "piezas": piezas,
                "pruebas_minimas": pruebas_min,
                "ivt": str(d.get("numero_ivt") or ""),
                "inconsistencias": len(novedades),
                "origen": op.origen,
            },
        )

        if op.es_sin_ods:
            messages.info(
                request,
                f"Orden {op.numero_orden} creada SIN ODS. Asígnele la ODS real cuando la tenga "
                "(menú «Órdenes sin ODS»).")
        elif instala:
            messages.success(request, f"Orden {op.numero_orden} registrada ({piezas} pieza(s)).")
        else:
            messages.success(
                request,
                f"Orden {op.numero_orden} registrada como {op.get_estado_display().lower()}.")
        if novedades:
            messages.warning(request, f"⚠ Se registraron {len(novedades)} inconsistencia(s) para revisar.")
        return redirect("almacen:tecnico_asignaciones", pk=self.tecnico.pk)
# ============================================================
# INSTALAR · UNA SOLA ASIGNACIÓN
# ============================================================
@method_decorator(never_cache, name="dispatch")
class AsignacionInstalarView(FormPageMixin, RolRequeridoMixin, FormView):
    roles_permitidos = ROLES_DIGITADOR
    form_class = InstalacionForm
    template_name = "almacen/asignacion_instalar.html"

    titulo = "Registrar instalación"
    boton = "Registrar instalación"

    def dispatch(self, request, *args, **kwargs):
        self.asignacion = get_object_or_404(
            Asignacion.objects.select_related(
                "tecnico", "producto", "producto__categoria",
                "unidad_serializada",
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

        cancelar_url = (
            self.request.GET.get("next")
            or self.request.POST.get("next")
            or self.request.META.get("HTTP_REFERER")
        )

        if not cancelar_url or not url_has_allowed_host_and_scheme(
            cancelar_url,
            allowed_hosts={self.request.get_host()},
            require_https=self.request.is_secure(),
        ):
            cancelar_url = self.get_success_url()

        ctx["cancelar_url"] = cancelar_url
        return ctx

    def form_valid(self, form):
        d = form.cleaned_data
        cantidad = d.get("cantidad")
        tecnico = self.asignacion.tecnico

        try:
            with transaction.atomic():
                op = services.obtener_o_crear_operacion(
                    numero_orden=d["numero_orden"] or services.generar_orden_sin_ods(d["tipo_operacion"]),
                    cuenta_contrato=d["numero_cuenta_contrato"],
                    tipo=d["tipo_operacion"],
                    tecnico=tecnico,
                    usuario=self.request.user,
                    fecha=d["fecha_instalacion"],
                    estado=d.get("estado") or Operacion.Estado.EJECUTADA,
                    anomalia=d.get("anomalia"),
                    observaciones=d.get("observaciones") or "",
                    archivo=d.get("archivo_soporte"),
                )
                _fijar_actividad(op, d.get("actividad"))
                services.instalar(
                    asignacion=self.asignacion,
                    usuario=self.request.user,
                    cantidad=cantidad,
                    fecha_instalacion=d["fecha_instalacion"],
                    operacion=op,
                    observaciones=d.get("observaciones") or "",
                )
        except services.ErrorInventario as e:
            form.add_error(None, str(e))
            return self.form_invalid(form)
        invalidar_burbuja_sin_ods()

        if cantidad:
            self.asignacion.refresh_from_db()
            restan = self.asignacion.cantidad_pendiente
            messages.success(
                self.request,
                f"Instaladas {cantidad} × {self.asignacion.producto.nombre_completo}. "
                f"Al técnico le quedan {restan} por instalar.",
            )
        else:
            messages.success(self.request, "Instalación registrada.")

        return super().form_valid(form)

import json

def _payload_desde_bd(operacion):
    """
    Lo que hoy tiene la orden, en el formato del formulario:
    lo instalado + lo reportado como inconsistencia (no instalado).
    """
    asignaciones = (
        Asignacion.objects
        .filter(operacion=operacion, estado=Asignacion.Estado.INSTALADA)
        .select_related("producto", "unidad_serializada")
        .order_by("pk")
    )

    seriales = {}      # producto_id -> item
    cantidades = {}    # producto_id -> item

    for a in asignaciones:
        if a.unidad_serializada_id:
            u = a.unidad_serializada
            item = seriales.setdefault(a.producto_id, {
                "producto": a.producto_id,
                "categoria": a.producto.categoria_id,
                "nombre": a.producto.nombre_completo,
                "serializado": True,
                "cantidad": 0,
                "unidades": [],
            })
            item["unidades"].append({
                "id": u.id,
                "serial": u.numero_serie or u.codigo,
            })
        else:
            item = cantidades.setdefault(a.producto_id, {
                "producto": a.producto_id,
                "nombre": a.producto.nombre_completo,
                "serializado": False,
                "cantidad": 0,
                "unidades": [],
            })
            item["cantidad"] += a.cantidad

    # Lo reportado pero NO instalado (inconsistencias pendientes)
    reportadas = (
        Inconsistencia.objects
        .filter(
            operacion=operacion,
            estado=Inconsistencia.Estado.PENDIENTE,
            instalado=False,
            producto__isnull=False,
        )
        .exclude(tipo__in=services.TIPOS_IVT)
        .select_related("producto")
    )

    for n in reportadas:
        p = n.producto
        if n.serial_reportado:
            item = seriales.setdefault(p.pk, {
                "producto": p.pk,
                "categoria": p.categoria_id,
                "nombre": p.nombre_completo,
                "serializado": True,
                "cantidad": 0,
                "unidades": [],
            })
            item["unidades"].append({
                "id": n.unidad_serializada_id,
                "serial": n.serial_reportado,
                "tipo": n.tipo,
                "instala": False,
                "aviso": n.detalle,
            })
        else:
            item = cantidades.setdefault(p.pk, {
                "producto": p.pk,
                "nombre": p.nombre_completo,
                "serializado": False,
                "cantidad": 0,
                "unidades": [],
            })
            item["cantidad"] += n.cantidad

    return list(seriales.values()) + list(cantidades.values())


def _leer_payload(request):
    """
    Lee el JSON enviado por el formulario y lo normaliza al formato que
    espera services.sincronizar_materiales_operacion:
        [{"producto_id": int, "seriales": [str], "cantidad": int}]
    Acepta tanto "producto"/"unidades" como "producto_id"/"seriales".
    """
    crudo = request.POST.get("payload") or request.POST.get("materiales") or "[]"
    try:
        datos = json.loads(crudo)
    except (TypeError, ValueError):
        raise services.ErrorInventario("Los materiales enviados no son válidos.")

    if not isinstance(datos, list):
        raise services.ErrorInventario("Los materiales enviados no son válidos.")

    normalizado = []
    for item in datos:
        if not isinstance(item, dict):
            raise services.ErrorInventario("Los materiales enviados no son válidos.")

        seriales = item.get("seriales")
        if seriales is None:
            seriales = [
                (u.get("serial") if isinstance(u, dict) else u)
                for u in (item.get("unidades") or [])
            ]

        normalizado.append({
            "producto_id": item.get("producto_id", item.get("producto")),
            "seriales": seriales,
            "cantidad": item.get("cantidad") or 0,
        })
    return normalizado
def _previo_desde_post(request):
    crudo = request.POST.get("payload_ui")
    if crudo:
        try:
            datos = json.loads(crudo)
            if isinstance(datos, list):
                return datos          # solo repinta; lo que se guarda sale de "payload"
        except ValueError:
            pass
    """Reconstruye lo que el usuario escribió para repintar el formulario."""
    try:
        payload = _leer_payload(request)
    except services.ErrorInventario:
        return None
    nombres = {
        p.pk: p.nombre_completo
        for p in Producto.objects.filter(
            pk__in=[int(p["producto_id"]) for p in payload
                    if str(p["producto_id"]).isdigit()]
        )
    }
    return [
        {
            "producto": p["producto_id"],
            "nombre": nombres.get(int(p["producto_id"]), "Material"),
            "serializado": bool(p["seriales"]),
            "cantidad": 0 if p["seriales"] else p["cantidad"],
            "unidades": [{"id": None, "serial": s} for s in p["seriales"]],
        }
        for p in payload if str(p["producto_id"]).isdigit()
    ]

@never_cache
@login_required
def editar_instalacion(request, pk):
    operacion = get_object_or_404(
        Operacion.objects.select_related("tecnico", "tipo_operacion"), pk=pk,
    )
    payload_previo = _payload_desde_bd(operacion)
    form = None

    if request.method == "POST":
        form = InstalacionMasivaForm(
            request.POST, request.FILES,
            tecnico=operacion.tecnico, operacion=operacion,
        )
        previo = _previo_desde_post(request)
        if previo is not None:
            payload_previo = previo

        if form.is_valid():
            d = form.cleaned_data
            tipo, estado = d["tipo_operacion"], d["estado"]
            instala = tipo.instala(estado)
            pide_ivt = tipo.pide_ivt(estado)
            pruebas_min = bool(d.get("pruebas_minimas"))

            # Capturar los valores originales ANTES de actualizar.
            antes = {
                "orden": operacion.numero_orden,
                "estado": operacion.estado,
                "tipo": operacion.tipo_operacion.nombre,
                "cuenta": operacion.numero_cuenta_contrato,
            }

            try:
                payload = _leer_payload(request) if instala else []
                if instala and not payload and not pruebas_min:     # <-- cambia
                    raise services.ErrorInventario(
                        "Agregue al menos un material o serial para instalar, "
                        "o marque «Pruebas mínimas»."
                    )

                with transaction.atomic():
                    operacion = services.actualizar_operacion(
                        operacion=operacion,
                        numero_orden=d["numero_orden"],
                        cuenta_contrato=d["numero_cuenta_contrato"],
                        tipo=tipo,
                        estado=estado,
                        anomalia=d["anomalia"],
                        observaciones=d["observaciones"],
                        fecha=d["fecha_instalacion"],
                        archivo=d.get("archivo_soporte"),
                    )
                    _fijar_actividad(operacion, d.get("actividad"))
                    novedades = services.sincronizar_materiales_operacion(
                        operacion=operacion,
                        payload=payload,
                        usuario=request.user,
                    )
                    _fijar_pruebas_minimas(operacion, pruebas_min)

                    ivt_numero = str(d.get("numero_ivt") or "").strip()
                    if pide_ivt and ivt_numero:
                        n_ivt = services.vincular_ivt_o_reportar(
                            numero=ivt_numero,
                            operacion=operacion,
                            tecnico=operacion.tecnico,
                            usuario=request.user,
                        )
                        if n_ivt:
                            novedades.append(n_ivt)
                    else:
                        services.liberar_ivt_operacion(
                            operacion=operacion,
                            usuario=request.user,
                            motivo="IVT quitado al editar la orden",
                        )

            except services.ErrorInventario as e:
                operacion.refresh_from_db()
                messages.error(request, str(e))

            else:
                invalidar_burbuja_sin_ods()
                # Se registra siempre que la actualización haya sido exitosa.
                registrar(
                    request,
                    RegistroActividad.Accion.ORDEN_EDITADA,
                    operacion.numero_orden,
                    detalle=d["observaciones"][:300],
                    tecnico=operacion.tecnico,
                    datos={
                        "antes": antes,
                        "despues": {
                            "orden": operacion.numero_orden,
                            "estado": operacion.estado,
                            "tipo": operacion.tipo_operacion.nombre,
                            "cuenta": operacion.numero_cuenta_contrato,
                        },
                        "novedades": len(novedades),
                    },
                )

                if novedades:
                    messages.warning(
                        request,
                        f"Orden actualizada con {len(novedades)} novedad(es) "
                        "pendiente(s) de revisar.",
                    )
                else:
                    messages.success(request, "Orden actualizada.")

                return redirect("operaciones:detalle", pk=operacion.pk)

        else:
            messages.error(request, "Revise los campos señalados.")

    pend_cantidad, pend_seriales = _datos_pendientes_tecnico(
        operacion.tecnico, operacion=operacion
    )
    ivt_actual = operacion.ivts.filter(
        estado=IVT.Estado.INSTALADO
    ).first()

    if form is None:
        form = InstalacionMasivaForm(
            tecnico=operacion.tecnico,
            operacion=operacion,
            initial={
                "numero_orden": operacion.numero_orden,
                "numero_cuenta_contrato": operacion.numero_cuenta_contrato,
                "tipo_operacion": operacion.tipo_operacion_id,
                "actividad": operacion.actividad_id,
                "estado": operacion.estado,
                "anomalia": operacion.anomalia_id,
                "pruebas_minimas": (operacion.datos_origen or {}).get("pruebas_minimas", False),
                "observaciones": operacion.observaciones,
                "fecha_instalacion": (
                    operacion.fecha_ejecucion or operacion.fecha_operacion
                ),
                "numero_ivt": ivt_actual.numero if ivt_actual else "",
            },
        )

    context = {
        "operacion": operacion,
        "tecnico": operacion.tecnico,
        "payload_previo": payload_previo,
        "modo_edicion": True,
        "pend_cantidad": pend_cantidad,
        "pend_seriales": pend_seriales,
        "productos_serializados": Producto.objects.filter(
            estado=True,
            tipo_inventario=Producto.TipoInventario.SERIALIZADO,
        ).select_related("categoria").order_by("nombre", "medida"),
        "cancelar_url": reverse(
            "operaciones:detalle", kwargs={"pk": operacion.pk}
        ),
        "tipos_inconsistencia": dict(Inconsistencia.Tipo.choices),
        "estados_instalan": sorted(
            e.value for e in Operacion.ESTADOS_QUE_INSTALAN
        ),
        "form": form,
    }
    context.update(contexto_tipos())
    return render(request, "operaciones/registrar_operacion.html", context)


# ============================================================
# DEVOLVER ASIGNACIÓN
# ============================================================
@method_decorator(never_cache, name="dispatch")
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

        total_devuelto = total_solicitado

        registrar(
            request, RegistroActividad.Accion.DEVOLUCION, documento.codigo,
            tecnico=self.tecnico,
            detalle=obs[:300],
            datos={"piezas": total_devuelto},
        )

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

        # Solo se instala si está todo bien, o si se regulariza automáticamente.
        instala = tipo is None or (
            services.REGULARIZAR_AUTOMATICO
            and tipo not in services.TIPOS_SOLO_REPORTE
        )

        return JsonResponse({
            "tipo": tipo or "OK",
            "mensaje": detalle or "Lo tiene asignado.",
            "instala": instala,
            "id": unidad.pk if unidad else None,
            "serial": (unidad.numero_serie or unidad.codigo) if unidad else serial,
            "producto": unidad.producto_id if unidad else None,
            "categoria": unidad.producto.categoria_id if unidad else None,
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
                    {
                        "id": u.id,
                        "texto": u.etiqueta,
                        "serial": u.numero_serie or u.codigo
                    }
                    for u in unidades
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
        u = self.request.user
        ctx["puede_consola"] = (u.is_superuser or tiene_rol(u, ADMINISTRADOR)
                                or tiene_rol(u, COORDINADOR))

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
@method_decorator(never_cache, name="dispatch")
class EntregaCreateView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "almacen/entrega_form.html"

    def get_context_data(self, **kwargs):
        super_ctx = super().get_context_data(**kwargs)

        super_ctx["form"] = kwargs.get("form") or EntregaForm()
        super_ctx["productos"] = Producto.objects.filter(
            estado=True
        ).order_by("nombre", "medida")

        super_ctx["payload_previo"] = kwargs.get("payload", [])
        super_ctx["payload_ivt_previo"] = kwargs.get("payload_ivt", [])

        disp = IVT.objects.filter(estado=IVT.Estado.DISPONIBLE)

        super_ctx["ivt_disponibles"] = disp.count()
        super_ctx["ivt_primero"] = disp.aggregate(m=Min("numero"))["m"] or ""
        super_ctx["ivt_rangos_disp"] = services.rangos_ivt_disponibles()

        super_ctx["seleccionar_ivt"] = (
            self.request.GET.get("tipo") == "ivt"
        )

        return super_ctx
    
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

        rangos_txt = ", ".join(f"{i:05d}-{f:05d}" for i, f in ivt_resueltos)

        registrar(
            request,
            RegistroActividad.Accion.ENTREGA,
            entrega.codigo,
            tecnico=d["tecnico"],
            detalle=d["observaciones"][:300],
            datos={
                "piezas": entrega.total_cantidad,
                "ivt": rangos_txt,
            },
        )

        messages.success(
            request,
            f"Entrega {entrega.codigo} registrada."
            + (f" IVT asignados: {rangos_txt}." if rangos_txt else ""),
        )

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
        from reportlab.lib.pagesizes import LETTER
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table, TableStyle

        e = get_object_or_404(
            EntregaMaterial.objects
            .select_related("tecnico", "entregado_por")
            .prefetch_related("items__producto", "items__unidad"),
            pk=pk,
        )

        es_devolucion = e.tipo == EntregaMaterial.Tipo.DEVOLUCION
        titulo_izq = (
            "DEVOLUCION DE MATERIAL A TECNICOS ATC"
            if es_devolucion
            else "ASIGNACION DE MATERIAL A TECNICOS ATC"
        )
        titulo_der = (
            "DEVOLUCION DE ACCESORIOS A TECNICOS ATC"
            if es_devolucion
            else "ASIGNACION DE ACCESORIOS A TECNICOS ATC"
        )
        etiqueta_fecha = "FECHA DEVOLUCION" if es_devolucion else "FECHA ASIGNACION"

        nombre_tecnico = escape(str(_nombre(e.tecnico)))
        cedula_tecnico = escape(str(_cedula(e.tecnico)))
        nombre_asigna = escape(str(_nombre(e.entregado_por)))
        fecha_txt = e.fecha.strftime("%d/%m/%Y")

        asignados = EntregaExcelView._agrupar_items(e)

        buf = BytesIO()
        # Hoja Carta VERTICAL (portrait), aprovechando al máximo el área imprimible
        doc = SimpleDocTemplate(
            buf,
            pagesize=LETTER,
            title=f"Planilla {e.codigo}",
            leftMargin=0.5 * cm,
            rightMargin=0.5 * cm,
            topMargin=0.4 * cm,
            bottomMargin=0.4 * cm,
        )

        gris_hdr = colors.HexColor("#D9E1F2")
        gris_cat = colors.HexColor("#F2F4F8")

        st_tit = ParagraphStyle(
            "Tit", fontName="Helvetica-Bold", fontSize=12, leading=13.5, alignment=1
        )
        st_sub = ParagraphStyle(
            "Sub", fontName="Helvetica-Bold", fontSize=7.5, leading=8.5, alignment=1
        )
        st_hdr = ParagraphStyle(
            "Hdr", fontName="Helvetica-Bold", fontSize=8.0, leading=9.0
        )
        st_txt = ParagraphStyle(
            "Txt", fontName="Helvetica", fontSize=7.4, leading=8.4
        )
        st_txt_b = ParagraphStyle(
            "TxtB", fontName="Helvetica-Bold", fontSize=7.4, leading=8.4
        )
        st_cat = ParagraphStyle(
            "Cat", fontName="Helvetica-Bold", fontSize=7.5, leading=8.5, alignment=1
        )
        st_legal = ParagraphStyle(
            "Leg", fontName="Helvetica", fontSize=8.0, leading=9.8, alignment=4
        )

        def construir_tabla_bloque(titulo_bloque, categorias, es_derecho=False):
            filas = [
                [
                    "",
                    Paragraph(f"NUMERO PLANILLA<br/>{escape(str(e.codigo))}", st_sub),
                    Paragraph(titulo_bloque, st_tit),
                    "",
                    "",
                ],
                [
                    Paragraph("TECNICO", st_sub),
                    Paragraph(nombre_tecnico, st_hdr),
                    "",
                    Paragraph(etiqueta_fecha, st_sub),
                    Paragraph(fecha_txt, st_sub),
                ],
                [
                    "",
                    Paragraph("CODIGO MATERIAL", st_sub),
                    Paragraph("TEXTO MATERIALES", st_sub),
                    Paragraph("CANTIDAD", st_sub),
                    Paragraph("SERIAL", st_sub),
                ],
            ]

            estilos = [
                ("GRID", (0, 0), (-1, 2), 0.8, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), gris_hdr),
                ("BACKGROUND", (0, 2), (-1, 2), gris_hdr),
                ("SPAN", (2, 0), (4, 0)),
                ("SPAN", (1, 1), (2, 1)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 2.3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.3),
                ("LEFTPADDING", (0, 0), (-1, -1), 4.0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4.0),
            ]

            fila_idx = 3
            for nombre_cat, items_cat in categorias:
                inicio_cat = fila_idx
                for cod, texto in items_cat:
                    info = asignados.get(cod.upper(), {})
                    cant = info.get("cantidad", "")
                    serial_txt = ", ".join(info.get("seriales", []))
                    estilo_fila = st_txt_b if cant else st_txt

                    filas.append([
                        "",
                        Paragraph(escape(cod), estilo_fila),
                        Paragraph(escape(texto), estilo_fila),
                        Paragraph(str(cant) if cant else "", st_txt_b),
                        Paragraph(escape(serial_txt), st_txt),
                    ])
                    estilos.append(("GRID", (1, fila_idx), (4, fila_idx), 0.4, colors.black))
                    estilos.append(("ALIGN", (1, fila_idx), (1, fila_idx), "CENTER"))
                    estilos.append(("ALIGN", (3, fila_idx), (3, fila_idx), "CENTER"))
                    fila_idx += 1

                fin_cat = fila_idx - 1
                filas[inicio_cat][0] = Paragraph(nombre_cat, st_cat)
                if fin_cat > inicio_cat:
                    estilos.append(("SPAN", (0, inicio_cat), (0, fin_cat)))
                estilos.append(("BOX", (0, inicio_cat), (0, fin_cat), 0.8, colors.black))
                estilos.append(("BACKGROUND", (0, inicio_cat), (0, fin_cat), gris_cat))

            if not es_derecho:
                filas.append(["", "", "", "", ""])
                filas.append([
                    "",
                    "",
                    Paragraph("<b>FIRMA:</b> ______________________________________________________", st_hdr),
                    "",
                    "",
                ])
                estilos.append(("SPAN", (2, fila_idx + 1), (4, fila_idx + 1)))
                estilos.append(("TOPPADDING", (2, fila_idx + 1), (4, fila_idx + 1), 6))
            else:
                texto_legal = (
                    f"Yo <b>{nombre_tecnico or '_______________________________'}</b> "
                    f"identificado con numero de cedula <b>{cedula_tecnico or '_______________'}</b> "
                    "de la ciudad _________________ he recibido a partir de la fecha los elementos de trabajo "
                    "consignados en este formato, asumiendo la responsabilidad sobre el uso y cuidado de las "
                    "mismas, entiendo que estas son para uso exclusivo de las actividades dispuestas por "
                    "<b>BSI COLOMBIA S.A.</b> Así mismo autorizo a que sea descontado el valor de los elementos "
                    "de trabajo en caso de pérdida, robo o cuando se presente un daño en las mismas, siempre y "
                    "cuando su causa no sea por defecto de fábrica."
                )
                filas.append([Paragraph(texto_legal, st_legal), "", "", "", ""])
                estilos.append(("SPAN", (0, fila_idx), (4, fila_idx)))
                estilos.append(("BOX", (0, fila_idx), (4, fila_idx + 4), 0.8, colors.black))
                estilos.append(("TOPPADDING", (0, fila_idx), (4, fila_idx), 5))
                estilos.append(("BOTTOMPADDING", (0, fila_idx), (4, fila_idx), 5))

                filas.append([
                    "",
                    Paragraph("<b>FIRMA:</b>", st_hdr),
                    Paragraph("__________________________________________________________", st_txt),
                    "",
                    "",
                ])
                estilos.append(("SPAN", (2, fila_idx + 1), (4, fila_idx + 1)))
                estilos.append(("TOPPADDING", (0, fila_idx + 1), (4, fila_idx + 1), 6))

                filas.append([
                    "",
                    Paragraph("<b>NOMBRE DE QUIEN ASIGNA:</b>", st_txt_b),
                    Paragraph(nombre_asigna or "____________________________________", st_txt),
                    Paragraph("<b>NOMBRE DE QUIEN ASIGNA:</b>", st_txt_b),
                    Paragraph("_____________________________", st_txt),
                ])
                estilos.append(("TOPPADDING", (0, fila_idx + 2), (4, fila_idx + 2), 6))

                filas.append([
                    "",
                    Paragraph("<b>CARGO:</b>", st_txt_b),
                    Paragraph("ALMACENISTA", st_txt),
                    Paragraph("<b>CARGO:</b>", st_txt_b),
                    Paragraph("_____________________________", st_txt),
                ])
                estilos.append(("TOPPADDING", (0, fila_idx + 3), (4, fila_idx + 3), 5))

                filas.append([
                    "",
                    Paragraph("<b>FIRMA:</b>", st_txt_b),
                    Paragraph("____________________________________", st_txt),
                    Paragraph("<b>FIRMA:</b>", st_txt_b),
                    Paragraph("_____________________________", st_txt),
                ])
                estilos.append(("TOPPADDING", (0, fila_idx + 4), (4, fila_idx + 4), 6))
                estilos.append(("BOTTOMPADDING", (0, fila_idx + 4), (4, fila_idx + 4), 6))

            # Ancho total = 20.59 cm (ocupa todo el ancho útil de Carta vertical)
            anchos_col = [1.80 * cm, 2.30 * cm, 9.20 * cm, 1.90 * cm, 5.39 * cm]
            tbl = Table(filas, colWidths=anchos_col)
            tbl.setStyle(TableStyle(estilos))
            return tbl

        # Página 1 (Frente): Materiales | Página 2 (Respaldo): Accesorios
        tabla_izq = construir_tabla_bloque(
            titulo_izq, EntregaExcelView.BLOQUE_IZQUIERDO, es_derecho=False
        )
        tabla_der = construir_tabla_bloque(
            titulo_der, EntregaExcelView.BLOQUE_DERECHO, es_derecho=True
        )

        doc.build([tabla_izq, PageBreak(), tabla_der])

        response = HttpResponse(buf.getvalue(), content_type="application/pdf")
        response["Content-Disposition"] = f'inline; filename="{e.codigo}.pdf"'
        return response
    
class EntregaExcelView(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN

    # Estructura exacta de la planilla izquierda (Materiales)
    BLOQUE_IZQUIERDO = [
        ("MEDIDORES", [
            ("2000005", "Medidor G4,0 Diafragma Izq 6m³/h Digital"),
            ("2000007", "Medidor G6,0 Diafragma Izq 10m³/hDigital"),
            ("2000009", "Medidor G10 Diafragma Izq 16m³/h"),
            ("2000013", "Medidor G40 Diafragma Izq 60m³/h Digital"),
            ("2000063", "Medidor G25 Diafragma Izq 40m³/h Digital"),
            ("2000064", "Medidor G2,5 Diafragma Izq 4m³/h Torn"),
            ("2000905", "Medidor G1,6 Diafragma Der 2,5m³/h Torn"),
            ("2000906", "Medidor G1,6 Diafragma Izq 2,5m³/h Torn"),
            ("2000925", "Medidor G4 Diaf. Izq 6m³/h Digital 152mm"),
            ("2000926", "Medidor G10 Diaf.Izq16m³/h Digital 209mm"),
            ("2000927", "Medidor G16 Diaf.Izq25m³/h Digital 209mm"),
            ("2001006", "Medidor G10 Diaf.Izq16m³/h Digital Slim"),
        ]),
        ("REGULADORES", [
            ("2600017", "Reg unica etapa cortexbaja 6m3/h"),
            ("2600951", "Reg primera etapa C1 0 - 10 m3/h TapaSeg"),
            ("2600952", "Reg primera etapa C1 0 - 40 m3/h TapaSeg"),
            ("2600953", "Reg primera etapa C1 11 - 25m3/h TapaSeg"),
            ("2600954", "Reg primera etapa C2 26 - 40m3/h TapaSeg"),
            ("2600955", "Reg primera etapa C3>81m3/h RP50 TapaSeg"),
            ("2600956", "Reg segunda etapa C1 0 - 10 m3/h TapaSeg"),
            ("2600957", "Reg segunda etapa C1 11 - 20m3/h TapaSeg"),
            ("2600958", "Reg segunda etapa C2 > 31 m3/h TapaSeg"),
            ("2600959", 'Reg unica etapa 0 - 4m3/h ½" h-h TapaSeg'),
            ("2600960", 'Reg unica etapa 0 - 4m3/h ½" h-m26 TapaS'),
            ("2600961", "Reg unica etapa C1 5 - 10 m3/h TapaSeg"),
            ("2600962", "Reg unica etapa C2 11 - 25 m3/h TapaSeg"),
            ("2600963", "Reg unica etapa C3 26 - 40 m3/h TapaSeg"),
            ("2600964", "Reg unica etapa cortexbaja 6m3/h TapaSeg"),
        ]),
        ("VALVULAS", [
            ("3000072", 'Valv bola BR ½"hNPTx½"hNPT maneral corto'),
            ("3000074", 'Valv bola BR ½"mNPTx½"hNPT tuerca m26'),
            ("3000075", 'Valv bola BR ½"mNPTx½"hNPT tuerca m32'),
            ("3000076", 'Valv bola BR ¾"hNPTx¾"hNPT'),
            ("3000078", 'Valv bola BR 1/8"mNPT toma de presion'),
        ]),
        ("CONECTORES", [
            ("3000009", "Conector medidor G1,6 - G2,5"),
            ("3000010", "Conector medidor G1,6 toma de presion"),
            ("3000012", "Conector medidor G4,0 toma de presion"),
            ("3000013", 'Conector medidor rosca 1¼"'),
            ("3000014", 'Conector medidor rosca 1½"'),
            ("3001513", 'Conector medidor G25 G2 ½"'),
            ("3001517", 'Conector medidor G25 G2 ½" toma presion'),
        ]),
        ("ADAPTADORES", [
            ("3000003", 'Adaptador BR ¾" mNPT x m26'),
            ("3000004", 'Adaptador BR elevador ½"m26 x ½"mNPT'),
        ]),
        ("T8", [
            ("T81", "Dispositivo T8 1/2"),
            ("T82", "Dispositivo T8 3/4"),
        ]),
        ("EMPAQUES", [
            ("EM01", "Empaque 3/4"),
            ("EM02", "Empaque residencial"),
            ("EM03", "Empaque universal"),
        ]),
    ]

    # Estructura exacta de la planilla derecha (Accesorios)
    BLOQUE_DERECHO = [
        ("NIPLES", [
            ("N01", 'Niple galv ½"x01"'),
            ("N02", 'Niple galv ½"x01½"'),
            ("N03", 'Niple galv ½"x02"'),
            ("N04", 'Niple galv ½"x03"'),
            ("N05", 'Niple galv ½"x04"'),
            ("N06", 'Niple galv ½"x05"'),
            ("N07", 'Niple galv ½"x06"'),
            ("N08", 'Niple galv ½"x07"'),
            ("N09", 'Niple galv ¾"x01"'),
            ("N10", 'Niple galv ¾"x01½"'),
            ("N11", 'Niple galv ¾"x02"'),
            ("N12", 'Niple galv ¾"x03"'),
            ("N13", 'Niple galv ¾"x04"'),
            ("N14", 'Niple galv ¾"x05"'),
            ("N15", 'Niple galv ¾"x06"'),
            ("N16", 'Niple galv ¾"x07"'),
        ]),
        ("UNIONES", [
            ("U01", 'Union galv ½" roscable'),
            ("U02", 'Union galv ¾" roscable'),
            ("UHH1", 'Universal ½ plana h-h'),
        ]),
        ("CODOS", [
            ("C1", "Codo ½"),
            ("C2", "Codo ¾"),
            ("CC1", "Codo calle ½"),
            ("CC2", "Codo calle ¾"),
        ]),
        ("B", [
            ("B1", 'Bushing galv ¾" mnpt x ½" hnpt'),
        ]),
        ("AC", [
            ("AC", "Anticorrosivo"),
        ]),
        ("PEGANTE", [
            ("FA", "Fuerza alta"),
            ("FM", "Fuerza media"),
        ]),
    ]

    @staticmethod
    def _agrupar_items(entrega):
        """Agrupa cantidades y seriales por código de material."""
        datos = {}
        for it in entrega.items.all():
            cod = str(it.codigo_material or "").strip().upper()
            if not cod:
                continue
            reg = datos.setdefault(cod, {"cantidad": 0, "seriales": []})
            reg["cantidad"] += it.cantidad
            serial = it.rango_ivt if it.es_ivt else (it.serial or "").strip()
            if serial:
                reg["seriales"].append(str(serial))
        return datos

    def get(self, request, pk):
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.worksheet.pagebreak import Break

        e = get_object_or_404(
            EntregaMaterial.objects
            .select_related("tecnico", "entregado_por")
            .prefetch_related("items__producto", "items__unidad"),
            pk=pk,
        )

        es_devolucion = e.tipo == EntregaMaterial.Tipo.DEVOLUCION
        titulo_izq = (
            "DEVOLUCION DE MATERIAL A TECNICOS ATC"
            if es_devolucion
            else "ASIGNACION DE MATERIAL A TECNICOS ATC"
        )
        titulo_der = (
            "DEVOLUCION DE ACCESORIOS A TECNICOS ATC"
            if es_devolucion
            else "ASIGNACION DE ACCESORIOS A TECNICOS ATC"
        )
        etiqueta_fecha = "FECHA DEVOLUCION" if es_devolucion else "FECHA ASIGNACION"
        nombre_tecnico = str(_nombre(e.tecnico))
        cedula_tecnico = str(_cedula(e.tecnico))
        nombre_asigna = str(_nombre(e.entregado_por))
        fecha_txt = e.fecha.strftime("%d/%m/%Y")

        asignados = self._agrupar_items(e)

        wb = Workbook()
        ws = wb.active
        ws.title = "ASIG IMPRIMIR"
        ws.views.sheetView[0].showGridLines = False

        thin = Side(style="thin", color="000000")
        med = Side(style="medium", color="000000")
        borde_celda = Border(left=thin, right=thin, top=thin, bottom=thin)
        borde_caja = Border(left=med, right=med, top=med, bottom=med)

        fill_gris = PatternFill("solid", fgColor="D9E1F2")
        fill_cat = PatternFill("solid", fgColor="F2F4F8")

        f_titulo = Font(name="Calibri", size=18, bold=True)
        f_sub = Font(name="Calibri", size=10, bold=True)
        f_dato_hdr = Font(name="Calibri", size=11, bold=True)
        f_cat = Font(name="Calibri", size=13, bold=True)
        f_item = Font(name="Calibri", size=11, bold=False)
        f_item_bold = Font(name="Calibri", size=11, bold=True)
        f_legal = Font(name="Arial", size=11, bold=False)
        f_firma = Font(name="Calibri", size=12, bold=True)

        def pintar_bloque_encabezado(col_ini, titulo_bloque):
            ws.cell(2, col_ini + 1, f"NUMERO PLANILLA\n{e.codigo}").font = f_sub
            ws.cell(2, col_ini + 1).alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            ws.merge_cells(start_row=2, start_column=col_ini + 2, end_row=2, end_column=col_ini + 8)
            c_tit = ws.cell(2, col_ini + 2, titulo_bloque)
            c_tit.font = f_titulo
            c_tit.alignment = Alignment(horizontal="center", vertical="center")

            for c in range(col_ini, col_ini + 9):
                ws.cell(2, c).border = borde_caja
                ws.cell(2, c).fill = fill_gris

            ws.cell(3, col_ini, "TECNICO").font = f_dato_hdr
            ws.cell(3, col_ini).alignment = Alignment(horizontal="center", vertical="center")
            ws.merge_cells(start_row=3, start_column=col_ini + 1, end_row=3, end_column=col_ini + 3)
            ws.cell(3, col_ini + 1, nombre_tecnico).font = f_dato_hdr
            ws.cell(3, col_ini + 1).alignment = Alignment(horizontal="left", vertical="center")

            ws.merge_cells(start_row=3, start_column=col_ini + 4, end_row=3, end_column=col_ini + 5)
            ws.cell(3, col_ini + 4, etiqueta_fecha).font = f_dato_hdr
            ws.cell(3, col_ini + 4).alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells(start_row=3, start_column=col_ini + 6, end_row=3, end_column=col_ini + 8)
            ws.cell(3, col_ini + 6, fecha_txt).font = f_dato_hdr
            ws.cell(3, col_ini + 6).alignment = Alignment(horizontal="center", vertical="center")

            for c in range(col_ini, col_ini + 9):
                ws.cell(3, c).border = borde_caja

            ws.cell(4, col_ini, "").border = borde_caja
            ws.cell(4, col_ini + 1, "CODIGO MATERIAL").font = f_sub
            ws.cell(4, col_ini + 2, "TEXTO MATERIALES").font = f_sub
            ws.cell(4, col_ini + 3, "CANTIDAD").font = f_sub
            ws.merge_cells(start_row=4, start_column=col_ini + 4, end_row=4, end_column=col_ini + 8)
            ws.cell(4, col_ini + 4, "SERIAL").font = f_sub

            for c in range(col_ini, col_ini + 9):
                cel = ws.cell(4, c)
                cel.border = borde_caja
                cel.fill = fill_gris
                cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        def pintar_filas_materiales(col_ini, categorias):
            fila = 5
            for nombre_cat, items_cat in categorias:
                fila_inicio_cat = fila
                for cod, texto in items_cat:
                    info = asignados.get(cod.upper(), {})
                    cant = info.get("cantidad", "")
                    serial_txt = ", ".join(info.get("seriales", []))

                    c_cod = ws.cell(fila, col_ini + 1, cod)
                    c_txt = ws.cell(fila, col_ini + 2, texto)
                    c_cnt = ws.cell(fila, col_ini + 3, cant if cant else None)

                    ws.merge_cells(
                        start_row=fila, start_column=col_ini + 4,
                        end_row=fila, end_column=col_ini + 8,
                    )
                    c_ser = ws.cell(fila, col_ini + 4, serial_txt or None)

                    c_cod.font = f_item_bold if cant else f_item
                    c_txt.font = f_item_bold if cant else f_item
                    c_cnt.font = f_item_bold
                    c_ser.font = f_item

                    c_cod.alignment = Alignment(horizontal="center", vertical="center")
                    c_txt.alignment = Alignment(horizontal="left", vertical="center")
                    c_cnt.alignment = Alignment(horizontal="center", vertical="center")
                    c_ser.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                    for c in range(col_ini + 1, col_ini + 9):
                        ws.cell(fila, c).border = borde_celda

                    ws.row_dimensions[fila].height = 24
                    fila += 1

                fila_fin_cat = fila - 1
                if fila_fin_cat > fila_inicio_cat:
                    ws.merge_cells(
                        start_row=fila_inicio_cat, start_column=col_ini,
                        end_row=fila_fin_cat, end_column=col_ini,
                    )
                c_cat = ws.cell(fila_inicio_cat, col_ini, nombre_cat)
                c_cat.font = f_cat
                rotacion = 0 if len(nombre_cat) <= 2 else 90
                c_cat.alignment = Alignment(
                    horizontal="center", vertical="center",
                    textRotation=rotacion, wrap_text=True,
                )
                for r_cat in range(fila_inicio_cat, fila_fin_cat + 1):
                    ws.cell(r_cat, col_ini).border = borde_caja
                    ws.cell(r_cat, col_ini).fill = fill_cat

            return fila

        pintar_bloque_encabezado(2, titulo_izq)
        pintar_bloque_encabezado(12, titulo_der)

        ws.row_dimensions[2].height = 34
        ws.row_dimensions[3].height = 28
        ws.row_dimensions[4].height = 28

        pintar_filas_materiales(2, self.BLOQUE_IZQUIERDO)
        pintar_filas_materiales(12, self.BLOQUE_DERECHO)

        # Firma pie izquierdo (fila 53)
        ws.cell(53, 4, "FIRMA:").font = f_firma
        ws.merge_cells("E53:J53")
        ws.cell(53, 5, "____________________________________________________").font = f_item
        ws.row_dimensions[53].height = 24

        # Cláusula legal y firmas en bloque derecho (filas 34 a 50)
        ws.merge_cells("L34:T37")
        texto_legal = (
            f"Yo {nombre_tecnico or '_______________________________'} "
            f"identificado con numero de cedula {cedula_tecnico or '_______________'} "
            "de la ciudad _________________ he recibido a partir de la fecha los elementos de trabajo "
            "consignados en este formato, asumiendo la responsabilidad sobre el uso y cuidado de las "
            "mismas, entiendo que estas son para uso exclusivo de las actividades dispuestas por "
            "BSI COLOMBIA S.A. Así mismo autorizo a que sea descontado el valor de los elementos de "
            "trabajo en caso de pérdida, robo o cuando se presente un daño en las mismas, siempre y "
            "cuando su causa no sea por defecto de fábrica."
        )
        c_legal = ws.cell(34, 12, texto_legal)
        c_legal.font = f_legal
        c_legal.alignment = Alignment(horizontal="justify", vertical="center", wrap_text=True)

        ws.cell(38, 13, "FIRMA:").font = f_firma
        ws.merge_cells("N38:T38")
        ws.cell(38, 14, "____________________________________________________________").font = f_item

        ws.cell(41, 13, "NOMBRE DE QUIEN ASIGNA:").font = f_firma
        ws.cell(42, 14, nombre_asigna or "______________________________________").font = f_item_bold
        ws.cell(43, 13, "CARGO:").font = f_firma
        ws.cell(44, 14, "ALMACENISTA").font = f_item
        ws.cell(46, 13, "FIRMA:").font = f_firma
        ws.cell(49, 13, "_____________________________________________").font = f_item

        ws.cell(41, 17, "NOMBRE DE QUIEN ASIGNA:").font = f_firma
        ws.cell(42, 18, "_______________________________").font = f_item
        ws.cell(43, 17, "CARGO:").font = f_firma
        ws.cell(44, 18, "_______________________________").font = f_item
        ws.cell(46, 17, "FIRMA:").font = f_firma
        ws.cell(49, 17, "_______________________________________").font = f_item

        for r in range(34, 51):
            for c in range(12, 21):
                cel = ws.cell(r, c)
                cel.border = Border(
                    left=med if c == 12 else None,
                    right=med if c == 20 else None,
                    top=med if r == 34 else None,
                    bottom=med if r == 50 else None,
                )

        anchos = {
            "A": 2, "B": 11, "C": 14, "D": 48, "E": 14,
            "F": 12, "G": 12, "H": 12, "I": 12, "J": 12,
            "K": 2,
            "L": 14, "M": 15, "N": 40, "O": 14,
            "P": 12, "Q": 12, "R": 12, "S": 12, "T": 12,
        }
        for col_let, w in anchos.items():
            ws.column_dimensions[col_let].width = w

        # Márgenes estrechos y 2 páginas lado a lado en tamaño Carta
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.25
        ws.page_margins.bottom = 0.25
        ws.col_breaks.append(Break(id=11))
        ws.print_area = "B2:T54"
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 2
        ws.page_setup.fitToHeight = 1

        buf = BytesIO()
        wb.save(buf)

        response = HttpResponse(
            buf.getvalue(),
            content_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )
        response["Content-Disposition"] = f'attachment; filename="{e.codigo}.xlsx"'
        return response


class _ColumnaIVT:
    """Columna ficticia para mostrar el IVT junto a los materiales."""
    pk = "IVT"
    nombre = "Informe de Visita Técnica (IVT)"
    nombre_completo = "Informe de Visita Técnica (IVT)"
    medida = ""
    codigo_producto = "IVT"


COLUMNA_IVT = _ColumnaIVT()


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

            if item.es_ivt:
                clave, columna = "IVT", COLUMNA_IVT
            elif item.producto_id:
                clave, columna = item.producto_id, item.producto
            else:
                continue

            columnas.setdefault(clave, columna)
            cantidades[clave] = cantidades.get(clave, 0) + item.cantidad

        filas.append({
            "entrega": documento,
            "cants": cantidades,
            "total": sum(cantidades.values()),
        })

    # El IVT siempre va de último
    cols = sorted(
        columnas.values(),
        key=lambda p: (p is COLUMNA_IVT, p.nombre or "", p.medida or ""),
    )

    for fila in filas:
        fila["celdas"] = [fila["cants"].get(c.pk, 0) for c in cols]

    totales = [
        sum(fila["cants"].get(c.pk, 0) for fila in filas)
        for c in cols
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

        matriz = _matriz_entregas(self.request)

        tipo_devolucion = EntregaMaterial.Tipo.DEVOLUCION
        tipo_entrega = EntregaMaterial.Tipo.ENTREGA

        filas_entregas = [
            f for f in matriz["filas"]
            if f["entrega"].tipo == tipo_entrega
        ]

        filas_devoluciones = [
            f for f in matriz["filas"]
            if f["entrega"].tipo == tipo_devolucion
        ]

        def calcular_totales(filas, cols):
            totales = [
                sum(
                    fila["cants"].get(
                        "IVT" if c is COLUMNA_IVT else c.pk, 0
                    )
                    for fila in filas
                )
                for c in cols
            ]

            return {
                "totales": totales,
                "total_general": sum(totales),
            }

        ctx.update(matriz)

        ctx["filas_entregas"] = filas_entregas
        ctx["filas_devoluciones"] = filas_devoluciones

        ctx.update({
            "totales_entregas": calcular_totales(
                filas_entregas, matriz["cols"]
            )["totales"],

            "total_general_entregas": calcular_totales(
                filas_entregas, matriz["cols"]
            )["total_general"],

            "totales_devoluciones": calcular_totales(
                filas_devoluciones, matriz["cols"]
            )["totales"],

            "total_general_devoluciones": calcular_totales(
                filas_devoluciones, matriz["cols"]
            )["total_general"],
        })

        ctx["tecnicos_lista"] = tecnicos_queryset()

        return ctx

    def _excel(self, request):
        m = _matriz_entregas(request)

        wb = Workbook()

        encabezado = PatternFill("solid", fgColor="1F3A5F")
        negrita = Font(bold=True)

        tipo_entrega = EntregaMaterial.Tipo.ENTREGA
        tipo_devolucion = EntregaMaterial.Tipo.DEVOLUCION

        columnas = (
            ["Código", "Fecha", "Técnico", "C.C.", "Responsable almacén"]
            + [p.nombre_completo for p in m["cols"]]
            + ["Total"]
        )

        def crear_hoja(nombre, filas, color):
            ws = wb.create_sheet(nombre)

            encabezado_local = PatternFill("solid", fgColor=color)

            for c, texto in enumerate(columnas, 1):
                celda = ws.cell(1, c, texto)
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = encabezado_local
                celda.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

            for r, fila in enumerate(filas, 2):
                e = fila["entrega"]

                valores = [
                    e.codigo,
                    e.fecha.strftime("%d/%m/%Y"),
                    _nombre(e.tecnico),
                    _cedula(e.tecnico),
                    _nombre(e.entregado_por),
                ]

                valores += [v or None for v in fila["celdas"]]
                valores.append(fila["total"])

                for c, valor in enumerate(valores, 1):
                    ws.cell(r, c, valor)

            fila_total = len(filas) + 2

            ws.cell(fila_total, 5, "TOTALES").font = negrita

            totales = [
                sum(
                    fila["cants"].get(
                        "IVT" if col is COLUMNA_IVT else col.pk, 0
                    )
                    for fila in filas
                )
                for col in m["cols"]
            ]

            total_general = sum(totales)

            for c, valor in enumerate(
                totales + [total_general],
                6,
            ):
                celda = ws.cell(fila_total, c, valor)
                celda.font = negrita

            ws.freeze_panes = "F2"
            ws.sheet_view.showGridLines = False

            for c in range(1, len(columnas) + 1):
                ws.column_dimensions[
                    get_column_letter(c)
                ].width = 14 if c > 5 else 22

            ws.auto_filter.ref = (
                f"A1:{get_column_letter(len(columnas))}{max(1, len(filas) + 1)}"
            )

            return ws

        # Eliminar hoja inicial
        ws = wb.active
        wb.remove(ws)

        filas_entregas = [
            f for f in m["filas"]
            if f["entrega"].tipo == tipo_entrega
        ]

        filas_devoluciones = [
            f for f in m["filas"]
            if f["entrega"].tipo == tipo_devolucion
        ]

        crear_hoja(
            "Entregas",
            filas_entregas,
            "1F3A5F",
        )

        crear_hoja(
            "Devoluciones",
            filas_devoluciones,
            "8B4513",
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

            registrar(
                request,
                RegistroActividad.Accion.INCONSISTENCIA,
                i.operacion.numero_orden if i.operacion else "",
                tecnico=i.tecnico,
                detalle=i.nota_resolucion[:300],
                datos={
                    "tipo": i.tipo,
                    "estado": i.estado,
                    "resuelto_por": request.user.username,
                },
            )

            messages.success(
                request,
                "Inconsistencia marcada como resuelta."
            )

        return redirect(
            request.POST.get("next") or "almacen:inconsistencias"
        )
    
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

        registrar(
            self.request,
            RegistroActividad.Accion.INGRESO,
            d["numero_orden"],
            detalle="Ingreso de IVT",
            datos={
                "cajas": len(cajas),
                "total_ivt": total,
                "remision": d["numero_orden"],
            },
        )

        messages.success(
            self.request,
            f"Ingresaron {total} IVT en {len(cajas)} caja(s) con la remisión {d['numero_orden']}.",
        )

        return redirect(
            f"{reverse('almacen:ivt_lista')}?q={d['numero_orden']}"
        )

@method_decorator(never_cache, name="dispatch")
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
                tecnico=d["tecnico"],
                inicial=d["consecutivo_inicial"],
                final=d["consecutivo_final"],
                usuario=self.request.user,
                fecha=d["fecha"],
                observacion=d["observaciones"],
            )
        except services.ErrorInventario as e:
            form.add_error(None, str(e))
            return self.form_invalid(form)

        registrar(
            self.request,
            RegistroActividad.Accion.IVT,
            str(d["tecnico"]),
            tecnico=d["tecnico"],
            detalle=d["observaciones"][:300],
            datos={
                "desde": d["consecutivo_inicial"],
                "hasta": d["consecutivo_final"],
                "total": n,
            },
        )

        messages.success(
            self.request,
            f"{n} IVT asignados a {d['tecnico']}."
        )
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
                n = services.marcar_ivt_fisico_perdido(
                    ids=ids,
                    usuario=request.user,
                    motivo=motivo,
                )
            else:
                n = services.cambiar_estado_ivt(
                    ids=ids,
                    nuevo_estado=accion,
                    usuario=request.user,
                    motivo=motivo,
                )

            registrar(
                request,
                RegistroActividad.Accion.IVT,
                accion,
                detalle=motivo[:300],
                datos={
                    "cantidad": n,
                    "accion": accion,
                    "ids": ids,
                },
            )

            messages.success(request, f"{n} IVT actualizados.")

        except services.ErrorInventario as e:
            messages.error(request, str(e))

        siguiente = request.POST.get("next", "")
        if siguiente and url_has_allowed_host_and_scheme(
            siguiente, allowed_hosts={request.get_host()}
        ):
            return redirect(siguiente)

        return redirect("almacen:ivt_lista")


class ValidarIVTJson(RolRequeridoMixin, View):
    roles_permitidos = ROLES_DIGITADOR

    def get(self, request):
        tecnico = get_object_or_404(User, pk=request.GET.get("tecnico"), rol="TECNICO")

        numero = (request.GET.get("numero") or "").strip()
        if not numero:
            return JsonResponse({"ok": False, "mensaje": "Escriba el número del IVT."})

        op_id = request.GET.get("operacion") or ""
        op = Operacion.objects.filter(pk=op_id).first() if op_id.isdigit() else None

        tipo, ivt, detalle = services.clasificar_ivt(
            numero=numero, tecnico=tecnico, operacion=op,
        )

        if tipo is None:
            return JsonResponse({
                "ok": True,
                "mensaje": f"✓ IVT {ivt.codigo} asignado a este técnico.",
            })

        return JsonResponse({
            "ok": False,
            "inconsistencia": True,
            "tipo": tipo,
            "mensaje": f"⚠ {detalle} La orden se registrará con una inconsistencia.",
        })

from django.db.models import Count  # ya lo tienes
from operaciones.models import TipoOperacion

from .sin_ods import ordenes_sin_ods_qs, con_cantidades

class OrdenesSinOdsView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_DIGITADOR
    template_name = "almacen/ordenes_sin_ods.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        q, tecnico = g.get("q", "").strip(), g.get("tecnico", "").strip()

        qs = (
            con_cantidades(ordenes_sin_ods_qs())
            .select_related("tecnico", "tipo_operacion__tipo_con_ods")
            .order_by("-fecha_operacion", "-id")
        )
        if tecnico.isdigit():
            qs = qs.filter(tecnico_id=int(tecnico))
        if q:
            qs = qs.filter(Q(numero_orden__icontains=q)
                           | Q(numero_cuenta_contrato__icontains=q))

        ordenes = list(qs)      # una sola consulta; los totales se calculan en Python

        por_tecnico = {}
        for o in ordenes:
            t = por_tecnico.setdefault(o.tecnico_id, {
                "id": o.tecnico_id,
                "nombre": o.tecnico.get_full_name() or o.tecnico.username,
                "ordenes": 0, "piezas": 0,
            })
            t["ordenes"] += 1
            t["piezas"] += o.piezas

        ctx.update(
            ordenes=ordenes,
            total=len(ordenes),
            total_piezas=sum(o.piezas for o in ordenes),
            sin_material=sum(1 for o in ordenes if not o.piezas),
            resumen_tecnicos=sorted(por_tecnico.values(),
                                    key=lambda t: (-t["ordenes"], t["nombre"])),
            q=q, tecnico_sel=tecnico, tecnicos_lista=tecnicos_queryset(),
        )
        return ctx


class AsignarOdsView(RolRequeridoMixin, View):
    """Reemplaza RSO0001 por la ODS real y pasa la orden a su tipo definitivo."""
    roles_permitidos = ROLES_DIGITADOR

    def post(self, request, pk):
        op = get_object_or_404(
            Operacion.objects.select_related("tipo_operacion__tipo_con_ods"), pk=pk
        )
        nueva = request.POST.get("numero_orden", "").strip().upper()

        if not op.es_sin_ods:
            messages.error(request, "Esta orden ya tiene ODS real.")
            return redirect("almacen:ordenes_sin_ods")
        if not nueva or len(nueva) > 10:
            messages.error(request, "Escriba la ODS (máximo 10 caracteres).")
            return redirect("almacen:ordenes_sin_ods")
        if Operacion.objects.filter(numero_orden=nueva).exclude(pk=op.pk).exists():
            messages.error(request, f"La ODS {nueva} ya está registrada en otra orden.")
            return redirect("almacen:ordenes_sin_ods")

        anterior = op.numero_orden
        tipo_anterior = op.tipo_operacion
        tipo_nuevo = tipo_anterior.tipo_con_ods
        campos = ["numero_orden", "fecha_modificacion"]

        if tipo_nuevo:
            # El estado actual debe ser válido para el tipo nuevo
            controlados = Operacion.ESTADOS_EJECUTADOS | Operacion.ESTADOS_NEGATIVOS
            validos = {e.value for e in tipo_nuevo.estados_validos()}
            if op.estado in controlados and op.estado not in validos:
                messages.error(
                    request,
                    f"«{tipo_nuevo.nombre}» no admite el estado "
                    f"{op.get_estado_display()}. Edite la orden primero.",
                )
                return redirect("almacen:ordenes_sin_ods")
            op.tipo_operacion = tipo_nuevo
            campos.append("tipo_operacion")

        op.numero_orden = nueva
        op.save(update_fields=campos)
        invalidar_burbuja_sin_ods()

        registrar(
            request, RegistroActividad.Accion.ODS_ASIGNADA, nueva,
            detalle=f"{anterior} → {nueva}",
            tecnico=op.tecnico,
            datos={
                "ods_anterior": anterior,
                "ods_nueva": nueva,
                "tipo_anterior": tipo_anterior.nombre,
                "tipo_nuevo": op.tipo_operacion.nombre,
            },
        )

        msg = f"{anterior} ahora es la ODS {nueva}."
        if tipo_nuevo:
            msg += f" Tipo: {tipo_nuevo.nombre}."
        else:
            messages.warning(
                request,
                f"El tipo «{tipo_anterior.nombre}» no tiene configurado su tipo con ODS; "
                "se dejó igual.",
            )
        messages.success(request, msg)
        return redirect("almacen:ordenes_sin_ods")

class ConsolaView(RolRequeridoMixin, ListView):
    """Registro de todo lo que hace cada usuario. Solo administrador y coordinador."""
    roles_permitidos = (ADMINISTRADOR, COORDINADOR)
    model = RegistroActividad
    template_name = "almacen/consola.html"
    context_object_name = "registros"
    paginate_by = 50

    def get_queryset(self):
        g = self.request.GET
        qs = RegistroActividad.objects.select_related("usuario", "tecnico")

        if g.get("usuario", "").isdigit():
            qs = qs.filter(usuario_id=int(g["usuario"]))
        if g.get("tecnico", "").isdigit():
            qs = qs.filter(tecnico_id=int(g["tecnico"]))
        if g.get("accion") in RegistroActividad.Accion.values:
            qs = qs.filter(accion=g["accion"])

        desde, hasta = parse_date(g.get("desde", "") or ""), parse_date(g.get("hasta", "") or "")
        if desde:
            qs = qs.filter(fecha__date__gte=desde)
        if hasta:
            qs = qs.filter(fecha__date__lte=hasta)

        q = g.get("q", "").strip()
        if q:
            qs = qs.filter(Q(referencia__icontains=q) | Q(detalle__icontains=q)
                           | Q(usuario_nombre__icontains=q))
        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        p = g.copy(); p.pop("page", None)
        ctx.update(
            querystring=p.urlencode(),
            usuarios=User.objects.filter(actividades__isnull=False).distinct()
                         .order_by("first_name", "username"),
            tecnicos_lista=tecnicos_queryset(),
            acciones=RegistroActividad.Accion.choices,
            usuario_sel=g.get("usuario", ""), tecnico_sel=g.get("tecnico", ""),
            accion_sel=g.get("accion", ""), desde=g.get("desde", ""),
            hasta=g.get("hasta", ""), q=g.get("q", ""),
        )
        return ctx

class CargaInventarioView(FormPageMixin, RolRequeridoMixin, FormView):
    roles_permitidos = ROLES_ALMACEN
    form_class = CargaInventarioForm
    template_name = "almacen/carga_inventario_form.html"
    titulo = "Carga IQ09 · medidores"
    subtitulo = "Compara el IQ09 (todos los medidores, en cualquier estado) contra el sistema."
    boton = "Cargar y comparar"
    cancel_url = reverse_lazy("almacen:inventario")

    def form_valid(self, form):
        d = form.cleaned_data
        try:
            carga = conciliacion_inventario.procesar_carga(
                archivo=d["archivo"], categoria=d["categoria"],
                usuario=self.request.user, observaciones=d["observaciones"],
            )
        except conciliacion_inventario.ErrorArchivo as exc:
            form.add_error("archivo", str(exc))
            return self.form_invalid(form)
        except Exception:
            logging.getLogger(__name__).exception("Error procesando carga de inventario")
            form.add_error("archivo", "No fue posible procesar el archivo. "
                                      "Verifique su contenido e inténtelo nuevamente.")
            return self.form_invalid(form)

        if carga.archivo_repetido:
            messages.warning(self.request, "Este mismo archivo ya se había cargado antes.")

        resumen = (f"Carga #{carga.pk}: {carga.filas_ok} coinciden, {carga.filas_error} con error "
                   f"y {carga.faltantes_archivo} disponibles en sistema que no vinieron en el archivo.")
        (messages.success if carga.coincide_todo else messages.warning)(self.request, resumen)

        registrar(self.request, RegistroActividad.Accion.CARGA_INVENTARIO, f"Carga inv. #{carga.pk}",
                  detalle=resumen,
                  datos={"categoria": str(carga.categoria), "archivo": carga.nombre_original,
                         "ok": carga.filas_ok, "error": carga.filas_error,
                         "faltantes": carga.faltantes_archivo})
        
        invalidar_burbujas_conciliacion()
        return redirect("almacen:carga_inventario_detalle", pk=carga.pk)


@login_required
def descargar_plantilla_inventario(request):
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"
    ws.append(["numero_serie", "codigo_producto", "codigo_unidad"])
    ws.append(["A1234567", "MEDf-G4", ""])
    ws.append(["A1234568", "MED-G4", "DET-00012"])
    resp = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp["Content-Disposition"] = 'attachment; filename="plantilla_inventario.xlsx"'
    wb.save(resp)
    return resp


class HistorialCargaInventarioView(RolRequeridoMixin, ListView):
    roles_permitidos = ROLES_ALMACEN
    model = CargaInventario
    template_name = "almacen/carga_inventario_historial.html"
    context_object_name = "cargas"
    paginate_by = 25

    def get_queryset(self):
        qs = CargaInventario.objects.select_related("usuario", "categoria")
        g = self.request.GET
        if g.get("categoria", "").isdigit():
            qs = qs.filter(categoria_id=int(g["categoria"]))
        if g.get("desde"):
            qs = qs.filter(fecha__date__gte=g["desde"])
        if g.get("hasta"):
            qs = qs.filter(fecha__date__lte=g["hasta"])
        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        p = g.copy(); p.pop("page", None)
        ctx.update(querystring=p.urlencode(), categorias=Categoria.objects.exclude(prefijo=""),
                   categoria_sel=g.get("categoria", ""), desde=g.get("desde", ""), hasta=g.get("hasta", ""))
        return ctx

class CargaInventarioDetalleView(RolRequeridoMixin, DetailView):
    roles_permitidos = ROLES_ALMACEN
    model = CargaInventario
    template_name = "almacen/carga_inventario_detalle.html"
    context_object_name = "carga"

    MAPA_VER = {
        "ok": RegistroInventario.Resultado.OK,
        "errores": RegistroInventario.Resultado.ERROR,
        "faltantes": RegistroInventario.Resultado.FALTANTE,
    }

    def _ver(self):
        ver = self.request.GET.get("ver", "")
        if ver in self.MAPA_VER:
            return ver
        c = self.object
        if c.filas_error:
            return "errores"
        if c.faltantes_archivo:
            return "faltantes"
        return "ok"

    def _registros(self, ver=None):
        g = self.request.GET
        qs = self.object.registros.select_related("producto", "unidad__producto")
        if ver in self.MAPA_VER:
            qs = qs.filter(resultado=self.MAPA_VER[ver])
        tipo = g.get("tipo_error", "").strip()
        if tipo:
            qs = qs.filter(tipo_error=tipo)
        q = g.get("q", "").strip()
        if q:
            qs = qs.filter(
                Q(numero_serie__icontains=q) | Q(codigo_producto__icontains=q)
                | Q(descripcion__icontains=q) | Q(detalle__icontains=q)
                | Q(orden_sap__icontains=q)
            )
        return qs

    def get(self, request, *args, **kwargs):
        self.object = self.get_object()
        if request.GET.get("export") == "xlsx":
            return self._excel()
        return super().get(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        ver = self._ver()

        # Obtener TODOS los registros, sin paginación
        filas = list(self._registros(ver))

        for r in filas:
            r.accion = conciliacion_inventario.ACCIONES.get(
                r.tipo_error, ""
            )

        etiquetas = dict(RegistroInventario.TipoError.choices)

        resumen_errores = [
            {
                "codigo": x["tipo_error"],
                "label": etiquetas.get(
                    x["tipo_error"],
                    x["tipo_error"]
                ),
                "n": x["n"],
            }
            for x in (
                self.object.registros
                .filter(resultado=RegistroInventario.Resultado.ERROR)
                .values("tipo_error")
                .annotate(n=Count("id"))
                .order_by("-n")
            )
        ]

        tipo_error = g.get("tipo_error", "").strip()
        q = g.get("q", "").strip()

        qs_base = urlencode({
            k: v
            for k, v in (
                ("ver", ver),
                ("tipo_error", tipo_error),
                ("q", q),
            )
            if v
        })

        en_archivo = self.object.registros.exclude(resultado=RegistroInventario.Resultado.FALTANTE)
        nombres_estado = dict(UnidadSerializada.Estado.choices)

        resumen_estados = [
            {"label": nombres_estado.get(x["estado_sistema"], x["estado_sistema"] or "No existe en CICB"),
             "n": x["n"]}
            for x in en_archivo.values("estado_sistema").annotate(n=Count("id")).order_by("-n")
        ]
        resumen_status = [
            {"status": x["status_sap"] or "(vacío)", "n": x["n"]}
            for x in en_archivo.values("status_sap").annotate(n=Count("id")).order_by("-n")
        ]

        ctx.update(
            ver=ver,
            filas=filas,
            total_filtrado=len(filas),
            resumen_errores=resumen_errores,
            tipo_error_sel=tipo_error,
            q=q,
            qs_base=qs_base,
        )

        return ctx

    def _excel(self):
        c = self.object
        ver = self.request.GET.get("ver", "")
        wb = Workbook()
        ws = wb.active
        ws.title = "Comparación"
        ws.append(["Fila", "Resultado", "Problema", "Serie", "Material", "Descripción",
                   "Status SAP", "Orden SAP", "Estado en sistema", "Detalle", "Qué hacer"])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F3A5F")
        for r in self._registros(ver):
            ws.append([
                r.fila, r.get_resultado_display(),
                r.get_tipo_error_display() if r.tipo_error else "",
                r.numero_serie, r.codigo_producto, r.descripcion, r.status_sap, r.orden_sap,
                r.unidad.get_estado_display() if r.unidad else "",
                r.detalle, conciliacion_inventario.ACCIONES.get(r.tipo_error, ""),
            ])
        ws.freeze_panes = "A2"
        for col, ancho in zip("ABCDEFGHIJK", (7, 14, 34, 16, 12, 34, 11, 14, 18, 50, 50)):
            ws.column_dimensions[col].width = ancho
        resp = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        resp["Content-Disposition"] = f'attachment; filename="comparacion_inventario_{c.pk}.xlsx"'
        wb.save(resp)
        return resp

class CargaStockView(FormPageMixin, RolRequeridoMixin, FormView):
    roles_permitidos = ROLES_ALMACEN
    form_class = CargaStockForm
    template_name = "almacen/carga_stock_form.html"
    titulo = "Carga MB52 · stock por material"
    subtitulo = "Compara el stock que SAP tiene en el almacén contra lo que CICB espera tener."
    boton = "Cargar y comparar"
    cancel_url = reverse_lazy("almacen:inventario")

    def form_valid(self, form):
        d = form.cleaned_data
        try:
            carga = conciliacion_stock.procesar_stock(
                archivo=d["archivo"], usuario=self.request.user,
                observaciones=d["observaciones"],
            )
        except conciliacion_stock.ErrorArchivo as exc:
            form.add_error("archivo", str(exc))
            return self.form_invalid(form)
        except Exception:
            logging.getLogger(__name__).exception("Error procesando MB52")
            form.add_error("archivo", "No fue posible procesar el archivo. "
                                      "Verifique su contenido e inténtelo nuevamente.")
            return self.form_invalid(form)

        if carga.archivo_repetido:
            messages.warning(self.request, "Este mismo archivo ya se había cargado antes.")

        resumen = (f"MB52 #{carga.pk}: {carga.coinciden} materiales coinciden y "
                   f"{carga.con_diferencia} con diferencia.")
        (messages.success if carga.coincide_todo else messages.warning)(self.request, resumen)

        registrar(self.request, RegistroActividad.Accion.CARGA_INVENTARIO, f"MB52 #{carga.pk}",
                  detalle=resumen,
                  datos={"archivo": carga.nombre_original, "ok": carga.coinciden,
                         "diferencias": carga.con_diferencia})
        invalidar_burbujas_conciliacion()
        return redirect("almacen:carga_stock_detalle", pk=carga.pk)
_REGEX_EXCLUIDOS = r"^\s*(" + "|".join(sorted(CODIGOS_EXCLUIDOS_VANTI)) + r")\s*$"

class HistorialCargaStockView(RolRequeridoMixin, ListView):
    roles_permitidos = ROLES_ALMACEN
    model = CargaStock
    template_name = "almacen/carga_stock_historial.html"
    context_object_name = "cargas"
    paginate_by = 25

    def get_queryset(self):
        R = RegistroStock
        # Mismos criterios que CargaStockDetalleView._base()
        vanti = ~Q(_cod__in=CODIGOS_EXCLUIDOS_VANTI)

        return (
            CargaStock.objects
            .select_related("usuario")
            .annotate(
                total_vanti=Count(
                    "registros",
                    filter=~Q(registros__codigo_producto__iregex=r"^\s*(" + "|".join(sorted(CODIGOS_EXCLUIDOS_VANTI)) + r")\s*$"),
                ),
                ok_vanti=Count(
                    "registros",
                    filter=Q(registros__resultado=R.Resultado.OK)
                    & ~Q(registros__codigo_producto__iregex=r"^\s*(" + "|".join(sorted(CODIGOS_EXCLUIDOS_VANTI)) + r")\s*$"),
                ),
                dif_vanti=Count(
                    "registros",
                    filter=Q(registros__resultado=R.Resultado.ERROR)
                    & ~Q(registros__tipo_error=R.TipoError.NO_EN_MB52)
                    & ~Q(registros__codigo_producto__iregex=_REGEX_EXCLUIDOS),
                ),
                faltan_vanti=Count(
                    "registros",
                    filter=Q(registros__tipo_error=R.TipoError.NO_EN_MB52)
                    & ~Q(registros__codigo_producto__iregex=_REGEX_EXCLUIDOS),
                ),
            )
            .order_by("-fecha", "-pk")
        )
class CargaStockDetalleView(RolRequeridoMixin, DetailView):
    roles_permitidos = ROLES_ALMACEN
    model = CargaStock
    template_name = "almacen/carga_stock_detalle.html"
    context_object_name = "carga"

    FALTANTE = RegistroStock.TipoError.NO_EN_MB52

    def _base(self):
        """Registros de la carga SIN los códigos excluidos (solo se compara VANTI)."""
        return (
            self.object.registros
            .annotate(_cod=Upper(Trim("codigo_producto")))
            .exclude(_cod__in=CODIGOS_EXCLUIDOS_VANTI)
        )

    def _ver(self):
        ver = self.request.GET.get("ver", "")
        if ver in ("ok", "errores", "faltantes", "todos"):
            return ver
        regs = self._base()
        if (regs.filter(resultado=RegistroStock.Resultado.ERROR)
                .exclude(tipo_error=self.FALTANTE).exists()):
            return "errores"
        if regs.filter(tipo_error=self.FALTANTE).exists():
            return "faltantes"
        return "ok"

    def _registros(self, ver):
        g = self.request.GET
        qs = self._base().select_related("producto")
        if ver == "ok":
            qs = qs.filter(resultado=RegistroStock.Resultado.OK)
        elif ver == "errores":
            qs = (qs.filter(resultado=RegistroStock.Resultado.ERROR)
                    .exclude(tipo_error=self.FALTANTE))
        elif ver == "faltantes":
            qs = qs.filter(tipo_error=self.FALTANTE)
        if g.get("tipo_error", "").strip():
            qs = qs.filter(tipo_error=g["tipo_error"].strip())
        q = g.get("q", "").strip()
        if q:
            qs = qs.filter(Q(codigo_producto__icontains=q) | Q(descripcion__icontains=q))
        return qs.annotate(_abs=Abs("diferencia")).order_by("-_abs", "codigo_producto")

    def get(self, request, *args, **kwargs):
        self.object = self.get_object()
        if request.GET.get("export") == "xlsx":
            return self._excel()
        return super().get(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        ver = self._ver()

        filas = list(self._registros(ver))
        for r in filas:
            r.accion = conciliacion_stock.ACCIONES.get(r.tipo_error, "")

        base = self._base()
        etiquetas = dict(RegistroStock.TipoError.choices)
        resumen_errores = [
            {"codigo": x["tipo_error"],
             "label": etiquetas.get(x["tipo_error"], x["tipo_error"]),
             "n": x["n"]}
            for x in (base
                      .filter(resultado=RegistroStock.Resultado.ERROR)
                      .exclude(tipo_error=self.FALTANTE)
                      .values("tipo_error").annotate(n=Count("id")).order_by("-n"))
        ]

        # Contadores recalculados SOLO con materiales VANTI
        n_ok = base.filter(resultado=RegistroStock.Resultado.OK).count()
        n_faltantes = base.filter(tipo_error=self.FALTANTE).count()
        n_errores = (base.filter(resultado=RegistroStock.Resultado.ERROR)
                         .exclude(tipo_error=self.FALTANTE).count())
        n_total = base.count()

        tipo_error, q = g.get("tipo_error", "").strip(), g.get("q", "").strip()

        ctx.update(
            ver=ver, filas=filas, total_filtrado=len(filas),
            resumen_errores=resumen_errores, tipo_error_sel=tipo_error, q=q,
            n_ok=n_ok, n_errores=n_errores, n_faltantes=n_faltantes,
            n_total=n_total,
            todo_coincide=(n_errores == 0 and n_faltantes == 0),
            qs_base=urlencode({k: v for k, v in (
                ("ver", ver), ("tipo_error", tipo_error), ("q", q)) if v}),
        )
        return ctx


    def _excel(self):
        c = self.object
        ver = self._ver()                      # respeta la pestaña activa
        wb = Workbook()
        ws = wb.active
        ws.title = "MB52 vs CICB"
        ws.append(["Material", "Descripción", "SAP libre", "CICB disponible", "CICB asignado",
                   "CICB instalado sin ODS", "CICB total", "Diferencia (SAP - CICB)",
                   "SAP calidad", "SAP bloqueado", "SAP tránsito", "Problema", "Qué hacer"])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F3A5F")
        for r in self._registros(ver):
            ws.append([
                r.codigo_producto, r.descripcion, r.sap_libre, r.five_disponible, r.five_asignado,
                r.five_sin_ods, r.five_total, r.diferencia, r.sap_calidad, r.sap_bloqueado,
                r.sap_transito,
                r.get_tipo_error_display() if r.tipo_error else "Coincide",
                conciliacion_stock.ACCIONES.get(r.tipo_error, ""),
            ])
        ws.freeze_panes = "A2"
        for col, ancho in zip("ABCDEFGHIJKLM", (14, 40, 11, 14, 13, 18, 11, 18, 11, 12, 12, 38, 60)):
            ws.column_dimensions[col].width = ancho
        resp = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        resp["Content-Disposition"] = f'attachment; filename="mb52_vs_cicb_{c.pk}_{ver}.xlsx"'
        wb.save(resp)
        return resp