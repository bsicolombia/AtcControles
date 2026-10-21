import logging
from collections import defaultdict
from almacen.views import FormPageMixin
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Count, Q
from django.http import HttpResponse, HttpResponseRedirect
from django.urls import reverse, reverse_lazy
from django.views.generic import DetailView, FormView, ListView, TemplateView
from openpyxl import Workbook
from usuarios.mixins import RolRequeridoMixin
from almacen.forms import CargaSAPForm                      # ajuste el import
from almacen.models import (
    Asignacion, Categoria, UnidadSerializada,
)
from operaciones.models import Operacion, RegistroSAP, CargaSAP
from . import services_sap
from usuarios.roles import ROLES_ALMACEN, DIGITADOR, COORDINADOR,ROLES_DIGITADOR
# + sus imports de RolRequeridoMixin y FormPageMixin

E = RegistroSAP.Estado


def _con_label(qs):
    return [
        {"codigo": x["tipo_error"],
         "label": services_sap.TIPOS_ERROR.get(x["tipo_error"], x["tipo_error"]),
         "n": x["n"]}
        for x in qs
    ]


class QuerystringMixin:
    """Deja en el contexto los filtros sin 'page' para armar la paginación."""
    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        p = self.request.GET.copy()
        p.pop("page", None)
        ctx["querystring"] = p.urlencode()
        return ctx


# ------------------------------------------------------------------ carga
class CargaSAPView(FormPageMixin, RolRequeridoMixin, FormView):
    roles_permitidos = ROLES_ALMACEN
    form_class = CargaSAPForm
    template_name = "operaciones/carga_sap_form.html"
    titulo = "Carga masiva SAP"
    boton = "Cargar y validar"
    cancel_url = reverse_lazy("almacen:asignaciones")

    def form_valid(self, form):
        try:
            carga = services_sap.procesar_carga(
                form.cleaned_data["archivo"], self.request.user
            )
        except services_sap.ErrorArchivo as exc:
            form.add_error("archivo", str(exc))
            return self.form_invalid(form)
        except Exception:
            logging.getLogger(__name__).exception("Error procesando archivo de SAP")
            form.add_error(
                "archivo",
                "No fue posible procesar el archivo. "
                "Verifique su contenido e inténtelo nuevamente.",
            )
            return self.form_invalid(form)

        resumen = (
            f"Carga #{carga.pk}: {carga.filas_ok} conciliados, "
            f"{carga.filas_error} con errores y "
            f"{carga.filas_ignoradas} ignorados (ya cargados antes)."
        )
        if carga.filas_error:
            messages.warning(self.request, resumen)
        else:
            messages.success(self.request, resumen)

        return HttpResponseRedirect(
            reverse("operaciones:carga_sap_detalle", kwargs={"pk": carga.pk})
        )


@login_required
def descargar_plantilla_sap(request):
    wb = Workbook()
    ws = wb.active
    ws.title = "SAP"
    ws.append(["numero_orden","cuenta_contrato","tipo_operacion",
        "numero_serie","codigo_unidad","cantidad",
    ])

    ws.append(["1234567890", "9876543210", "Instalación", "SER123", "", ""])
    ws.append(["1234567891", "9876543211", "Instalación", "", "REG-0001", ""])
    ws.append(["1234567892", "9876543212", "Instalación", "", "", "5"])
    resp = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    resp["Content-Disposition"] = 'attachment; filename="plantilla_sap.xlsx"'
    wb.save(resp)
    return resp


# ------------------------------------------------------------------ historial
class HistorialCargaSAPView(QuerystringMixin, RolRequeridoMixin, ListView):
    roles_permitidos = ROLES_ALMACEN
    model = CargaSAP
    template_name = "operaciones/carga_sap_historial.html"
    context_object_name = "cargas"
    paginate_by = 25

    def get_queryset(self):
        qs = CargaSAP.objects.select_related("usuario").order_by("-fecha", "-pk")
        g = self.request.GET
        desde, hasta = g.get("desde", "").strip(), g.get("hasta", "").strip()
        estado, q = g.get("estado", "").strip(), g.get("q", "").strip()

        if desde:
            qs = qs.filter(fecha__date__gte=desde)
        if hasta:
            qs = qs.filter(fecha__date__lte=hasta)
        if estado == "OK":
            qs = qs.filter(filas_error=0, filas_ok__gt=0)
        elif estado == "ERROR":
            qs = qs.filter(filas_error__gt=0)
        if q:
            qs = qs.filter(Q(archivo__icontains=q) | Q(usuario__username__icontains=q))
        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        ctx.update(desde=g.get("desde", ""), hasta=g.get("hasta", ""),
                   estado_sel=g.get("estado", ""), q=g.get("q", ""))

        cargas = list(ctx["cargas"])
        filas = (
            RegistroSAP.objects
            .filter(carga_id__in=[c.pk for c in cargas], estado=E.ERROR)
            .values("carga_id", "tipo_error")
            .annotate(n=Count("id")).order_by("-n")
        )
        por_carga = defaultdict(list)
        for x in filas:
            por_carga[x["carga_id"]].append({
                "codigo": x["tipo_error"],
                "label": services_sap.TIPOS_ERROR.get(x["tipo_error"], x["tipo_error"]),
                "n": x["n"],
            })
        for c in cargas:
            c.errores_resumen = por_carga.get(c.pk, [])
        ctx["cargas"] = cargas
        return ctx


# ------------------------------------------------------------------ detalle
class CargaSAPDetalleView(QuerystringMixin, RolRequeridoMixin, DetailView):
    roles_permitidos = ROLES_ALMACEN
    model = CargaSAP
    template_name = "operaciones/carga_sap_detalle.html"
    context_object_name = "carga"

    def get_queryset(self):
        return CargaSAP.objects.select_related("usuario")

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        estado = g.get("estado", "").strip()
        tipo_error = g.get("tipo_error", "").strip()
        q = g.get("q", "").strip()

        base = self.object.registros
        regs = base.select_related(
            "operacion", "tecnico", "producto", "unidad_serializada"
        ).order_by("fila")

        if estado in RegistroSAP.Estado.values:
            regs = regs.filter(estado=estado)
        if tipo_error:
            regs = regs.filter(tipo_error=tipo_error)
        if q:
            regs = regs.filter(
                Q(numero_orden__icontains=q)
                | Q(codigo_unidad__icontains=q)
                | Q(numero_serie__icontains=q)
                | Q(cuenta_contrato__icontains=q)
                | Q(tecnico__username__icontains=q)
                | Q(tecnico__first_name__icontains=q)
                | Q(tecnico__last_name__icontains=q)
                | Q(tecnico__cedula__icontains=q)
                | Q(error__icontains=q)
                | Q(diagnostico__icontains=q)
            )

        ctx["page_obj"] = Paginator(regs, 50).get_page(g.get("page"))
        ctx["resumen_errores"] = _con_label(
            base.filter(estado=E.ERROR).values("tipo_error")
            .annotate(n=Count("id")).order_by("-n")
        )
        ctx.update(estado_sel=estado, tipo_error_sel=tipo_error, q=q)
        return ctx


# ------------------------------------------------------------------ Five vs SAP
class ConciliacionSAPView(QuerystringMixin, RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "operaciones/conciliacion_sap.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET

        carga_sel = g.get("carga")
        if carga_sel is None:                       # por defecto, la última carga
            ultima = CargaSAP.objects.order_by("-fecha", "-pk").first()
            carga_sel = str(ultima.pk) if ultima else "todas"

        qs = (
            RegistroSAP.objects
            .exclude(numero_serie="")               # solo serializados
            .exclude(estado=E.IGNORADO)
            .select_related("carga", "producto__categoria")
            .order_by("-carga__fecha", "fila")
        )
        if carga_sel != "todas":
            qs = qs.filter(carga_id=carga_sel)

        tipo = g.get("tipo", "")
        if tipo.isdigit():
            qs = qs.filter(producto__categoria_id=tipo)

        resumen = {x["estado"]: x["n"] for x in qs.values("estado").annotate(n=Count("id"))}

        estado, error, q = g.get("estado", ""), g.get("tipo_error", ""), g.get("q", "").strip()
        if estado in (E.OK, E.ERROR):
            qs = qs.filter(estado=estado)
        if error:
            qs = qs.filter(tipo_error=error)
        if q:
            qs = qs.filter(
                Q(numero_serie__icontains=q) | Q(numero_orden__icontains=q)
                | Q(cuenta_contrato__icontains=q)
            )

        ctx.update(
            page_obj=Paginator(qs, 50).get_page(g.get("page")),
            categorias=Categoria.objects.exclude(prefijo="").order_by("nombre"),
            cargas=CargaSAP.objects.order_by("-fecha", "-pk")[:30],
            tipos_error=[{"codigo": k, "label": v} for k, v in services_sap.TIPOS_ERROR.items()],
            carga_sel=carga_sel, tipo_sel=tipo, estado_sel=estado,
            error_sel=error, q=q,
            n_ok=resumen.get(E.OK, 0), n_error=resumen.get(E.ERROR, 0),
        )
        return ctx


# ------------------------------------------------------------------ consulta por orden / serie
class ConsultaOrdenView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "operaciones/consulta_orden.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        q = self.request.GET.get("q", "").strip()
        ctx["q"] = q
        if not q:
            return ctx

        operacion = (
            Operacion.objects.select_related("tecnico").filter(numero_orden=q).first()
        )
        if operacion:
            registros = list(
                RegistroSAP.objects.filter(numero_orden=q)
                .select_related("carga", "producto").order_by("-carga__fecha", "fila")
            )
            ultimo_por_serie = {}
            for r in registros:
                if r.numero_serie:
                    ultimo_por_serie.setdefault(r.numero_serie, r)

            materiales = list(
                operacion.materiales
                .select_related("producto", "unidad_serializada", "tecnico")
                .prefetch_related("evidencias")
                .order_by("producto__nombre", "id")
            )
            for m in materiales:
                serie = m.unidad_serializada.numero_serie if m.unidad_serializada_id else ""
                m.sap = ultimo_por_serie.get(serie)
            ctx.update(operacion=operacion, materiales=materiales, registros=registros)
            return ctx

        # No es una orden exacta: ¿es una serie / código de unidad?
        unidades = list(
            UnidadSerializada.objects
            .filter(Q(numero_serie=q) | Q(codigo=q)).select_related("producto")
        )
        bloques = []
        for u in unidades:
            bloques.append({
                "unidad": u,
                "asignaciones": (
                    u.asignaciones.select_related("operacion", "tecnico")
                    .prefetch_related("evidencias").order_by("-id")
                ),
                "registros": (
                    RegistroSAP.objects.filter(numero_serie=u.numero_serie)
                    .select_related("carga").order_by("-carga__fecha")
                ),
            })
        ctx["bloques"] = bloques
        if not bloques:
            ctx["candidatas"] = Operacion.objects.filter(numero_orden__icontains=q)[:10]
        return ctx
    
class ErroresActualesView(QuerystringMixin, RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "operaciones/errores_actuales.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        vista = g.get("vista", "actuales")          # actuales | resueltos | falta_sap
        tipo, sev, q = g.get("tipo_error", ""), g.get("severidad", ""), g.get("q", "").strip()
        ctx.update(vista=vista, tipo_sel=tipo, sev_sel=sev, q=q)

        if vista == "falta_sap":
            qs = services_sap.instalados_sin_sap()
            if q:
                qs = qs.filter(Q(unidad_serializada__numero_serie__icontains=q)
                               | Q(operacion__numero_orden__icontains=q))
            ctx.update(page_obj=Paginator(qs, 50).get_page(g.get("page")),
                       limite=services_sap.limite_gracia(),
                       dias=services_sap.DIAS_GRACIA_SAP)
            return ctx

        base = RegistroSAP.objects.select_related("carga", "tecnico", "producto", "resuelto_en")
        if vista == "resueltos":
            qs = base.filter(resuelto=True).order_by("-resuelto_en__fecha", "-pk")
        else:
            qs = base.filter(estado=E.ERROR, vigente=True).order_by("-carga__fecha", "fila")

        resumen = [
            {**services_sap.meta(x["tipo_error"]), "codigo": x["tipo_error"], "n": x["n"]}
            for x in qs.values("tipo_error").annotate(n=Count("id"))
        ]
        resumen.sort(key=lambda r: (services_sap.ORDEN_SEV[r["severidad"]], -r["n"]))

        if tipo:
            qs = qs.filter(tipo_error=tipo)
        if sev in services_sap.ORDEN_SEV:
            qs = qs.filter(tipo_error__in=services_sap.tipos_por_severidad(sev))
        if q:
            qs = qs.filter(Q(numero_serie__icontains=q) | Q(numero_orden__icontains=q)
                           | Q(cuenta_contrato__icontains=q))

        ctx.update(
            resumen=resumen,
            n_alta=sum(r["n"] for r in resumen if r["severidad"] == "ALTA"),
            n_media=sum(r["n"] for r in resumen if r["severidad"] == "MEDIA"),
            n_baja=sum(r["n"] for r in resumen if r["severidad"] == "BAJA"),
            page_obj=Paginator(qs, 50).get_page(g.get("page")),
        )
        return ctx