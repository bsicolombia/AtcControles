import logging
from django.views.generic import (
    CreateView,
    DetailView,
    FormView,
    ListView,
    TemplateView,
    UpdateView,
)
from django.utils.dateparse import parse_date
from .resumen import MESES, url_con, resumen_ejecucion
from django.utils import timezone
from django.shortcuts import get_object_or_404, redirect
from almacen import services as services_inventario
from django.db import transaction
from django.core.exceptions import ValidationError
import json
import json
from django.db import transaction
from django.core.exceptions import ValidationError
from django.contrib import messages
from usuarios.roles import ADMINISTRADOR
from almacen.models import (
    Asignacion,
    Producto,
    Inconsistencia,
)
from django.contrib.auth import get_user_model
User = get_user_model() 
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
import os
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404
from django.views import View
from operaciones.models import Anomalia, Operacion, RegistroSAP, CargaSAP
from .forms import OperacionEditarForm

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
    ws.append(["numero_orden", "cuenta_contrato", "tipo_operacion",
               "numero_serie", "codigo_unidad", "codigo_producto", "cantidad"])
    ws.append(["1234567890", "9876543210", "Instalación", "SER123", "", "", ""])
    ws.append(["1234567890", "9876543210", "Instalación", "", "REG-001", "", ""])
    ws.append(["1234567890", "9876543210", "Instalación", "", "", "COD-001", "5"])
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



# ------------------------------------------------------------------
# Five vs SAP
# ------------------------------------------------------------------


# ------------------------------------------------------------------
# Five vs SAP
# ------------------------------------------------------------------


class ConciliacionSAPView(
    QuerystringMixin,
    RolRequeridoMixin,
    TemplateView,
):
    roles_permitidos = ROLES_ALMACEN
    template_name = "operaciones/conciliacion_sap.html"

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        E = RegistroSAP.Estado

        # --------------------------------------------------
        # Valores de los filtros
        # --------------------------------------------------
        carga_sel = g.get("carga")
        if carga_sel is None:
            ultima = (
                CargaSAP.objects
                .order_by("-fecha", "-pk")
                .first()
            )
            carga_sel = str(ultima.pk) if ultima else "todas"

        tipo = g.get("tipo", "").strip()
        estado = g.get("estado", "").strip().upper()
        error = g.get("tipo_error", "").strip()
        q = g.get("q", "").strip()

        # --------------------------------------------------
        # Base de registros SAP
        # No incluir PENDIENTE aquí: viene de FIVE.
        # --------------------------------------------------
        qs = (
            RegistroSAP.objects
            .exclude(estado=E.IGNORADO)
            .filter(
                Q(estado=E.OK)
                | Q(estado=E.ERROR, vigente=True)
            )
            .select_related(
                "carga",
                "producto__categoria",
                "operacion",
                "unidad_serializada",
            )
            .order_by("-carga__fecha", "fila")
        )

        # Filtros compartidos por los contadores SAP
        # y por los resultados de SAP.
        if carga_sel != "todas":
            if carga_sel.isdigit():
                qs = qs.filter(carga_id=int(carga_sel))
            else:
                qs = qs.none()

        if tipo.isdigit():
            qs = qs.filter(producto__categoria_id=int(tipo))

        if error:
            qs = qs.filter(tipo_error=error)

        if q:
            qs = qs.filter(
                Q(numero_serie__icontains=q)
                | Q(numero_orden__icontains=q)
                | Q(cuenta_contrato__icontains=q)
                | Q(codigo_producto__icontains=q)
                | Q(codigo_unidad__icontains=q)
            )

        # --------------------------------------------------
        # CONTADORES SAP: antes del filtro de resultado.
        # Así las tarjetas no se ponen en cero al seleccionarlas.
        # --------------------------------------------------
        n_ok = qs.filter(estado=E.OK).count()
        n_error = qs.filter(
            estado=E.ERROR,
            vigente=True,
        ).count()

        # --------------------------------------------------
        # PENDIENTES FIVE
        # --------------------------------------------------
        pendientes = list(
            services_sap.instalados_sin_sap()
        )

        # Categoría de producto
        if tipo.isdigit():
            categoria_id = int(tipo)
            pendientes = [
                a for a in pendientes
                if a.producto
                and a.producto.categoria_id == categoria_id
            ]

        # Búsqueda de pendientes por identificador, producto,
        # orden, cuenta de contrato o técnico.
        if q:
            termino = q.casefold()
            filtrados = []

            for a in pendientes:
                unidad = getattr(a, "unidad_serializada", None)
                operacion = getattr(a, "operacion", None)
                tecnico = getattr(a, "tecnico", None)
                producto = getattr(a, "producto", None)

                valores = [
                    getattr(unidad, "numero_serie", ""),
                    getattr(unidad, "codigo", ""),
                    getattr(producto, "codigo_producto", ""),
                    getattr(producto, "nombre_completo", ""),
                    getattr(operacion, "numero_orden", ""),
                    getattr(operacion, "numero_cuenta_contrato", ""),
                    getattr(tecnico, "first_name", ""),
                    getattr(tecnico, "last_name", ""),
                    getattr(tecnico, "username", ""),
                ]

                if any(
                    termino in str(valor or "").casefold()
                    for valor in valores
                ):
                    filtrados.append(a)

            pendientes = filtrados

        n_pendientes = len(pendientes)

        # --------------------------------------------------
        # FILTRO DE RESULTADO
        # OK y ERROR muestran SAP.
        # PENDIENTE muestra instalaciones FIVE sin conciliar.
        # --------------------------------------------------
        if estado == "PENDIENTE":
            resultados = qs.none()
        elif estado in (E.OK, E.ERROR):
            resultados = qs.filter(estado=estado)
        else:
            resultados = qs

        ctx.update(
            page_obj=Paginator(resultados, 50).get_page(
                g.get("page")
            ),
            categorias=Categoria.objects.exclude(
                prefijo=""
            ).order_by("nombre"),
            cargas=CargaSAP.objects.order_by(
                "-fecha", "-pk"
            )[:30],
            tipos_error=[
                {"codigo": k, "label": v}
                for k, v in services_sap.TIPOS_ERROR.items()
            ],
            carga_sel=carga_sel,
            tipo_sel=tipo,
            estado_sel=estado,
            error_sel=error,
            q=q,
            n_ok=n_ok,
            n_error=n_error,
            instalados_pendientes=pendientes,
            n_pendientes=n_pendientes,
            mostrar_pendientes=(estado == "PENDIENTE"),
        )

        return ctx


class ConsultaOrdenView(RolRequeridoMixin, TemplateView):
    roles_permitidos = ROLES_ALMACEN
    template_name = "operaciones/consulta_orden.html"
    por_pagina = 25

    @staticmethod
    def _n(valor):
        return str(valor or "").strip().upper()

    # ---------- filtros comunes (listado + resumen) ----------
    @staticmethod
    def _filtrar(qs, q, estado, tipo, tecnico, anomalia):
        """Búsqueda, estado, tipo, técnico y anomalía. NO toca fechas."""
        if q:
            qs = qs.filter(
                Q(numero_orden__icontains=q)
                | Q(numero_cuenta_contrato__icontains=q)
            )
        if estado:
            qs = qs.filter(estado=estado)
        if tipo:
            qs = qs.filter(tipo_operacion=tipo)
        if tecnico.isdigit():
            qs = qs.filter(tecnico_id=int(tecnico))
        if anomalia.isdigit():
            qs = qs.filter(anomalia_id=int(anomalia))
        return qs

    # ---------- detalle de una orden ----------
    def _detalle(self, ctx, operacion):
        OK = RegistroSAP.Estado.OK
        IGN = RegistroSAP.Estado.IGNORADO

        registros = list(
            RegistroSAP.objects
            .filter(
                Q(numero_orden=operacion.numero_orden)
                | Q(operacion=operacion)
            )
            .exclude(estado=IGN)
            .select_related(
                "carga", "producto", "operacion", "tecnico", "unidad_serializada",
            )
            .order_by("-carga__fecha", "-fila", "-pk")
            .distinct()
        )

        por_serie, por_unidad, por_producto = {}, {}, {}

        def guardar(indice, clave, r):
            """Un OK siempre gana; entre iguales gana el más reciente."""
            actual = indice.get(clave)
            if actual is None or (actual.estado != OK and r.estado == OK):
                indice[clave] = r

        for r in registros:                 # del más nuevo al más viejo
            if r.numero_serie:
                guardar(por_serie, self._n(r.numero_serie), r)
            if r.codigo_unidad:
                guardar(por_unidad, self._n(r.codigo_unidad), r)
            if r.codigo_producto:
                guardar(por_producto, self._n(r.codigo_producto), r)

        materiales = list(
            operacion.materiales
            .exclude(estado=Asignacion.Estado.DEVUELTA)
            .select_related("producto", "unidad_serializada", "tecnico")
            .order_by("producto__nombre", "pk")
        )

        for m in materiales:
            u = m.unidad_serializada if m.unidad_serializada_id else None

            if u and u.numero_serie:                       # medidor
                m.identificador_consulta = f"{u.numero_serie} / {u.codigo}"
                m.tipo_identificador = "Serie / código de unidad"
                m.sap = (por_serie.get(self._n(u.numero_serie))
                         or por_unidad.get(self._n(u.codigo)))
            elif u:                                        # regulador, válvula
                m.identificador_consulta = u.codigo
                m.tipo_identificador = "Código de unidad"
                m.sap = por_unidad.get(self._n(u.codigo))
            else:                                          # por cantidad
                m.identificador_consulta = m.producto.codigo_producto
                m.tipo_identificador = "Código de producto"
                m.sap = por_producto.get(self._n(m.producto.codigo_producto))

        inconsistencias = (
            Inconsistencia.objects
            .filter(operacion=operacion)
            .select_related("producto", "unidad_serializada", "tecnico")
            .order_by("estado", "-fecha")
        )

        ctx.update(
            operacion=operacion,
            materiales=materiales,
            registros=registros,
            inconsistencias=inconsistencias,
        )

    # ---------- consulta por serie o código de unidad ----------
    def _bloques(self, q):
        bloques = []
        for u in (UnidadSerializada.objects
                  .filter(Q(numero_serie__iexact=q) | Q(codigo__iexact=q))
                  .select_related("producto")):
            filtro = Q(codigo_unidad=u.codigo)
            if u.numero_serie:
                filtro |= Q(numero_serie=u.numero_serie)
            bloques.append({
                "unidad": u,
                "asignaciones": (u.asignaciones
                                 .select_related("operacion", "tecnico")
                                 .order_by("-id")),
                "registros": (RegistroSAP.objects.filter(filtro)
                              .select_related("carga").order_by("-carga__fecha")),
            })
        return bloques

    # ---------- contexto ----------
    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        g = self.request.GET
        hoy = timezone.localdate()

        q = g.get("q", "").strip()
        estado = g.get("estado", "").strip()
        tipo = g.get("tipo", "").strip()
        tecnico = g.get("tecnico", "").strip()
        anomalia = g.get("anomalia", "").strip()
        detalle = g.get("detalle", "").strip()

        anio_param = g.get("anio", "").strip()
        mes_param = g.get("mes", "").strip()
        fecha_param = g.get("fecha", "").strip()

        # parse_date lanza ValueError con fechas imposibles (2026-02-31)
        try:
            fecha_dt = parse_date(fecha_param) if fecha_param else None
        except ValueError:
            fecha_dt = None

        # --- Año y mes efectivos del resumen ---
        anio = (int(anio_param)
                if anio_param.isdigit() and 1900 <= int(anio_param) <= 2100
                else hoy.year)
        mes = (int(mes_param)
               if mes_param.isdigit() and 1 <= int(mes_param) <= 12
               else None)
        if fecha_dt:                      # si eligió un día, el resumen se ubica en su mes
            anio = fecha_dt.year
            mes = fecha_dt.month

        # --- Años disponibles (chips del resumen y select del filtro) ---
        anios_db = sorted(
            {d.year for d in Operacion.objects.dates("fecha_operacion", "year")},
            reverse=True,
        )
        if anio not in anios_db:
            anios_db.insert(0, anio)
        anios = [
            {"valor": a, "activo": a == anio,
             "url": url_con(g, anio=a, mes=None, fecha=None)}
            for a in anios_db
        ]

        ctx.update(
            q=q,
            estado_sel=estado,
            tipo_sel=tipo,
            tecnico_sel=tecnico,
            anomalia_sel=anomalia,
            anio_sel=anio_param,
            mes_sel=str(mes) if mes else "",
            fecha_sel=fecha_dt.isoformat() if fecha_dt else "",
            anios=anios,
            meses=MESES,
            anomalias=Anomalia.objects.filter(activa=True),
            estados_orden=Operacion.Estado.choices,
            tipos_operacion=Operacion.TipoOperacion.choices,
            tecnicos=User.objects.filter(is_active=True, rol="TECNICO")
                         .order_by("first_name", "last_name", "username"),
            mostrar_listado=True,
        )

        # Una orden escrita exacta en el buscador abre el detalle directo
        if not detalle and q and Operacion.objects.filter(numero_orden=q).exists():
            detalle = q

        # ---------------- Detalle de una orden ----------------
        if detalle:
            ctx["mostrar_listado"] = False
            op = (Operacion.objects.select_related("tecnico", "anomalia")
                  .filter(numero_orden=detalle).first())
            if op:
                self._detalle(ctx, op)
            else:
                ctx["error_detalle"] = f"No se encontró la orden {detalle}."
            return ctx

        # ---------------- Serie o código de unidad exactos ----------------
        if q:
            bloques = self._bloques(q)
            if bloques:
                ctx.update(bloques=bloques, mostrar_listado=False)
                return ctx

        # ---------------- Resumen + listado ----------------
        # 1) Base: filtros normales, SIN fechas y SIN annotate
        base = self._filtrar(
            Operacion.objects.select_related("tecnico", "anomalia"),
            q, estado, tipo, tecnico, anomalia,
        )

        # 2) Resumen por técnico (año -> mes -> día)
        ctx["resumen"] = resumen_ejecucion(base, g, anio, mes)

        # 3) Filtros de fecha sobre el LISTADO
        qs = base
        if fecha_dt:
            qs = qs.filter(fecha_operacion=fecha_dt)
        else:
            if anio_param.isdigit() or mes:
                qs = qs.filter(fecha_operacion__year=anio)
            if mes:
                qs = qs.filter(fecha_operacion__month=mes)

        # 4) Ahora sí, el annotate de inconsistencias
        qs = qs.annotate(
            n_inc=Count(
                "inconsistencias",
                filter=Q(inconsistencias__estado="PENDIENTE"),
                distinct=True,
            )
        ).order_by("-fecha_operacion", "-pk")

        # 5) Migas de pan (cada una quita su filtro)
        migas = []
        if anio_param.isdigit() or mes or fecha_dt:
            migas.append({"label": f"Año {anio}",
                          "quitar": url_con(g, anio=None, mes=None, fecha=None)})
        if mes:
            migas.append({"label": dict(MESES)[mes],
                          "quitar": url_con(g, mes=None, fecha=None)})
        if fecha_dt:
            migas.append({"label": fecha_dt.strftime("%d/%m/%Y"),
                          "quitar": url_con(g, fecha=None)})
        if tecnico.isdigit():
            tec = User.objects.filter(pk=int(tecnico)).first()
            if tec:
                migas.append({"label": tec.get_full_name() or tec.username,
                              "quitar": url_con(g, tecnico=None)})

        # 6) Paginación (el querystring conserva todos los filtros)
        params = g.copy()
        params.pop("page", None)
        params.pop("detalle", None)

        paginator = Paginator(qs, self.por_pagina)
        ctx.update(
            page_obj=paginator.get_page(g.get("page")),
            total_ordenes=paginator.count,
            migas=migas,
            querystring=params.urlencode(),
        )

        # Sugerencias si la búsqueda no trajo nada
        if q and paginator.count == 0:
            ctx["candidatas"] = Operacion.objects.filter(
                numero_orden__startswith=q[:6]
            ).order_by("numero_orden")[:5]

        return ctx
    
class SoporteOperacionView(RolRequeridoMixin, View):
    roles_permitidos = ROLES_ALMACEN   # agregue ROLES_DIGITADOR si ellos también deben verlo

    def get(self, request, pk):
        op = get_object_or_404(Operacion, pk=pk)
        if not op.archivo_soporte:
            raise Http404("Esta operación no tiene soporte.")
        return FileResponse(
            op.archivo_soporte.open("rb"),
            filename=os.path.basename(op.archivo_soporte.name),
        )
        
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
    



class OperacionUpdateView(
    FormPageMixin,
    RolRequeridoMixin,
    UpdateView,
):
    roles_permitidos = (ADMINISTRADOR, COORDINADOR)

    model = Operacion
    form_class = OperacionEditarForm
    template_name = "almacen/asignacion_instalar.html"

    titulo = "Editar operación"
    boton = "Guardar cambios"
    


    def get_queryset(self):
        return (
            Operacion.objects
            .select_related(
                "tecnico",
                "anomalia",
            )
            .prefetch_related(
                "materiales__producto",
                "materiales__unidad_serializada",
                "materiales__tecnico",
            )
        )

    def form_valid(self, form):

        try:
            # ------------------------------------------------
            # VALIDAR PAYLOAD
            # ------------------------------------------------

            payload_raw = self.request.POST.get(
                "payload",
                "[]",
            )

            try:
                payload = json.loads(payload_raw)
            except json.JSONDecodeError:
                raise ValidationError(
                    "Los materiales enviados no tienen "
                    "un formato JSON válido."
                )

            if not isinstance(payload, list):
                raise ValidationError(
                    "El formato de materiales no es válido."
                )

            if not all(
                isinstance(item, dict)
                for item in payload
            ):
                raise ValidationError(
                    "Cada material debe tener un formato válido."
                )

            # ------------------------------------------------
            # TRANSACCIÓN COMPLETA
            # ------------------------------------------------

            with transaction.atomic():

                # Bloquear operación.
                operacion = (
                    Operacion.objects
                    .select_for_update()
                    .get(pk=self.get_object().pk)
                )

                # Guardar datos anteriores antes de actualizar.
                self.object = operacion

                # Guardar los datos generales.
                self.object = form.save()
                
                if form.cleaned_data["estado"] in Operacion.ESTADOS_QUE_INSTALAN:
                    services_inventario.vincular_ivt(
                        numero=form.cleaned_data["numero_ivt"],
                        operacion=self.object,
                        tecnico=self.object.tecnico,
                        usuario=self.request.user,
                    )
                else:
                    services_inventario.liberar_ivt_operacion(
                        operacion=self.object,
                        usuario=self.request.user,
                        motivo="La orden dejó de estar ejecutada",
                    )
                
                logging.warning(
                    "EDITANDO OPERACION %s | PAYLOAD: %s",
                    self.object.pk,
                    json.dumps(payload, ensure_ascii=False)
                )

                # Sincronizar materiales.
                services_inventario.sincronizar_materiales_operacion(
                    operacion=self.object,
                    payload=payload,
                    usuario=self.request.user,
                )

            # ------------------------------------------------
            # CONFIRMACIÓN
            # ------------------------------------------------

            messages.success(
                self.request,
                "La operación y sus materiales "
                "fueron actualizados correctamente."
            )

            return HttpResponseRedirect(
                self.get_success_url()
            )

        except (
            ValidationError,
            ValueError,
            TypeError,
            services_inventario.ErrorInventario,
        ) as exc:

            logging.exception(
                "ERROR VALIDANDO MATERIALES DE OPERACION %s",
                getattr(self.object, "pk", None)
            )

            form.add_error(None, str(exc))
            return self.form_invalid(form)

        except Exception as exc:
            logging.getLogger(__name__).exception(
                "Error actualizando operación y materiales."
            )
            form.add_error(
                None,
                f"No fue posible actualizar la operación ({type(exc).__name__}: {exc}). "
                "No se confirmaron los cambios de inventario."
            )
            return self.form_invalid(form)

    # --------------------------------------------------------
    # MATERIALES PENDIENTES DEL TÉCNICO
    # --------------------------------------------------------

    def _datos_pendientes(self):

        tecnico = self.object.tecnico

        pendientes = (
            Asignacion.objects
            .filter(
                tecnico=tecnico,
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

        cantidades = {}
        seriales = []

        for asignacion in pendientes:

            if asignacion.cantidad_pendiente <= 0:
                continue

            if asignacion.unidad_serializada_id:

                unidad = asignacion.unidad_serializada

                seriales.append({
                    "id": unidad.pk,
                    "serial": (
                        unidad.numero_serie
                        or unidad.codigo
                    ),
                    "codigo": unidad.codigo,
                    "producto": asignacion.producto_id,
                    "nombre": asignacion.producto.nombre_completo,
                })

            else:

                cantidades[asignacion.producto_id] = (
                    cantidades.get(
                        asignacion.producto_id,
                        0,
                    )
                    + asignacion.cantidad_pendiente
                )

        pend_cantidad = [
            {
                "id": producto.pk,
                "nombre": producto.nombre_completo,
                "pendiente": cantidades.get(
                    producto.pk,
                    0,
                ),
            }
            for producto in Producto.objects.filter(
                estado=True,
                tipo_inventario=Producto.TipoInventario.CANTIDAD,
            ).order_by(
                "nombre",
                "medida",
            )
        ]

        return pend_cantidad, seriales

    # --------------------------------------------------------
    # MATERIALES EXISTENTES
    # --------------------------------------------------------

    def _payload_existente(self):
        agrupados = {}

        def grupo(producto, serializado):
            return agrupados.setdefault(
                (producto.pk, serializado),
                {
                    "producto": producto.pk,
                    "nombre": producto.nombre_completo,
                    "serializado": serializado,
                    "cantidad": 0,
                    "unidades": [],
                },
            )

        # Lo que realmente quedó instalado en la orden
        instalados = (
            self.object.materiales
            .filter(estado=Asignacion.Estado.INSTALADA)
            .select_related("producto", "unidad_serializada")
            .order_by("producto__nombre", "pk")
        )
        for m in instalados:
            if m.unidad_serializada_id:
                u = m.unidad_serializada
                grupo(m.producto, True)["unidades"].append({
                    "id": u.pk,
                    "serial": u.numero_serie or u.codigo,
                })
            else:
                grupo(m.producto, False)["cantidad"] += m.cantidad

        # Lo que se reportó pero no se instaló (inconsistencias abiertas)
        reportadas = (
            Inconsistencia.objects
            .filter(
                operacion=self.object,
                estado=Inconsistencia.Estado.PENDIENTE,
                instalado=False,
                producto__isnull=False,
            )
            .select_related("producto", "unidad_serializada")
        )
        for i in reportadas:
            if i.unidad_serializada_id or i.serial_reportado:
                serial = i.serial_reportado or i.unidad_serializada.codigo
                grupo(i.producto, True)["unidades"].append({
                    "id": i.unidad_serializada_id,
                    "serial": serial,
                    "tipo": i.tipo,
                    "instala": False,
                    "aviso": i.detalle,
                })
            else:
                grupo(i.producto, False)["cantidad"] += i.cantidad

        return list(agrupados.values())

    # --------------------------------------------------------
    # CONTEXTO
    # --------------------------------------------------------

    def get_context_data(self, **kwargs):

        ctx = super().get_context_data(**kwargs)

        operacion = self.object

        pend_cantidad, pend_seriales = (
            self._datos_pendientes()
        )

        ctx.update({

            "operacion": operacion,

            "tecnico": operacion.tecnico,

            "payload_previo": self._payload_existente(),

            "pend_cantidad": pend_cantidad,

            "pend_seriales": pend_seriales,

            "productos_serializados": (
                Producto.objects.filter(
                    estado=True,
                    tipo_inventario=(
                        Producto.TipoInventario.SERIALIZADO
                    ),
                ).order_by(
                    "nombre",
                    "medida",
                )
            ),

            "estados_instalan": sorted(
                estado.value
                for estado in Operacion.ESTADOS_QUE_INSTALAN
            ),

            "tipos_inconsistencia": dict(
                Inconsistencia.Tipo.choices
            ),

            "instalaciones": (
                operacion.materiales.select_related(
                    "producto",
                    "unidad_serializada",
                    "tecnico",
                )
            ),

            "cancel_url": reverse_lazy(
                "operaciones:consulta_orden"
            ),

            "modo_edicion": True,
        })

        return ctx

    def get_success_url(self):

        return reverse_lazy(
            "operaciones:consulta_orden"
        )

class OperacionReabrirView(RolRequeridoMixin, View):
    roles_permitidos = (ADMINISTRADOR, COORDINADOR)

    def post(self, request, pk):
        operacion = get_object_or_404(Operacion, pk=pk)
        try:
            numero = services_inventario.reabrir_operacion(
                operacion=operacion, usuario=request.user
            )
            messages.success(
                request,
                f"Orden {numero} reabierta: los materiales volvieron a su estado "
                "anterior y ya puede registrarla de nuevo.",
            )
        except services_inventario.ErrorInventario as exc:
            messages.error(request, str(exc))
        return redirect("operaciones:consulta_orden")