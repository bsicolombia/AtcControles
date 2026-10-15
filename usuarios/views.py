from io import BytesIO
from django.views.generic import (
    CreateView,
    FormView,
    ListView,
    UpdateView,
    View,
)
from django.contrib import messages
from django.db import transaction
from django.db.models import Q
from django.http import HttpResponse
from django.shortcuts import render
from django.urls import reverse_lazy
from django.views.generic import CreateView, FormView, ListView, UpdateView

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .forms import CargaMasivaTecnicosForm, TecnicoForm
from .mixins import RolRequeridoMixin
from .models import Tecnicos
from .roles import ADMINISTRADOR


def error_403(request, exception=None):
    return render(request, "403.html", status=403)


class TecnicoListView(RolRequeridoMixin, ListView):
    model = Tecnicos
    template_name = "usuarios/tecnico_list.html"
    paginate_by = 20

    def get_queryset(self):
        qs = super().get_queryset()

        q = self.request.GET.get("q", "").strip()

        if q:
            qs = qs.filter(
                Q(tecnico_apellido_nombres__icontains=q)
                | Q(tecnico_cedula__icontains=q)
                | Q(supervisor__icontains=q)
            )

        return qs

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        params = self.request.GET.copy()
        params.pop("page", None)

        ctx["qs"] = params.urlencode()

        return ctx


class TecnicoFormMixin(RolRequeridoMixin):
    roles_permitidos = (ADMINISTRADOR,)

    model = Tecnicos
    form_class = TecnicoForm
    template_name = "form.html"
    success_url = reverse_lazy("tecnicos:lista")

    def form_valid(self, form):
        messages.success(
            self.request,
            "Técnico guardado correctamente."
        )

        return super().form_valid(form)

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)

        ctx["cancel_url"] = self.success_url

        return ctx


class TecnicoCreateView(TecnicoFormMixin, CreateView):
    extra_context = {
        "titulo": "Nuevo técnico",
        "boton": "Guardar técnico",
    }


class TecnicoUpdateView(TecnicoFormMixin, UpdateView):
    extra_context = {
        "titulo": "Editar técnico",
        "boton": "Guardar cambios",
    }


class DescargarPlantillaTecnicosView(
    RolRequeridoMixin,
    View,
):
    roles_permitidos = (ADMINISTRADOR,)

    def get(self, request, *args, **kwargs):

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Técnicos"

        encabezados = [
            "SUPERVISOR",
            "CEDULA",
            "APELLIDOS Y NOMBRES",
        ]

        worksheet.append(encabezados)

        color_encabezado = "2D6A4F"

        for columna, encabezado in enumerate(
            encabezados,
            start=1,
        ):
            celda = worksheet.cell(
                row=1,
                column=columna,
                value=encabezado,
            )

            celda.fill = PatternFill(
                fill_type="solid",
                fgColor=color_encabezado,
            )

            celda.font = Font(
                color="FFFFFF",
                bold=True,
            )

            celda.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        worksheet.column_dimensions["A"].width = 30
        worksheet.column_dimensions["B"].width = 20
        worksheet.column_dimensions["C"].width = 40

        worksheet.row_dimensions[1].height = 28

        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = "A1:C1000"

        # La cédula se maneja como texto
        # para evitar que Excel la convierta.
        for fila in range(2, 1001):
            worksheet.cell(
                row=fila,
                column=2,
            ).number_format = "@"

        # Hoja de instrucciones
        instrucciones = workbook.create_sheet(
            "Instrucciones"
        )

        instrucciones["A1"] = "CARGA MASIVA DE TÉCNICOS"

        instrucciones["A1"].font = Font(
            bold=True,
            size=16,
            color="2D6A4F",
        )

        instrucciones["A3"] = "COLUMNAS"
        instrucciones["A3"].font = Font(bold=True)

        instrucciones["A4"] = "SUPERVISOR"
        instrucciones["B4"] = (
            "Nombre del supervisor del técnico."
        )

        instrucciones["A5"] = "CEDULA"
        instrucciones["B5"] = (
            "Número de cédula del técnico."
        )

        instrucciones["A6"] = "APELLIDOS Y NOMBRES"
        instrucciones["B6"] = (
            "Apellidos y nombres completos del técnico."
        )

        instrucciones["A8"] = "IMPORTANTE"
        instrucciones["A8"].font = Font(
            bold=True,
            color="C62828",
        )

        instrucciones["A9"] = (
            "No cambie los nombres de las columnas."
        )

        instrucciones["A10"] = (
            "La cédula debe ser única."
        )

        instrucciones["A11"] = (
            "Los técnicos cargados quedarán activos."
        )

        instrucciones["A12"] = (
            "No incluya la columna ACTIVO."
        )

        instrucciones.column_dimensions["A"].width = 30
        instrucciones.column_dimensions["B"].width = 75

        archivo = BytesIO()

        workbook.save(archivo)

        archivo.seek(0)

        response = HttpResponse(
            archivo.getvalue(),
            content_type=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
        )

        response["Content-Disposition"] = (
            'attachment; filename="plantilla_tecnicos.xlsx"'
        )

        return response

class CargaMasivaTecnicosView(
    RolRequeridoMixin,
    FormView,
):
    roles_permitidos = (ADMINISTRADOR,)

    template_name = "usuarios/tecnico_carga_masiva.html"
    form_class = CargaMasivaTecnicosForm
    success_url = reverse_lazy("tecnicos:lista")

    def form_valid(self, form):

        archivo = form.cleaned_data["archivo"]

        try:
            workbook = load_workbook(
                archivo,
                read_only=True,
                data_only=True,
            )

            worksheet = workbook.active

        except Exception:
            form.add_error(
                "archivo",
                "No fue posible leer el archivo Excel."
            )

            return self.form_invalid(form)

        encabezados_esperados = [
            "SUPERVISOR",
            "CEDULA",
            "APELLIDOS Y NOMBRES",
        ]

        encabezados_excel = []

        for celda in worksheet[1]:
            valor = celda.value

            if valor is None:
                encabezados_excel.append("")
            else:
                encabezados_excel.append(
                    str(valor).strip().upper()
                )

        if encabezados_excel != encabezados_esperados:

            form.add_error(
                "archivo",
                (
                    "El formato del Excel no es válido. "
                    "Las columnas deben ser exactamente: "
                    "SUPERVISOR, CEDULA y "
                    "APELLIDOS Y NOMBRES."
                ),
            )

            workbook.close()

            return self.form_invalid(form)

        registros = []
        cedulas_excel = set()
        errores = []

        for numero_fila, fila in enumerate(
            worksheet.iter_rows(
                min_row=2,
                values_only=True,
            ),
            start=2,
        ):

            # Ignorar filas completamente vacías
            if not any(
                valor is not None
                and str(valor).strip()
                for valor in fila[:3]
            ):
                continue

            supervisor = (
                str(fila[0]).strip()
                if fila[0] is not None
                else ""
            )

            cedula = (
                str(fila[1]).strip()
                if fila[1] is not None
                else ""
            )

            nombre = (
                str(fila[2]).strip()
                if fila[2] is not None
                else ""
            )

            if not supervisor:
                errores.append(
                    f"Fila {numero_fila}: "
                    "el supervisor está vacío."
                )
                continue

            if not cedula:
                errores.append(
                    f"Fila {numero_fila}: "
                    "la cédula está vacía."
                )
                continue

            if not cedula.isdigit():
                errores.append(
                    f"Fila {numero_fila}: "
                    f"la cédula '{cedula}' "
                    "debe contener únicamente números."
                )
                continue

            if len(cedula) > 10:
                errores.append(
                    f"Fila {numero_fila}: "
                    f"la cédula '{cedula}' "
                    "supera los 10 dígitos."
                )
                continue

            if not nombre:
                errores.append(
                    f"Fila {numero_fila}: "
                    "los apellidos y nombres están vacíos."
                )
                continue

            if cedula in cedulas_excel:
                errores.append(
                    f"Fila {numero_fila}: "
                    f"la cédula '{cedula}' "
                    "está repetida dentro del Excel."
                )
                continue

            cedulas_excel.add(cedula)

            registros.append(
                {
                    "supervisor": supervisor,
                    "tecnico_cedula": cedula,
                    "tecnico_apellido_nombres": nombre,
                }
            )

        workbook.close()

        if errores:
            for error in errores:
                form.add_error(
                    "archivo",
                    error,
                )

            return self.form_invalid(form)

        if not registros:
            form.add_error(
                "archivo",
                "El Excel no contiene técnicos para importar.",
            )

            return self.form_invalid(form)

        # Verificar cédulas que ya existen en la base de datos
        cedulas_existentes = set(
            Tecnicos.objects.filter(
                tecnico_cedula__in=cedulas_excel
            ).values_list(
                "tecnico_cedula",
                flat=True,
            )
        )

        if cedulas_existentes:
            lista = ", ".join(
                sorted(cedulas_existentes)
            )

            form.add_error(
                "archivo",
                (
                    "Las siguientes cédulas ya están "
                    f"registradas: {lista}"
                ),
            )

            return self.form_invalid(form)

        # Crear técnicos
        with transaction.atomic():

            Tecnicos.objects.bulk_create(
                [
                    Tecnicos(
                        supervisor=registro["supervisor"],
                        tecnico_cedula=registro[
                            "tecnico_cedula"
                        ],
                        tecnico_apellido_nombres=registro[
                            "tecnico_apellido_nombres"
                        ],
                        activo=True,
                    )
                    for registro in registros
                ]
            )

        messages.success(
            self.request,
            (
                f"Se cargaron correctamente "
                f"{len(registros)} técnico(s)."
            ),
        )

        return super().form_valid(form)
