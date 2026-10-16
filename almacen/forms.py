from django import forms
from django.utils import timezone

from operaciones.models import Operacion
from usuarios.models import Tecnicos
from . import services
from .models import Categoria, Inventario, Producto, UnidadSerializada


class EstiloMixin:
    """Aplica las clases CSS del diseño a todos los campos."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        for field in self.fields.values():
            widget = field.widget

            if isinstance(widget, forms.CheckboxInput):
                widget.attrs["class"] = "check"
            else:
                widget.attrs["class"] = "input"


class DateInput(forms.DateInput):
    input_type = "date"

    def __init__(self, **kwargs):
        super().__init__(format="%Y-%m-%d", **kwargs)


# ---------------------------------------------------------
# CATEGORÍAS
# ---------------------------------------------------------
class CategoriaForm(EstiloMixin, forms.ModelForm):
    class Meta:
        model = Categoria
        fields = ["nombre", "prefijo", "estado"]
        labels = {
            "estado": "Material activo",
        }

    def clean_prefijo(self):
        return self.cleaned_data["prefijo"].strip().upper()


# ---------------------------------------------------------
# PRODUCTOS
# ---------------------------------------------------------
class ProductoForm(EstiloMixin, forms.ModelForm):
    codigo_producto = forms.CharField(
        label="Código del tipo de material",
        max_length=30,
        required=False,
        disabled=True,
        widget=forms.TextInput(
            attrs={
                "placeholder": "Se genera automáticamente según la material",
                "class": "campo-automatico",
            }
        ),
        help_text="Este código será generado automáticamente.",
    )

    class Meta:
        model = Producto
        fields = [
            "categoria",
            "nombre",
            "codigo_producto",
            "tipo_inventario",
            "requiere_numero_serie",
            "marca",
            "medida",
            "referencia",
            "unidad_medida",
            "color",
            "presentacion",
            "estado",
        ]

        labels = {
            "categoria": "Material",
            "nombre": "Nombre del producto",
            "codigo_producto": "Código del producto",
            "tipo_inventario": "Tipo de inventario",
            "requiere_numero_serie": "Requiere número de serie físico",
            "medida": "Medida",
            "marca": "Marca",
            "referencia": "Referencia",
            "unidad_medida": "Unidad de medida",
            "color": "Color",
            "presentacion": "Presentación (tarro)",
            "estado": "Producto activo",
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.fields["categoria"].queryset = Categoria.objects.filter(
            estado=True
        )

    def clean_codigo_producto(self):
        codigo = (self.cleaned_data.get("codigo_producto") or "").strip().upper()

        if codigo:
            existe = Producto.objects.filter(
                codigo_producto=codigo
            ).exclude(pk=self.instance.pk)

            if existe.exists():
                raise forms.ValidationError(
                    "Ya existe un producto con este código."
                )

        return codigo

    def clean(self):
        cleaned_data = super().clean()

        tipo = cleaned_data.get("tipo_inventario")
        requiere_serial = cleaned_data.get("requiere_numero_serie")

        if tipo == Producto.TipoInventario.CANTIDAD and requiere_serial:
            self.add_error(
                "requiere_numero_serie",
                "Un producto manejado por cantidad "
                "no puede requerir número de serie físico.",
            )

        return cleaned_data

    def save(self, commit=True):
        obj = super().save(commit=False)

        if not obj.codigo_producto:
            obj.codigo_producto = services.generar_codigo_producto(
                obj.categoria
            )

        if commit:
            obj.save()

        return obj


class ProductoSelect(forms.Select):
    """Select que expone los datos del producto como atributos data-*."""

    def __init__(self, *args, **kwargs):
        self.productos = {}
        super().__init__(*args, **kwargs)

    def create_option(
        self,
        name,
        value,
        label,
        selected,
        index,
        subindex=None,
        attrs=None,
    ):
        option = super().create_option(
            name,
            value,
            label,
            selected,
            index,
            subindex=subindex,
            attrs=attrs,
        )

        # Opción vacía: "Seleccione un producto"
        if not value:
            return option

        # Django puede entregar ModelChoiceIteratorValue
        producto_id = getattr(value, "value", value)

        try:
            producto_id = int(producto_id)
        except (TypeError, ValueError):
            return option

        producto = self.productos.get(producto_id)

        if not producto:
            return option

        option["attrs"].update({
            "data-codigo": producto.codigo_producto or "",
            "data-nombre": producto.nombre or "",
            "data-marca": producto.marca or "",
            "data-referencia": producto.referencia or "",
            "data-medida": producto.medida or "",
            "data-unidad-medida": producto.get_unidad_medida_display() or "",
            "data-presentacion": producto.presentacion or "",
            "data-serializado": "true" if producto.es_serializado else "false",
            "data-requiere-serial": (
                "true" if producto.requiere_numero_serie else "false"
            ),
        })

        return option


# ---------------------------------------------------------
# INGRESOS
# ---------------------------------------------------------
from django import forms
from django.utils import timezone
from .models import Producto, UnidadSerializada, RemisionIngreso  # <-- Asegúrate de importar RemisionIngreso aquí

class IngresoForm(EstiloMixin, forms.Form):
    # Campo para seleccionar o escribir la remisión
    numero_orden = forms.CharField(
        label="Número de orden / Remisión",
        max_length=30,
        help_text="Seleccione o escriba el número de remisión. Si es nuevo, deberá adjuntar la foto."
    )
    
    comprobante_imagen = forms.ImageField(
        label="Foto de la remisión / orden",
        required=False,
        help_text="Obligatorio solo si el número de remisión es nuevo."
    )

    producto = forms.ModelChoiceField(
        queryset=Producto.objects.none(),
        label="Producto",
        empty_label="Seleccione un producto",
        widget=ProductoSelect(),
    )

    fecha_ingreso = forms.DateField(
        label="Fecha de ingreso",
        required=False,
        disabled=True,
        initial=timezone.localdate,
        widget=forms.DateInput(
            format="%Y-%m-%d",
            attrs={
                "type": "date",
                "class": "input campo-automatico",
            },
        ),
    )

    cantidad = forms.IntegerField(
        label="Cantidad",
        min_value=1,
        help_text="Cantidad que ingresa al almacén.",
    )

    marca = forms.CharField(required=False, disabled=True)
    referencia = forms.CharField(required=False, disabled=True)
    medida = forms.CharField(required=False, disabled=True)

    numeros_serie = forms.CharField(
        required=False,
        label="Números de serie físicos",
        widget=forms.Textarea(
            attrs={
                "rows": 5,
                "placeholder": "Un número de serie físico por línea",
            }
        ),
    )

    observacion = forms.CharField(
        required=False,
        label="Observaciones",
        widget=forms.Textarea(attrs={"rows": 4}),
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        productos = (
            Producto.objects
            .filter(estado=True, categoria__estado=True)
            .select_related("categoria")
            .order_by("nombre", "medida")
        )

        self.fields["producto"].queryset = productos

        self.fields["producto"].widget.productos = {
            producto.pk: producto for producto in productos
        }

        self.fields["producto"].label_from_instance = (
            lambda producto: (
                f"{producto.nombre_completo} · {producto.codigo_producto}"
            )
        )

        # La fecha siempre será la fecha actual.
        self.fields["fecha_ingreso"].initial = timezone.localdate()
        
    def clean_numero_orden(self):
        return self.cleaned_data["numero_orden"].strip().upper()

    def clean(self):
        cleaned_data = super().clean()
        
        num_orden = cleaned_data.get("numero_orden")
        imagen = cleaned_data.get("comprobante_imagen")
        producto = cleaned_data.get("producto")
        cantidad = cleaned_data.get("cantidad")
        texto_seriales = cleaned_data.get("numeros_serie", "")
        
        if num_orden:
            num_orden = num_orden.strip().upper()
            cleaned_data["numero_orden"] = num_orden
            
            # Verificamos si la remisión ya existe en la BD
            existe = RemisionIngreso.objects.filter(numero_remision=num_orden).exists()
            
            # Si NO existe y tampoco subió imagen, lanzamos error
            if not existe and not imagen:
                self.add_error(
                    "comprobante_imagen", 
                    "Este número de remisión es nuevo. Debe adjuntar la foto del comprobante."
                )

        if not producto or not cantidad:
            return cleaned_data

        seriales = [
            linea.strip().upper()
            for linea in texto_seriales.splitlines()
            if linea.strip()
        ]

        if producto.requiere_numero_serie:

            if len(seriales) != cantidad:
                self.add_error(
                    "numeros_serie",
                    f"Debe registrar exactamente {cantidad} "
                    "número(s) de serie físico(s).",
                )
                return cleaned_data

            if len(seriales) != len(set(seriales)):
                self.add_error(
                    "numeros_serie",
                    "No puede repetir números de serie.",
                )
                return cleaned_data

            existentes = set(
                UnidadSerializada.objects
                .filter(numero_serie__in=seriales)
                .values_list("numero_serie", flat=True)
            )

            if existentes:
                self.add_error(
                    "numeros_serie",
                    "Los siguientes números de serie ya existen: "
                    + ", ".join(sorted(existentes)),
                )
                return cleaned_data

            cleaned_data["seriales"] = seriales

        else:

            if seriales:
                self.add_error(
                    "numeros_serie",
                    "Este producto no requiere números de serie físicos.",
                )
                return cleaned_data

            cleaned_data["seriales"] = []

        return cleaned_data  # <-- Corregido aquí


# ---------------------------------------------------------
# UNIDADES CON CÓDIGO
# ---------------------------------------------------------
class UnidadForm(EstiloMixin, forms.ModelForm):
    class Meta:
        model = UnidadSerializada
        fields = [
            "numero_serie",
            "marca",
            "referencia",
            "medida",
            "observaciones",
        ]
        widgets = {
            "observaciones": forms.Textarea(attrs={"rows": 3}),
        }


class CambioEstadoForm(EstiloMixin, forms.Form):
    """Cambio manual de estado de una unidad."""

    estado = forms.ChoiceField(label="Nuevo estado")

    motivo = forms.CharField(
        label="Motivo",
        widget=forms.Textarea(attrs={"rows": 3}),
        help_text="Quedará en el historial de la unidad.",
    )

    def __init__(self, *args, unidad, **kwargs):
        self.unidad = unidad
        super().__init__(*args, **kwargs)

        permitidos = services.TRANSICIONES.get(unidad.estado, ())

        self.fields["estado"].choices = [
            (e.value, e.label) for e in permitidos
        ]


# ---------------------------------------------------------
# ASIGNACIONES
# ---------------------------------------------------------
class AsignacionForm(EstiloMixin, forms.Form):
    tecnico = forms.ModelChoiceField(
        queryset=Tecnicos.objects.none(),
        label="Técnico",
    )

    producto = forms.ModelChoiceField(
        queryset=Producto.objects.none(),
        label="Producto",
    )

    unidad = forms.ModelChoiceField(
        queryset=UnidadSerializada.objects.none(),
        required=False,
        label="Unidad (código)",
        help_text=(
            "Solo para productos serializados: "
            "elija el código que se entrega."
        ),
    )

    cantidad = forms.IntegerField(
        min_value=1,
        initial=1,
        label="Cantidad",
    )

    fecha_asignacion = forms.DateField(
        widget=DateInput(),
        initial=timezone.localdate,
        label="Fecha de asignación",
    )

    observaciones = forms.CharField(
        widget=forms.Textarea(attrs={"rows": 3}),
        required=False,
        label="Observaciones",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.fields["tecnico"].queryset = Tecnicos.objects.filter(activo=True)
        self.fields["producto"].queryset = Producto.objects.filter(estado=True)

        self.fields["producto"].label_from_instance = (
            lambda p: f"{p.nombre_completo} · {p.codigo_producto}"
        )

        self.fields["unidad"].queryset = UnidadSerializada.objects.filter(
            estado=UnidadSerializada.Estado.DISPONIBLE
        )
        self.fields["unidad"].label_from_instance = lambda u: u.etiqueta

    def clean(self):
        data = super().clean()

        producto = data.get("producto")
        unidad = data.get("unidad")
        cantidad = data.get("cantidad")

        if producto and producto.es_serializado:
            if not unidad:
                self.add_error("unidad", "Seleccione la unidad que se entrega.")
            elif unidad.producto_id != producto.id:
                self.add_error("unidad", "Esa unidad no pertenece al producto elegido.")
            data["cantidad"] = 1

        elif producto and cantidad:
            inv = Inventario.objects.filter(producto=producto).first()
            disponible = inv.cantidad_actual if inv else 0

            if cantidad > disponible:
                self.add_error(
                    "cantidad",
                    f"Solo hay {disponible} disponible(s) de "
                    f"{producto.nombre_completo} y está intentando asignar {cantidad}.",
                )

        return data



# ---------------------------------------------------------
# INSTALACIÓN
# ---------------------------------------------------------
class DatosOperacionForm(EstiloMixin, forms.Form):
    """Datos de la orden. Comunes a toda instalación."""

    tipo_operacion = forms.ChoiceField(
        label="Tipo de operación",
        choices=Operacion.TipoOperacion.choices,
    )

    numero_orden = forms.CharField(
        label="Número de orden",
        max_length=10,
    )

    numero_cuenta_contrato = forms.CharField(
        label="Número de cuenta contrato",
        max_length=10,
    )

    fecha_instalacion = forms.DateField(
        widget=DateInput(),
        initial=timezone.localdate,
        label="Fecha de instalación",
    )

    observaciones = forms.CharField(
        label="Observaciones",
        widget=forms.Textarea(attrs={"rows": 3}),
        required=False,
    )

    field_order = [
        "tipo_operacion",
        "numero_orden",
        "numero_cuenta_contrato",
        "fecha_instalacion",
        "observaciones",
    ]

    def clean_numero_orden(self):
        return self.cleaned_data["numero_orden"].strip().upper()

    def clean_numero_cuenta_contrato(self):
        return self.cleaned_data["numero_cuenta_contrato"].strip()


class InstalacionForm(DatosOperacionForm):
    """Instalación de UNA asignación (con su cantidad)."""

    cantidad = forms.IntegerField(
        label="Cantidad instalada",
        min_value=1,
    )

    def __init__(self, *args, asignacion=None, **kwargs):
        # Se guarda ANTES de super() para que EstiloMixin no lo reciba
        self.asignacion = asignacion

        super().__init__(*args, **kwargs)

        if asignacion is None:
            return

        if asignacion.unidad_serializada_id:
            # Una unidad con código siempre es exactamente 1
            del self.fields["cantidad"]
        else:
            pendiente = asignacion.cantidad_pendiente
            field = self.fields["cantidad"]
            field.initial = pendiente
            field.widget.attrs.update({"min": 1, "max": pendiente})
            field.help_text = (
                f"{asignacion.producto.nombre_completo}: "
                f"tiene {pendiente} pendiente(s) por instalar."
            )

    def clean_cantidad(self):
        cantidad = self.cleaned_data["cantidad"]

        if self.asignacion and cantidad > self.asignacion.cantidad_pendiente:
            raise forms.ValidationError(
                f"Solo tiene {self.asignacion.cantidad_pendiente} "
                "pendiente(s) por instalar."
            )

        return cantidad


class InstalacionMasivaForm(DatosOperacionForm):
    """Datos comunes de la orden. Las cantidades van por fila en la plantilla."""


# ---------------------------------------------------------
# INVENTARIO
# ---------------------------------------------------------
class InventarioAlmacenForm(EstiloMixin, forms.ModelForm):
    class Meta:
        model = Inventario
        fields = ["inventario_almacen"]
        labels = {
            "inventario_almacen": "Inventario de almacén",
        }
        widgets = {
            "inventario_almacen": forms.NumberInput(
                attrs={"min": "0", "placeholder": "Cantidad"}
            )
        }


class MinimoForm(EstiloMixin, forms.ModelForm):
    class Meta:
        model = Inventario
        fields = ["cantidad_minima"]
        labels = {
            "cantidad_minima": "Cantidad mínima",
        }
        widgets = {
            "cantidad_minima": forms.NumberInput(
                attrs={"min": "0", "placeholder": "Cantidad mínima"}
            )
        }
class CargaSAPForm(forms.Form):
    archivo = forms.FileField(
        label="Archivo Excel de SAP (.xlsx)",
        help_text="Seleccione un archivo Excel con las columnas requeridas.",
    )

    def clean_archivo(self):
        archivo = self.cleaned_data["archivo"]

        if not archivo.name.lower().endswith(".xlsx"):
            raise forms.ValidationError(
                "El archivo debe tener formato .xlsx."
            )

        return archivo