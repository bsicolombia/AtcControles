from django import forms

from .models import Tecnicos


class TecnicoForm(forms.ModelForm):

    class Meta:
        model = Tecnicos
        fields = [
            "tecnico_apellido_nombres",
            "tecnico_cedula",
            "supervisor",
            "activo",
        ]

        labels = {
            "tecnico_apellido_nombres": "Apellidos y nombres",
            "tecnico_cedula": "Cédula",
            "supervisor": "Supervisor",
            "activo": "Técnico activo",
        }

        widgets = {
            "tecnico_apellido_nombres": forms.TextInput(
                attrs={
                    "placeholder": "Apellidos y nombres completos",
                }
            ),
            "tecnico_cedula": forms.TextInput(
                attrs={
                    "placeholder": "Número de cédula",
                    "inputmode": "numeric",
                }
            ),
            "supervisor": forms.TextInput(
                attrs={
                    "placeholder": "Nombre del supervisor",
                }
            ),
        }

    def clean_tecnico_cedula(self):
        cedula = self.cleaned_data["tecnico_cedula"].strip()

        if not cedula.isdigit():
            raise forms.ValidationError(
                "La cédula debe contener únicamente números."
            )

        if len(cedula) > 10:
            raise forms.ValidationError(
                "La cédula no puede tener más de 10 dígitos."
            )

        queryset = Tecnicos.objects.filter(
            tecnico_cedula=cedula
        ).exclude(
            pk=self.instance.pk
        )

        if queryset.exists():
            raise forms.ValidationError(
                "Ya existe un técnico registrado con esta cédula."
            )

        return cedula


class CargaMasivaTecnicosForm(forms.Form):

    archivo = forms.FileField(
        label="Archivo Excel",
        help_text=(
            "Seleccione la plantilla Excel diligenciada. "
            "Debe ser un archivo .xlsx."
        ),
        widget=forms.ClearableFileInput(
            attrs={
                "accept": ".xlsx",
                "class": "input",
            }
        ),
    )

    def clean_archivo(self):
        archivo = self.cleaned_data["archivo"]

        if not archivo.name.lower().endswith(".xlsx"):
            raise forms.ValidationError(
                "El archivo debe estar en formato Excel .xlsx."
            )

        return archivo
    