from django import forms



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
    