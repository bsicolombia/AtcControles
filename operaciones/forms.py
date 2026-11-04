from django import forms
from .models import Operacion


class OperacionEditarForm(forms.ModelForm):

    class Meta:
        model = Operacion
        fields = [
            "numero_orden",
            "numero_cuenta_contrato",
            "tecnico",
            "tipo_operacion",
            "estado",
            "observaciones",
        ]