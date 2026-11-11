from django import forms
from almacen import services
from .models import Operacion


class OperacionEditarForm(forms.ModelForm):
    numero_ivt = forms.IntegerField(
        label="N.º de IVT", min_value=1, required=False,
        help_text="Obligatorio cuando la orden instala material.",
    )

    class Meta:
        model = Operacion
        fields = [
            "numero_orden", "numero_cuenta_contrato", "tecnico",
            "tipo_operacion", "estado", "observaciones",
        ]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if self.instance.pk:
            ivt = self.instance.ivts.filter(estado="INSTALADO").first()
            if ivt:
                self.fields["numero_ivt"].initial = ivt.numero

    def clean(self):
        data = super().clean()
        if data.get("estado") in Operacion.ESTADOS_QUE_INSTALAN:
            numero = data.get("numero_ivt")
            if not numero:
                self.add_error("numero_ivt", "Este campo es obligatorio: indique el IVT de la visita.")
            else:
                tecnico = data.get("tecnico") or self.instance.tecnico
                try:
                    services.validar_ivt_para_operacion(
                        numero=numero, tecnico=tecnico, operacion=self.instance)
                except services.ErrorInventario as e:
                    self.add_error("numero_ivt", str(e))
        return data