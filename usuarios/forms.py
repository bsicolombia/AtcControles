from django import forms
from django.contrib.auth.forms import UserCreationForm
from .models import User


class UserForm(UserCreationForm):

    class Meta:
        model = User
        fields = [
            'username',
            'first_name',
            'last_name',
            'email',
            'cedula',
            'rol',
            'supervisor',
        ]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.fields['supervisor'].queryset = User.objects.filter(
            rol='SUPERVISOR'
        )

        self.fields['supervisor'].required = (
            self.data.get('rol', getattr(self.instance, 'rol', '')) == 'TECNICO'
        )

    def clean(self):
        cleaned_data = super().clean()

        rol = cleaned_data.get('rol')
        supervisor = cleaned_data.get('supervisor')

        if rol == 'TECNICO' and not supervisor:
            self.add_error(
                'supervisor',
                'Debes seleccionar un supervisor para el técnico.'
            )

        if rol != 'TECNICO' and supervisor:
            self.add_error(
                'supervisor',
                'Solo los técnicos pueden tener supervisor.'
            )

        return cleaned_data