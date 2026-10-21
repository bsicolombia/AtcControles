from django.contrib.auth.mixins import LoginRequiredMixin, UserPassesTestMixin

from .roles import TODOS, tiene_rol


class RolRequeridoMixin(LoginRequiredMixin, UserPassesTestMixin):

    roles_permitidos = TODOS

    def test_func(self):
        return tiene_rol(
            self.request.user,
            *self.roles_permitidos,
        )
