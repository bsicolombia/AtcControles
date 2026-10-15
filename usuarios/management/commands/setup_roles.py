from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group
from django.core.management.base import BaseCommand

from usuarios.roles import ADMINISTRADOR, ALMACENISTA


class Command(BaseCommand):
    help = "Crea los grupos (roles) del sistema. Con --demo crea usuarios de prueba."

    def add_arguments(self, parser):
        parser.add_argument("--demo", action="store_true")

    def handle(self, *args, **opts):
        for nombre in (ADMINISTRADOR, ALMACENISTA):
            _, creado = Group.objects.get_or_create(name=nombre)
            self.stdout.write(f"Rol {nombre}: {'creado' if creado else 'ya existía'}")

        if opts["demo"]:
            User = get_user_model()
            for username, rol, staff in (("admin", ADMINISTRADOR, True), ("almacen", ALMACENISTA, False)):
                u, creado = User.objects.get_or_create(username=username, defaults={"is_staff": staff})
                if creado:
                    u.set_password("cambiar123")
                    u.save()
                u.groups.add(Group.objects.get(name=rol))
                self.stdout.write(f"Usuario {username} ({rol}) - clave: cambiar123")
