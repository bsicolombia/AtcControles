from django.core.management.base import BaseCommand

from almacen.models import Categoria, Inventario, Producto

S, C = Producto.TipoInventario.SERIALIZADO, Producto.TipoInventario.CANTIDAD

CATEGORIAS = [
    ("MEDIDOR", "DET"), ("REGULADOR", "REG"), ("VALVULAS", "VAL"), ("CODOS", ""),
    ("CONECTOR", ""), ("NIPLE", ""), ("PRECINTOS", ""), ("FUERZA MEDIA", ""), ("FUERZA ALTA", ""),
]
# (categoria, nombre, codigo, tipo, medida, color, presentacion, minimo)
PRODUCTOS = [
    ("MEDIDOR", "Medidor", "MED-001", S, "", "", "", 0),
    ("REGULADOR", "Regulador", "REG-P01", S, "", "", "", 0),
    ("VALVULAS", "Válvula", "VAL-P01", S, "", "", "", 0),
    ("CODOS", "Codo", "COD-12", C, "1/2", "", "", 20),
    ("CODOS", "Codo", "COD-14", C, "1/4", "", "", 20),
    ("CONECTOR", "Conector", "CON-12", C, "1/2", "", "", 10),
    ("NIPLE", "Niple", "NIP-12", C, "1/2", "", "", 10),
    ("PRECINTOS", "Precinto", "PRE-ROJO", C, "", "Rojo", "", 20),
    ("FUERZA MEDIA", "Fuerza media", "FME-8G", C, "", "", "8g", 3),
    ("FUERZA ALTA", "Fuerza alta", "FAL-20G", C, "", "", "20g", 3),
]


class Command(BaseCommand):
    help = "Carga las categorías y productos iniciales del almacén (se puede ejecutar varias veces)."

    def handle(self, *args, **opts):
        cats = {}
        for nombre, prefijo in CATEGORIAS:
            cats[nombre], _ = Categoria.objects.get_or_create(nombre=nombre, defaults={"prefijo": prefijo})
        for cat, nombre, cod, tipo, medida, color, pres, minimo in PRODUCTOS:
            p, _ = Producto.objects.get_or_create(
                codigo_producto=cod,
                defaults=dict(categoria=cats[cat], nombre=nombre, tipo_inventario=tipo,
                              medida=medida or None, color=color or None, presentacion=pres or None),
            )
            inv, _ = Inventario.objects.get_or_create(producto=p)
            if minimo and not inv.cantidad_minima:
                inv.cantidad_minima = minimo
                inv.save()
        self.stdout.write(self.style.SUCCESS("Almacén inicial cargado."))
