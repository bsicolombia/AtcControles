from django import template
from operaciones import services_sap

register = template.Library()

@register.filter
def sap_meta(codigo, campo):
    return services_sap.meta(codigo).get(campo, "")
@register.filter
def sap_item(comparacion):
    """Devuelve el bloque principal de la comparación:
    'medidor' (serie o código de unidad) o 'cantidad' (productos por cantidad)."""
    comparacion = comparacion or {}
    return comparacion.get("medidor") or comparacion.get("cantidad") or {}