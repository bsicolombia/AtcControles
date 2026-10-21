from django import template
from operaciones import services_sap

register = template.Library()

@register.filter
def sap_meta(codigo, campo):
    return services_sap.meta(codigo).get(campo, "")