# almacen/etiquetas.py
#
# Etiquetas cortas de cada tipo de inconsistencia.
# Se usan en la tabla de asignaciones y en la burbuja flotante.
# En views.py borre el diccionario ETIQUETAS_ERRORES y use:
#     from .etiquetas import ETIQUETAS_ERRORES

ETIQUETAS_ERRORES = {
    # Material
    "NO_ASIGNADO": "INSTALADO NO ASIGNADO A ESTE TECNICO",
    "OTRO_TECNICO": "INSTALADO ASIGNADO A OTRO TECNICO",
    "NO_REGISTRADO": "SERIAL INEXISTENTE",
    "YA_INSTALADO": "YA INSTALADO",
    "NO_UTILIZABLE": "DAÑADO / BAJA",
    "EXCEDE_CANTIDAD": "EXCEDE CANTIDAD ASIGNADA",
    "SIN_STOCK": "SIN STOCK",

    # IVT (Informe de Visita Técnica)
    "IVT_NO_REGISTRADO": "IVT INEXISTENTE",
    "IVT_NO_ASIGNADO": "IVT NO ASIGNADO AL TECNICO",
    "IVT_OTRO_TECNICO": "IVT DE OTRO TECNICO",
    "IVT_YA_USADO": "IVT YA USADO EN OTRA ORDEN",
    "IVT_NO_UTILIZABLE": "IVT ANULADO / PERDIDO",
}