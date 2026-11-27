from django.core.management.base import BaseCommand
from django.db import transaction

from operaciones.models import ActividadMtto, TipoOperacion

DATOS = {
    ("ZCES", "WM-Cese de servicio"): [
        ("S10", "Cese dispositivo por falta de pago"),
        ("S20", "Cese dispo resid-pyme control medicion"),
        ("S30", "Cese dispo resid por consumo cero"),
        ("S40", "ODS Cese dispo resid-pyme solicit client"),
        ("S70", "ODS Cese Administrativo"),
    ],
    ("ZRCM", "WM-Reparación centro de medición"): [
        ("F10", "Garantia por reparacion centro medicion"),
        ("F10", "Garantia por sustitucion medidor"),
    ],
    ("ZSCO", "WM-Reparación CM / Sustit MED"): [
        ("J10", "Reparacion centro medicion SAC"),
        ("J10", "Reparacion fuga N2/N3 Mantenimiento Red"),
        ("J10", "Reparacion/ Polietileno a la vista"),
        ("J10", "Reparacion centro medicion RPO"),
        ("J10", "Reparacion centro medicion Sx Rx"),
        ("J20", "Sustitucion medidor por revision RPO"),
        ("J30", "Aumento de capacidad SAC"),
        ("J30", "Sustitucion medidor por robo"),
        ("J35", "ODS Sustitucion medidor Emergencias"),
        ("J40", "Reparacion centro medicion client cesado"),
        ("J50", "Cambio de uso de Pyme a residencial"),
        ("J60", "Reparacion obra civil y/o PE a la vista"),
    ],
    ("ZVCL", "WM-Visitas a clientes"): [
        ("K10", "Visita integral"),
        ("K10", "V_Integral desviac significativa residen"),
        ("K10", "V_Integral desviac significativa pyme"),
        ("K20", "Visita Suspension por error"),
        ("K20", "Visita por daño en nicho"),
        ("K20", "Visita reconectado no pasa gas"),
        ("K20", "Visita validacion lectura"),
        ("K20", "Visita verificacion datos basicos"),
        ("K20", "Visita por daño en candado"),
        ("K20", "Visita cambio dispositivo de seguridad"),
        ("K20", "ODS Retiro dispositivo de seguridad"),
        ("K30", "Cotizacion cambio uso y/o aumento capaci"),
        ("K40", "Actualizacion de presion contratada"),
        ("K60", "Verfi desviac significativa residencial"),
        ("K60", "Verfi desviac significativa comercial"),
        ("K70", "Gestion cobranza"),
        ("K80", "Reconexion visita ajuste de presion GYF"),
    ],
}


class Command(BaseCommand):
    help = "Carga las actividades de mantenimiento (Act Mtto) de SAP por tipo de operación."

    @transaction.atomic
    def handle(self, *args, **options):
        for (grupo, nombre), actividades in DATOS.items():
            tipo = TipoOperacion.objects.filter(nombre=nombre).first()

            if tipo is None:
                tipo = TipoOperacion.objects.create(nombre=nombre, codigo_ods=grupo)
                self.stdout.write(self.style.WARNING(
                    f"⚠ TIPO CREADO (no existía): {nombre}"
                ))
            elif not tipo.codigo_ods:
                tipo.codigo_ods = grupo
                tipo.save(update_fields=["codigo_ods"])

            nuevas = 0
            for codigo, descripcion in actividades:
                _, creada = ActividadMtto.objects.get_or_create(
                    tipo=tipo, codigo=codigo, descripcion=descripcion,
                    defaults={"activa": True},
                )
                nuevas += creada

            self.stdout.write(self.style.SUCCESS(
                f"{grupo} · {nombre}: {nuevas} nuevas, {len(actividades)} en total"
            ))