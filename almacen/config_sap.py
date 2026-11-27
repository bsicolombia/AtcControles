# almacen/config_sap.py
import re
import unicodedata

# «Status sistema» de SAP que significa "está en el almacén"
STATUS_ALMACEN = {"ALMA"}
# almacen/config_sap.py
CODIGOS_EXCLUIDOS_VANTI = {
    "N01", "N02", "N03", "N04", "N05", "N06", "N07", "N08", "N09", "N10",
    "N11", "N12", "N13", "N14", "N15", "N16", "U01", "U02", "UHH1",
    "C1", "C2", "CC1", "CC2", "B1", "AC", "FA", "FM",
}
# Tipos de operación cuyo material instalado SAP todavía ve en el almacén
# (se instala sin ODS, así que SAP no lo ha descontado).
TIPOS_SIN_ODS = ["WM-Reparación sin ods centro de medición"]


class ErrorConfiguracion(Exception):
    pass


def normalizar(texto):
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def tipos_sin_ods(estricto=False):
    """TipoOperacion que corresponden a TIPOS_SIN_ODS (ignora mayúsculas y tildes)."""
    from operaciones.models import TipoOperacion

    buscados = {normalizar(n): n for n in TIPOS_SIN_ODS}
    todos = list(TipoOperacion.objects.all())
    encontrados = [t for t in todos if normalizar(t.nombre) in buscados]

    if estricto:
        hallados = {normalizar(t.nombre) for t in encontrados}
        faltan = [n for k, n in buscados.items() if k not in hallados]
        if faltan:
            parecidos = [t.nombre for t in todos if "sin ods" in normalizar(t.nombre)][:8]
            raise ErrorConfiguracion(
                "No encontré el tipo de operación «" + "», «".join(faltan) + "»."
                + (" Tipos parecidos: " + "; ".join(parecidos) + "." if parecidos else "")
                + " Ajuste TIPOS_SIN_ODS en almacen/config_sap.py con el nombre exacto."
            )
    return encontrados