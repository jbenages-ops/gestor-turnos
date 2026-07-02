"""Lectura de ficheros .xml (DatosEnergeticosDelEdificio) y .cex (CE3X) para
extraer los datos necesarios para rellenar la Hoja de Encargo de Castilla y León.

La fuente principal es el .xml: es un formato oficial, documentado y estable
(el que exige el registro de certificados energéticos), a diferencia del .cex,
que es un volcado interno de CE3X (pickle de Python 2, posicional, no
documentado). El .xml no incluye la identidad del cliente (nombre/NIF del
propietario), así que si existe un .cex hermano con el mismo nombre se usa
sólo para rellenar ese dato.
"""
import io
import os
import pickle
import re
import xml.etree.ElementTree as ET


class SafeUnpickler(pickle.Unpickler):
    """Unpickler que sólo permite tipos básicos (str/list/dict/...), nunca clases."""

    def find_class(self, module, name):
        raise pickle.UnpicklingError(f"Clase no permitida en .cex: {module}.{name}")


# Índices conocidos dentro de la lista principal de un .cex "CEXv2.3 Residencial"
IDX_NOMBRE_EDIFICIO = 0
IDX_DIRECCION = 1
IDX_PROVINCIA_EDIFICIO = 3
IDX_MUNICIPIO_EDIFICIO = 4
IDX_CLIENTE_NOMBRE_NIF = 5
IDX_CLIENTE_DOMICILIO = 7
IDX_CLIENTE_TELEFONO = 8
IDX_CLIENTE_EMAIL = 9
IDX_TECNICO_RAZON_SOCIAL = 10
IDX_TECNICO_NOMBRE = 11
IDX_TECNICO_TELEFONO = 12
IDX_TECNICO_EMAIL = 13
IDX_EDIFICIO_POSTAL = 14
IDX_TECNICO_NIF = 19
IDX_TECNICO_NIF_ENTIDAD = 20
IDX_TECNICO_DOMICILIO = 21
IDX_TECNICO_MUNICIPIO = 22
IDX_TECNICO_PROVINCIA = 23
IDX_TECNICO_POSTAL = 24
IDX_TECNICO_TITULACION = 25


def _get(lst, idx, default=""):
    try:
        v = lst[idx]
    except (IndexError, TypeError):
        return default
    return v.strip() if isinstance(v, str) else default


def parse_cex(path):
    """Lee un .cex y devuelve un dict con los campos crudos identificados."""
    with open(path, "rb") as f:
        raw = f.read()
    raw = raw.replace(b"\r\n", b"\n")
    buf = io.BytesIO(raw)

    unpickler = SafeUnpickler(buf, encoding="latin-1")
    header = unpickler.load()
    datos = SafeUnpickler(buf, encoding="latin-1").load()

    if not isinstance(datos, list):
        raise ValueError("Formato de .cex inesperado (no es una lista de datos).")

    nombre_nif = _get(datos, IDX_CLIENTE_NOMBRE_NIF)
    cliente_nombre, cliente_nif = "", ""
    if "_" in nombre_nif:
        cliente_nombre, cliente_nif = nombre_nif.rsplit("_", 1)
    else:
        cliente_nombre = nombre_nif

    return {
        "formato": header,
        "edificio_nombre": _get(datos, IDX_NOMBRE_EDIFICIO),
        "edificio_direccion": _get(datos, IDX_DIRECCION),
        "edificio_provincia": _get(datos, IDX_PROVINCIA_EDIFICIO),
        "edificio_municipio": _get(datos, IDX_MUNICIPIO_EDIFICIO),
        "edificio_postal": _get(datos, IDX_EDIFICIO_POSTAL),
        "cliente_nombre": cliente_nombre.strip(),
        "cliente_nif": cliente_nif.strip(),
        "cliente_domicilio": _get(datos, IDX_CLIENTE_DOMICILIO),
        "cliente_telefono": _get(datos, IDX_CLIENTE_TELEFONO),
        "cliente_email": _get(datos, IDX_CLIENTE_EMAIL),
        "tecnico_razon_social": _get(datos, IDX_TECNICO_RAZON_SOCIAL),
        "tecnico_nombre": _get(datos, IDX_TECNICO_NOMBRE),
        "tecnico_telefono": _get(datos, IDX_TECNICO_TELEFONO),
        "tecnico_email": _get(datos, IDX_TECNICO_EMAIL),
        "tecnico_nif": _get(datos, IDX_TECNICO_NIF),
        "tecnico_nif_entidad": _get(datos, IDX_TECNICO_NIF_ENTIDAD),
        "tecnico_domicilio": _get(datos, IDX_TECNICO_DOMICILIO),
        "tecnico_municipio": _get(datos, IDX_TECNICO_MUNICIPIO),
        "tecnico_provincia": _get(datos, IDX_TECNICO_PROVINCIA),
        "tecnico_postal": _get(datos, IDX_TECNICO_POSTAL),
        "tecnico_titulacion": _get(datos, IDX_TECNICO_TITULACION),
    }


def find_sibling_xml(cex_path):
    stem = os.path.splitext(cex_path)[0]
    xml_path = stem + ".xml"
    return xml_path if os.path.isfile(xml_path) else None


def find_sibling_cex(xml_path):
    stem = os.path.splitext(xml_path)[0]
    cex_path = stem + ".cex"
    return cex_path if os.path.isfile(cex_path) else None


def _xml_text(node, path, default=""):
    if node is None:
        return default
    return (node.findtext(path, default=default) or default).strip()


def parse_xml(path):
    """Lee el .xml DatosEnergeticosDelEdificio y devuelve el mismo dict que
    parse_cex, salvo cliente_nombre/cliente_nif (no existen en el .xml)."""
    tree = ET.parse(path)
    root = tree.getroot()

    cert = root.find("DatosDelCertificador")
    ident = root.find("IdentificacionEdificio")
    visita = root.find("PruebasComprobacionesInspecciones/Visita")

    return {
        "formato": root.get("version", ""),
        "edificio_nombre": _xml_text(ident, "NombreDelEdificio"),
        "edificio_direccion": _xml_text(ident, "Direccion"),
        "edificio_provincia": _xml_text(ident, "Provincia"),
        "edificio_municipio": _xml_text(ident, "Municipio"),
        "edificio_postal": _xml_text(ident, "CodigoPostal"),
        "cliente_nombre": "",
        "cliente_nif": "",
        "tecnico_razon_social": _xml_text(cert, "RazonSocial"),
        "tecnico_nombre": _xml_text(cert, "NombreyApellidos"),
        "tecnico_telefono": _xml_text(cert, "Telefono"),
        "tecnico_email": _xml_text(cert, "Email"),
        "tecnico_nif": _xml_text(cert, "NIF"),
        "tecnico_nif_entidad": _xml_text(cert, "NIFEntidad"),
        "tecnico_domicilio": _xml_text(cert, "Domicilio"),
        "tecnico_municipio": _xml_text(cert, "Municipio"),
        "tecnico_provincia": _xml_text(cert, "Provincia"),
        "tecnico_postal": _xml_text(cert, "CodigoPostal"),
        "tecnico_titulacion": _xml_text(cert, "Titulacion"),
        "alcance": _xml_text(ident, "AlcanceInformacionXML"),
        "tipo_edificio": _xml_text(ident, "TipoDeEdificio"),
        "fecha_visita": _xml_text(visita, "FechaVisita"),
    }


def guess_grupo2(alcance):
    """Deduce la opción de trámite (radio Grupo2) a partir de AlcanceInformacionXML.
    Es sólo una sugerencia: el usuario debe confirmarla en el formulario."""
    a = (alcance or "").lower()
    if "proyecto" in a:
        if "modific" in a:
            return "1"
        if "anul" in a:
            return "2"
        return "0"
    if "anul" in a:
        return "5"
    if "renov" in a or "actualiz" in a:
        return "4"
    return "3"  # certificado de edificio terminado: caso más habitual


_TIPOS_VIA = {
    "CL", "AV", "AVDA", "PZ", "PZA", "PLAZA", "CTRA", "CM", "CAMINO", "PS", "PASEO",
    "TR", "TRAV", "TRAVESIA", "RD", "RONDA", "URB", "PJE", "PASAJE", "GL", "GLORIETA",
}


def split_direccion(direccion):
    """Divide 'CL MIESES 1' en (tipo_via, nombre_via, numero).
    Best-effort: si no encaja el patrón, devuelve todo en nombre_via."""
    direccion = (direccion or "").strip()
    if not direccion:
        return "", "", ""
    m = re.match(r"^(\S+)\s+(.+?)\s+(\d+\S*)\s*$", direccion)
    if m and m.group(1).upper().rstrip(".") in _TIPOS_VIA:
        return m.group(1), m.group(2), m.group(3)
    if m:
        # No reconocemos el tipo de vía, pero el patrón "tipo nombre numero" encaja igual
        return m.group(1), m.group(2), m.group(3)
    return "", direccion, ""
