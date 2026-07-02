"""Relleno del formulario PDF 'HOJA ENCARGO CyL.pdf' a partir de un diccionario
de valores (nombre de campo AcroForm -> texto)."""
import pypdf
from pypdf.generic import BooleanObject, NameObject, TextStringObject

# Todos los campos de texto/choice del formulario que la app puede rellenar.
# Los que no se indiquen se dejan en blanco (evita arrastrar valores de un
# rellenado anterior si se reutiliza la misma plantilla).
CAMPOS_TEXTO = [
    "nombre técnico", "NIF 1", "domiciliado en", "razón social",
    "Nombre y apellidos", "con NIF n_2", "domicilio", "nombre vía",
    "Razón social", "con NIF n_3", "domicilio 2", "representante legal", "con NIF n_4",
    "VÍA", "nkm", "bloque", "planta", "puerta",
    "localidad", "provincia", "escalera", "postal",
    "Fdo_2", "Fdo",
]

CAMPOS_CHOICE = ["uso del edificio 2", "promotor/prop", "día", "mes", "AÑO"]


def _separar_nombre_via(writer):
    """En la plantilla, el campo 'domicilio' tiene dos widgets: el 'domiciliado
    en' del cliente y el 'nombre de la vía' del inmueble, de modo que ambos
    muestran siempre el mismo texto. Convierte el widget inferior (nombre de la
    vía) en un campo independiente llamado 'nombre vía' para poder rellenarlos
    por separado."""
    fields = writer._root_object["/AcroForm"]["/Fields"]
    for f in fields:
        campo = f.get_object()
        if campo.get("/T") != "domicilio":
            continue
        kids = campo.get("/Kids")
        if not kids or len(kids) < 2:
            return  # plantilla sin el widget duplicado: nada que separar
        # El widget del "nombre de la vía" es el más bajo en la página
        kid_via = min(kids, key=lambda k: float(k.get_object()["/Rect"][1]))
        widget = kid_via.get_object()
        widget[NameObject("/T")] = TextStringObject("nombre vía")
        widget[NameObject("/FT")] = NameObject("/Tx")
        del widget["/Parent"]
        kids.remove(kid_via)
        fields.append(kid_via)
        return


def fill_pdf(template_path, values, output_path):
    """values: dict con claves de CAMPOS_TEXTO/CAMPOS_CHOICE, más opcionalmente
    'Grupo2' con el número de opción ('0'..'5')."""
    writer = pypdf.PdfWriter(clone_from=template_path)
    _separar_nombre_via(writer)

    campos = {}
    for name in CAMPOS_TEXTO + CAMPOS_CHOICE:
        campos[name] = values.get(name, "") or ""

    grupo2 = values.get("Grupo2")
    if grupo2 is not None and grupo2 != "":
        campos["Grupo2"] = f"/{grupo2}" if not str(grupo2).startswith("/") else grupo2

    for page in writer.pages:
        writer.update_page_form_field_values(page, campos, auto_regenerate=False)

    root = writer._root_object
    if root.get("/AcroForm"):
        root["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    with open(output_path, "wb") as f:
        writer.write(f)
