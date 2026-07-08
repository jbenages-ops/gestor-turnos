"""Relleno del formulario PDF 'HOJA ENCARGO CyL.pdf' a partir de un diccionario
de valores (nombre de campo AcroForm -> texto)."""
import pypdf
from pypdf.generic import BooleanObject, NameObject

# Todos los campos de texto/choice del formulario que la app puede rellenar.
# Los que no se indiquen se dejan en blanco (evita arrastrar valores de un
# rellenado anterior si se reutiliza la misma plantilla).
CAMPOS_TEXTO = [
    "nombre técnico", "NIF 1", "domiciliado en", "razón social",
    "Nombre y apellidos", "con NIF n_2", "domicilio",
    "Razón social", "con NIF n_3", "domicilio 2", "representante legal", "con NIF n_4",
    "VÍA", "nkm", "bloque", "planta", "puerta",
    "localidad", "provincia", "escalera", "postal",
    "Fdo_2", "Fdo",
]

CAMPOS_CHOICE = ["uso del edificio 2", "promotor/prop", "día", "mes", "AÑO"]


def fill_pdf(template_path, values, output_path):
    """values: dict con claves de CAMPOS_TEXTO/CAMPOS_CHOICE, más opcionalmente
    'Grupo2' con el número de opción ('0'..'5')."""
    writer = pypdf.PdfWriter(clone_from=template_path)

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
