"""Relleno del formulario PDF 'MODELO REPRESENTACION CyL.pdf' (Modelo de
representación voluntaria, Decreto 5/2026) a partir de un diccionario de
valores (nombre de campo AcroForm -> texto)."""
import pypdf
from pypdf.generic import BooleanObject, NameObject

# Todos los campos de texto del formulario que la app puede rellenar. Los que
# no se indiquen se dejan en blanco (evita arrastrar valores de un rellenado
# anterior si se reutiliza la misma plantilla). Las claves son los nombres
# reales de los campos AcroForm, aunque algunos sean poco descriptivos.
CAMPOS_TEXTO = [
    # De una parte: propietaria/promotora — persona física
    "Nombre y apellidos", "con NIF n", "dirección",
    # De una parte: propietaria/promotora — otros casos (empresa)
    "Razón social", "con NIF n_2",
    "representada legalmente por 1 se deberá acreditar",  # domicilio de la empresa
    "dicha representación",                                # representante legal
    "con NIF n_3",                                         # NIF del representante legal
    # De otra parte: persona representante (el técnico)
    "De otra parte la persona representante cuyo nombre y apellidos son",
    "con NIF n_4", "domiciliada en", "razón social", "con NIF n_5",
    # Inmueble
    "tipo de vía", "nombre de la vía", "nkm", "bloque", "escalera",
    "planta", "puerta", "localidad", "provincia", "código postal",
    "uso del edificio 2",
    # Lugar, fecha y firmas
    "ciudad", "día", "mes", "año", "Fdo", "Fdo_2",
]

# Trámite: casillas de verificación (antes eran opciones de un radio). La clave
# lógica -> nombre del campo AcroForm de la casilla correspondiente.
CAMPOS_TRAMITE = {
    "existente":      "Certificado de eficiencia energética de edificio existente",
    "proyecto":       "Certificado de eficiencia energética de proyecto",
    "obra_terminada": "Certificado de eficiencia energética de obra terminada",
    "modificacion":   "Modificación de certificado de eficiencia energética inscrito",
    "renovacion":     "Renovación  actualización de certificado de eficiencia energética de edificio",
    "anul_proyecto":  "Anulación de certificado de eficiencia energética de proyecto",
    "anul_edificio":  "Anulación de certificado de eficiencia energética de edificio",
}


def fill_pdf(template_path, values, output_path):
    """values: dict con claves de CAMPOS_TEXTO, más opcionalmente 'tramites'
    con un iterable de claves de CAMPOS_TRAMITE ('existente', 'proyecto', ...)
    para marcar esas casillas."""
    writer = pypdf.PdfWriter(clone_from=template_path)

    campos = {}
    for name in CAMPOS_TEXTO:
        campos[name] = values.get(name, "") or ""

    seleccionados = set(values.get("tramites") or [])
    for clave, field_name in CAMPOS_TRAMITE.items():
        campos[field_name] = NameObject("/On") if clave in seleccionados else NameObject("/Off")

    for page in writer.pages:
        writer.update_page_form_field_values(page, campos, auto_regenerate=False)

    root = writer._root_object
    if root.get("/AcroForm"):
        root["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    with open(output_path, "wb") as f:
        writer.write(f)
