import os
import subprocess
import sys
from datetime import date, datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from cex_parser import (
    parse_xml, parse_cex, find_sibling_cex, guess_tramite,
    split_direccion,
)
from pdf_filler import fill_pdf

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PLANTILLA_PDF = os.path.join(APP_DIR, "MODELO REPRESENTACION CyL.pdf")

# -- Paleta de colores (coherente con el resto de apps de Proyectos/) --
BG_APP        = "#F2F2F7"
BG_CARD       = "#FFFFFF"
BG_HEADER     = "#1C1C1E"
FG_HEADER     = "#FFFFFF"
FG_TITLE      = "#1C1C1E"
FG_LABEL      = "#3A3A3C"
FG_MUTED      = "#8E8E93"
COLOR_ACCENT  = "#007AFF"
COLOR_HOVER   = "#0062CC"
COLOR_BORDER  = "#D1D1D6"
COLOR_WARN    = "#FF9500"

USOS_EDIFICIO = [
    "Bloque de viviendas completo",
    "Vivienda individual en bloque",
    "Vivienda unifamiliar aislada",
    "Viviendas unifamiliares adosadas",
    "Viviendas unifamiliares pareadas",
]

# Trámite: cada casilla del formulario (clave lógica -> etiqueta que se muestra).
# Las claves coinciden con pdf_filler.CAMPOS_TRAMITE y con guess_tramite().
TRAMITES = [
    ("existente",      "Certificado de eficiencia energética de edificio existente"),
    ("proyecto",       "Certificado de eficiencia energética de proyecto"),
    ("obra_terminada", "Certificado de eficiencia energética de obra terminada"),
    ("modificacion",   "Modificación de certificado de eficiencia energética inscrito"),
    ("renovacion",     "Renovación / actualización de certificado de eficiencia energética de edificio"),
    ("anul_proyecto",  "Anulación de certificado de eficiencia energética de proyecto"),
    ("anul_edificio",  "Anulación de certificado de eficiencia energética de edificio"),
]

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
          "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# Campos de la propietaria/promotora según sea persona física o empresa. Se
# usan para vaciar el bloque no aplicable al generar el PDF.
CAMPOS_FISICA = ["Nombre y apellidos", "con NIF n", "dirección"]
CAMPOS_EMPRESA = ["Razón social", "con NIF n_2",
                  "representada legalmente por 1 se deberá acreditar",
                  "dicha representación", "con NIF n_3"]


def _guess_uso(tipo_edificio):
    """Sugiere el uso del edificio a partir de TipoDeEdificio del .xml."""
    t = (tipo_edificio or "").lower().replace(" ", "")
    if "bloque" in t and ("completo" in t or "edificio" in t):
        return USOS_EDIFICIO[0]
    if "individual" in t or "enbloque" in t:
        return USOS_EDIFICIO[1]
    if "adosad" in t:
        return USOS_EDIFICIO[3]
    if "paread" in t:
        return USOS_EDIFICIO[4]
    if "unifamiliar" in t or "aislad" in t:
        return USOS_EDIFICIO[2]
    return USOS_EDIFICIO[1]


def _card(parent, title=""):
    outer = tk.Frame(parent, bg=COLOR_BORDER, padx=1, pady=1)
    inner = tk.Frame(outer, bg=BG_CARD, padx=16, pady=12)
    inner.pack(fill="x")
    if title:
        tk.Label(inner, text=title, font=("Helvetica", 10, "bold"),
                  bg=BG_CARD, fg=FG_MUTED).pack(anchor="w", pady=(0, 8))
    return outer, inner


class _HoverButton(tk.Button):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self._normal_bg = kw.get("bg", COLOR_ACCENT)
        self.bind("<Enter>", lambda _: self.config(bg=COLOR_HOVER))
        self.bind("<Leave>", lambda _: self.config(bg=self._normal_bg))


class _ScrollableFrame(tk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        canvas = tk.Canvas(self, bg=BG_APP, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.body = tk.Frame(canvas, bg=BG_APP)

        self.body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        window_id = canvas.create_window((0, 0), window=self.body, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(window_id, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _wheel(event):
            delta = event.delta
            if sys.platform == "darwin":
                canvas.yview_scroll(-1 * delta, "units")
            else:
                canvas.yview_scroll(-1 * int(delta / 120), "units")

        canvas.bind_all("<MouseWheel>", _wheel)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))


def _entry_row(parent, label, var, width=40):
    row = tk.Frame(parent, bg=BG_CARD)
    row.pack(fill="x", pady=3)
    tk.Label(row, text=label, font=("Helvetica", 11), bg=BG_CARD, fg=FG_LABEL,
              width=22, anchor="w").pack(side="left")
    tk.Entry(row, textvariable=var, font=("Helvetica", 11), width=width,
              relief="solid", bd=1).pack(side="left", fill="x", expand=True)
    return row


class RepresentacionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Representación Voluntaria CyL")
        self.root.configure(bg=BG_APP)
        self.root.geometry("720x820")

        self.xml_path = None
        self.vars = {}
        self.tramite_vars = {}
        self.es_empresa_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Selecciona un archivo .xml para empezar.")

        self._build_header()
        self._build_body()

    # ------------------------------------------------------------------ UI
    def _build_header(self):
        hdr = tk.Frame(self.root, bg=BG_HEADER, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="Modelo de Representación Voluntaria — CyL",
                  font=("Helvetica", 14, "bold"), bg=BG_HEADER, fg=FG_HEADER
                  ).place(relx=0.5, rely=0.5, anchor="center")

    def _build_body(self):
        scroll = _ScrollableFrame(self.root, bg=BG_APP)
        scroll.pack(fill="both", expand=True)
        body = tk.Frame(scroll.body, bg=BG_APP)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # --- Selección de archivo ---
        outer, card = _card(body, "1. ARCHIVO .XML")
        outer.pack(fill="x", pady=(0, 10))
        row = tk.Frame(card, bg=BG_CARD)
        row.pack(fill="x")
        _HoverButton(row, text="Elegir .xml…", command=self._elegir_xml,
                      bg=COLOR_ACCENT, fg="white", relief="flat", cursor="hand2",
                      font=("Helvetica", 11, "bold"), padx=14, pady=6).pack(side="left")
        tk.Label(row, textvariable=self.status_var, font=("Helvetica", 10),
                  bg=BG_CARD, fg=FG_MUTED, wraplength=520, justify="left"
                  ).pack(side="left", padx=(12, 0))

        # --- Propietaria / promotora (de una parte) ---
        outer, card = _card(body, "2. PROPIETARIA / PROMOTORA")
        outer.pack(fill="x", pady=(0, 10))
        tk.Checkbutton(card, text="Es persona jurídica (empresa)",
                         variable=self.es_empresa_var, bg=BG_CARD, fg=FG_LABEL,
                         font=("Helvetica", 10), command=self._toggle_empresa
                         ).pack(anchor="w", pady=(0, 6))

        self.frame_persona_fisica = tk.Frame(card, bg=BG_CARD)
        self.frame_persona_fisica.pack(fill="x")
        for key, label in [("Nombre y apellidos", "Nombre y apellidos"),
                             ("con NIF n", "NIF"),
                             ("dirección", "Domiciliada en")]:
            self.vars[key] = tk.StringVar()
            _entry_row(self.frame_persona_fisica, label, self.vars[key])

        self.frame_persona_juridica = tk.Frame(card, bg=BG_CARD)
        for key, label in [("Razón social", "Razón social"),
                             ("con NIF n_2", "NIF empresa"),
                             ("representada legalmente por 1 se deberá acreditar", "Domicilio empresa"),
                             ("dicha representación", "Representante legal"),
                             ("con NIF n_3", "NIF representante")]:
            self.vars[key] = tk.StringVar()
            _entry_row(self.frame_persona_juridica, label, self.vars[key])

        # --- Persona representante / técnico (de otra parte) ---
        outer, card = _card(body, "3. PERSONA REPRESENTANTE (TÉCNICO)")
        outer.pack(fill="x", pady=(0, 10))
        for key, label in [
            ("De otra parte la persona representante cuyo nombre y apellidos son", "Nombre y apellidos"),
            ("con NIF n_4", "NIF"),
            ("domiciliada en", "Domiciliada en"),
            ("razón social", "Razón social"),
            ("con NIF n_5", "NIF razón social"),
        ]:
            self.vars[key] = tk.StringVar()
            _entry_row(card, label, self.vars[key])

        # --- Inmueble ---
        outer, card = _card(body, "4. INMUEBLE")
        outer.pack(fill="x", pady=(0, 10))
        for key, label in [("tipo de vía", "Tipo de vía (CL, AV…)"),
                             ("nombre de la vía", "Nombre de la vía"),
                             ("nkm", "Nº / Km"),
                             ("bloque", "Bloque"),
                             ("escalera", "Escalera"),
                             ("planta", "Planta"),
                             ("puerta", "Puerta"),
                             ("localidad", "Localidad"),
                             ("provincia", "Provincia"),
                             ("código postal", "Código postal")]:
            self.vars[key] = tk.StringVar()
            _entry_row(card, label, self.vars[key], width=30)

        uso_row = tk.Frame(card, bg=BG_CARD)
        uso_row.pack(fill="x", pady=(6, 0))
        tk.Label(uso_row, text="Uso del edificio (revisar)", font=("Helvetica", 11, "bold"),
                  bg=BG_CARD, fg=COLOR_WARN, width=22, anchor="w").pack(side="left")
        self.vars["uso del edificio 2"] = tk.StringVar(value=USOS_EDIFICIO[1])
        ttk.Combobox(uso_row, textvariable=self.vars["uso del edificio 2"],
                      values=USOS_EDIFICIO, width=38).pack(side="left", fill="x", expand=True)

        # --- Trámite ---
        outer, card = _card(body, "5. TRÁMITE (sugerido, revisar)")
        outer.pack(fill="x", pady=(0, 10))
        for key, label in TRAMITES:
            self.tramite_vars[key] = tk.BooleanVar(value=False)
            tk.Checkbutton(card, text=label, variable=self.tramite_vars[key],
                             bg=BG_CARD, fg=FG_LABEL, font=("Helvetica", 10),
                             anchor="w", justify="left", wraplength=620
                             ).pack(fill="x", anchor="w")

        # --- Lugar y fecha de firma ---
        outer, card = _card(body, "6. LUGAR Y FECHA DE FIRMA")
        outer.pack(fill="x", pady=(0, 10))
        self.vars["ciudad"] = tk.StringVar()
        _entry_row(card, "Ciudad", self.vars["ciudad"])

        hoy = date.today()
        fecha_row = tk.Frame(card, bg=BG_CARD)
        fecha_row.pack(fill="x", pady=3)
        tk.Label(fecha_row, text="Fecha", font=("Helvetica", 11),
                  bg=BG_CARD, fg=FG_LABEL, width=22, anchor="w").pack(side="left")
        self.vars["día"] = tk.StringVar(value=str(hoy.day))
        self.vars["mes"] = tk.StringVar(value=MESES[hoy.month - 1])
        self.vars["año"] = tk.StringVar(value=str(hoy.year)[-2:])
        ttk.Combobox(fecha_row, textvariable=self.vars["día"],
                      values=[str(i) for i in range(1, 32)], width=4, state="readonly").pack(side="left")
        ttk.Combobox(fecha_row, textvariable=self.vars["mes"],
                      values=MESES, width=11, state="readonly").pack(side="left", padx=4)
        tk.Label(fecha_row, text="de 20", font=("Helvetica", 11),
                  bg=BG_CARD, fg=FG_LABEL).pack(side="left")
        ttk.Combobox(fecha_row, textvariable=self.vars["año"],
                      values=[str(y)[-2:] for y in range(2018, 2036)], width=4, state="readonly").pack(side="left")

        # --- Botón generar ---
        self.btn_generar = _HoverButton(
            body, text="Generar PDF", command=self._generar,
            bg=COLOR_ACCENT, fg="white", font=("Helvetica", 13, "bold"),
            relief="flat", cursor="hand2", padx=20, pady=10,
            activebackground=COLOR_HOVER, activeforeground="white",
        )
        self.btn_generar.pack(fill="x", pady=(4, 20))

        self._toggle_empresa()

    def _toggle_empresa(self):
        if self.es_empresa_var.get():
            self.frame_persona_fisica.pack_forget()
            self.frame_persona_juridica.pack(fill="x")
        else:
            self.frame_persona_juridica.pack_forget()
            self.frame_persona_fisica.pack(fill="x")

    # ------------------------------------------------------------- lógica
    def _elegir_xml(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo .xml del certificado",
            filetypes=[("XML", "*.xml"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            datos = parse_xml(path)
        except Exception as e:
            messagebox.showerror("Error al leer el .xml", str(e))
            return

        # El .xml no incluye la identidad del cliente; si hay un .cex con el
        # mismo nombre en la misma carpeta, se usa sólo para ese dato.
        cliente_encontrado = False
        cex_path = find_sibling_cex(path)
        if cex_path:
            try:
                cex_datos = parse_cex(cex_path)
                datos["cliente_nombre"] = cex_datos["cliente_nombre"]
                datos["cliente_nif"] = cex_datos["cliente_nif"]
                datos["cliente_domicilio"] = cex_datos["cliente_domicilio"]
                cliente_encontrado = bool(cex_datos["cliente_nombre"])
            except Exception:
                pass

        self.xml_path = path
        fecha_ok = self._rellenar_desde_datos(datos)
        avisos = []
        if cliente_encontrado:
            avisos.append("propietaria tomada del .cex")
        else:
            avisos.append("nombre/NIF de la propietaria no disponibles: rellénalos a mano")
        if not fecha_ok:
            avisos.append("fecha no encontrada: revísala")
        self.status_var.set(
            f"Cargado: {os.path.basename(path)}. Revisa 'Trámite' y 'Uso del edificio'. "
            + "; ".join(avisos) + "."
        )

    def _rellenar_desde_datos(self, d):
        # Persona representante (el técnico certificador)
        self.vars["De otra parte la persona representante cuyo nombre y apellidos son"].set(d["tecnico_nombre"])
        self.vars["con NIF n_4"].set(d["tecnico_nif"])
        self.vars["domiciliada en"].set(
            f"{d['tecnico_domicilio']} - {d['tecnico_municipio']}".strip(" -"))
        self.vars["razón social"].set(d["tecnico_razon_social"])
        self.vars["con NIF n_5"].set(d["tecnico_nif_entidad"])

        # Propietaria / promotora (persona física por defecto)
        self.vars["Nombre y apellidos"].set(d["cliente_nombre"])
        self.vars["con NIF n"].set(d["cliente_nif"])
        # El domicilio de la propietaria no viene en el .xml; se usa el del .cex
        # si existe y, si no, el del inmueble como valor de partida.
        self.vars["dirección"].set(d.get("cliente_domicilio") or d["edificio_direccion"])
        self.es_empresa_var.set(False)
        self._toggle_empresa()

        # Inmueble
        via, nombre_via, nkm = split_direccion(d["edificio_direccion"])
        self.vars["tipo de vía"].set(via)
        self.vars["nombre de la vía"].set(nombre_via)
        self.vars["nkm"].set(nkm)
        self.vars["localidad"].set(d["edificio_municipio"])
        self.vars["provincia"].set(d["edificio_provincia"].upper())
        self.vars["código postal"].set(d["edificio_postal"])
        self.vars["uso del edificio 2"].set(_guess_uso(d.get("tipo_edificio")))

        # Trámite sugerido (una casilla)
        sugerido = guess_tramite(d.get("alcance"))
        for key, var in self.tramite_vars.items():
            var.set(key == sugerido)

        # Lugar y fecha de firma
        self.vars["ciudad"].set(d["edificio_municipio"])
        fecha = d.get("fecha_visita", "")
        if fecha:
            try:
                dt = datetime.strptime(fecha.strip(), "%d/%m/%Y")
                self.vars["día"].set(str(dt.day))
                self.vars["mes"].set(MESES[dt.month - 1])
                self.vars["año"].set(str(dt.year)[-2:])
                return True
            except ValueError:
                pass
        self.vars["día"].set("")
        self.vars["mes"].set("")
        self.vars["año"].set("")
        return False

    def _generar(self):
        if not self.xml_path:
            messagebox.showwarning("Falta el .xml", "Primero elige un archivo .xml.")
            return

        values = {k: v.get() for k, v in self.vars.items()}
        values["tramites"] = [k for k, v in self.tramite_vars.items() if v.get()]

        if self.es_empresa_var.get():
            for k in CAMPOS_FISICA:
                values[k] = ""
        else:
            for k in CAMPOS_EMPRESA:
                values[k] = ""

        cliente = values.get("Nombre y apellidos") or values.get("Razón social") or "cliente"
        sugerido = f"Representacion CyL - {cliente}.pdf"
        destino = filedialog.asksaveasfilename(
            initialfile=sugerido, defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not destino:
            return

        try:
            fill_pdf(PLANTILLA_PDF, values, destino)
        except Exception as e:
            messagebox.showerror("Error al generar el PDF", str(e))
            return

        if messagebox.askyesno("Listo", "PDF generado correctamente.\n¿Deseas abrirlo ahora?"):
            if sys.platform == "darwin":
                subprocess.call(["open", destino])
            elif sys.platform == "win32":
                os.startfile(destino)
            else:
                subprocess.call(["xdg-open", destino])


if __name__ == "__main__":
    root = tk.Tk()
    RepresentacionApp(root)
    root.mainloop()
