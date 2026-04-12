import calendar
import os
import subprocess
import sys
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# -- Paleta de colores --
BG_APP        = "#F2F2F7"   # fondo general (estilo macOS)
BG_CARD       = "#FFFFFF"   # fondo de tarjetas
BG_HEADER     = "#1C1C1E"   # barra de título
FG_HEADER     = "#FFFFFF"
FG_TITLE      = "#1C1C1E"
FG_LABEL      = "#3A3A3C"
FG_MUTED      = "#8E8E93"
COLOR_ACCENT  = "#007AFF"
COLOR_HOVER   = "#0062CC"
COLOR_BORDER  = "#D1D1D6"

# -- Excel --
COLOR_HEADER_XL = "6C4C87"
COLOR_ROW_ODD   = "C2B4CA"
COLOR_ROW_EVN   = "DCD3E1"


def _card(parent, title=""):
    """Frame tipo tarjeta con borde sutil y esquinas simuladas."""
    outer = tk.Frame(parent, bg=COLOR_BORDER, padx=1, pady=1)
    inner = tk.Frame(outer, bg=BG_CARD, padx=16, pady=12)
    inner.pack(fill="x")
    if title:
        tk.Label(
            inner, text=title,
            font=("Helvetica", 10, "bold"),
            bg=BG_CARD, fg=FG_MUTED
        ).pack(anchor="w", pady=(0, 6))
    return outer, inner


class _HoverButton(tk.Button):
    """Botón con efecto hover."""
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self._normal_bg = kw.get("bg", COLOR_ACCENT)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _):
        self.config(bg=COLOR_HOVER)

    def _on_leave(self, _):
        self.config(bg=self._normal_bg)


class AppTurnosNativa:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Turnos")
        self.root.resizable(False, False)
        self.root.configure(bg=BG_APP)

        self._build_header()
        self._build_body()
        self._centrar_ventana(480, 460)

    # ------------------------------------------------------------------ #
    #  UI BUILD                                                            #
    # ------------------------------------------------------------------ #

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=BG_HEADER, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="Gestor de Limpieza",
            font=("Helvetica", 15, "bold"),
            bg=BG_HEADER, fg=FG_HEADER
        ).place(relx=0.5, rely=0.5, anchor="center")

    def _build_body(self):
        body = tk.Frame(self.root, bg=BG_APP)
        body.pack(fill="both", expand=True, padx=20, pady=18)

        # Defaults dinámicos
        today = datetime.today()
        m_end = today.month + 2
        y_end = today.year
        if m_end > 12:
            m_end -= 12
            y_end += 1
        last_day_end = calendar.monthrange(y_end, m_end)[1]

        # --- Tarjeta FECHA DE INICIO ---
        card_outer, card_ini = _card(body, "FECHA DE INICIO")
        card_outer.pack(fill="x", pady=(0, 10))

        row_ini = tk.Frame(card_ini, bg=BG_CARD)
        row_ini.pack(anchor="w")

        self.d_ini, self.m_ini, self.a_ini = self._date_row(
            row_ini,
            str(today.day).zfill(2),
            str(today.month).zfill(2),
            str(today.year),
        )

        # --- Tarjeta FECHA DE FIN ---
        card_outer2, card_fin = _card(body, "FECHA DE FIN")
        card_outer2.pack(fill="x", pady=(0, 10))

        row_fin = tk.Frame(card_fin, bg=BG_CARD)
        row_fin.pack(anchor="w")

        self.d_fin, self.m_fin, self.a_fin = self._date_row(
            row_fin,
            str(last_day_end).zfill(2),
            str(m_end).zfill(2),
            str(y_end),
        )

        # --- Preview de semanas ---
        self._preview_var = tk.StringVar(value="")
        tk.Label(
            body,
            textvariable=self._preview_var,
            font=("Helvetica", 11),
            bg=BG_APP, fg=FG_MUTED
        ).pack(pady=(4, 14))

        # --- Botón principal ---
        self.btn = _HoverButton(
            body,
            text="Generar y Guardar Excel",
            command=self.procesar,
            bg=COLOR_ACCENT, fg="#FFFFFF",
            font=("Helvetica", 13, "bold"),
            relief="flat", cursor="hand2",
            padx=20, pady=10,
            activebackground=COLOR_HOVER, activeforeground="#FFFFFF",
        )
        self.btn.pack(fill="x")

        # Bind para preview
        for widget in (self.d_ini, self.m_ini, self.a_ini,
                        self.d_fin, self.m_fin, self.a_fin):
            widget.bind("<<ComboboxSelected>>", lambda _: self._actualizar_preview())

        self._actualizar_preview()

    def _date_row(self, parent, day, month, year):
        """Fila compacta DD / MM / AAAA con estilo coherente."""
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Date.TCombobox",
            fieldbackground=BG_APP,
            background=BG_APP,
            foreground=FG_TITLE,
            selectbackground=COLOR_ACCENT,
            selectforeground="#FFFFFF",
            bordercolor=COLOR_BORDER,
            relief="flat",
        )

        combo_day   = ttk.Combobox(parent, values=[str(i).zfill(2) for i in range(1, 32)],
                                    width=3, state="readonly", style="Date.TCombobox")
        combo_month = ttk.Combobox(parent, values=[str(i).zfill(2) for i in range(1, 13)],
                                    width=3, state="readonly", style="Date.TCombobox")
        combo_year  = ttk.Combobox(parent, values=[str(i) for i in range(2024, 2036)],
                                    width=5, state="readonly", style="Date.TCombobox")

        combo_day.set(day)
        combo_month.set(month)
        combo_year.set(year)

        sep_kw = dict(bg=BG_CARD, fg=FG_MUTED, font=("Helvetica", 13))

        combo_day.pack(side="left", padx=(0, 2))
        tk.Label(parent, text="/", **sep_kw).pack(side="left")
        combo_month.pack(side="left", padx=(2, 2))
        tk.Label(parent, text="/", **sep_kw).pack(side="left")
        combo_year.pack(side="left", padx=(2, 0))

        return combo_day, combo_month, combo_year

    # ------------------------------------------------------------------ #
    #  HELPERS                                                             #
    # ------------------------------------------------------------------ #

    def _centrar_ventana(self, ancho, alto):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - ancho) // 2
        y = (self.root.winfo_screenheight() - alto)  // 2
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")

    def _actualizar_preview(self):
        inicio = self.obtener_fecha(self.d_ini.get(), self.m_ini.get(), self.a_ini.get())
        fin    = self.obtener_fecha(self.d_fin.get(), self.m_fin.get(), self.a_fin.get())
        if inicio and fin and fin > inicio:
            curr = inicio - timedelta(days=inicio.weekday())
            n = 0
            while curr <= fin:
                n += 1
                curr += timedelta(days=7)
            self._preview_var.set(f"{n} semana{'s' if n != 1 else ''} a generar")
        else:
            self._preview_var.set("Comprueba las fechas")

    def obtener_fecha(self, d, m, a):
        try:
            return datetime(int(a), int(m), int(d))
        except ValueError:
            return None

    # ------------------------------------------------------------------ #
    #  LÓGICA DE TURNOS (sin cambios)                                      #
    # ------------------------------------------------------------------ #

    def generar_juan(self, semanas):
        estancias = ["Hab. Juan", "Hab. AP", "Baño Juan"]
        staff = ["Lina", "Eva", "Guillem", "Valentina", "Miguel"]

        dias_disponibles = {
            "Lina":      ["Lunes"],
            "Eva":       ["Miércoles", "Jueves", "Viernes"],
            "Valentina": ["Martes", "Miércoles", "Jueves", "Viernes"],
            "Miguel":    ["Lunes", "Martes"],
            "Guillem":   ["Sábado"],
        }

        rows = []
        idx = 0
        idx_comunes = 0
        miguel_dia_idx = 0

        for i, (f_ini, f_fin) in enumerate(semanas):
            row = {
                "Semana":    f"{f_ini.strftime('%d/%m')} -\n{f_fin.strftime('%d/%m')}",
                "Lunes":     "-", "Martes":    "-", "Miércoles": "-",
                "Jueves":    "-", "Viernes":   "-", "Sábado":    "-",
            }

            p1 = staff[idx % 5]
            p2 = staff[(idx + 1) % 5]
            p3 = staff[(idx + 2) % 5]
            idx += 3

            asignaciones = [(p1, estancias[0]), (p2, estancias[1]), (p3, estancias[2])]

            dias_ocupados = set()
            for persona, tarea in asignaciones:
                if persona == "Miguel":
                    dias_p = ["Lunes", "Martes"] if miguel_dia_idx % 2 == 0 else ["Martes", "Lunes"]
                    miguel_dia_idx += 1
                else:
                    dias_p = dias_disponibles[persona]

                dia_elegido = dias_p[0]
                for d in dias_p:
                    if d not in dias_ocupados:
                        dia_elegido = d
                        break
                dias_ocupados.add(dia_elegido)

                etiqueta = f"{tarea} ({persona})"
                if row[dia_elegido] == "-":
                    row[dia_elegido] = etiqueta
                else:
                    row[dia_elegido] += f"\n| {etiqueta}"

            juan_comunes_tareas = {
                0: [("Miércoles", "Entrada")],
                1: [("Viernes",   "Cocina")],
                2: [("Lunes",     "Salón"), ("Martes", "Cocina")],
            }[i % 3]

            for dia, tarea in juan_comunes_tareas:
                for offset in range(len(staff)):
                    candidato = staff[(idx_comunes + offset) % len(staff)]
                    if dia in dias_disponibles[candidato]:
                        etiqueta = f"{tarea} ({candidato})"
                        if row[dia] == "-":
                            row[dia] = etiqueta
                        else:
                            row[dia] += f"\n| {etiqueta}"
                        idx_comunes = (idx_comunes + offset + 1) % len(staff)
                        break

            rows.append(row)

        return pd.DataFrame(rows)

    def generar_comunes(self, semanas):
        staff = ["Lina", "Angie", "Juan"]
        rows = []

        for i, (f_ini, f_fin) in enumerate(semanas):
            p_A = staff[i % 3]
            p_B = staff[(i + 1) % 3]
            p_C = staff[(i + 2) % 3]

            row = {
                "Semana":    f"{f_ini.strftime('%d/%m')} -\n{f_fin.strftime('%d/%m')}",
                "Lunes":     f"Salón ({p_A})",
                "Martes":    f"Cocina ({p_A})",
                "Miércoles": f"Entrada ({p_C})",
                "Viernes":   f"Cocina ({p_B})",
            }
            rows.append(row)

        return pd.DataFrame(rows)

    def aplicar_estilos_excel(self, ruta, df_comunes, df_juan):
        with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
            df_comunes.to_excel(writer, sheet_name='Zonas Comunes', index=False)
            df_juan.to_excel(writer, sheet_name='Zonas de Juan', index=False)

            workbook = writer.book

            fill_header = PatternFill(start_color=COLOR_HEADER_XL, end_color=COLOR_HEADER_XL, fill_type="solid")
            fill_odd    = PatternFill(start_color=COLOR_ROW_ODD,   end_color=COLOR_ROW_ODD,   fill_type="solid")
            fill_even   = PatternFill(start_color=COLOR_ROW_EVN,   end_color=COLOR_ROW_EVN,   fill_type="solid")

            font_header = Font(color="FFFFFF", bold=True,  size=12)
            font_normal = Font(color="000000", bold=False, size=12)

            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

            white_side   = Side(border_style="thin", color="FFFFFF")
            border_white = Border(left=white_side, right=white_side, top=white_side, bottom=white_side)

            for sheet_name in ['Zonas Comunes', 'Zonas de Juan']:
                ws = workbook[sheet_name]

                for col_idx in range(1, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = 20

                for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row,
                                                            min_col=1, max_col=ws.max_column)):
                    ws.row_dimensions[row[0].row].height = 45

                    for cell in row:
                        cell.alignment = align_center
                        cell.border = border_white

                        if row_idx == 0 or cell.column == 1:
                            cell.fill = fill_header
                            cell.font = font_header
                        else:
                            cell.font = font_normal
                            cell.fill = fill_odd if row_idx % 2 != 0 else fill_even

    def procesar(self):
        d_i, m_i, a_i = self.d_ini.get(), self.m_ini.get(), self.a_ini.get()
        d_f, m_f, a_f = self.d_fin.get(), self.m_fin.get(), self.a_fin.get()

        inicio = self.obtener_fecha(d_i, m_i, a_i)
        fin    = self.obtener_fecha(d_f, m_f, a_f)

        if not inicio or not fin:
            messagebox.showerror("Error", "La fecha seleccionada no existe.")
            return
        if fin <= inicio:
            messagebox.showerror("Error", "La fecha de fin debe ser posterior a la de inicio.")
            return

        curr = inicio - timedelta(days=inicio.weekday())

        semanas = []
        while curr <= fin:
            semanas.append((curr, curr + timedelta(days=6)))
            curr += timedelta(days=7)

        nombre_archivo_sugerido = (
            f"Turnos limpieza {inicio.strftime('%d-%m-%Y')} a {fin.strftime('%d-%m-%Y')}.xlsx"
        )

        ruta = filedialog.asksaveasfilename(
            initialfile=nombre_archivo_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )

        if not ruta:
            return

        self.btn.config(state="disabled", text="Generando…")
        self.root.config(cursor="watch")
        self.root.update()

        try:
            df_c = self.generar_comunes(semanas)
            df_j = self.generar_juan(semanas)
            self.aplicar_estilos_excel(ruta, df_c, df_j)

            if messagebox.askyesno("Listo", "El Excel se ha generado correctamente.\n¿Deseas abrirlo ahora?"):
                if sys.platform == "darwin":
                    subprocess.call(["open", ruta])
                elif sys.platform == "win32":
                    os.startfile(ruta)
                else:
                    subprocess.call(["xdg-open", ruta])

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar: {e}")
        finally:
            self.btn.config(state="normal", text="Generar y Guardar Excel")
            self.root.config(cursor="")


if __name__ == "__main__":
    root = tk.Tk()
    AppTurnosNativa(root)
    root.mainloop()
