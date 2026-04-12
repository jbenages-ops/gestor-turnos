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
COLOR_BTN_BG  = "#007AFF"
COLOR_BTN_FG  = "#FFFFFF"
COLOR_HEADER  = "6C4C87"
COLOR_ROW_ODD = "C2B4CA"
COLOR_ROW_EVN = "DCD3E1"


class AppTurnosNativa:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Limpieza - Horarios Reales")
        self.root.resizable(False, False)

        tk.Label(root, text="Configuración de Calendario", font=('Arial', 14, 'bold')).pack(pady=15)

        # Defaults dinámicos: hoy → último día del mes +2
        today = datetime.today()
        m_end = today.month + 2
        y_end = today.year
        if m_end > 12:
            m_end -= 12
            y_end += 1
        last_day_end = calendar.monthrange(y_end, m_end)[1]

        # --- FECHA DE INICIO ---
        frame_ini = tk.LabelFrame(root, text=" Fecha de Inicio ", padx=10, pady=10)
        frame_ini.pack(pady=8, fill="x", padx=20)

        self.d_ini = ttk.Combobox(frame_ini, values=[str(i).zfill(2) for i in range(1, 32)], width=3, state="readonly")
        self.m_ini = ttk.Combobox(frame_ini, values=[str(i).zfill(2) for i in range(1, 13)], width=3, state="readonly")
        self.a_ini = ttk.Combobox(frame_ini, values=[str(i) for i in range(2026, 2035)], width=5, state="readonly")

        self.d_ini.set(str(today.day).zfill(2))
        self.m_ini.set(str(today.month).zfill(2))
        self.a_ini.set(str(today.year))

        self.d_ini.pack(side="left", padx=5)
        tk.Label(frame_ini, text="/").pack(side="left")
        self.m_ini.pack(side="left", padx=5)
        tk.Label(frame_ini, text="/").pack(side="left")
        self.a_ini.pack(side="left", padx=5)

        # --- FECHA DE FIN ---
        frame_fin = tk.LabelFrame(root, text=" Fecha de Fin ", padx=10, pady=10)
        frame_fin.pack(pady=8, fill="x", padx=20)

        self.d_fin = ttk.Combobox(frame_fin, values=[str(i).zfill(2) for i in range(1, 32)], width=3, state="readonly")
        self.m_fin = ttk.Combobox(frame_fin, values=[str(i).zfill(2) for i in range(1, 13)], width=3, state="readonly")
        self.a_fin = ttk.Combobox(frame_fin, values=[str(i) for i in range(2026, 2035)], width=5, state="readonly")

        self.d_fin.set(str(last_day_end).zfill(2))
        self.m_fin.set(str(m_end).zfill(2))
        self.a_fin.set(str(y_end))

        self.d_fin.pack(side="left", padx=5)
        tk.Label(frame_fin, text="/").pack(side="left")
        self.m_fin.pack(side="left", padx=5)
        tk.Label(frame_fin, text="/").pack(side="left")
        self.a_fin.pack(side="left", padx=5)

        # --- PREVIEW DE SEMANAS ---
        self._preview_var = tk.StringVar(value="")
        tk.Label(root, textvariable=self._preview_var,
                 font=('Arial', 10), fg="#555555").pack(pady=4)

        # --- BOTÓN ---
        self.btn = tk.Button(
            root, text="Generar y Guardar Excel", command=self.procesar,
            bg=COLOR_BTN_BG, fg=COLOR_BTN_FG, padx=15, pady=6,
            font=('Arial', 12, 'bold'), relief="flat", cursor="hand2"
        )
        self.btn.pack(pady=15)

        # Actualizar preview cuando cambien las fechas
        for widget in (self.d_ini, self.m_ini, self.a_ini, self.d_fin, self.m_fin, self.a_fin):
            widget.bind("<<ComboboxSelected>>", lambda _: self._actualizar_preview())

        self._actualizar_preview()
        self._centrar_ventana(500, 430)

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

    def generar_juan(self, semanas):
        estancias = ["Hab. Juan", "Hab. AP", "Baño Juan"]
        staff = ["Lina", "Dora", "Pablo", "Valentina", "Miguel"]

        # Días disponibles por persona (domingos excluidos, Pablo solo sábado)
        dias_disponibles = {
            "Lina":      ["Lunes"],
            "Dora":       ["Miércoles", "Jueves", "Viernes"],
            "Valentina": ["Martes", "Miércoles", "Jueves", "Viernes"],
            "Miguel":    ["Lunes", "Martes"],
            "Pablo":   ["Sábado"],
        }

        # Rotación: 5 personas con avance de 3 por semana.
        # mcd(3, 5) = 1 → ciclo completo cada 5 semanas; cada persona
        # ocupa cada posición (p1/p2/p3) exactamente una vez por ciclo.
        rows = []
        idx = 0
        idx_comunes = 0  # rotación independiente para tareas de comunes de Juan
        miguel_dia_idx = 0  # alterna Lunes (par) / Martes (impar) para Miguel

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

            # Cada persona ocupa el primer día disponible aún libre en la semana.
            # Miguel alterna su preferencia entre Lunes y Martes en cada asignación.
            dias_ocupados = set()
            for persona, tarea in asignaciones:
                if persona == "Miguel":
                    dias_p = ["Lunes", "Martes"] if miguel_dia_idx % 2 == 0 else ["Martes", "Lunes"]
                    miguel_dia_idx += 1
                else:
                    dias_p = dias_disponibles[persona]

                dia_elegido = dias_p[0]  # fallback
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

            # Tareas de comunes de Juan repartidas entre el staff en rotación propia.
            # Se asigna al siguiente candidato disponible que pueda trabajar ese día.
            # i%3==0 → Miércoles Entrada | i%3==1 → Viernes Cocina | i%3==2 → Lunes Salón + Martes Cocina
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
        # Rotación de 3 personas (Lina, Angie, Juan) sin restricciones de día.
        # Ciclo de 3 semanas: p_A → Lunes Salón + Martes Cocina
        #                     p_B → Viernes Cocina
        #                     p_C → Miércoles Entrada
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

            fill_header = PatternFill(start_color=COLOR_HEADER,  end_color=COLOR_HEADER,  fill_type="solid")
            fill_odd    = PatternFill(start_color=COLOR_ROW_ODD, end_color=COLOR_ROW_ODD, fill_type="solid")
            fill_even   = PatternFill(start_color=COLOR_ROW_EVN, end_color=COLOR_ROW_EVN, fill_type="solid")

            font_header = Font(color="FFFFFF", bold=True,  size=12)
            font_normal = Font(color="000000", bold=False, size=12)

            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

            white_side   = Side(border_style="thin", color="FFFFFF")
            border_white = Border(left=white_side, right=white_side, top=white_side, bottom=white_side)

            for sheet_name in ['Zonas Comunes', 'Zonas de Juan']:
                ws = workbook[sheet_name]

                for col_idx in range(1, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = 20

                for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column)):
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

        # Ajustar al lunes de la semana de inicio para alinear columnas de días
        curr = inicio - timedelta(days=inicio.weekday())

        semanas = []
        while curr <= fin:
            semanas.append((curr, curr + timedelta(days=6)))
            curr += timedelta(days=7)

        nombre_archivo_sugerido = f"Turnos limpieza {inicio.strftime('%d-%m-%Y')} a {fin.strftime('%d-%m-%Y')}.xlsx"

        ruta = filedialog.asksaveasfilename(
            initialfile=nombre_archivo_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )

        if not ruta:
            return

        self.btn.config(state="disabled", text="Generando...")
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
