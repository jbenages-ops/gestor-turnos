"""Permite importar generador_turnos sin display ni tkinter instalado.

La lógica de turnos no depende de la GUI, pero el módulo importa tkinter
arriba del todo; si el entorno de test no lo tiene (CI, contenedores),
se inyectan stubs mínimos antes del import.
"""
import sys
import types
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

try:
    import tkinter  # noqa: F401
except ImportError:
    class _DummyWidget:
        def __init__(self, *args, **kwargs):
            pass

        def __getattr__(self, name):
            return lambda *args, **kwargs: None

    tk = types.ModuleType("tkinter")
    for nombre in ("Frame", "Label", "Button", "Canvas", "Tk", "StringVar"):
        setattr(tk, nombre, _DummyWidget)

    ttk = types.ModuleType("tkinter.ttk")
    ttk.Combobox = _DummyWidget
    ttk.Style = _DummyWidget

    filedialog = types.ModuleType("tkinter.filedialog")
    filedialog.asksaveasfilename = lambda **kwargs: ""

    messagebox = types.ModuleType("tkinter.messagebox")
    messagebox.showerror = lambda *args, **kwargs: None
    messagebox.askyesno = lambda *args, **kwargs: False

    tk.ttk = ttk
    tk.filedialog = filedialog
    tk.messagebox = messagebox

    sys.modules["tkinter"] = tk
    sys.modules["tkinter.ttk"] = ttk
    sys.modules["tkinter.filedialog"] = filedialog
    sys.modules["tkinter.messagebox"] = messagebox
