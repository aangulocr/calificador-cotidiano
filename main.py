"""
main.py — Interfaz gráfica principal del Calificador de Trabajos Cotidianos.
Usa ttkbootstrap para un diseño moderno.

Ejecución:
    python main.py          (desde d:\\REVISION_TC)
    python -m main          (desde d:\\REVISION_TC)
"""

from __future__ import annotations
import sys
import os
import threading
from pathlib import Path

# Asegurar que el paquete calificador sea encontrable
sys.path.insert(0, str(Path(__file__).parent))

try:
    import ttkbootstrap as ttk                     # type: ignore
    from ttkbootstrap.constants import *           # type: ignore
    from tkinter.scrolledtext import ScrolledText
except ImportError:
    print(
        "ERROR: ttkbootstrap no está instalado.\n"
        "Ejecuta:  pip install ttkbootstrap\n"
    )
    sys.exit(1)

import tkinter as tk
from tkinter import filedialog, messagebox

from calificador.config import TEMAS, MAX_HOJAS  # type: ignore
from calificador.evaluador import cargar_patron, calificar_carpeta  # type: ignore
from calificador.exportador import generar_reporte  # type: ignore


# ===========================================================================
# Aplicación principal
# ===========================================================================

class App(ttk.Window):  # type: ignore
    def __init__(self):
        super().__init__(themename="flatly")  # type: ignore
        self.title("Calificador de Trabajos Cotidianos")
        self.geometry("980x720")
        self.minsize(800, 600)
        self.resizable(True, True)

        # Icono (ignorar si no existe)
        try:
            self.iconbitmap("icono.ico")
        except Exception:
            pass

        # Variable de carpeta
        self.carpeta_var = tk.StringVar(value="")

        # Variables de checkboxes (tema_id → BooleanVar)
        self.temas_vars: dict[str, tk.BooleanVar] = {
            t["id"]: tk.BooleanVar(value=False) for t in TEMAS
        }

        self._build_ui()

    # -----------------------------------------------------------------------
    # Construcción de la UI
    # -----------------------------------------------------------------------
    def _build_ui(self):
        # ---- Cabecera -------------------------------------------------------
        header = ttk.Frame(self, style="primary.TFrame")
        header.pack(fill=tk.X)
        ttk.Label(
            header,
            text="📋  Calificador de Trabajos Cotidianos",
            font=("Calibri", 20, "bold"),
            style="inverse-primary.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Label(
            header,
            text="MEP — Escala 0-3 por hoja",
            font=("Calibri", 11),
            style="inverse-primary.TLabel",
            foreground="#AACFE8",
        ).pack(side=tk.LEFT, padx=20)

        # Botón de Reiniciar app
        ttk.Button(
            header,
            text="🔄 Reiniciar",
            style="secondary.TButton",
            command=self._reiniciar_app,
        ).pack(side=tk.RIGHT, padx=15, pady=8)

        # ---- Contenido principal (columna izq + derecha) --------------------
        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=10)
        body.columnconfigure(0, weight=2)
        body.columnconfigure(1, weight=3)
        body.rowconfigure(0, weight=1)

        # ============ Panel izquierdo: carpeta + temas ========================
        izq = ttk.Frame(body)
        izq.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 12))
        izq.rowconfigure(1, weight=1)

        # Selector de carpeta
        frame_dir = ttk.LabelFrame(izq, text="  📁  Carpeta Trabajos_Cotidianos")
        frame_dir.grid(row=0, column=0, sticky=tk.EW, pady=(0, 10))
        frame_dir.columnconfigure(0, weight=1)

        self.entry_carpeta = ttk.Entry(
            frame_dir, textvariable=self.carpeta_var,
            font=("Calibri", 10), state="readonly"
        )
        self.entry_carpeta.grid(row=0, column=0, sticky=tk.EW, padx=(8, 8), pady=8)

        ttk.Button(
            frame_dir, text="Examinar…",
            style="outline.TButton", command=self._seleccionar_carpeta,
            width=10
        ).grid(row=0, column=1, padx=(0, 8), pady=8)

        # --- Checkboxes de temas ---
        frame_temas = ttk.LabelFrame(izq, text="  ✅  Temas a evaluar")
        frame_temas.grid(row=1, column=0, sticky=tk.NSEW)
        frame_temas.columnconfigure(0, weight=1)
        frame_temas.rowconfigure(0, weight=1)

        canvas_cb = tk.Canvas(frame_temas, highlightthickness=0)
        scrollbar  = ttk.Scrollbar(frame_temas, orient=tk.VERTICAL, command=canvas_cb.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas_cb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas_cb.configure(yscrollcommand=scrollbar.set)

        inner = ttk.Frame(canvas_cb)
        canvas_cb.create_window((0, 0), window=inner, anchor=tk.NW)
        inner.bind(
            "<Configure>",
            lambda e: canvas_cb.configure(scrollregion=canvas_cb.bbox("all"))
        )
        # Scroll con rueda
        canvas_cb.bind_all("<MouseWheel>",
            lambda e: canvas_cb.yview_scroll(int(-1*(e.delta/120)), "units"))

        categoria_actual = None
        for tema in TEMAS:
            cat = tema["categoria"]
            if cat != categoria_actual:
                categoria_actual = cat
                ttk.Label(
                    inner, text=f"── {cat} ──",
                    font=("Calibri", 9, "bold"), foreground="#0056B3"
                ).pack(anchor=tk.W, padx=4, pady=(8, 2))

            cb = ttk.Checkbutton(
                inner,
                text=tema["nombre"],
                variable=self.temas_vars[tema["id"]],
                style="primary.TCheckbutton",
                padding=(4, 3),
            )
            cb.pack(anchor=tk.W, padx=12)

        # Botones rápidos
        frame_botones_cb = ttk.Frame(izq)
        frame_botones_cb.grid(row=2, column=0, sticky=tk.EW, pady=(6, 0))

        ttk.Button(
            frame_botones_cb, text="Seleccionar todos",
            style="info.Outline.TButton",
            command=lambda: self._marcar_todos(True),
        ).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))

        ttk.Button(
            frame_botones_cb, text="Desmarcar todos",
            style="secondary.Outline.TButton",
            command=lambda: self._marcar_todos(False),
        ).pack(side=tk.LEFT, expand=True, fill=tk.X)

        # ============ Panel derecho: log + botón calificar ====================
        der = ttk.Frame(body)
        der.grid(row=0, column=1, sticky=tk.NSEW)
        der.rowconfigure(1, weight=1)
        der.columnconfigure(0, weight=1)

        ttk.Label(
            der, text="📝  Registro de evaluación",
            font=("Calibri", 11, "bold"), foreground="#0056B3"
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 4))

        self.log_area = ScrolledText(
            der,
            font=("Consolas", 10),
            state="disabled",
            height=22,
        )
        self.log_area.grid(row=1, column=0, sticky=tk.NSEW)

        # Barra de progreso
        self.progress = ttk.Progressbar(
            der, mode="indeterminate", style="primary.Striped.Horizontal.TProgressbar"
        )
        self.progress.grid(row=2, column=0, sticky=tk.EW, pady=(8, 4))

        # Botón calificar
        self.btn_calificar = ttk.Button(
            der,
            text="▶  CALIFICAR",
            style="primary.TButton",
            command=self._iniciar_calificacion,
            width=22,
        )
        self.btn_calificar.grid(row=3, column=0, sticky=tk.EW, ipady=10)

        # ---- Status bar ------------------------------------------------------
        self.status_var = tk.StringVar(value="Listo — Selecciona la carpeta y los temas.")
        ttk.Label(
            self, textvariable=self.status_var,
            font=("Calibri", 9), relief=tk.SUNKEN, anchor=tk.W,
            padding=(8, 3),
        ).pack(fill=tk.X, side=tk.BOTTOM)

    # -----------------------------------------------------------------------
    # Acciones
    # -----------------------------------------------------------------------
    def _seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta Trabajos_Cotidianos")
        if carpeta:
            self.carpeta_var.set(carpeta)
            self._log(f"📂 Carpeta seleccionada: {carpeta}\n")
            self.status_var.set(f"Carpeta: {carpeta}")

    def _marcar_todos(self, valor: bool):
        for v in self.temas_vars.values():
            v.set(valor)

    def _reiniciar_app(self):
        """Limpia el estado de la aplicación para iniciar una nueva calificación."""
        # 1. Limpiar carpeta
        self.carpeta_var.set("")
        self.entry_carpeta.configure(state="normal")
        self.entry_carpeta.delete(0, tk.END)
        self.entry_carpeta.configure(state="readonly")
        
        # 2. Desmarcar temas
        self._marcar_todos(False)
        
        # 3. Limpiar log
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", tk.END)
        self.log_area.configure(state="disabled")
        
        # 4. Status bar
        self.status_var.set("Listo — Selecciona la carpeta y los temas.")
        self._log("🔄 Aplicación reiniciada. Lista para calificar un nuevo grupo.\n")

    def _log(self, mensaje: str):
        """Escribe en el área de log de forma thread-safe."""
        def _write():
            self.log_area.configure(state="normal")
            self.log_area.insert(tk.END, mensaje + "\n")
            self.log_area.see(tk.END)
            self.log_area.configure(state="disabled")
        self.after(0, _write)

    def _iniciar_calificacion(self):
        carpeta = self.carpeta_var.get().strip()
        if not carpeta:
            messagebox.showwarning("Carpeta requerida", "Por favor selecciona la carpeta de trabajos.")
            return

        temas_activos = {
            tid for tid, var in self.temas_vars.items() if var.get()
        }
        if not temas_activos:
            messagebox.showwarning("Sin temas", "Selecciona al menos un tema para evaluar.")
            return

        self.btn_calificar.configure(state="disabled")
        self.progress.start(10)
        self.status_var.set("Procesando… por favor espera.")
        self._log("=" * 60)
        self._log(f"▶ Iniciando calificación")
        self._log(f"   Temas activos: {len(temas_activos)}")
        self._log("=" * 60)

        thread = threading.Thread(
            target=self._ejecutar_calificacion,
            args=(carpeta, temas_activos),
            daemon=True,
        )
        thread.start()

    def _ejecutar_calificacion(self, carpeta: str, temas_activos: set):
        """Se ejecuta en un hilo secundario."""
        try:
            self._log("📌 Cargando PLANTILLA.xlsx …")
            patron = cargar_patron(carpeta)
            n_hojas_patron = len(patron["orden_hojas"])
            self._log(f"   Hojas del patrón: {n_hojas_patron}")

            resultados = calificar_carpeta(
                carpeta, patron, temas_activos,
                log_callback=self._log,
            )

            self._log(f"\n💾 Generando reporte para {len(resultados)} estudiantes …")
            ruta_salida = generar_reporte(resultados, carpeta)
            self._log(f"✅ Reporte guardado en:\n   {ruta_salida}\n")

            self.after(0, self._calificacion_exitosa, ruta_salida, len(resultados))

        except Exception as exc:
            self._log(f"\n❌ ERROR: {exc}\n")
            self.after(0, self._calificacion_error, str(exc))

    def _calificacion_exitosa(self, ruta: str, n: int):
        self.progress.stop()
        self.btn_calificar.configure(state="normal")
        self.status_var.set(f"✅ {n} estudiantes calificados. Archivo: {Path(ruta).name}")
        if messagebox.askyesno(
            "Calificación completada",
            f"Se calificaron {n} estudiantes.\n\n"
            f"Archivo generado:\n{ruta}\n\n"
            "¿Deseas abrir el archivo ahora?"
        ):
            os.startfile(ruta)  # type: ignore

    def _calificacion_error(self, msg: str):
        self.progress.stop()
        self.btn_calificar.configure(state="normal")
        self.status_var.set("❌ Error durante la calificación.")
        messagebox.showerror("Error", msg)


# ===========================================================================
# Punto de entrada
# ===========================================================================
if __name__ == "__main__":
    app = App()
    app.mainloop()
