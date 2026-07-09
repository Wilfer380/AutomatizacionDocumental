from __future__ import annotations

import logging
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox, ttk

from .. import __version__
from ..config import APP_NAME, DEFAULT_CONFLICT_STRATEGY, DEFAULT_EXCEL_PATH, DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE_PATH, PLACEHOLDER
from ..models import BatchOptions, ConflictStrategy, SeriesRow, SheetInfo, ValidationStatus, WorkbookInfo
from ..services.batch_service import BatchService
from ..services.excel_service import ExcelService
from ..services.pdf_converter import PdfConverter
from ..services.report_service import ReportService
from ..services.template_service import TemplateService
from .dossier_panel import DossierPanel
from ..utils.files import open_folder


logger = logging.getLogger(__name__)


DISPLAY_STATUS = {
    ValidationStatus.VALID: "Válida",
    ValidationStatus.EMPTY: "Vacía",
    ValidationStatus.DUPLICATE: "Duplicada",
    ValidationStatus.INVALID: "Inválida",
}


PALETTE = {
    "background": "#edf3f8",
    "surface": "#ffffff",
    "surface_soft": "#f6f9fc",
    "line": "#d9e3ee",
    "header": "#16395f",
    "header_dark": "#0f2743",
    "accent": "#1e6cf2",
    "accent_dark": "#1559c7",
    "accent_light": "#eaf2ff",
    "text": "#17324d",
    "muted": "#657385",
    "success": "#eaf8ef",
    "success_accent": "#2f9e62",
    "warning": "#fff5e3",
    "warning_accent": "#d79b16",
    "danger": "#fdecec",
    "danger_accent": "#cf4b5b",
    "info": "#eef4ff",
    "info_accent": "#4b72f0",
    "green": "#1bb36a",
}


class GenerationCancelled(Exception):
    pass


class MainWindow(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1366x768")
        self.minsize(1280, 720)
        self._active_tab = "phase2"

        self.excel_service = ExcelService()
        self.template_service = TemplateService()
        self.report_service = ReportService()
        self.pdf_converter = PdfConverter()
        self.batch_service = BatchService(self.template_service, self.pdf_converter, self.report_service)

        self.workbook_info: WorkbookInfo | None = None
        self.series_rows: list[SeriesRow] = []
        self.visible_rows: list[SeriesRow] = []
        self.report_path: Path | None = None

        self.queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self._generation_running = False
        self._cancel_requested = False

        self.template_var = StringVar(value=str(DEFAULT_TEMPLATE_PATH))
        self.excel_var = StringVar(value=str(DEFAULT_EXCEL_PATH))
        self.output_var = StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.sheet_var = StringVar(value="")
        self.column_var = StringVar(value="")
        self.search_var = StringVar(value="")
        self.status_var = StringVar(value="Listo para cargar la vista previa")
        self.report_var = StringVar(value="Sin reporte generado")

        self.preview_total_var = StringVar(value="0")
        self.preview_valid_var = StringVar(value="0")
        self.preview_warning_var = StringVar(value="0")
        self.preview_selected_var = StringVar(value="0")

        self.execution_total_var = StringVar(value="0")
        self.execution_done_var = StringVar(value="0")
        self.execution_failed_var = StringVar(value="0")
        self.execution_skipped_var = StringVar(value="0")

        self.exclude_empty_var = BooleanVar(value=True)
        self.exclude_duplicate_var = BooleanVar(value=True)
        self.process_only_selected_var = BooleanVar(value=True)
        self.conflict_strategy_var = StringVar(value=DEFAULT_CONFLICT_STRATEGY)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._configure_style()
        self._build_ui()
        self.after(0, self._maximize_to_screen)
        self.after(80, self._bring_to_front)
        self.after(120, self._poll_queue)

    def _maximize_to_screen(self) -> None:
        try:
            self.state("zoomed")
        except Exception:
            pass

    def _bring_to_front(self) -> None:
        try:
            self.deiconify()
            self.lift()
            self.attributes("-topmost", True)
            self.after(250, lambda: self.attributes("-topmost", False))
            self.focus_force()
        except Exception:
            pass

    def _on_close(self) -> None:
        try:
            self.destroy()
        except Exception:
            pass

    def _configure_style(self) -> None:
        self.configure(background=PALETTE["background"])
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("App.TFrame", background=PALETTE["background"])
        style.configure("Surface.TFrame", background=PALETTE["surface"])
        style.configure("SurfaceSoft.TFrame", background=PALETTE["surface_soft"])
        style.configure("CardBody.TFrame", background=PALETTE["surface"])
        style.configure("Header.TFrame", background=PALETTE["header"])
        style.configure("Header.TLabel", background=PALETTE["header"], foreground="white")
        style.configure("CardTitle.TLabel", background=PALETTE["surface"], foreground=PALETTE["text"], font=("Segoe UI Semibold", 12))
        style.configure("CardSubtitle.TLabel", background=PALETTE["surface"], foreground=PALETTE["muted"], font=("Segoe UI", 9))
        style.configure("Muted.TLabel", background=PALETTE["surface"], foreground=PALETTE["muted"], font=("Segoe UI", 9))
        style.configure("Body.TLabel", background=PALETTE["background"], foreground=PALETTE["text"], font=("Segoe UI", 10))
        style.configure("BodyBold.TLabel", background=PALETTE["background"], foreground=PALETTE["text"], font=("Segoe UI Semibold", 10))
        style.configure("SummaryTitle.TLabel", background=PALETTE["surface"], foreground=PALETTE["muted"], font=("Segoe UI", 9))
        style.configure("SummaryValue.TLabel", background=PALETTE["surface"], foreground=PALETTE["text"], font=("Segoe UI Semibold", 18))
        style.configure("SectionNote.TLabel", background=PALETTE["surface"], foreground=PALETTE["muted"], font=("Segoe UI", 9))
        style.configure("Footer.TLabel", background=PALETTE["header_dark"], foreground="#d7e5f7", font=("Segoe UI", 9))
        style.configure("FooterValue.TLabel", background=PALETTE["header_dark"], foreground="white", font=("Segoe UI Semibold", 10))
        style.configure("App.TEntry", padding=7)
        style.configure("App.TCombobox", padding=5)
        style.configure("App.TCheckbutton", background=PALETTE["surface"], foreground=PALETTE["text"], font=("Segoe UI", 10))
        style.configure("Primary.TButton", padding=(18, 10), font=("Segoe UI Semibold", 10), background=PALETTE["accent"], foreground="white")
        style.map(
            "Primary.TButton",
            background=[("active", PALETTE["accent_dark"]), ("pressed", PALETTE["accent_dark"])],
            foreground=[("disabled", "#dbe6f3"), ("active", "white")],
        )
        style.configure("Secondary.TButton", padding=(14, 8), font=("Segoe UI Semibold", 10), background="#e7eef7", foreground=PALETTE["text"])
        style.map("Secondary.TButton", background=[("active", "#d7e2f0"), ("pressed", "#d7e2f0")])
        style.configure("Ghost.TButton", padding=(12, 7), font=("Segoe UI Semibold", 10), background=PALETTE["surface"], foreground=PALETTE["text"])
        style.map("Ghost.TButton", background=[("active", PALETTE["surface_soft"]), ("pressed", PALETTE["surface_soft"])])
        style.configure(
            "App.Treeview",
            background=PALETTE["surface"],
            fieldbackground=PALETTE["surface"],
            foreground=PALETTE["text"],
            rowheight=31,
            borderwidth=0,
        )
        style.configure(
            "App.Treeview.Heading",
            background=PALETTE["surface_soft"],
            foreground=PALETTE["text"],
            font=("Segoe UI Semibold", 10),
            padding=7,
        )
        style.map("App.Treeview", background=[("selected", PALETTE["accent_light"])], foreground=[("selected", PALETTE["text"])])
        style.configure("TProgressbar", troughcolor=PALETTE["surface_soft"], background=PALETTE["accent"], bordercolor=PALETTE["line"], lightcolor=PALETTE["accent"], darkcolor=PALETTE["accent"])

    def _build_ui(self) -> None:
        self.header = tk.Frame(self, bg=PALETTE["header"], highlightthickness=0)
        self.header.pack(fill="x")

        header_inner = tk.Frame(self.header, bg=PALETTE["header"], padx=22, pady=16)
        header_inner.pack(fill="x")

        left = tk.Frame(header_inner, bg=PALETTE["header"])
        left.pack(side="left", fill="x", expand=True)
        tk.Label(left, text="Automatización documental", bg=PALETTE["header"], fg="white", font=("Segoe UI Semibold", 20)).pack(anchor="w")
        tk.Label(
            left,
            text="Automatiza la distribución de documentos en los dossiers según CP y Serie.",
            bg=PALETTE["header"],
            fg="#d7e5f7",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        actions = tk.Frame(header_inner, bg=PALETTE["header"])
        actions.pack(side="right", anchor="n")
        self._header_button(actions, "Acerca de", self._show_about).pack(side="right", padx=(8, 0))
        self._header_button(actions, "Ayuda", self._show_help).pack(side="right")

        tab_strip = tk.Frame(self, bg=PALETTE["background"])
        tab_strip.pack(fill="x")
        tabs_inner = tk.Frame(tab_strip, bg=PALETTE["background"])
        tabs_inner.pack(fill="x", padx=22, pady=(10, 0))

        self.tab_shells: dict[str, dict[str, tk.Widget]] = {}
        self._create_tab(tabs_inner, "phase1", "Fase 1 - Generación Word → PDF")
        self._create_tab(tabs_inner, "phase2", "Fase 2 - Gestión de Dossier")

        tk.Frame(tab_strip, bg=PALETTE["line"], height=1).pack(fill="x", pady=(8, 0))

        self.content_host = tk.Frame(self, bg=PALETTE["background"])
        self.content_host.pack(fill="both", expand=True)

        self.phase1_container = ttk.Frame(self.content_host, style="App.TFrame")
        self.phase2_container = ttk.Frame(self.content_host, style="App.TFrame")
        for frame in (self.phase1_container, self.phase2_container):
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)

        self._build_phase1_content(self.phase1_container)
        self.dossier_panel = DossierPanel(self.phase2_container)
        self.dossier_panel.pack(fill="both", expand=True)

        footer = tk.Frame(self, bg=PALETTE["header_dark"], height=44)
        footer.pack(fill="x", side="bottom")
        footer_inner = tk.Frame(footer, bg=PALETTE["header_dark"], padx=22, pady=10)
        footer_inner.pack(fill="both", expand=True)
        footer_inner.columnconfigure(0, weight=1)
        footer_inner.columnconfigure(1, weight=1)
        footer_inner.columnconfigure(2, weight=1)

        self.footer_user_var = StringVar(value="wandica")
        self.footer_profile_var = StringVar(value="Administrador")
        self.footer_mode_var = StringVar(value="Simulación")
        self.footer_version_var = StringVar(value="1.0.0")

        footer_left = tk.Frame(footer_inner, bg=PALETTE["header_dark"])
        footer_left.grid(row=0, column=0, sticky="w")
        self._footer_field(footer_left, "Usuario", self.footer_user_var).pack(side="left", padx=(0, 18))
        self._footer_field(footer_left, "Perfil", self.footer_profile_var).pack(side="left")

        footer_center = tk.Frame(footer_inner, bg=PALETTE["header_dark"])
        footer_center.grid(row=0, column=1)
        self._footer_field(footer_center, "Modo", self.footer_mode_var).pack()

        footer_right = tk.Frame(footer_inner, bg=PALETTE["header_dark"])
        footer_right.grid(row=0, column=2, sticky="e")
        self._footer_field(footer_right, "Versión", self.footer_version_var).pack(anchor="e")

        self._show_tab("phase2")

    def _header_button(self, parent: tk.Widget, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg="#214a7b",
            fg="white",
            activebackground="#2b5c95",
            activeforeground="white",
            relief="flat",
            padx=16,
            pady=6,
            cursor="hand2",
            font=("Segoe UI Semibold", 10),
        )

    def _footer_field(self, parent: tk.Widget, label: str, variable: StringVar) -> tk.Widget:
        frame = tk.Frame(parent, bg=PALETTE["header_dark"])
        tk.Label(frame, text=f"{label}: ", bg=PALETTE["header_dark"], fg="#d7e5f7", font=("Segoe UI", 9)).pack(side="left")
        tk.Label(frame, textvariable=variable, bg=PALETTE["header_dark"], fg="white", font=("Segoe UI Semibold", 10)).pack(side="left")
        return frame

    def _create_tab(self, parent: tk.Widget, key: str, label: str) -> None:
        shell = tk.Frame(parent, bg=PALETTE["background"])
        shell.pack(side="left", padx=(0, 28))
        button = tk.Button(
            shell,
            text=label,
            command=lambda tab=key: self._show_tab(tab),
            bg=PALETTE["background"],
            fg=PALETTE["text"],
            activebackground=PALETTE["background"],
            activeforeground=PALETTE["accent"],
            relief="flat",
            bd=0,
            padx=0,
            pady=0,
            font=("Segoe UI Semibold", 10),
            cursor="hand2",
        )
        button.pack(anchor="w")
        underline = tk.Frame(shell, bg=PALETTE["background"], height=3)
        underline.pack(fill="x", pady=(6, 0))
        self.tab_shells[key] = {"shell": shell, "button": button, "underline": underline}

    def _show_about(self) -> None:
        messagebox.showinfo(
            "Acerca de",
            f"{APP_NAME}\nVersión {__version__}\n\nFase 1 y Fase 2 comparten una misma base corporativa.",
        )

    def _show_tab(self, tab: str) -> None:
        self._active_tab = tab
        if tab == "phase1":
            self.phase1_container.lift()
            self.footer_mode_var.set("Generación")
        else:
            self.phase2_container.lift()
            self.footer_mode_var.set("Simulación")

        for key, widgets in self.tab_shells.items():
            active = key == tab
            widgets["button"].configure(fg=PALETTE["accent"] if active else PALETTE["text"])
            widgets["underline"].configure(bg=PALETTE["green"] if active else PALETTE["background"])

    def _build_phase1_content(self, parent: ttk.Frame) -> None:
        viewport = ttk.Frame(parent, style="App.TFrame")
        viewport.pack(fill="both", expand=True)
        viewport.columnconfigure(0, weight=1)
        viewport.rowconfigure(0, weight=1)

        self.content_canvas = tk.Canvas(viewport, bg=PALETTE["background"], highlightthickness=0)
        self.content_canvas.grid(row=0, column=0, sticky="nsew")
        content_scroll = ttk.Scrollbar(viewport, orient="vertical", command=self.content_canvas.yview)
        content_scroll.grid(row=0, column=1, sticky="ns")
        self.content_canvas.configure(yscrollcommand=content_scroll.set)

        self.content_container = ttk.Frame(self.content_canvas, style="App.TFrame", padding=18)
        self.content_window = self.content_canvas.create_window((0, 0), window=self.content_container, anchor="nw")
        self.content_container.bind("<Configure>", self._sync_scroll_region)
        self.content_canvas.bind("<Configure>", self._sync_content_width)
        self._bind_mousewheel(self.content_canvas)

        content = self.content_container
        content.columnconfigure(0, weight=0)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(1, weight=1)

        self._build_summary_strip(content)

        left_column = ttk.Frame(content, style="App.TFrame")
        left_column.grid(row=1, column=0, sticky="ns", padx=(0, 14), pady=(14, 0))
        left_column.columnconfigure(0, weight=1)

        right_column = ttk.Frame(content, style="App.TFrame")
        right_column.grid(row=1, column=1, sticky="nsew", pady=(14, 0))
        right_column.columnconfigure(0, weight=1)
        right_column.rowconfigure(0, weight=1)

        self._build_files_section(left_column)
        self._build_workbook_section(left_column)
        self._build_generation_options_section(left_column)
        self._build_actions_section(left_column)

        self._build_preview_section(right_column)
        self._build_progress_section(right_column)
        self._build_summary_section(right_column)

    def _sync_scroll_region(self, _event: tk.Event) -> None:
        if hasattr(self, "content_canvas"):
            self.content_canvas.configure(scrollregion=self.content_canvas.bbox("all"))

    def _sync_content_width(self, event: tk.Event) -> None:
        if hasattr(self, "content_window"):
            self.content_canvas.itemconfigure(self.content_window, width=event.width)

    def _bind_mousewheel(self, widget: tk.Widget) -> None:
        widget.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        widget.bind_all("<Button-4>", self._on_mousewheel, add="+")
        widget.bind_all("<Button-5>", self._on_mousewheel, add="+")

    def _on_mousewheel(self, event: tk.Event) -> None:
        if not hasattr(self, "content_canvas"):
            return
        if getattr(event, "num", None) == 4:
            self.content_canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            self.content_canvas.yview_scroll(1, "units")
        else:
            delta = int(-1 * (event.delta / 120))
            self.content_canvas.yview_scroll(delta, "units")

    def _build_summary_strip(self, parent: ttk.Frame) -> None:
        strip = ttk.Frame(parent, style="App.TFrame")
        strip.grid(row=0, column=0, columnspan=2, sticky="ew")
        for index in range(4):
            strip.columnconfigure(index, weight=1)

        self._summary_card(strip, 0, "Filas visibles", self.preview_total_var, PALETTE["info_accent"])
        self._summary_card(strip, 1, "Series válidas", self.preview_valid_var, PALETTE["success_accent"])
        self._summary_card(strip, 2, "Con advertencia", self.preview_warning_var, PALETTE["warning_accent"])
        self._summary_card(strip, 3, "Seleccionadas", self.preview_selected_var, PALETTE["accent"])

    def _build_files_section(self, parent: ttk.Frame) -> None:
        card, body = self._section_card(parent, "Archivos", "Plantilla Word, Excel y carpeta de salida.", PALETTE["info_accent"])
        card.grid(row=0, column=0, sticky="ew")
        body.columnconfigure(1, weight=1)
        self._add_path_row(body, 0, "Plantilla Word", self.template_var, self._browse_template)
        self._add_path_row(body, 1, "Archivo Excel", self.excel_var, self._browse_excel)
        self._add_path_row(body, 2, "Carpeta de salida", self.output_var, self._browse_output, browse_label="Examinar carpeta")

    def _build_workbook_section(self, parent: ttk.Frame) -> None:
        card, body = self._section_card(parent, "Hoja y serie", "Seleccione la hoja y la columna que contiene la serie.", PALETTE["success_accent"])
        card.grid(row=1, column=0, sticky="ew", pady=(12, 0))

        controls = ttk.Frame(body, style="CardBody.TFrame")
        controls.pack(fill="x")
        controls.columnconfigure(1, weight=1)
        controls.columnconfigure(3, weight=1)

        ttk.Label(controls, text="Hoja", style="Body.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        self.sheet_combo = ttk.Combobox(controls, textvariable=self.sheet_var, state="readonly", width=32, style="App.TCombobox")
        self.sheet_combo.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_sheet_changed())

        ttk.Label(controls, text="Columna de serie", style="Body.TLabel").grid(row=0, column=2, sticky="w", padx=(16, 8), pady=(0, 8))
        self.column_combo = ttk.Combobox(controls, textvariable=self.column_var, state="readonly", width=32, style="App.TCombobox")
        self.column_combo.grid(row=0, column=3, sticky="ew", pady=(0, 8))
        self.column_combo.bind("<<ComboboxSelected>>", lambda _event: self._load_rows())

        self.validate_button = ttk.Button(controls, text="Validar configuración", style="Primary.TButton", command=self.validate_configuration)
        self.validate_button.grid(row=0, column=4, sticky="e", padx=(16, 0), pady=(0, 8))

    def _build_generation_options_section(self, parent: ttk.Frame) -> None:
        card, body = self._section_card(parent, "Opciones de generación", "Filtro de filas antes de generar los documentos.", PALETTE["warning_accent"])
        card.grid(row=2, column=0, sticky="ew", pady=(12, 0))

        ttk.Checkbutton(body, text="Excluir vacías", variable=self.exclude_empty_var, style="App.TCheckbutton", command=self._refresh_tree).pack(anchor="w")
        ttk.Checkbutton(body, text="Excluir duplicadas", variable=self.exclude_duplicate_var, style="App.TCheckbutton", command=self._refresh_tree).pack(anchor="w", pady=(8, 0))
        ttk.Checkbutton(body, text="Procesar solo seleccionadas", variable=self.process_only_selected_var, style="App.TCheckbutton").pack(anchor="w", pady=(8, 0))

        conflict_row = ttk.Frame(body, style="CardBody.TFrame")
        conflict_row.pack(fill="x", pady=(10, 0))
        ttk.Label(conflict_row, text="Conflicto de nombre", style="Body.TLabel").pack(side="left")
        conflict_combo = ttk.Combobox(
            conflict_row,
            textvariable=self.conflict_strategy_var,
            values=[strategy.value for strategy in ConflictStrategy],
            state="readonly",
            width=18,
            style="App.TCombobox",
        )
        conflict_combo.pack(side="right")

        ttk.Label(body, text="Solo se ejecuta la fase 1: plantillas, Excel, vista previa y generación.", style="SectionNote.TLabel", wraplength=330).pack(anchor="w", pady=(12, 0))

    def _build_actions_section(self, parent: ttk.Frame) -> None:
        card, body = self._section_card(parent, "Acciones", "Generar, cancelar, abrir carpeta o revisar el reporte.", PALETTE["accent"])
        card.grid(row=3, column=0, sticky="ew", pady=(12, 0))

        self.generate_button = ttk.Button(body, text="Generar DOCX / PDF", style="Primary.TButton", command=self.start_generation)
        self.generate_button.pack(fill="x")

        self.cancel_button = ttk.Button(body, text="Cancelar", style="Secondary.TButton", command=self.cancel_generation, state="disabled")
        self.cancel_button.pack(fill="x", pady=(10, 0))

        buttons = ttk.Frame(body, style="CardBody.TFrame")
        buttons.pack(fill="x", pady=(10, 0))
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)

        ttk.Button(buttons, text="Abrir carpeta", style="Ghost.TButton", command=self.open_output_folder).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.view_report_button = ttk.Button(buttons, text="Ver reporte", style="Ghost.TButton", command=self.open_report, state="disabled")
        self.view_report_button.grid(row=0, column=1, sticky="ew", padx=(6, 0))

    def _build_preview_section(self, parent: ttk.Frame) -> None:
        card, body = self._section_card(parent, "Vista previa de series", "Filtre, busque y seleccione las filas que se van a generar.", PALETTE["accent"])
        card.grid(row=0, column=0, sticky="nsew")
        body.rowconfigure(2, weight=1)
        body.columnconfigure(0, weight=1)

        header = ttk.Frame(body, style="CardBody.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        left = ttk.Frame(header, style="CardBody.TFrame")
        left.grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="Vista previa", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(left, text="Use la columna Seleccionar para incluir o excluir filas antes de generar.", style="CardSubtitle.TLabel").pack(anchor="w", pady=(2, 0))

        search_bar = ttk.Frame(header, style="CardBody.TFrame")
        search_bar.grid(row=0, column=1, sticky="e")
        ttk.Label(search_bar, text="Buscar", style="Body.TLabel").pack(side="left", padx=(0, 8))
        search_entry = ttk.Entry(search_bar, textvariable=self.search_var, width=30, style="App.TEntry")
        search_entry.pack(side="left")
        search_entry.bind("<KeyRelease>", lambda _event: self._refresh_tree())
        ttk.Button(search_bar, text="Limpiar", style="Ghost.TButton", command=self._clear_search).pack(side="left", padx=(8, 0))

        table_frame = ttk.Frame(body, style="CardBody.TFrame")
        table_frame.grid(row=2, column=0, sticky="nsew", pady=(14, 0))
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        columns = ("select", "row", "series", "state", "observation")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse", style="App.Treeview")
        self.tree.heading("select", text="Seleccionar")
        self.tree.heading("row", text="Fila")
        self.tree.heading("series", text="Serie")
        self.tree.heading("state", text="Estado")
        self.tree.heading("observation", text="Observación")
        self.tree.column("select", width=100, anchor="center", stretch=False)
        self.tree.column("row", width=84, anchor="center", stretch=False)
        self.tree.column("series", width=280, anchor="w", stretch=True)
        self.tree.column("state", width=120, anchor="w", stretch=False)
        self.tree.column("observation", width=440, anchor="w", stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<space>", self._toggle_selected_from_keyboard)

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.tag_configure("valid", background=PALETTE["success"])
        self.tree.tag_configure("empty", background=PALETTE["warning"])
        self.tree.tag_configure("duplicate", background=PALETTE["danger"])
        self.tree.tag_configure("invalid", background="#f2f4f7")
        self.tree.tag_configure("selected", foreground=PALETTE["text"])

    def _build_progress_section(self, parent: ttk.Frame) -> None:
        card, body = self._section_card(parent, "Progreso de generación", "Seguimiento del lote mientras se crean los archivos.", PALETTE["success_accent"])
        card.grid(row=1, column=0, sticky="ew", pady=(12, 0))

        self.progress = ttk.Progressbar(body, mode="determinate", maximum=1)
        self.progress.pack(fill="x")
        ttk.Label(body, textvariable=self.status_var, style="Muted.TLabel").pack(anchor="w", pady=(6, 0))

    def _build_summary_section(self, parent: ttk.Frame) -> None:
        card, body = self._section_card(parent, "Resumen final", "Resultados del último proceso de generación.", PALETTE["danger_accent"])
        card.grid(row=2, column=0, sticky="ew", pady=(12, 0))

        metrics = ttk.Frame(body, style="CardBody.TFrame")
        metrics.pack(fill="x")
        for index in range(4):
            metrics.columnconfigure(index, weight=1)

        self._summary_card(metrics, 0, "A generar", self.execution_total_var, PALETTE["info_accent"])
        self._summary_card(metrics, 1, "Generados", self.execution_done_var, PALETTE["success_accent"])
        self._summary_card(metrics, 2, "Fallidos", self.execution_failed_var, PALETTE["danger_accent"])
        self._summary_card(metrics, 3, "Omitidos", self.execution_skipped_var, PALETTE["warning_accent"])

        ttk.Label(body, textvariable=self.report_var, style="CardSubtitle.TLabel", wraplength=820).pack(anchor="w", pady=(12, 0))

    def _section_card(self, parent: ttk.Frame, title: str, subtitle: str, accent: str) -> tuple[tk.Frame, ttk.Frame]:
        outer = tk.Frame(parent, bg=accent, bd=0, highlightthickness=0)
        inner = tk.Frame(outer, bg=PALETTE["surface"], bd=0, highlightthickness=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        header = tk.Frame(inner, bg=PALETTE["surface"])
        header.pack(fill="x", padx=16, pady=(14, 10))
        tk.Label(header, text=title, bg=PALETTE["surface"], fg=PALETTE["text"], font=("Segoe UI Semibold", 12)).pack(anchor="w")
        tk.Label(header, text=subtitle, bg=PALETTE["surface"], fg=PALETTE["muted"], font=("Segoe UI", 9), wraplength=340, justify="left").pack(anchor="w", pady=(2, 0))

        body = ttk.Frame(inner, style="CardBody.TFrame", padding=(16, 0, 16, 16))
        body.pack(fill="both", expand=True)
        return outer, body

    def _summary_card(self, parent: ttk.Frame, column: int, title: str, variable: StringVar, accent: str) -> None:
        outer = tk.Frame(parent, bg=accent, bd=0, highlightthickness=0)
        outer.grid(row=0, column=column, sticky="nsew", padx=6)
        inner = tk.Frame(outer, bg=PALETTE["surface"], bd=0, highlightthickness=0, padx=14, pady=10)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text=title, bg=PALETTE["surface"], fg=PALETTE["muted"], font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(inner, textvariable=variable, bg=PALETTE["surface"], fg=PALETTE["text"], font=("Segoe UI Semibold", 18)).pack(anchor="w", pady=(3, 0))

    def _add_path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: StringVar,
        browse_command,
        browse_label: str = "Examinar",
    ) -> None:
        ttk.Label(parent, text=label, style="Body.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        entry = ttk.Entry(parent, textvariable=variable, style="App.TEntry")
        entry.grid(row=row, column=1, sticky="ew", padx=8, pady=5)
        ttk.Button(parent, text=browse_label, style="Ghost.TButton", command=browse_command).grid(row=row, column=2, sticky="e", pady=5)

    def _browse_template(self) -> None:
        try:
            initialdir = str(Path(self.template_var.get()).expanduser().parent)
        except Exception:
            initialdir = str(DEFAULT_TEMPLATE_PATH.parent)
        path = filedialog.askopenfilename(filetypes=[("Documento de Word", "*.docx")], initialdir=initialdir)
        if path:
            self.template_var.set(path)

    def _browse_excel(self) -> None:
        try:
            initialdir = str(Path(self.excel_var.get()).expanduser().parent)
        except Exception:
            initialdir = str(DEFAULT_EXCEL_PATH.parent)
        path = filedialog.askopenfilename(filetypes=[("Libro de Excel", "*.xlsx")], initialdir=initialdir)
        if path:
            self.excel_var.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(initialdir=self.output_var.get() or str(DEFAULT_OUTPUT_DIR))
        if path:
            self.output_var.set(path)

    def _show_help(self) -> None:
        messagebox.showinfo(
            "Ayuda",
            "1. Cargue la plantilla Word, el archivo Excel y la carpeta de salida.\n"
            "2. Elija la hoja y la columna de serie.\n"
            "3. Pulse Validar configuración para cargar la vista previa.\n"
            "4. Filtre y seleccione las series que desea generar.\n"
            "5. Pulse Generar DOCX / PDF.",
        )

    def _ensure_workbook_context(self, *, reload_rows: bool) -> bool:
        template_path = Path(self.template_var.get())
        excel_path = Path(self.excel_var.get())

        if not template_path.exists():
            messagebox.showerror("Plantilla no encontrada", f"La plantilla no existe:\n{template_path}")
            return False
        if not excel_path.exists():
            messagebox.showerror("Archivo Excel no encontrado", f"El archivo de Excel no existe:\n{excel_path}")
            return False

        try:
            if not self.template_service.validate_placeholder_present(template_path, PLACEHOLDER):
                messagebox.showerror(
                    "Validación de plantilla",
                    f"La plantilla debe contener {PLACEHOLDER} en el cuerpo, tablas, encabezados o pies de página.",
                )
                return False
        except Exception as exc:
            messagebox.showerror("Falló la validación de la plantilla", str(exc))
            return False

        needs_reload = reload_rows or self.workbook_info is None or self.workbook_info.path != excel_path
        if needs_reload:
            try:
                self.workbook_info = self.excel_service.inspect_workbook(excel_path)
            except Exception as exc:
                messagebox.showerror("No se pudo cargar el libro", str(exc))
                return False

            sheets = [sheet.name for sheet in self.workbook_info.sheets]
            self.sheet_combo["values"] = sheets
            if not sheets:
                messagebox.showwarning("Libro sin hojas", "El libro no contiene hojas utilizables.")
                return False

            if self.sheet_var.get() not in sheets:
                self.sheet_var.set(sheets[0])

            sheet = self._selected_sheet_info()
            if sheet is None:
                return False

            self._populate_columns(sheet)
            if not self._load_rows():
                return False
        elif not self.series_rows:
            if not self._load_rows():
                return False
        return True

    def validate_configuration(self) -> None:
        if self._ensure_workbook_context(reload_rows=True):
            self.status_var.set("Configuración validada")
            messagebox.showinfo("Validación completa", "La plantilla, el Excel y la vista previa quedaron cargados.")

    def _selected_sheet_info(self) -> SheetInfo | None:
        if not self.workbook_info:
            return None
        for sheet in self.workbook_info.sheets:
            if sheet.name == self.sheet_var.get():
                return sheet
        return self.workbook_info.sheets[0] if self.workbook_info.sheets else None

    def _populate_columns(self, sheet: SheetInfo) -> None:
        columns = [column.name for column in sheet.columns]
        self.column_combo["values"] = columns
        if sheet.suggested_serial_column and sheet.suggested_serial_column in columns:
            self.column_var.set(sheet.suggested_serial_column)
        elif self.column_var.get() not in columns and columns:
            self.column_var.set(columns[0])

    def _on_sheet_changed(self) -> None:
        sheet = self._selected_sheet_info()
        if sheet:
            self._populate_columns(sheet)
            self._load_rows()

    def _load_rows(self) -> bool:
        if not self.workbook_info:
            return False

        sheet_name = self.sheet_var.get().strip()
        column_name = self.column_var.get().strip()
        if not sheet_name or not column_name:
            return False

        previous_selection = {(row.row_number, row.series): row.selected for row in self.series_rows}
        try:
            rows = self.excel_service.load_series_rows(Path(self.excel_var.get()), sheet_name, column_name)
        except Exception as exc:
            messagebox.showerror("No se pudo generar la vista previa", str(exc))
            return False

        for row in rows:
            row.selected = previous_selection.get((row.row_number, row.series), row.selected)

        self.series_rows = rows
        self.status_var.set(f"Vista previa cargada: {len(rows)} filas")
        self._refresh_tree()
        return True

    def _clear_search(self) -> None:
        self.search_var.set("")
        self._refresh_tree()

    def _refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        search = self.search_var.get().strip().lower()

        visible: list[SeriesRow] = []
        for row in self.series_rows:
            if self.exclude_empty_var.get() and row.status == ValidationStatus.EMPTY:
                continue
            if self.exclude_duplicate_var.get() and row.status == ValidationStatus.DUPLICATE:
                continue
            if search and search not in row.series.lower():
                continue
            visible.append(row)

        for row in visible:
            iid = str(row.row_number)
            tags = (self._row_tag(row.status), "selected") if row.selected else (self._row_tag(row.status),)
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=("☑" if row.selected else "☐", row.row_number, row.series, self._display_status(row.status), row.observation),
                tags=tags,
            )

        self.visible_rows = visible
        self._update_preview_metrics()

    def _update_preview_metrics(self) -> None:
        visible = self.visible_rows
        self.preview_total_var.set(str(len(visible)))
        self.preview_valid_var.set(str(sum(1 for row in visible if row.status == ValidationStatus.VALID)))
        self.preview_warning_var.set(str(sum(1 for row in visible if row.status in {ValidationStatus.EMPTY, ValidationStatus.DUPLICATE})))
        self.preview_selected_var.set(str(sum(1 for row in visible if row.selected)))

    def _row_tag(self, status: ValidationStatus) -> str:
        if status == ValidationStatus.EMPTY:
            return "empty"
        if status == ValidationStatus.DUPLICATE:
            return "duplicate"
        if status == ValidationStatus.INVALID:
            return "invalid"
        return "valid"

    @staticmethod
    def _display_status(status: ValidationStatus) -> str:
        return DISPLAY_STATUS.get(status, status.value)

    def _selected_rows_for_generation(self) -> list[SeriesRow]:
        rows = list(self.visible_rows)
        if self.process_only_selected_var.get():
            rows = [row for row in rows if row.selected]
        return rows

    def _build_options(self) -> BatchOptions:
        return BatchOptions(
            template_path=Path(self.template_var.get()),
            excel_path=Path(self.excel_var.get()),
            output_dir=Path(self.output_var.get()),
            sheet_name=self.sheet_var.get(),
            serial_column=self.column_var.get(),
            exclude_empties=self.exclude_empty_var.get(),
            exclude_duplicates=self.exclude_duplicate_var.get(),
            search_term=self.search_var.get().strip(),
            process_only_selected=self.process_only_selected_var.get(),
            libreoffice_path="",
            placeholder=PLACEHOLDER,
            conflict_strategy=ConflictStrategy(self.conflict_strategy_var.get()),
        )

    def start_generation(self) -> None:
        if self._generation_running:
            return
        if not self._ensure_workbook_context(reload_rows=False):
            return

        rows = self._selected_rows_for_generation()
        if not rows:
            messagebox.showwarning("Sin filas para procesar", "No hay filas disponibles con los filtros o selecciones actuales.")
            return

        self.progress["value"] = 0
        self.progress["maximum"] = max(len(rows), 1)
        self.status_var.set("Iniciando generación...")
        self.report_var.set("Generación en curso...")
        self.report_path = None
        self.view_report_button.configure(state="disabled")
        self.execution_total_var.set(str(len(rows)))
        self.execution_done_var.set("0")
        self.execution_failed_var.set("0")
        self.execution_skipped_var.set("0")

        self._cancel_requested = False
        self._generation_running = True
        self.generate_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")

        options = self._build_options()

        def worker() -> None:
            try:
                summary = self.batch_service.run(options, rows, progress_callback=self._enqueue_progress)
                self.queue.put(("done", summary))
            except GenerationCancelled:
                self.queue.put(("cancelled", None))
            except Exception as exc:  # pragma: no cover - background thread
                self.queue.put(("error", exc))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def cancel_generation(self) -> None:
        if not self._generation_running:
            self.destroy()
            return
        self._cancel_requested = True
        self.status_var.set("Cancelación solicitada...")

    def _enqueue_progress(self, current: int, total: int, message: str) -> None:
        if self._cancel_requested:
            raise GenerationCancelled()
        self.queue.put(("progress", current, total, message))

    def _poll_queue(self) -> None:
        try:
            while True:
                event = self.queue.get_nowait()
                kind = event[0]
                if kind == "progress":
                    _, current, total, message = event
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = current
                    self.status_var.set(message)
                elif kind == "done":
                    _, summary = event
                    self._generation_running = False
                    self.progress["value"] = self.progress["maximum"]
                    self.status_var.set("Generación completada")
                    self.report_path = summary.report_path
                    self.report_var.set(f"Reporte generado: {summary.report_path}")
                    self.execution_total_var.set(str(summary.total))
                    self.execution_done_var.set(str(summary.succeeded))
                    self.execution_failed_var.set(str(summary.failed))
                    self.execution_skipped_var.set(str(summary.skipped))
                    self.generate_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.view_report_button.configure(state="normal")
                    messagebox.showinfo(
                        "Generación completa",
                        f"Generados: {summary.succeeded}\nFallidos: {summary.failed}\nOmitidos: {summary.skipped}\nInforme: {summary.report_path}",
                    )
                elif kind == "cancelled":
                    self._generation_running = False
                    self.status_var.set("Generación cancelada")
                    self.report_var.set("Proceso cancelado por el usuario")
                    self.generate_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    messagebox.showinfo("Generación cancelada", "El proceso se canceló antes de terminar.")
                elif kind == "error":
                    _, exc = event
                    self._generation_running = False
                    self.status_var.set("La generación falló")
                    self.generate_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    messagebox.showerror("La generación falló", str(exc))
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def open_output_folder(self) -> None:
        output = Path(self.output_var.get())
        if not output.exists():
            messagebox.showwarning("Carpeta de salida", f"La carpeta aún no existe:\n{output}")
            return
        try:
            open_folder(output)
        except Exception as exc:
            messagebox.showerror("No se pudo abrir la carpeta", str(exc))

    def open_report(self) -> None:
        if self.report_path is None:
            messagebox.showwarning("Reporte", "Primero debe completar una generación para abrir el reporte.")
            return
        if not self.report_path.exists():
            messagebox.showwarning("Reporte", f"El reporte no existe:\n{self.report_path}")
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(self.report_path)  # type: ignore[attr-defined]
                return
            open_folder(self.report_path.parent)
        except Exception as exc:
            messagebox.showerror("No se pudo abrir el reporte", str(exc))

    def _on_tree_click(self, event: tk.Event) -> str | None:
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return None

        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item or column != "#1":
            return None

        self._toggle_row(item)
        return "break"

    def _toggle_selected_from_keyboard(self, _event: tk.Event) -> str | None:
        selection = self.tree.selection()
        if not selection:
            return None
        self._toggle_row(selection[0])
        return "break"

    def _toggle_row(self, item_id: str) -> None:
        if not item_id:
            return

        row_number = int(item_id)
        row = next((candidate for candidate in self.series_rows if candidate.row_number == row_number), None)
        if row is None:
            return

        row.selected = not row.selected
        if self.tree.exists(item_id):
            self.tree.set(item_id, "select", "☑" if row.selected else "☐")
        self._refresh_tree()
