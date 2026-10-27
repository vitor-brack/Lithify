"""GUI principal do Lithify."""
from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Optional

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    HAS_DND = True
except ImportError:
    HAS_DND = False
    TkinterDnD = None  # type: ignore[assignment]
    DND_FILES = None  # type: ignore[assignment]

from core.document_processor import comprimir_documento, detect_format
from core.docx_processor import DocxCompressionResult
from core.exceptions import (
    InvalidDocxError,
    LegacyDocFormatError,
    WordNotAvailableError,
)
from core.image_compressor import (
    SUPPORTED_EXTENSIONS,
    CompressionResult,
    comprimir_lote,
)

logger = logging.getLogger(__name__)


PRESETS: dict[str, Optional[tuple[int, int]]] = {
    "Tela / PDF (recomendado)": (1920, 82),
    "Impressão": (2500, 88),
    "Personalizado": None,
}

IMAGE_FILETYPES = [
    ("Imagens", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp"),
    ("Todos os arquivos", "*.*"),
]
DOCX_FILETYPES = [
    ("Documento Word", "*.doc *.docx"),
    ("Documento Word moderno (.docx)", "*.docx"),
    ("Documento Word antigo (.doc)", "*.doc"),
    ("Todos os arquivos", "*.*"),
]
DOC_EXTENSIONS = {".doc", ".docx"}


def make_root() -> tk.Tk:
    """Cria a janela raiz com suporte a DnD se disponível."""
    if HAS_DND:
        return TkinterDnD.Tk()
    return tk.Tk()


class MainWindow:
    """Janela principal com abas para os dois modos."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Lithify — Compressor de Imagens para Laudos")
        self.root.geometry("780x620")
        self.root.minsize(680, 540)

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.image_tab = ImageTab(notebook)
        self.docx_tab = DocxTab(notebook)

        notebook.add(self.image_tab, text="  Comprimir Imagens  ")
        notebook.add(self.docx_tab, text="  Comprimir Word  ")
        notebook.select(self.image_tab)


class _BaseTab(ttk.Frame):
    """Plumbing comum: queue de mensagens entre worker thread e main loop."""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=10)
        self._queue: "queue.Queue[tuple[str, Any]]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._poll_after_id: Optional[str] = None

    def _start_worker(self, target: Callable[[], None]) -> None:
        if self._worker and self._worker.is_alive():
            return
        self._worker = threading.Thread(target=self._wrap_worker(target), daemon=True)
        self._worker.start()
        self._poll_queue()

    def _wrap_worker(self, target: Callable[[], None]) -> Callable[[], None]:
        def runner() -> None:
            try:
                target()
            except Exception as exc:
                logger.exception("Erro no worker")
                self._queue.put(("error", exc))
            finally:
                self._queue.put(("done", None))

        return runner

    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self._queue.get_nowait()
                self._handle_message(msg)
        except queue.Empty:
            pass

        if self._worker and self._worker.is_alive():
            self._poll_after_id = self.after(100, self._poll_queue)
        else:
            try:
                while True:
                    self._handle_message(self._queue.get_nowait())
            except queue.Empty:
                pass
            self._poll_after_id = None

    def _handle_message(self, msg: tuple[str, Any]) -> None:
        kind, payload = msg
        if kind == "progress":
            cur, total, name = payload
            self._on_progress(cur, total, name)
        elif kind == "done":
            self._on_done()
        elif kind == "error":
            self._on_error(payload)
        elif kind == "result":
            self._on_result(payload)
        elif kind == "log":
            self._on_log(str(payload))
        elif kind == "stage":
            self._on_stage(str(payload))

    def _on_progress(self, cur: int, total: int, name: str) -> None: ...
    def _on_done(self) -> None: ...
    def _on_error(self, exc: BaseException) -> None: ...
    def _on_result(self, result: Any) -> None: ...
    def _on_log(self, message: str) -> None: ...
    def _on_stage(self, message: str) -> None: ...


def _parse_drop_paths(widget: tk.Misc, data: str) -> list[Path]:
    raw = widget.tk.splitlist(data)
    return [Path(p) for p in raw]


def _expand_image_paths(paths: list[Path]) -> list[Path]:
    """Expande pastas em arquivos de imagem suportados (recursivo)."""
    out: list[Path] = []
    for p in paths:
        if p.is_dir():
            for child in sorted(p.rglob("*")):
                if child.is_file() and child.suffix.lower() in SUPPORTED_EXTENSIONS:
                    out.append(child)
        elif p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
            out.append(p)
    return out


def _format_bytes(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024 * 1024:
        return f"{n / 1024:.1f} KB"
    return f"{n / (1024 * 1024):.2f} MB"


class _DropZone(tk.Frame):
    """Área de drop estilizada com label central."""

    def __init__(
        self,
        master: tk.Misc,
        on_drop: Callable[[list[Path]], None],
        instruction: str,
        click_action: Optional[Callable[[], None]] = None,
    ) -> None:
        super().__init__(master, bg="#eef3f8", relief="ridge", bd=2, highlightthickness=0)
        self._on_drop = on_drop

        self.label = tk.Label(
            self,
            text=instruction,
            bg="#eef3f8",
            fg="#34465c",
            font=("Segoe UI", 11),
            justify="center",
        )
        self.label.pack(expand=True, padx=20, pady=20)

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._handle_drop)
            self.dnd_bind("<<DropEnter>>", lambda e: self._set_active(True))
            self.dnd_bind("<<DropLeave>>", lambda e: self._set_active(False))

        if click_action is not None:
            for w in (self, self.label):
                w.bind("<Button-1>", lambda e: click_action())
                w.configure(cursor="hand2")

    def _handle_drop(self, event: tk.Event) -> None:
        self._set_active(False)
        paths = _parse_drop_paths(self, event.data)
        self._on_drop(paths)

    def _set_active(self, active: bool) -> None:
        color = "#d4e4f4" if active else "#eef3f8"
        self.configure(bg=color)
        self.label.configure(bg=color)

    def set_text(self, text: str) -> None:
        self.label.configure(text=text)


class ImageTab(_BaseTab):
    """Modo 1 — Comprimir imagens soltas."""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self.selected_paths: list[Path] = []
        self._build_ui()
        self._on_preset_change()

    def _build_ui(self) -> None:
        instruction = (
            "Arraste imagens ou uma pasta aqui\n\nou clique para selecionar arquivos"
            if HAS_DND
            else "Clique aqui para selecionar arquivos\n\n(arraste-e-solte indisponível)"
        )
        self.drop_zone = _DropZone(
            self,
            on_drop=self._on_drop_paths,
            instruction=instruction,
            click_action=self._select_files,
        )
        self.drop_zone.pack(fill="both", expand=True, pady=(0, 8))

        btn_row = ttk.Frame(self)
        btn_row.pack(fill="x", pady=(0, 8))
        ttk.Button(btn_row, text="Selecionar arquivos", command=self._select_files).pack(side="left")
        ttk.Button(btn_row, text="Selecionar pasta", command=self._select_folder).pack(side="left", padx=6)
        ttk.Button(btn_row, text="Limpar", command=self._clear).pack(side="left")

        cfg = ttk.LabelFrame(self, text="Configuração", padding=10)
        cfg.pack(fill="x", pady=(0, 8))
        self._build_preset_widgets(cfg)

        self.preserve_alpha_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            cfg,
            text="Preservar transparência (PNGs com alpha continuam como PNG)",
            variable=self.preserve_alpha_var,
        ).pack(anchor="w", pady=(4, 2))

        out_row = ttk.Frame(cfg)
        out_row.pack(fill="x", pady=(4, 0))
        ttk.Label(out_row, text="Pasta de saída:").pack(side="left")
        self.output_dir_var = tk.StringVar(value="")
        ttk.Entry(out_row, textvariable=self.output_dir_var).pack(
            side="left", fill="x", expand=True, padx=6
        )
        ttk.Button(out_row, text="...", width=3, command=self._select_output_dir).pack(side="left")
        ttk.Label(
            cfg,
            text="(vazio = subpasta 'comprimidas/' ao lado dos originais)",
            foreground="#666",
            font=("Segoe UI", 8),
        ).pack(anchor="w")

        self.overwrite_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            cfg,
            text="Sobrescrever os arquivos originais (perigoso)",
            variable=self.overwrite_var,
        ).pack(anchor="w", pady=(4, 0))

        self.compress_btn = ttk.Button(self, text="Comprimir", command=self._on_compress)
        self.compress_btn.pack(fill="x", ipady=6, pady=(0, 4))

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x")

        self.status_var = tk.StringVar(value="Nenhum arquivo selecionado.")
        ttk.Label(self, textvariable=self.status_var, foreground="#445566").pack(
            anchor="w", pady=(4, 0)
        )

        self.log = _LogBox(self, height=5)
        self.log.pack(fill="both", expand=False, pady=(4, 0))

    def _build_preset_widgets(self, parent: ttk.Frame) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x")
        ttk.Label(row, text="Preset:").pack(side="left")
        self.preset_var = tk.StringVar(value="Tela / PDF (recomendado)")
        combo = ttk.Combobox(
            row,
            textvariable=self.preset_var,
            values=list(PRESETS.keys()),
            state="readonly",
            width=28,
        )
        combo.pack(side="left", padx=6)
        combo.bind("<<ComboboxSelected>>", lambda _e: self._on_preset_change())

        self.sliders = ttk.Frame(parent)
        self.max_dim_var = tk.IntVar(value=1920)
        self.quality_var = tk.IntVar(value=82)
        ttk.Label(self.sliders, text="Dimensão máxima:").grid(row=0, column=0, sticky="w")
        self.max_dim_label = ttk.Label(self.sliders, text="1920 px", width=8)
        self.max_dim_label.grid(row=0, column=2, sticky="w", padx=4)
        ttk.Scale(
            self.sliders,
            from_=720,
            to=4000,
            variable=self.max_dim_var,
            orient="horizontal",
            command=lambda v: self.max_dim_label.configure(text=f"{int(float(v))} px"),
        ).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Label(self.sliders, text="Qualidade JPEG:").grid(row=1, column=0, sticky="w")
        self.quality_label = ttk.Label(self.sliders, text="82", width=8)
        self.quality_label.grid(row=1, column=2, sticky="w", padx=4)
        ttk.Scale(
            self.sliders,
            from_=60,
            to=95,
            variable=self.quality_var,
            orient="horizontal",
            command=lambda v: self.quality_label.configure(text=f"{int(float(v))}"),
        ).grid(row=1, column=1, sticky="ew", padx=4)
        self.sliders.columnconfigure(1, weight=1)

    def _on_preset_change(self) -> None:
        preset = PRESETS.get(self.preset_var.get())
        if preset is None:
            self.sliders.pack(fill="x", pady=(6, 0))
        else:
            self.sliders.pack_forget()
            self.max_dim_var.set(preset[0])
            self.quality_var.set(preset[1])

    def _select_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Selecionar imagens",
            filetypes=IMAGE_FILETYPES,
        )
        if paths:
            self._add_paths([Path(p) for p in paths])

    def _select_folder(self) -> None:
        folder = filedialog.askdirectory(title="Selecionar pasta")
        if folder:
            self._add_paths([Path(folder)])

    def _select_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Pasta de saída")
        if folder:
            self.output_dir_var.set(folder)

    def _on_drop_paths(self, paths: list[Path]) -> None:
        self._add_paths(paths)

    def _add_paths(self, paths: list[Path]) -> None:
        expanded = _expand_image_paths(paths)
        if not expanded:
            messagebox.showwarning(
                "Nenhuma imagem encontrada",
                "Os itens selecionados não contêm imagens em formato suportado.",
            )
            return
        existing = set(self.selected_paths)
        for p in expanded:
            if p not in existing:
                self.selected_paths.append(p)
        self._refresh_status()

    def _clear(self) -> None:
        self.selected_paths.clear()
        self._refresh_status()
        self.progress.configure(value=0)
        self.log.clear()

    def _refresh_status(self) -> None:
        n = len(self.selected_paths)
        if n == 0:
            self.status_var.set("Nenhum arquivo selecionado.")
            self.drop_zone.set_text(
                "Arraste imagens ou uma pasta aqui\n\nou clique para selecionar arquivos"
                if HAS_DND
                else "Clique aqui para selecionar arquivos\n\n(arraste-e-solte indisponível)"
            )
        else:
            self.status_var.set(f"{n} arquivo(s) selecionado(s).")
            self.drop_zone.set_text(f"{n} arquivo(s) prontos para comprimir.\n\n(arraste mais para adicionar)")

    def _resolve_output_dir(self) -> Optional[Path]:
        if self.overwrite_var.get():
            return None
        text = self.output_dir_var.get().strip()
        if text:
            return Path(text)
        first_parent = self.selected_paths[0].parent
        return first_parent / "comprimidas"

    def _on_compress(self) -> None:
        if not self.selected_paths:
            messagebox.showinfo("Sem arquivos", "Selecione imagens antes de comprimir.")
            return

        if self.overwrite_var.get():
            confirmed = messagebox.askyesno(
                "Sobrescrever originais",
                "Tem certeza? Os arquivos originais serão substituídos. Esta ação não pode ser desfeita.\n\nFortemente recomendado: faça backup antes.",
                icon="warning",
            )
            if not confirmed:
                return

        output_dir = self._resolve_output_dir()
        max_dim = int(self.max_dim_var.get())
        quality = int(self.quality_var.get())
        preserve_alpha = self.preserve_alpha_var.get()
        paths = list(self.selected_paths)
        overwrite = self.overwrite_var.get()

        try:
            if output_dir is not None:
                output_dir.mkdir(parents=True, exist_ok=True)
                test = output_dir / ".lithify_write_test"
                test.write_text("ok", encoding="utf-8")
                test.unlink()
        except OSError as exc:
            messagebox.showerror(
                "Pasta de saída inválida",
                f"Não foi possível escrever na pasta de saída:\n{exc}",
            )
            return

        self._set_running(True)
        self.progress.configure(value=0, maximum=len(paths))
        self.log.clear()
        self.log.append(f"Iniciando compressão de {len(paths)} arquivo(s)...")

        def task() -> None:
            results: list[CompressionResult] = []

            def cb(cur: int, total: int, name: str) -> None:
                self._queue.put(("progress", (cur, total, name)))

            if overwrite:
                from core.image_compressor import comprimir_imagem
                from core.exceptions import (
                    CorruptImageError,
                    UnsupportedFormatError,
                )

                total = len(paths)
                for i, p in enumerate(paths, start=1):
                    cb(i, total, p.name)
                    try:
                        result = comprimir_imagem(
                            input_path=p,
                            output_path=p,
                            max_dim=max_dim,
                            quality=quality,
                            preserve_alpha=preserve_alpha,
                        )
                        results.append(result)
                    except (CorruptImageError, UnsupportedFormatError, FileNotFoundError) as exc:
                        self._queue.put(("log", f"Pulado: {p.name} — {exc}"))
            else:
                assert output_dir is not None
                results = comprimir_lote(
                    inputs=paths,
                    output_dir=output_dir,
                    max_dim=max_dim,
                    quality=quality,
                    preserve_alpha=preserve_alpha,
                    progress_callback=cb,
                )

            self._queue.put(("result", results))

        self._start_worker(task)

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.compress_btn.configure(state=state, text="Comprimindo..." if running else "Comprimir")

    def _on_progress(self, cur: int, total: int, name: str) -> None:
        self.progress.configure(value=cur, maximum=total)
        self.status_var.set(f"({cur}/{total}) {name}")

    def _on_log(self, message: str) -> None:
        self.log.append(message)

    def _on_result(self, result: Any) -> None:
        results: list[CompressionResult] = result
        if not results:
            self.log.append("Nenhum arquivo foi processado.")
            return
        total_orig = sum(r.original_size for r in results)
        total_new = sum(r.compressed_size for r in results)
        saved = total_orig - total_new
        ratio = (saved / total_orig * 100) if total_orig else 0
        skipped = sum(1 for r in results if r.was_skipped)
        compressed = len(results) - skipped
        self.log.append("")
        self.log.append(f"Concluído: {compressed} comprimidas, {skipped} já otimizadas (puladas).")
        self.log.append(
            f"Antes: {_format_bytes(total_orig)} → Depois: {_format_bytes(total_new)} "
            f"(economia: {_format_bytes(saved)}, {ratio:.1f}%)."
        )

    def _on_done(self) -> None:
        self._set_running(False)
        self.status_var.set(f"{len(self.selected_paths)} arquivo(s) selecionado(s).")

    def _on_error(self, exc: BaseException) -> None:
        self.log.append(f"Erro: {exc}")
        messagebox.showerror("Erro", str(exc))


class DocxTab(_BaseTab):
    """Modo 2 — Comprimir imagens dentro de um .docx existente."""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self.selected_docx: Optional[Path] = None
        self._build_ui()
        self._on_preset_change()

    def _build_ui(self) -> None:
        intro = ttk.Label(
            self,
            text=(
                "Use isto quando você esqueceu de comprimir as imagens antes de inserir no Word.\n"
                "Aceita .docx e .doc (formato antigo). O ideal é usar a aba 'Comprimir Imagens' "
                "antes de montar o documento."
            ),
            foreground="#445566",
            justify="left",
        )
        intro.pack(anchor="w", pady=(0, 8))

        instruction = (
            "Arraste um arquivo .doc ou .docx aqui\n\nou clique para selecionar"
            if HAS_DND
            else "Clique aqui para selecionar um .doc ou .docx\n\n(arraste-e-solte indisponível)"
        )
        self.drop_zone = _DropZone(
            self,
            on_drop=self._on_drop_paths,
            instruction=instruction,
            click_action=self._select_docx,
        )
        self.drop_zone.pack(fill="both", expand=True, pady=(0, 8))

        cfg = ttk.LabelFrame(self, text="Configuração", padding=10)
        cfg.pack(fill="x", pady=(0, 8))
        self._build_preset_widgets(cfg)

        self.preserve_alpha_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            cfg,
            text="Preservar transparência (PNGs com alpha continuam como PNG)",
            variable=self.preserve_alpha_var,
        ).pack(anchor="w", pady=(4, 2))

        out_row = ttk.Frame(cfg)
        out_row.pack(fill="x", pady=(4, 0))
        ttk.Label(out_row, text="Pasta de saída:").pack(side="left")
        self.output_dir_var = tk.StringVar(value="")
        ttk.Entry(out_row, textvariable=self.output_dir_var).pack(
            side="left", fill="x", expand=True, padx=6
        )
        ttk.Button(out_row, text="...", width=3, command=self._select_output_dir).pack(side="left")
        ttk.Label(
            cfg,
            text="(vazio = subpasta 'comprimidos/' ao lado do .docx original; mesmo nome de arquivo)",
            foreground="#666",
            font=("Segoe UI", 8),
        ).pack(anchor="w")

        self.compress_btn = ttk.Button(self, text="Comprimir Word", command=self._on_compress)
        self.compress_btn.pack(fill="x", ipady=6, pady=(0, 4))

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x")

        self.status_var = tk.StringVar(value="Nenhum .docx selecionado.")
        ttk.Label(self, textvariable=self.status_var, foreground="#445566").pack(
            anchor="w", pady=(4, 0)
        )

        self.log = _LogBox(self, height=6)
        self.log.pack(fill="both", expand=False, pady=(4, 0))

    def _build_preset_widgets(self, parent: ttk.Frame) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x")
        ttk.Label(row, text="Preset:").pack(side="left")
        self.preset_var = tk.StringVar(value="Tela / PDF (recomendado)")
        combo = ttk.Combobox(
            row,
            textvariable=self.preset_var,
            values=list(PRESETS.keys()),
            state="readonly",
            width=28,
        )
        combo.pack(side="left", padx=6)
        combo.bind("<<ComboboxSelected>>", lambda _e: self._on_preset_change())

        self.sliders = ttk.Frame(parent)
        self.max_dim_var = tk.IntVar(value=1920)
        self.quality_var = tk.IntVar(value=82)
        ttk.Label(self.sliders, text="Dimensão máxima:").grid(row=0, column=0, sticky="w")
        self.max_dim_label = ttk.Label(self.sliders, text="1920 px", width=8)
        self.max_dim_label.grid(row=0, column=2, sticky="w", padx=4)
        ttk.Scale(
            self.sliders,
            from_=720,
            to=4000,
            variable=self.max_dim_var,
            orient="horizontal",
            command=lambda v: self.max_dim_label.configure(text=f"{int(float(v))} px"),
        ).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Label(self.sliders, text="Qualidade JPEG:").grid(row=1, column=0, sticky="w")
        self.quality_label = ttk.Label(self.sliders, text="82", width=8)
        self.quality_label.grid(row=1, column=2, sticky="w", padx=4)
        ttk.Scale(
            self.sliders,
            from_=60,
            to=95,
            variable=self.quality_var,
            orient="horizontal",
            command=lambda v: self.quality_label.configure(text=f"{int(float(v))}"),
        ).grid(row=1, column=1, sticky="ew", padx=4)
        self.sliders.columnconfigure(1, weight=1)

    def _on_preset_change(self) -> None:
        preset = PRESETS.get(self.preset_var.get())
        if preset is None:
            self.sliders.pack(fill="x", pady=(6, 0))
        else:
            self.sliders.pack_forget()
            self.max_dim_var.set(preset[0])
            self.quality_var.set(preset[1])

    def _select_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecionar documento Word (.doc ou .docx)",
            filetypes=DOCX_FILETYPES,
        )
        if path:
            self._set_docx(Path(path))

    def _select_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Pasta de saída")
        if folder:
            self.output_dir_var.set(folder)

    def _on_drop_paths(self, paths: list[Path]) -> None:
        if not paths:
            return
        path = paths[0]
        if path.is_dir() or path.suffix.lower() not in DOC_EXTENSIONS:
            messagebox.showwarning(
                "Arquivo inválido",
                "Arraste um único arquivo .doc ou .docx.",
            )
            return
        self._set_docx(path)

    def _set_docx(self, path: Path) -> None:
        self.selected_docx = path
        size_mb = path.stat().st_size / (1024 * 1024) if path.exists() else 0
        self.status_var.set(f"{path.name}  ({size_mb:.2f} MB)")
        self.drop_zone.set_text(f"{path.name}\n\n({size_mb:.2f} MB)")

    def _resolve_output_path(self) -> Optional[Path]:
        if self.selected_docx is None:
            return None
        text = self.output_dir_var.get().strip()
        if text:
            return Path(text) / self.selected_docx.name
        return self.selected_docx.parent / "comprimidos" / self.selected_docx.name

    def _on_compress(self) -> None:
        if self.selected_docx is None:
            messagebox.showinfo("Sem arquivo", "Selecione um .docx antes de comprimir.")
            return

        output_path = self._resolve_output_path()
        assert output_path is not None
        if output_path.resolve() == self.selected_docx.resolve():
            messagebox.showerror(
                "Saída inválida",
                "A pasta de saída resultaria em sobrescrever o original. Escolha outra pasta.",
            )
            return

        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            test = output_path.parent / ".lithify_write_test"
            test.write_text("ok", encoding="utf-8")
            test.unlink()
        except OSError as exc:
            messagebox.showerror(
                "Pasta de saída inválida",
                f"Não foi possível escrever na pasta de saída:\n{exc}",
            )
            return

        max_dim = int(self.max_dim_var.get())
        quality = int(self.quality_var.get())
        preserve_alpha = self.preserve_alpha_var.get()
        src = self.selected_docx

        self._set_running(True)
        self.progress.configure(value=0, maximum=1)
        self.log.clear()
        self.log.append(f"Comprimindo: {src.name}")
        self.log.append(f"Saída: {output_path}")
        if src.suffix.lower() == ".doc":
            self.log.append(
                "Formato .doc detectado: o Word será aberto em segundo plano "
                "para a conversão. Isso pode levar alguns segundos."
            )

        def task() -> None:
            def progress_cb(cur: int, total: int, name: str) -> None:
                self._queue.put(("progress", (cur, total, name)))

            def stage_cb(message: str) -> None:
                self._queue.put(("stage", message))

            try:
                result = comprimir_documento(
                    input_path=src,
                    output_path=output_path,
                    max_dim=max_dim,
                    quality=quality,
                    preserve_alpha=preserve_alpha,
                    progress_callback=progress_cb,
                    stage_callback=stage_cb,
                )
            except (LegacyDocFormatError, InvalidDocxError, WordNotAvailableError) as exc:
                self._queue.put(("error", exc))
                return
            self._queue.put(("result", result))

        self._start_worker(task)

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.compress_btn.configure(
            state=state,
            text="Comprimindo..." if running else "Comprimir Word",
        )

    def _on_progress(self, cur: int, total: int, name: str) -> None:
        self.progress.configure(value=cur, maximum=max(total, 1))
        self.status_var.set(f"Imagem ({cur}/{total}): {name}")

    def _on_stage(self, message: str) -> None:
        self.status_var.set(message)
        self.log.append(message)

    def _on_log(self, message: str) -> None:
        self.log.append(message)

    def _on_result(self, result: Any) -> None:
        r: DocxCompressionResult = result
        self.log.append("")
        self.log.append(f"Antes: {_format_bytes(r.original_size)}")
        self.log.append(
            f"Depois: {_format_bytes(r.compressed_size)} "
            f"(economia: {_format_bytes(r.saved_bytes)}, {r.saved_ratio * 100:.1f}%)"
        )
        self.log.append(f"Imagens comprimidas: {r.images_processed}")
        if r.images_skipped_already_optimized:
            self.log.append(f"Já otimizadas (puladas): {r.images_skipped_already_optimized}")
        if r.images_skipped_non_raster:
            self.log.append(f"Vetoriais (mantidas intactas): {r.images_skipped_non_raster}")
        if r.images_failed:
            self.log.append(f"Falharam (mantido original): {r.images_failed}")
        if r.oles_processed:
            self.log.append(f"Bitmaps embutidos (OLE) comprimidos: {r.oles_processed}")
        if r.oles_skipped:
            self.log.append(f"Bitmaps embutidos pulados: {r.oles_skipped}")

    def _on_done(self) -> None:
        self._set_running(False)
        if self.selected_docx is not None:
            size_mb = self.selected_docx.stat().st_size / (1024 * 1024)
            self.status_var.set(f"{self.selected_docx.name}  ({size_mb:.2f} MB)")

    def _on_error(self, exc: BaseException) -> None:
        self.log.append(f"Erro: {exc}")
        if isinstance(exc, WordNotAvailableError):
            messagebox.showerror("Microsoft Word não disponível", str(exc))
        else:
            messagebox.showerror("Erro", str(exc))


class _LogBox(ttk.Frame):
    def __init__(self, master: tk.Misc, height: int = 6) -> None:
        super().__init__(master)
        self._text = tk.Text(
            self,
            height=height,
            wrap="word",
            state="disabled",
            background="#fbfcfd",
            relief="solid",
            borderwidth=1,
            font=("Consolas", 9),
        )
        scroll = ttk.Scrollbar(self, orient="vertical", command=self._text.yview)
        self._text.configure(yscrollcommand=scroll.set)
        self._text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def append(self, message: str) -> None:
        self._text.configure(state="normal")
        self._text.insert("end", message + "\n")
        self._text.see("end")
        self._text.configure(state="disabled")

    def clear(self) -> None:
        self._text.configure(state="normal")
        self._text.delete("1.0", "end")
        self._text.configure(state="disabled")
