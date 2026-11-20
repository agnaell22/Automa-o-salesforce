import ctypes
import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from automacao import (
    DEFAULT_COLUMN,
    DEFAULT_OUTPUT,
    DEFAULT_SHEET,
    DEFAULT_WHATSAPP_GROUP,
    DEFAULT_WORKBOOK,
    main as run_automation,
)

APP_INSTALL_DIRECTORY_NAME = "AutomacaoSalesforce"
APP_EXECUTABLE_NAME = "Automacao Salesforce.exe"
SOURCE_FILES = ("automacao.py", "interface.py", "salesforce.py", "whatsapp.py")


def build_cli_args(
    workbook: str,
    output_dir: str,
    group_name: str,
    sheet_name: str,
    column_name: str,
    limit: str,
    consult_salesforce: bool,
    send_whatsapp: bool,
) -> list[str]:
    args = [
        "--planilha",
        workbook,
        "--aba",
        sheet_name,
        "--coluna",
        column_name,
        "--saida",
        output_dir,
    ]
    if consult_salesforce:
        args.extend(("--consultar-salesforce", "--limite", limit))
    if send_whatsapp:
        args.extend(("--enviar-whatsapp", "--grupo-whatsapp", group_name))
    return args


def _create_desktop_shortcut(executable: Path, working_directory: Path) -> None:
    def quote_powershell(value: Path) -> str:
        return "'" + str(value).replace("'", "''") + "'"

    command = (
        "$ErrorActionPreference = 'Stop'; "
        "$desktop = [Environment]::GetFolderPath('Desktop'); "
        "if (-not (Test-Path -LiteralPath $desktop -PathType Container)) "
        "{ throw \"Desktop folder not found: $desktop\" }; "
        "$shell = New-Object -ComObject WScript.Shell; "
        "$shortcut = $shell.CreateShortcut((Join-Path $desktop "
        "'Automacao Salesforce.lnk')); "
        f"$shortcut.TargetPath = {quote_powershell(executable)}; "
        f"$shortcut.WorkingDirectory = {quote_powershell(working_directory)}; "
        "$shortcut.Save()"
    )
    subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            command,
        ],
        check=True,
        capture_output=True,
        text=True,
    )


def _install_packaged_app(
    executable: Path,
    install_directory: Path,
    bundled_source_directory: Path,
) -> bool:
    installed_executable = install_directory / APP_EXECUTABLE_NAME
    if executable.resolve() == installed_executable.resolve():
        return False

    missing_sources = [
        filename
        for filename in SOURCE_FILES
        if not (bundled_source_directory / filename).is_file()
    ]
    if missing_sources:
        raise FileNotFoundError(
            "O pacote não contém todos os fontes necessários: "
            + ", ".join(missing_sources)
        )

    install_directory.mkdir(parents=True, exist_ok=True)
    shutil.copy2(executable, installed_executable)
    source_directory = install_directory / "source"
    source_directory.mkdir(exist_ok=True)
    for filename in SOURCE_FILES:
        shutil.copy2(
            bundled_source_directory / filename,
            source_directory / filename,
        )

    _create_desktop_shortcut(installed_executable, install_directory)
    subprocess.Popen([str(installed_executable)], cwd=install_directory)
    return True


def _install_if_packaged() -> bool:
    if not getattr(sys, "frozen", False):
        return False
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise RuntimeError("A variável LOCALAPPDATA não está disponível.")

    install_directory = (
        Path(local_app_data) / APP_INSTALL_DIRECTORY_NAME / "app"
    )
    bundled_root = getattr(sys, "_MEIPASS", None)
    if not bundled_root:
        raise RuntimeError("Não foi possível localizar os arquivos do pacote.")
    return _install_packaged_app(
        Path(sys.executable),
        install_directory,
        Path(bundled_root) / "source",
    )


class AutomationApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Automação Salesforce e WhatsApp")
        self.root.geometry("820x690")
        self.root.minsize(700, 580)

        self.workbook_var = tk.StringVar(value=str(DEFAULT_WORKBOOK))
        self.output_var = tk.StringVar(value=str(DEFAULT_OUTPUT))
        self.sheet_var = tk.StringVar(value=DEFAULT_SHEET)
        self.column_var = tk.StringVar(value=DEFAULT_COLUMN)
        self.group_var = tk.StringVar(value=DEFAULT_WHATSAPP_GROUP)
        self.limit_var = tk.StringVar(value="1")
        self.consult_var = tk.BooleanVar(value=True)
        self.send_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Pronto para iniciar.")
        self.salesforce_progress_var = tk.StringVar(
            value="Salesforce: sem dados carregados."
        )
        self.whatsapp_progress_var = tk.StringVar(
            value="WhatsApp: sem dados carregados."
        )

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._close)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(7, weight=1)

        ttk.Label(
            container,
            text="Processamento de notas Salesforce",
            font=("", 16, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 4))
        ttk.Label(
            container,
            text=(
                "Compare a planilha com o relatório, consulte notas novas "
                "e publique os resultados quando autorizado."
            ),
            wraplength=760,
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(0, 14))

        self._path_row(
            container,
            row=2,
            label="Planilha Excel",
            variable=self.workbook_var,
            button_text="Selecionar...",
            command=self._browse_workbook,
        )
        self._path_row(
            container,
            row=3,
            label="Pasta de respostas",
            variable=self.output_var,
            button_text="Selecionar...",
            command=self._browse_output,
        )

        options = ttk.LabelFrame(container, text="Configuração", padding=10)
        options.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=(12, 8))
        options.columnconfigure(1, weight=1)
        ttk.Label(options, text="Aba").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(options, textvariable=self.sheet_var).grid(
            row=0, column=1, sticky=tk.EW, pady=3
        )
        ttk.Label(options, text="Coluna da nota").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 8)
        )
        ttk.Entry(options, textvariable=self.column_var).grid(
            row=1, column=1, sticky=tk.EW, pady=3
        )
        ttk.Label(options, text="Máximo de notas por consulta").grid(
            row=2, column=0, sticky=tk.W, padx=(0, 8)
        )
        ttk.Entry(options, textvariable=self.limit_var, width=10).grid(
            row=2, column=1, sticky=tk.W, pady=3
        )
        ttk.Checkbutton(
            options,
            text="Consultar notas pendentes no Salesforce e baixar imagens",
            variable=self.consult_var,
        ).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(7, 2))
        ttk.Checkbutton(
            options,
            text="Enviar notas concluídas e arquivo ZIP pelo WhatsApp",
            variable=self.send_var,
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Label(options, text="Nome exato do grupo WhatsApp").grid(
            row=5, column=0, sticky=tk.W, padx=(0, 8), pady=(5, 0)
        )
        ttk.Entry(options, textvariable=self.group_var).grid(
            row=5, column=1, sticky=tk.EW, pady=(5, 0)
        )

        counters = ttk.LabelFrame(container, text="Contadores", padding=8)
        counters.grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=(8, 0))
        ttk.Label(counters, textvariable=self.salesforce_progress_var).pack(
            anchor=tk.W
        )
        ttk.Label(counters, textvariable=self.whatsapp_progress_var).pack(
            anchor=tk.W
        )

        log_frame = ttk.LabelFrame(container, text="Andamento", padding=8)
        log_frame.grid(row=7, column=0, columnspan=3, sticky=tk.NSEW, pady=(8, 8))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=12, wrap=tk.WORD, state=tk.DISABLED)
        self.log.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.log.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(container)
        footer.grid(row=8, column=0, columnspan=3, sticky=tk.EW)
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status_var).grid(
            row=0, column=0, sticky=tk.W
        )
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=130)
        self.progress.grid(row=0, column=1, padx=8)
        self.report_button = ttk.Button(
            footer,
            text="Abrir relatório",
            command=self._open_report,
        )
        self.report_button.grid(row=0, column=2, padx=(0, 8))
        self.start_button = ttk.Button(
            footer,
            text="Iniciar processamento",
            command=self._start,
        )
        self.start_button.grid(row=0, column=3)

    @staticmethod
    def _path_row(
        parent: ttk.Frame,
        *,
        row: int,
        label: str,
        variable: tk.StringVar,
        button_text: str,
        command: Callable[[], None],
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky=tk.EW, pady=3
        )
        ttk.Button(parent, text=button_text, command=command).grid(
            row=row, column=2, padx=(8, 0)
        )

    def _browse_workbook(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione a planilha de notas",
            filetypes=(("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")),
            initialdir=str(Path(self.workbook_var.get()).parent),
        )
        if selected:
            self.workbook_var.set(selected)

    def _browse_output(self) -> None:
        selected = filedialog.askdirectory(
            title="Selecione a pasta de respostas",
            initialdir=self.output_var.get(),
            mustexist=False,
        )
        if selected:
            self.output_var.set(selected)

    def _open_report(self) -> None:
        report_path = Path(self.output_var.get()) / "relatorio_processamento.csv"
        if not report_path.is_file():
            messagebox.showinfo(
                "Relatório ainda não existe",
                "Execute o processamento para criar o relatório.",
                parent=self.root,
            )
            return
        try:
            os.startfile(str(report_path))
        except OSError as error:
            messagebox.showerror(
                "Não foi possível abrir o relatório",
                str(error),
                parent=self.root,
            )

    def _append_log(self, message: str) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, message.rstrip() + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def _gui_input(self, _prompt: str) -> str:
        result = ["n"]
        failures: list[tk.TclError] = []
        completed = threading.Event()

        def ask_for_confirmation() -> None:
            try:
                confirmed = messagebox.askyesno(
                    "Confirmar envio pelo WhatsApp",
                    (
                        f"Confirma o envio para o grupo "
                        f"'{self.group_var.get().strip()}'?\n\n"
                        "O envio publicará as notas pendentes e o ZIP."
                    ),
                    parent=self.root,
                )
                result[0] = "s" if confirmed else "n"
            except tk.TclError as error:
                failures.append(error)
            finally:
                completed.set()

        self.root.after(0, ask_for_confirmation)
        completed.wait()
        if failures:
            raise RuntimeError(
                "Não foi possível exibir a confirmação do envio pelo WhatsApp."
            ) from failures[0]
        return result[0]

    def _start(self) -> None:
        workbook = self.workbook_var.get().strip()
        output_dir = self.output_var.get().strip()
        group_name = self.group_var.get().strip()
        if not workbook or not output_dir:
            messagebox.showerror(
                "Dados obrigatórios",
                "Selecione a planilha e a pasta de respostas.",
                parent=self.root,
            )
            return
        if self.send_var.get() and not group_name:
            messagebox.showerror(
                "Grupo obrigatório",
                "Informe o nome exato do grupo do WhatsApp.",
                parent=self.root,
            )
            return
        if self.consult_var.get():
            try:
                if int(self.limit_var.get()) < 1:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "Limite inválido",
                    "O máximo de notas por consulta deve ser um inteiro maior que zero.",
                    parent=self.root,
                )
                return

        args = build_cli_args(
            workbook,
            output_dir,
            group_name,
            self.sheet_var.get().strip(),
            self.column_var.get().strip(),
            self.limit_var.get().strip(),
            self.consult_var.get(),
            self.send_var.get(),
        )
        self.start_button.configure(state=tk.DISABLED)
        self.progress.start(12)
        self.status_var.set("Processando... Não feche esta janela.")
        self._append_log("Iniciando. O navegador pode abrir para login ou processamento.")
        threading.Thread(target=self._run, args=(args,), daemon=True).start()

    def _run(self, args: list[str]) -> None:
        try:
            result = run_automation(
                args,
                input_fn=self._gui_input,
                output_fn=lambda message: self.root.after(
                    0, self._append_log, message
                ),
                error_fn=lambda message: self.root.after(
                    0, self._append_log, message
                ),
                progress_fn=lambda counts: self.root.after(
                    0, self._update_progress, counts
                ),
            )
            self.root.after(0, self._finish, result == 0)
        except Exception as error:
            self.root.after(0, self._append_log, f"Erro inesperado: {error}")
            self.root.after(0, self._finish, False)

    def _update_progress(self, counts: dict[str, int]) -> None:
        salesforce_done = counts.get("salesforce_done", 0)
        salesforce_total = counts.get("salesforce_total", 0)
        whatsapp_sent = counts.get("whatsapp_sent", 0)
        whatsapp_total = counts.get("whatsapp_total", 0)
        self.salesforce_progress_var.set(
            f"Salesforce: {salesforce_done}/{salesforce_total} concluídas; "
            f"faltam consultar/reprocessar "
            f"{salesforce_total - salesforce_done}."
        )
        self.whatsapp_progress_var.set(
            f"WhatsApp: {whatsapp_sent}/{whatsapp_total} enviadas e "
            f"confirmadas; faltam {whatsapp_total - whatsapp_sent}."
        )

    def _finish(self, succeeded: bool) -> None:
        self.progress.stop()
        self.start_button.configure(state=tk.NORMAL)
        self.status_var.set(
            "Processamento concluído." if succeeded else "Processamento terminou com erro."
        )

    def _close(self) -> None:
        if str(self.start_button["state"]) == tk.DISABLED:
            messagebox.showinfo(
                "Processamento em andamento",
                "Aguarde o processamento terminar antes de fechar esta janela.",
                parent=self.root,
            )
            return
        self.root.destroy()


def main() -> None:
    try:
        if _install_if_packaged():
            return
    except (OSError, RuntimeError, subprocess.SubprocessError) as error:
        ctypes.windll.user32.MessageBoxW(
            None,
            f"Não foi possível instalar a automação para este usuário:\n{error}",
            "Instalação da Automação Salesforce",
            0x10,
        )
        return

    root = tk.Tk()
    AutomationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
