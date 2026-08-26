from __future__ import annotations

import json
import os
import shutil
import threading
import traceback
from typing import TYPE_CHECKING
from datetime import datetime, timezone
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .paths import project_root, resource_root

if TYPE_CHECKING:
    from .processor import ProcessingResult


ROOT = project_root()
RESOURCE_ROOT = resource_root()
DEFAULT_CONFIG = RESOURCE_ROOT / "02_CONFIGURACAO" / "mapeamentos" / "mapeamento_template_oficial.json"
DEFAULT_OUTPUT = ROOT / "04_SAIDA"
DEFAULT_DISCOVERY = ROOT / "06_DOCUMENTACAO" / "descoberta_planilhas.md"
EXCEL_TYPES = (("Planilhas Excel", "*.xlsx *.xlsm *.xls"), ("Todos os arquivos", "*.*"))
APP_TITLE = "IAM Brasil | GovernanÃ§a SoD & SAT"
BRAND_BLUE = "#2448A6"
BRAND_TEAL = "#00AFC1"
INK = "#17233A"
MUTED = "#52647E"
BACKGROUND = "#F3F6FB"
SURFACE = "#FFFFFF"
ICON_PATH = RESOURCE_ROOT / "02_CONFIGURACAO" / "recursos" / "iam_brasil_access_governance.ico"
LOGO_PATH = RESOURCE_ROOT / "02_CONFIGURACAO" / "recursos" / "iam_brasil_access_governance.png"


class SodSatApplication(ttk.Frame):
    """Interface local. As fontes sÃ£o apenas lidas e cada execuÃ§Ã£o guarda cÃ³pias RAW."""

    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=20)
        self.master = master
        self.funcionalidades = tk.StringVar()
        self.usuarios = tk.StringVar()
        self.conflitos = tk.StringVar()
        self.config = tk.StringVar(value=str(DEFAULT_CONFIG))
        self.status = tk.StringVar(value="Selecione a base de funcionalidades e a base de usuÃ¡rios; elas podem conter um ou mais sistemas.")
        self.last_result_file: Path | None = None
        self.last_result_dir: Path | None = None
        self.last_diagnostics_file: Path | None = None
        self.last_email_draft_file: Path | None = None
        self.last_sod_analysis_file: Path | None = None
        self._busy = False
        self._header_icon: tk.PhotoImage | None = None
        self._configure_style()
        self.grid(sticky="nsew")
        self._build()

    def _configure_style(self) -> None:
        self.master.configure(background=BACKGROUND)
        style = ttk.Style(self.master)
        style.theme_use("clam")
        style.configure("App.TFrame", background=BACKGROUND)
        style.configure("Card.TLabelframe", background=SURFACE, borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", background=SURFACE, foreground=BRAND_BLUE, font=("Aptos", 10, "bold"))
        style.configure("Field.TLabel", background=SURFACE, foreground=INK, font=("Aptos", 10))
        style.configure("Hint.TLabel", background=SURFACE, foreground=MUTED, font=("Aptos", 9))
        style.configure("Status.TLabel", background=BACKGROUND, foreground=BRAND_BLUE, font=("Aptos", 10, "bold"))
        style.configure("Primary.TButton", font=("Aptos", 10, "bold"), foreground="white", background=BRAND_BLUE, padding=(14, 8))
        style.map("Primary.TButton", background=[("active", "#173880"), ("disabled", "#AAB7D5")], foreground=[("disabled", "#F8FAFF")])
        style.configure("Secondary.TButton", font=("Aptos", 10), foreground=BRAND_BLUE, background="#E7EEFF", padding=(12, 8))
        style.map("Secondary.TButton", background=[("active", "#D7E3FF"), ("disabled", "#EFF2F7")])
        style.configure("Action.TButton", font=("Aptos", 9), padding=(10, 6))

    def _build(self) -> None:
        self.master.title(APP_TITLE)
        if ICON_PATH.is_file():
            self.master.iconbitmap(default=str(ICON_PATH))
        self.master.minsize(980, 470)
        self.master.geometry("1080x500")
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.configure(style="App.TFrame")

        hero = tk.Frame(self, background=BRAND_BLUE, padx=24, pady=18)
        hero.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        hero.columnconfigure(1, weight=1)
        if LOGO_PATH.is_file():
            self._header_icon = tk.PhotoImage(file=str(LOGO_PATH)).subsample(15, 15)
            tk.Label(hero, image=self._header_icon, background=BRAND_BLUE).grid(row=0, column=0, rowspan=2, padx=(0, 16), sticky="w")
        tk.Label(hero, text="Governança de Acessos", background=BRAND_BLUE, foreground="white", font=("Aptos Display", 22, "bold")).grid(row=0, column=1, sticky="sw")
        tk.Label(hero, text="GovernanÃ§a SoD & SAT", background=BRAND_BLUE, foreground="#DFF8FF", font=("Aptos", 12, "bold")).grid(row=1, column=1, sticky="nw", pady=(2, 0))
        tk.Label(hero, text="Matriz funcional, anÃ¡lise de conflitos e classificaÃ§Ã£o SAT com rastreabilidade.", background=BRAND_BLUE, foreground="#DCE8FF", font=("Aptos", 10), wraplength=330, justify="right").grid(row=0, column=2, rowspan=2, sticky="e")

        sources = ttk.LabelFrame(self, text="  1. Fontes de entrada  ", padding=16, style="Card.TLabelframe")
        sources.grid(row=1, column=0, sticky="ew")
        sources.columnconfigure(1, weight=1)
        self._file_row(sources, 0, "Base de funcionalidades *", self.funcionalidades, self._choose_functionalities)
        self._file_row(sources, 1, "Base de usuÃ¡rios *", self.usuarios, self._choose_users)
        self._file_row(sources, 2, "Regras SoD (opcional)", self.conflitos, self._choose_conflicts)
        ttk.Label(sources, text="* As planilhas podem conter diversos sistemas. O sistema realiza o de-para por nome, configuraÃ§Ã£o interna ou perfil em comum e registra as pendÃªncias para tratativa.", style="Hint.TLabel", wraplength=910).grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))

        actions = ttk.LabelFrame(self, text="  2. Processar e consultar resultados  ", padding=12, style="Card.TLabelframe")
        actions.grid(row=2, column=0, sticky="ew", pady=(16, 12))
        self.discovery_button = ttk.Button(actions, text="Conferir fontes", command=self._discover, style="Secondary.TButton")
        self.discovery_button.grid(row=0, column=0, padx=(0, 8))
        self.process_button = ttk.Button(actions, text="Gerar anÃ¡lise no modelo oficial", command=self._process, style="Primary.TButton")
        self.process_button.grid(row=0, column=1, padx=(0, 8))
        self.open_button = ttk.Button(actions, text="Abrir Excel gerado", command=self._open_result, state="disabled", style="Action.TButton")
        self.open_button.grid(row=0, column=2, padx=(0, 8))
        self.download_button = ttk.Button(actions, text="Salvar cÃ³pia...", command=self._download_result, state="disabled", style="Action.TButton")
        self.download_button.grid(row=0, column=3)
        self.diagnostics_button = ttk.Button(actions, text="DiagnÃ³stico", command=self._open_diagnostics, state="disabled", style="Action.TButton")
        self.diagnostics_button.grid(row=0, column=4, padx=(8, 0))
        self.email_button = ttk.Button(actions, text="Minuta de e-mail", command=self._open_email_draft, state="disabled", style="Action.TButton")
        self.email_button.grid(row=0, column=5, padx=(8, 0))
        self.sod_analysis_button = ttk.Button(actions, text="AnÃ¡lise SoD", command=self._open_sod_analysis, state="disabled", style="Action.TButton")
        self.sod_analysis_button.grid(row=0, column=6, padx=(8, 0))
        ttk.Label(self, textvariable=self.status, style="Status.TLabel", wraplength=980).grid(row=3, column=0, sticky="w", pady=(0, 8))

    def _file_row(self, parent: ttk.LabelFrame, row: int, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(parent, textvariable=variable, state="readonly", width=82).grid(row=row, column=1, sticky="ew", pady=6, ipady=2)
        ttk.Button(parent, text="Selecionar", command=command, style="Secondary.TButton").grid(row=row, column=2, padx=(10, 0), pady=6)

    def _choose_functionalities(self) -> None:
        self._select_file(self.funcionalidades, "Selecionar base de funcionalidades/transaÃ§Ãµes")

    def _choose_users(self) -> None:
        self._select_file(self.usuarios, "Selecionar base de usuÃ¡rios e perfis")

    def _choose_conflicts(self) -> None:
        self._select_file(self.conflitos, "Selecionar regras SoD explÃ­citas (opcional)")

    def _choose_config(self) -> None:
        selected = filedialog.askopenfilename(title="Selecionar configuraÃ§Ã£o JSON", initialdir=str(DEFAULT_CONFIG.parent), filetypes=(("Arquivo JSON", "*.json"),))
        if selected:
            self.config.set(selected)

    def _select_file(self, variable: tk.StringVar, title: str) -> None:
        selected = filedialog.askopenfilename(title=title, filetypes=EXCEL_TYPES)
        if selected:
            variable.set(selected)

    def _validated_paths(self) -> tuple[Path, Path, Path | None, Path] | None:
        required = (("Base de funcionalidades", self.funcionalidades.get()), ("Base de usuÃ¡rios", self.usuarios.get()), ("ConfiguraÃ§Ã£o", self.config.get()))
        invalid = [name for name, value in required if not value or not Path(value).is_file()]
        if invalid:
            messagebox.showwarning("Dados necessÃ¡rios", "Selecione arquivos vÃ¡lidos para: " + ", ".join(invalid) + ".")
            return None
        functionalities = Path(self.funcionalidades.get()).resolve()
        users = Path(self.usuarios.get()).resolve()
        rules = Path(self.conflitos.get()).resolve() if self.conflitos.get() else None
        config = Path(self.config.get()).resolve()
        if functionalities == users:
            messagebox.showwarning("Fontes iguais", "Selecione arquivos distintos para funcionalidades e usuÃ¡rios.")
            return None
        if rules and rules in {functionalities, users}:
            messagebox.showwarning("Fonte repetida", "A planilha de regras SoD deve ser distinta das duas bases obrigatÃ³rias.")
            return None
        return functionalities, users, rules, config

    def _discover(self) -> None:
        paths = self._validated_paths()
        if not paths:
            return
        functionalities, users, rules, config = paths

        def work():
            from .discovery import discover_workbooks, write_discovery_report
            report = discover_workbooks(functionalities, users, rules, config)
            write_discovery_report(report, DEFAULT_DISCOVERY)
            return f"ConferÃªncia concluÃ­da. RelatÃ³rio: {DEFAULT_DISCOVERY}", None

        self._run("Conferindo estrutura e mapeamento das fontes...", work)

    def _process(self) -> None:
        paths = self._validated_paths()
        if not paths:
            return
        functionalities, users, rules, config = paths
        confirm = messagebox.askokcancel(
            "Gerar resultado",
            "O sistema confrontarÃ¡ os perfis por sistema nos dois sentidos. Havendo inconsistÃªncia, o template ainda serÃ¡ gerado e a execuÃ§Ã£o incluirÃ¡ diagnÃ³stico, de-para e minuta de e-mail para tratativa. Nenhuma fonte ou template serÃ¡ sobrescrito. Continuar?",
        )
        if not confirm:
            return

        def work():
            from .processor import process_workbooks
            result = process_workbooks(functionalities, users, rules, config, DEFAULT_OUTPUT)
            email_message = f"Minuta de e-mail: {result.email_draft_path}\n" if result.email_required else "Minuta de e-mail: nÃ£o habilitada, pois nÃ£o foram comprovadas divergÃªncias de perfis em uma conciliaÃ§Ã£o vÃ¡lida.\n"
            message = (f"Status: {result.status}\n"
                       f"Resultado: {result.workbook_path or 'nÃ£o gerado â€” conciliaÃ§Ã£o bloqueada'}\n"
                       f"Teste de consistÃªncia: {result.validation_path}\n"
                       f"DiagnÃ³stico: {result.diagnostics_path}\n"
                       f"{email_message}"
                       f"AnÃ¡lise SoD detalhada: {result.sod_analysis_path or 'nÃ£o gerada'}\n"
                       f"Log: {result.log_path}\n"
                       f"PendÃªncias: {result.pending_path}")
            return message, result

        self._run("Validando as bases e gerando uma nova execuÃ§Ã£o...", work)

    def _run(self, started_message: str, work) -> None:
        if self._busy:
            return
        self._busy = True
        self.status.set(started_message)
        self.discovery_button.configure(state="disabled")
        self.process_button.configure(state="disabled")
        self._set_result_actions({"open": "disabled", "download": "disabled", "diagnostics": "disabled", "email": "disabled", "sod": "disabled"})

        def worker() -> None:
            try:
                message, payload = work()
                self.master.after(0, lambda: self._complete(message, payload))
            except Exception as exc:
                details = "".join(traceback.format_exception_only(type(exc), exc)).strip()
                self.master.after(0, lambda: self._fail(details))

        threading.Thread(target=worker, daemon=True).start()

    def _complete(self, message: str, payload) -> None:
        self._busy = False
        self.discovery_button.configure(state="normal")
        self.process_button.configure(state="normal")
        if payload is not None and hasattr(payload, "run_id"):
            self.last_result_dir = payload.log_path.parent
            self.last_result_file = payload.workbook_path
            self.last_diagnostics_file = payload.diagnostics_path
            self.last_email_draft_file = payload.email_draft_path
            self.last_sod_analysis_file = payload.sod_analysis_path
            states = self._result_action_states(payload)
            self._set_result_actions(states)
            if payload.status == "BLOQUEADO_SEM_CONCILIACAO":
                self.status.set("Erro de conciliaÃ§Ã£o: nÃ£o hÃ¡ sistema comum mapeado entre as bases. Nenhuma aÃ§Ã£o foi liberada.")
            elif payload.status == "BLOQUEADO_POR_DIVERGENCIA_DE_PERFIS":
                self.status.set("Matriz Funcional nÃ£o liberada: regularize as divergÃªncias de perfil e consulte o diagnÃ³stico, a minuta e a anÃ¡lise SoD.")
            else:
                self.status.set("ConcluÃ­do. Resultado liberado apÃ³s a conciliaÃ§Ã£o das bases.")
        else:
            self._set_result_actions({"open": "disabled", "download": "disabled", "diagnostics": "disabled", "email": "disabled", "sod": "disabled"})
            self.status.set("ConcluÃ­do.")
        self._append(message)
        if getattr(payload, "status", None) == "BLOQUEADO_SEM_CONCILIACAO":
            messagebox.showerror("Bases nÃ£o conciliadas", message)
        elif getattr(payload, "status", None) == "BLOQUEADO_POR_DIVERGENCIA_DE_PERFIS":
            messagebox.showwarning("Matriz Funcional nÃ£o liberada", message)
        else:
            messagebox.showinfo(APP_TITLE, message)

    def _fail(self, message: str) -> None:
        self._busy = False
        self.discovery_button.configure(state="normal")
        self.process_button.configure(state="normal")
        self.status.set("A operaÃ§Ã£o nÃ£o foi concluÃ­da; consulte a mensagem abaixo.")
        self._append("ERRO: " + message)
        messagebox.showerror("Processamento nÃ£o concluÃ­do", message)

    def _open_result(self) -> None:
        if self.last_result_file and self.last_result_file.is_file():
            os.startfile(self.last_result_file)  # noqa: S606 - botÃ£o local acionado pelo usuÃ¡rio.

    def _download_result(self) -> None:
        if not self.last_result_file or not self.last_result_file.is_file():
            return
        destination = filedialog.asksaveasfilename(
            title="Salvar cÃ³pia do resultado gerado", initialfile=self.last_result_file.name,
            defaultextension=".xlsx", filetypes=(("Planilha Excel", "*.xlsx"),),
        )
        if not destination:
            return
        target = Path(destination).resolve()
        if target == self.last_result_file.resolve():
            messagebox.showwarning("Destino invÃ¡lido", "Escolha uma pasta diferente da execuÃ§Ã£o original.")
            return
        shutil.copy2(self.last_result_file, target)
        if self.last_result_dir:
            record = self.last_result_dir / "registro_entrega.json"
            record.write_text(json.dumps({"copiado_em_utc": datetime.now(timezone.utc).isoformat(), "arquivo_origem": str(self.last_result_file), "copia_entregue": str(target)}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        self._append(f"CÃ³pia do resultado salva em: {target}")
        messagebox.showinfo("Resultado salvo", f"CÃ³pia salva em:\n{target}")

    def _open_diagnostics(self) -> None:
        if self.last_diagnostics_file and self.last_diagnostics_file.is_file():
            os.startfile(self.last_diagnostics_file)  # noqa: S606 - botÃ£o local acionado pelo usuÃ¡rio.

    def _open_email_draft(self) -> None:
        if self.last_email_draft_file and self.last_email_draft_file.is_file():
            os.startfile(self.last_email_draft_file)  # noqa: S606 - botÃ£o local acionado pelo usuÃ¡rio.

    def _open_sod_analysis(self) -> None:
        if self.last_sod_analysis_file and self.last_sod_analysis_file.is_file():
            os.startfile(self.last_sod_analysis_file)  # noqa: S606 - botÃ£o local acionado pelo usuÃ¡rio.

    def _append(self, message: str) -> None:
        # O histÃ³rico visual foi removido para manter a interface objetiva.
        # Os logs completos continuam sendo salvos dentro de cada execuÃ§Ã£o.
        return None

    @staticmethod
    def _result_action_states(result: "ProcessingResult") -> dict[str, str]:
        if getattr(result, "status", None) == "BLOQUEADO_SEM_CONCILIACAO":
            return {"open": "disabled", "download": "disabled", "diagnostics": "disabled", "email": "disabled", "sod": "disabled"}
        has_workbook = bool(result.workbook_path and result.workbook_path.is_file())
        has_sod = bool(result.sod_analysis_path and result.sod_analysis_path.is_file())
        return {
            "open": "normal" if has_workbook else "disabled",
            "download": "normal" if has_workbook else "disabled",
            "diagnostics": "normal" if result.diagnostics_path.is_file() else "disabled",
            "email": "normal" if result.email_required and result.email_draft_path.is_file() else "disabled",
            "sod": "normal" if has_sod else "disabled",
        }

    def _set_result_actions(self, states: dict[str, str]) -> None:
        self.open_button.configure(state=states["open"])
        self.download_button.configure(state=states["download"])
        self.diagnostics_button.configure(state=states["diagnostics"])
        self.email_button.configure(state=states["email"])
        self.sod_analysis_button.configure(state=states["sod"])


def start() -> None:
    root = tk.Tk()
    SodSatApplication(root)
    root.mainloop()


