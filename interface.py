import contextlib
import io
import os
import re
import shutil
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd

import paths
from preencher import (
    build_mes_variants,
    build_codigo_mes,
    gerar_ppt,
    load_cloud_incidents,
    load_excel_reports,
    load_throttling_logs,
    listar_templates_ppt,
    normalizar,
    sanitize_filename,
    split_codigo_mes,
)


class QueueWriter(io.TextIOBase):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def write(self, text):
        if text:
            self.callback(text.rstrip())
        return len(text)

    def flush(self):
        return None


class GeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de relatórios PPTX")
        self.root.geometry("1000x820")
        self.root.minsize(860, 700)

        self.mode = tk.StringVar(value="single")
        self.year = tk.StringVar(value="2026")
        self.month = tk.StringVar(value="09")
        self.client = tk.StringVar()
        self.include_mex = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Pronto para gerar")
        self.selected_logs = []
        self.updated_inputs = []
        self.result_paths = []
        self.preview_status = tk.StringVar(value="Selecione um período e atualize o preview.")

        self._configure_style()
        self._build_ui()
        self._update_mode()

    def _configure_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 22, "bold"), foreground="#17324d")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#5f7182")
        style.configure("Section.TLabel", font=("Segoe UI", 11, "bold"), foreground="#17324d")
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 9))
        style.configure("TNotebook", tabmargins=(0, 0, 0, 0))

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=28)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=0)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(2, weight=1)

        ttk.Label(outer, text="Gerador de relatórios PPTX", style="Title.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            outer,
            text="Selecione o período, adicione os logs e acompanhe os arquivos gerados.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 22))

        controls_shell = ttk.Frame(outer)
        controls_shell.grid(row=2, column=0, sticky="nsew", padx=(0, 20))
        controls_shell.rowconfigure(0, weight=1)
        controls_shell.columnconfigure(0, weight=1)
        controls_canvas = tk.Canvas(controls_shell, highlightthickness=0, borderwidth=0)
        controls_canvas.grid(row=0, column=0, sticky="nsew")
        controls_scrollbar = ttk.Scrollbar(
            controls_shell, orient="vertical", command=controls_canvas.yview
        )
        controls_scrollbar.grid(row=0, column=1, sticky="ns")
        controls_canvas.configure(yscrollcommand=controls_scrollbar.set)
        controls = ttk.Frame(controls_canvas, padding=20, relief="groove")
        controls_window = controls_canvas.create_window((0, 0), window=controls, anchor="nw")

        def update_controls_scroll_region(_event=None):
            controls_canvas.configure(scrollregion=controls_canvas.bbox("all"))

        def resize_controls_width(event):
            controls_canvas.itemconfigure(controls_window, width=event.width)

        controls.bind("<Configure>", update_controls_scroll_region)
        controls_canvas.bind("<Configure>", resize_controls_width)

        def scroll_controls(event):
            controls_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        controls_canvas.bind("<Enter>", lambda _event: controls_canvas.bind_all("<MouseWheel>", scroll_controls))
        controls_canvas.bind("<Leave>", lambda _event: controls_canvas.unbind_all("<MouseWheel>"))

        ttk.Label(controls, text="1. O que gerar", style="Section.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12)
        )
        ttk.Radiobutton(
            controls,
            text="Um período",
            variable=self.mode,
            value="single",
            command=self._update_mode,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=3)
        ttk.Radiobutton(
            controls,
            text="Todos os períodos disponíveis",
            variable=self.mode,
            value="all",
            command=self._update_mode,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=3)
        ttk.Radiobutton(
            controls,
            text="Cliente específico",
            variable=self.mode,
            value="client",
            command=self._update_mode,
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=3)

        ttk.Separator(controls).grid(row=4, column=0, columnspan=2, sticky="ew", pady=18)
        ttk.Label(controls, text="Período", style="Section.TLabel").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )
        ttk.Label(controls, text="Ano").grid(row=6, column=0, sticky="w", pady=4)
        self.year_box = ttk.Combobox(
            controls, textvariable=self.year, values=["2025", "2026"], state="readonly", width=12
        )
        self.year_box.grid(row=6, column=1, sticky="ew", pady=4)
        ttk.Label(controls, text="Mês").grid(row=7, column=0, sticky="w", pady=4)
        self.month_box = ttk.Combobox(
            controls,
            textvariable=self.month,
            values=[f"{month:02d}" for month in range(1, 13)],
            state="readonly",
            width=12,
        )
        self.month_box.grid(row=7, column=1, sticky="ew", pady=4)
        ttk.Label(controls, text="Cliente").grid(row=8, column=0, sticky="w", pady=4)
        self.client_box = ttk.Combobox(
            controls, textvariable=self.client, state="disabled", width=28
        )
        self.client_box.grid(row=8, column=1, sticky="ew", pady=4)
        self.year_box.bind("<<ComboboxSelected>>", lambda _event: self._load_client_options())
        self.month_box.bind("<<ComboboxSelected>>", lambda _event: self._load_client_options())
        self.mex_check = ttk.Checkbutton(
            controls,
            text="Gerar também versão México (-MEX)",
            variable=self.include_mex,
        )
        self.mex_check.grid(row=9, column=0, columnspan=2, sticky="w", pady=(10, 0))

        ttk.Separator(controls).grid(row=10, column=0, columnspan=2, sticky="ew", pady=18)
        ttk.Label(controls, text="2. Logs de throttling", style="Section.TLabel").grid(
            row=11, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )
        ttk.Button(controls, text="Adicionar arquivos CSV", command=self._add_logs).grid(
            row=12, column=0, columnspan=2, sticky="ew", pady=(0, 8)
        )
        self.logs_label = ttk.Label(controls, text="Nenhum arquivo novo selecionado", wraplength=235)
        self.logs_label.grid(row=13, column=0, columnspan=2, sticky="w")

        ttk.Separator(controls).grid(row=14, column=0, columnspan=2, sticky="ew", pady=18)
        ttk.Label(controls, text="3. Atualizar dados de entrada", style="Section.TLabel").grid(
            row=15, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )
        ttk.Button(
            controls,
            text="Atualizar Storage",
            command=lambda: self._add_input_files("storage", "Selecione o CSV de Storage", False),
        ).grid(row=16, column=0, columnspan=2, sticky="ew", pady=3)
        ttk.Button(
            controls,
            text="Atualizar Cloud Incidents",
            command=lambda: self._add_input_files("cloud", "Selecione os CSVs de Cloud Incidents", True),
        ).grid(row=17, column=0, columnspan=2, sticky="ew", pady=3)
        self.inputs_label = ttk.Label(controls, text="Nenhum arquivo de entrada atualizado", wraplength=235)
        self.inputs_label.grid(row=18, column=0, columnspan=2, sticky="w", pady=(4, 0))

        self.generate_button = ttk.Button(
            controls,
            text="Gerar apresentações",
            style="Primary.TButton",
            command=self._start_generation,
        )
        self.generate_button.grid(row=19, column=0, columnspan=2, sticky="ew", pady=(28, 0))
        ttk.Button(
            controls,
            text="Abrir manual de instruções",
            command=self._open_manual,
        ).grid(row=20, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        controls.columnconfigure(1, weight=1)

        results = ttk.Frame(outer)
        results.grid(row=2, column=1, sticky="nsew")
        results.rowconfigure(2, weight=1)
        results.columnconfigure(0, weight=1)
        ttk.Label(results, textvariable=self.status).grid(row=1, column=0, sticky="w", pady=(4, 10))

        self.results_notebook = ttk.Notebook(results)
        self.results_notebook.grid(row=2, column=0, sticky="nsew")

        files_tab = ttk.Frame(self.results_notebook, padding=12)
        files_tab.rowconfigure(1, weight=1)
        files_tab.columnconfigure(0, weight=1)
        self.results_notebook.add(files_tab, text="Arquivos gerados")
        ttk.Label(files_tab, text="Arquivos gerados", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 10)
        )

        list_frame = ttk.Frame(files_tab)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)
        self.output_list = tk.Listbox(
            list_frame,
            font=("Consolas", 10),
            activestyle="dotbox",
            selectmode=tk.SINGLE,
            background="#f6f9fb",
            foreground="#17324d",
            borderwidth=1,
            relief="solid",
        )
        self.output_list.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.output_list.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.output_list.configure(yscrollcommand=scrollbar.set)
        self.output_list.bind("<Double-Button-1>", self._open_selected_result)

        ttk.Label(
            files_tab,
            text="Clique duas vezes em um arquivo para abri-lo no Explorer.",
            style="Subtitle.TLabel",
        ).grid(row=2, column=0, sticky="w", pady=(10, 0))

        logs_tab = ttk.Frame(self.results_notebook, padding=12)
        logs_tab.rowconfigure(1, weight=1)
        logs_tab.columnconfigure(0, weight=1)
        self.results_notebook.add(logs_tab, text="Logs da execução")
        ttk.Label(logs_tab, text="Acompanhamento da execução", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 10)
        )
        log_frame = ttk.Frame(logs_tab)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_frame,
            state="disabled",
            wrap="word",
            font=("Consolas", 9),
            background="#f6f9fb",
            foreground="#17324d",
            borderwidth=1,
            relief="solid",
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        preview_tab = ttk.Frame(self.results_notebook, padding=12)
        preview_tab.rowconfigure(2, weight=1)
        preview_tab.columnconfigure(0, weight=1)
        self.results_notebook.add(preview_tab, text="Preview dos dados")
        preview_toolbar = ttk.Frame(preview_tab)
        preview_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        preview_toolbar.columnconfigure(0, weight=1)
        ttk.Label(preview_toolbar, textvariable=self.preview_status).grid(
            row=0, column=0, sticky="w"
        )
        self.preview_button = ttk.Button(
            preview_toolbar, text="Atualizar preview", command=self._refresh_preview
        )
        self.preview_button.grid(row=0, column=1, sticky="e")
        ttk.Label(
            preview_tab,
            text="Os dados exibidos seguem as mesmas fontes usadas no preenchimento do PPTX.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 8))

        preview_frame = ttk.Frame(preview_tab)
        preview_frame.grid(row=2, column=0, sticky="nsew")
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_columns = (
            "cliente",
            "usuarios",
            "s3",
            "database",
            "incidentes",
            "indisponibilidade",
            "logs",
            "status",
        )
        self.preview_tree = ttk.Treeview(
            preview_frame, columns=preview_columns, show="headings", selectmode="browse"
        )
        headings = {
            "cliente": ("Cliente", 210),
            "usuarios": ("Usuários", 75),
            "s3": ("S3", 90),
            "database": ("Database", 100),
            "incidentes": ("Incidentes", 85),
            "indisponibilidade": ("Indisp. (min)", 100),
            "logs": ("Logs QTD", 80),
            "status": ("Status", 170),
        }
        for column in preview_columns:
            heading, width = headings[column]
            self.preview_tree.heading(column, text=heading)
            self.preview_tree.column(column, width=width, anchor="w")
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        preview_y_scroll = ttk.Scrollbar(
            preview_frame, orient="vertical", command=self.preview_tree.yview
        )
        preview_y_scroll.grid(row=0, column=1, sticky="ns")
        preview_x_scroll = ttk.Scrollbar(
            preview_frame, orient="horizontal", command=self.preview_tree.xview
        )
        preview_x_scroll.grid(row=1, column=0, sticky="ew")
        self.preview_tree.configure(
            yscrollcommand=preview_y_scroll.set, xscrollcommand=preview_x_scroll.set
        )

        manual_tab = ttk.Frame(self.results_notebook, padding=12)
        manual_tab.rowconfigure(1, weight=1)
        manual_tab.columnconfigure(0, weight=1)
        self.results_notebook.add(manual_tab, text="Manual de instruções")
        ttk.Label(manual_tab, text="Manual de instruções", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 10)
        )
        manual_frame = ttk.Frame(manual_tab)
        manual_frame.grid(row=1, column=0, sticky="nsew")
        manual_frame.rowconfigure(0, weight=1)
        manual_frame.columnconfigure(0, weight=1)
        self.manual_text = tk.Text(
            manual_frame,
            wrap="word",
            state="disabled",
            font=("Segoe UI", 10),
            background="#f6f9fb",
            foreground="#17324d",
            padx=14,
            pady=12,
            borderwidth=1,
            relief="solid",
        )
        self.manual_text.grid(row=0, column=0, sticky="nsew")
        manual_scrollbar = ttk.Scrollbar(
            manual_frame, orient="vertical", command=self.manual_text.yview
        )
        manual_scrollbar.grid(row=0, column=1, sticky="ns")
        self.manual_text.configure(yscrollcommand=manual_scrollbar.set)
        self._load_manual_content()

    def _update_mode(self):
        is_period_mode = self.mode.get() in ("single", "client")
        is_client_mode = self.mode.get() == "client"
        state = "readonly" if is_period_mode else "disabled"
        self.year_box.configure(state=state)
        self.month_box.configure(state=state)
        self.mex_check.configure(state="normal" if is_period_mode else "disabled")
        self.client_box.configure(state="readonly" if is_client_mode else "disabled")
        if is_client_mode:
            self._load_client_options()

    def _add_logs(self):
        files = filedialog.askopenfilenames(
            title="Selecione os logs de throttling",
            initialdir=str(paths.LOGS_DIR),
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not files:
            return
        paths.ensure_directories()
        copied = []
        for file_name in files:
            source = Path(file_name)
            target = paths.LOGS_DIR / source.name
            if source.resolve() != target.resolve():
                shutil.copy2(source, target)
            copied.append(source.name)
        self.selected_logs.extend(copied)
        self.logs_label.configure(text=f"{len(self.selected_logs)} arquivo(s) selecionado(s) e copiado(s).")
        self._write_log("Logs adicionados: " + ", ".join(copied))

    def _load_client_options(self):
        if self.mode.get() != "client":
            return
        try:
            reports = load_excel_reports()
            df = reports.get(self.year.get())
            if df is None or df.empty:
                clients = []
            else:
                target_date = pd.Timestamp(
                    year=int(self.year.get()), month=int(self.month.get()), day=1
                ).date()
                clients = sorted(df.loc[df["Mês"] == target_date, "Cliente"].dropna().unique())
            self.client_box.configure(values=clients)
            if self.client.get() not in clients:
                self.client.set(clients[0] if clients else "")
        except Exception as error:
            self.client_box.configure(values=[])
            self.client.set("")
            self._write_log(f"Não foi possível carregar clientes: {error}")

    def _add_input_files(self, category, title, multiple):
        ask_files = filedialog.askopenfilenames if multiple else filedialog.askopenfilename
        selected = ask_files(
            title=title,
            initialdir=str(paths.INPUTS_DIR),
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not selected:
            return
        if isinstance(selected, str):
            selected = (selected,)

        invalid = [Path(file_name).name for file_name in selected if not self._matches_input_category(file_name, category)]
        if invalid:
            messagebox.showerror(
                "Arquivo incompatível",
                "O nome do arquivo não corresponde ao tipo selecionado:\n\n" + "\n".join(invalid),
            )
            return

        paths.ensure_directories()
        copied = []
        for file_name in selected:
            source = Path(file_name)
            target, obsolete_files = self._input_destination(source, category)
            for obsolete_file in obsolete_files:
                if obsolete_file != target and obsolete_file.exists():
                    obsolete_file.unlink()
            if source.resolve() != target.resolve():
                shutil.copy2(source, target)
            copied.append(source.name)
        self.updated_inputs.extend(copied)
        self.inputs_label.configure(
            text=f"{len(self.updated_inputs)} arquivo(s) de entrada atualizado(s)."
        )
        self._write_log("Entradas atualizadas: " + ", ".join(copied))

    @staticmethod
    def _matches_input_category(file_name, category):
        name = Path(file_name).name.lower()
        if not name.endswith(".csv"):
            return False
        if category == "storage":
            return "storage" in name
        return "cloud" in name and "incident" in name

    @staticmethod
    def _input_destination(source, category):
        year_match = re.search(r"(?:19|20)\d{2}", source.name)
        if not year_match:
            raise ValueError(f"Não foi possível identificar o ano em {source.name}.")
        year = year_match.group(0)
        if category == "storage":
            target = paths.INPUTS_DIR / f"Cloud - Storage {year}.csv"
            obsolete = [
                paths.INPUTS_DIR / f"Cloud Storage {year}.csv",
                paths.INPUTS_DIR / f"Storage {year}.csv",
            ]
        else:
            target = paths.INPUTS_DIR / f"Cloud Incidents(Registro {year}).csv"
            obsolete = list(paths.INPUTS_DIR.glob(f"Cloud Incidents(Registro*{year}*.csv"))
        return target, obsolete

    def _start_generation(self):
        if self.mode.get() in ("single", "client") and not self.year.get().isdigit():
            messagebox.showerror("Período inválido", "Selecione um ano e um mês válidos.")
            return
        if self.mode.get() == "client" and not self.client.get():
            messagebox.showerror("Cliente não selecionado", "Escolha um cliente antes de gerar.")
            return
        self.generate_button.configure(state="disabled")
        self.output_list.delete(0, tk.END)
        self.result_paths = []
        self.status.set("Gerando apresentações...")
        self._clear_execution_log()
        threading.Thread(target=self._generate, daemon=True).start()

    def _generate(self):
        writer = QueueWriter(self._write_log)
        try:
            paths.ensure_directories()
            if self.mode.get() == "all":
                template_codes = listar_templates_ppt(paths.PPTX_DIR)
                base_codes = sorted({split_codigo_mes(code)[0] for code in template_codes})
                if not base_codes:
                    raise FileNotFoundError("Nenhum template .pptx ou .odp foi encontrado em data/pptx.")
                codes = []
                for base_code in base_codes:
                    codes.extend(self._variants_for(base_code))
            else:
                code = build_codigo_mes(self.year.get(), self.month.get())
                codes = self._variants_for(code)

            generated = []
            target_client = self.client.get() if self.mode.get() == "client" else None
            with contextlib.redirect_stdout(writer):
                for code in codes:
                    final_ppt, final_dir, clients = gerar_ppt(code, target_client=target_client)
                    generated.append(final_ppt)
                    for client_name in clients:
                        generated.append(
                            final_dir / f"{code}_{sanitize_filename(client_name)}.pptx"
                        )
            self.root.after(0, lambda: self._generation_finished(generated))
        except Exception as error:
            self.root.after(0, lambda: self._generation_failed(error))

    def _variants_for(self, code):
        variants = build_mes_variants(code)
        return variants if self.include_mex.get() or self.mode.get() == "all" else variants[:1]

    def _generation_finished(self, generated):
        unique_paths = []
        seen = set()
        for path in generated:
            path = Path(path)
            if path.exists() and path not in seen:
                seen.add(path)
                unique_paths.append(path)
        self.result_paths = unique_paths
        for path in unique_paths:
            self.output_list.insert(tk.END, str(path))
        self.status.set(f"Concluído: {len(unique_paths)} arquivo(s) PPTX disponível(is).")
        self.generate_button.configure(state="normal")
        if not unique_paths:
            messagebox.showwarning("Sem arquivos", "A execução terminou, mas nenhum PPTX foi encontrado.")

    def _generation_failed(self, error):
        self.status.set("A execução terminou com erro.")
        self.generate_button.configure(state="normal")
        self._write_log(f"ERRO: {error}")
        messagebox.showerror("Erro na geração", str(error))

    def _write_log(self, text):
        if text:
            self.root.after(0, lambda: self._append_execution_log(text))

    def _clear_execution_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def _append_execution_log(self, text):
        self.status.set(text[-180:])
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _refresh_preview(self):
        if self.mode.get() != "single":
            self.preview_status.set("Preview disponível para um período por vez; selecione Um período.")
            return
        self.preview_status.set("Carregando dados do período...")
        try:
            year = int(self.year.get())
            month = int(self.month.get())
            reports = load_excel_reports()
            df_excel = reports.get(str(year))
            if df_excel is None or df_excel.empty:
                self._clear_preview("Nenhum relatório Excel disponível para este ano.")
                return

            target_date = pd.Timestamp(year=year, month=month, day=1).date()
            rows = df_excel[df_excel["Mês"] == target_date].copy()
            if rows.empty:
                self._clear_preview("Nenhum dado Excel encontrado para este período.")
                self.preview_status.set(f"Nenhum cliente encontrado para {year}-{month:02d}.")
                return
            df_incidents = load_cloud_incidents()
            df_logs = load_throttling_logs()
            mes_code = f"{year}-{month:02d}"
            self._clear_preview()

            for _, row in rows.sort_values("Cliente").iterrows():
                cliente = row["Cliente"]
                incidents = df_incidents[
                    (df_incidents["Customer_norm"] == cliente)
                    & (df_incidents["Mes_code"] == mes_code)
                ] if not df_incidents.empty else df_incidents
                minutes = 0
                if not incidents.empty:
                    minutes = pd.to_numeric(
                        incidents.loc[
                            incidents["Consider Unavailability?"].str.lower() == "yes",
                            "Total Time (Minutes)",
                        ],
                        errors="coerce",
                    ).fillna(0).sum()
                logs = df_logs[
                    (df_logs["cliente"] == cliente) & (df_logs["mes_code"] == mes_code)
                ] if not df_logs.empty else df_logs
                log_total = int(logs.iloc[0]["total_qtd"]) if not logs.empty else 0
                required_fields = [row.get("Usuários"), row.get("S3 (Gb)"), row.get("Database (Gb)")]
                ready = all(not pd.isna(value) and str(value).strip() for value in required_fields)
                status = "Pronto para preencher" if ready else "Sem dados Excel"
                self.preview_tree.insert(
                    "",
                    tk.END,
                    values=(
                        cliente,
                        self._display_value(row.get("Usuários")),
                        self._display_value(row.get("S3 (Gb)")),
                        self._display_value(row.get("Database (Gb)")),
                        len(incidents),
                        int(minutes),
                        log_total,
                        status,
                    ),
                )
            self.preview_status.set(
                f"{len(rows)} cliente(s) encontrados para {year}-{month:02d}."
            )
        except Exception as error:
            self._clear_preview(f"Erro ao carregar preview: {error}")
            self.preview_status.set("Não foi possível carregar o preview.")

    def _clear_preview(self, message=None):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        if message:
            self.preview_tree.insert("", tk.END, values=(message, "", "", "", "", "", "", ""))

    @staticmethod
    def _display_value(value):
        if value is None or pd.isna(value):
            return ""
        return str(value)

    def _open_selected_result(self, _event=None):
        selection = self.output_list.curselection()
        if not selection:
            return
        path = Path(self.output_list.get(selection[0]))
        if path.exists():
            os.startfile(path)

    def _open_manual(self):
        for tab_id in self.results_notebook.tabs():
            if self.results_notebook.tab(tab_id, "text") == "Manual de instruções":
                self.results_notebook.select(tab_id)
                return

    def _load_manual_content(self):
        manual_path = Path(__file__).resolve().parent / "MANUAL.md"
        try:
            content = manual_path.read_text(encoding="utf-8")
        except OSError as error:
            content = f"Não foi possível carregar o manual.\n\nDetalhes: {error}"
        self.manual_text.configure(state="normal")
        self.manual_text.insert("1.0", content)
        self.manual_text.configure(state="disabled")


def main():
    root = tk.Tk()
    GeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
