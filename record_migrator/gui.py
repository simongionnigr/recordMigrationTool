import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from simple_salesforce import Salesforce
import yaml

import pandas as pd

from config import load_config, save_config, CONFIG_FILE
from sf_client import SalesforceClient
from excel_utils import read_spreadsheet, write_spreadsheet

# Carica la configurazione YAML
CONFIG = load_config()

class MigrationToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Salesforce Data Migrator")
        self.geometry("900x600")

        self.sf_client = None
        self.query_df = None

        self._create_widgets()

    def _create_widgets(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(2, weight=1)

        # Fonte dati
        src_frame = ttk.LabelFrame(main, text="Fonte dati", padding=10)
        src_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10))
        self.source_mode = tk.StringVar(value="env")
        ttk.Radiobutton(src_frame, text="Ambiente di partenza", variable=self.source_mode,
                        value="env", command=self._toggle_source).pack(side="left", padx=5)
        ttk.Radiobutton(src_frame, text="File CSV/Excel", variable=self.source_mode,
                        value="file", command=self._toggle_source).pack(side="left", padx=5)

        # Ambiente di partenza
        self.env_frame = ttk.LabelFrame(main, text="Ambiente di partenza", padding=10)
        self.env_frame.grid(row=1, column=0, sticky="nsew", padx=(0,5))
        self._make_cred_fields(self.env_frame, "source")

        # Caricamento file generico
        self.file_frame = ttk.LabelFrame(main, text="Carica file dati", padding=10)
        self.file_type = tk.StringVar(value="CSV")
        ttk.Radiobutton(self.file_frame, text="CSV", variable=self.file_type, value="CSV").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(self.file_frame, text="Excel (.xlsx)", variable=self.file_type, value="XLSX").grid(row=0, column=1, padx=5)
        ttk.Button(self.file_frame, text="Seleziona file…", command=self._select_file).grid(row=0, column=2, padx=10)
        self.file_label = ttk.Label(self.file_frame, text="Nessun file selezionato")
        self.file_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(5,0))

        # Ambiente di destinazione
        self.target_frame = ttk.LabelFrame(main, text="Ambiente di destinazione", padding=10)
        self.target_frame.grid(row=1, column=1, sticky="nsew", padx=(5,0))
        self._make_cred_fields(self.target_frame, "target")

        # Query per SObject (span colonne)
        self.query_frame = ttk.LabelFrame(main, text="Query per SObject", padding=10)
        self.query_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10,0))
        self.query_frame.rowconfigure(1, weight=1)
        self.query_frame.columnconfigure(0, weight=1)

        # Bottone per caricamento query CSV
        ttk.Button(self.query_frame, text=CONFIG["query_csv"]["prompt"],
                   command=self._load_query_csv).grid(row=0, column=0, sticky="w")

        # Tabella scrollabile
        tf = ttk.Frame(self.query_frame)
        tf.grid(row=1, column=0, sticky="nsew", pady=(5,0))
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        cols = CONFIG["query_csv"]["columns"]
        self.query_table = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.query_table.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self.query_table.xview)
        self.query_table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.query_table.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        for col in cols:
            self.query_table.heading(col, text=col)
            self.query_table.column(col, width=200, anchor="w")

        # Bottone esecuzione
        self.run_button = ttk.Button(self.query_frame, text="Esegui Query e Esporta Excel", 
                                     command=self._start_run_queries, state="disabled")
        self.run_button.grid(row=2, column=0, sticky="ew", pady=(10,0))

        # Progress bar (inizialmente nascosta)
        self.progress = ttk.Progressbar(self, mode="indeterminate")

        # Menu contestuale per copia cella
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Copy", command=self._copy_cell)
        self.query_table.bind("<Button-3>", self._show_context_menu)

        # Abilitazione run_button al cambiamento credenziali
        for attr in ["source_username", "source_password", "source_security_token"]:
            ent = getattr(self, attr)
            ent.bind("<KeyRelease>", lambda e: self._update_run_button())

        self._toggle_source()

    def _make_cred_fields(self, frame, prefix):
        ttk.Label(frame, text="Tipo ambiente").grid(row=0, column=0, sticky="w")
        comb = ttk.Combobox(frame, values=["Sandbox", "Production"], state="readonly")
        comb.current(0); comb.grid(row=0, column=1, sticky="ew")
        if prefix == "source":
            self.source_env_type = comb
        labels = ["Username", "Password", "Security Token"]
        for i, label in enumerate(labels, start=1):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w")
            show = "*" if label != "Username" else None
            entry = ttk.Entry(frame, show=show) if show else ttk.Entry(frame)
            entry.grid(row=i, column=1, sticky="ew")
            setattr(self, f"{prefix}_{label.lower().replace(' ','_')}", entry)
        for i in range(4): frame.rowconfigure(i, pad=5)
        frame.columnconfigure(1, weight=1)

    def _toggle_source(self):
        if self.source_mode.get() == "env":
            self.env_frame.grid()
            self.query_frame.grid()
            self.file_frame.grid_remove()
        else:
            self.file_frame.grid(row=1, column=0, sticky="nsew", padx=(0,5))
            self.env_frame.grid_remove()
            self.query_frame.grid_remove()

    def _load_query_csv(self):
        path = filedialog.askopenfilename(filetypes=CONFIG["query_csv"]["filetypes"])
        if not path: return
        try:
            df = pd.read_csv(path, sep=CONFIG["query_csv"]["separator"], usecols=CONFIG["query_csv"]["columns"] )
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile leggere il file:\n{e}")
            return
        # Salva dataframe
        self.query_df = df
        # Popola tabella
        for row in self.query_table.get_children():
            self.query_table.delete(row)
        for _, r in df.iterrows():
            self.query_table.insert("", "end", values=[r[c] for c in CONFIG["query_csv"]["columns"]])
        self._update_run_button()

    def _copy_cell(self):
        col_idx = int(self._clicked_col.replace('#','')) - 1
        col_id = CONFIG["query_csv"]["columns"][col_idx]
        val = self.query_table.set(self._clicked_row, col_id)
        self.clipboard_clear(); self.clipboard_append(val)

    def _show_context_menu(self, event):
        region = self.query_table.identify_region(event.x, event.y)
        if region == "cell":
            self._clicked_row = self.query_table.identify_row(event.y)
            self._clicked_col = self.query_table.identify_column(event.x)
            self.menu.tk_popup(event.x_root, event.y_root)

    def _update_run_button(self):
        if self.source_mode.get() == "env" and self.query_df is not None:
            u = self.source_username.get().strip()
            p = self.source_password.get().strip()
            t = self.source_security_token.get().strip()
            if u and p and t:
                self.run_button.config(state="normal")
                return
        self.run_button.config(state="disabled")

    def _on_mode_change(self):
        if self.source_mode.get() == "env":
            self.env_frame.pack(fill="x", padx=10, pady=(5,0))
            self.file_frame.pack_forget()
        else:
            self.env_frame.pack_forget()
            self.file_frame.pack(fill="x", padx=10, pady=(5,0))
        self._update_run_button_state()

    def _update_run_button_state(self):
        if self.source_mode.get() == "env":
            creds_ok = all([
                self.source_username.get().strip(),
                self.source_password.get().strip(),
                self.source_security_token.get().strip(),
                self.query_df is not None
            ])
            state = "normal" if creds_ok else "disabled"
        else:
            txt = self.file_label.cget("text")
            state = "normal" if txt.endswith((".csv", ".xls", ".xlsx")) else "disabled"
        self.run_button.config(state=state)

    def _select_file(self):
        types = [("CSV file", "*.csv")] if self.file_type.get()=="CSV" else [("Excel file", "*.xlsx")]
        path = filedialog.askopenfilename(filetypes=types)
        display = path or "Nessun file selezionato"
        self.file_label.config(text=display)

        if path and self.source_mode.get()=="env":
            cols = CONFIG["query_csv"]["columns"]
            sep  = CONFIG["query_csv"].get("separator", ",")
            self.query_df = pd.read_csv(path, sep=sep, usecols=cols)

        if path:
            cfg = load_config()
            cfg["input_tables"] = path
            save_config(cfg)

        if self.source_mode.get()=="file" and path.lower().endswith((".xls", ".xlsx")):
            self._show_excel_tabs(path)

        self._update_run_button_state()

    def _start_run_queries(self):
        self.run_button.config(state="disabled")
        self.progress.pack(fill="x", padx=10, pady=(5,0))
        self.progress.start(10)

        thread = threading.Thread(target=self._run_queries_thread, daemon=True)
        thread.start()

    def _run_queries_thread(self):
        try:
            self._run_queries_logic()
        except Exception as e:
            # se c’è un errore non previsto, lo mostriamo nella GUI
            self.after(0, lambda: messagebox.showerror("Errore", str(e)))
        finally:
            # al termine, anche in caso di errore, ripristiniamo la UI
            self.after(0, self._on_queries_complete)

    def _run_queries_logic(self):
        # ---- autenticazione Salesforce ----
        src_type = self.source_env_type.get().lower()  # "sandbox" o "production"
        domain   = "test" if src_type=="sandbox" else "login"
        self.sf = Salesforce(
            username=self.source_username.get(),
            password=self.source_password.get(),
            security_token=self.source_security_token.get(),
            domain=domain
        )

        # ---- scelta del file di output ----
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file","*.xlsx")]
        )
        if not save_path:
            return
        self._update_config_with_input_table_path(save_path)

        # ---- esecuzione delle query ----
        col_sobj, col_soql = CONFIG["query_csv"]["columns"]
        sheets_written = []

        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            for _, row in self.query_df.iterrows():
                sobject = row[col_sobj]
                soql    = row[col_soql]
                try:
                    resp    = self.sf.query_all(soql)
                    records = resp.get("records", [])
                    # escludo l’attributo “attributes”
                    data    = [
                        {k:v for k,v in r.items() if k!="attributes"}
                        for r in records
                    ]
                    df_res = pd.DataFrame(data)
                except Exception as err:
                    # warning non blocca il ciclo
                    self.after(0, lambda e=err, so=sobject:
                               messagebox.showwarning("Query fallita",
                                                      f"{so}: {e}"))
                    continue

                sheet_name = sobject[:31]
                df_res.to_excel(writer, sheet_name=sheet_name, index=False)
                sheets_written.append(sheet_name)

            # fallback se nessun foglio creato
            if not sheets_written:
                pd.DataFrame(
                    ["Nessuna query eseguita correttamente"]
                ).to_excel(writer, sheet_name="Callback", index=False, header=False)

        # Informo al termine
        def finish_and_show():
            messagebox.showinfo("Completato", f"File salvato in:\n{save_path}")
            self._show_excel_tabs(save_path)
        self.after(0, finish_and_show)

    
    def _on_queries_complete(self):
        self.progress.stop()
        self.progress.pack_forget()
        self._update_run_button_state()


    def _show_excel_tabs(self, path):
        """
        Apre una finestra con un Notebook: ogni tab è un foglio Excel (o il CSV).
        In ciascuna tab permette di scegliere 'insert' o 'upsert', e per upsert
        di selezionare il campo External ID. Infine un pulsante per confermare
        le impostazioni e salvarle in config.yaml.
        """
        try:
            sheets = read_spreadsheet(path)
        except Exception as e:
            messagebox.showerror("Errore lettura file", str(e))
            return

        # Carica config attuale (per non sovrascrivere altre sezioni)
        cfg = load_config()

        # Resetto le impostazioni correnti
        self.sheet_settings = {}

        win = tk.Toplevel(self)
        win.title(f"Anteprima: {os.path.basename(path)}")
        win.geometry("900x600")

        notebook = ttk.Notebook(win)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        for sheet_name, df in sheets.items():
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=sheet_name[:31])

            # ─── Action Frame ─────────────────────────────────
            action_frame = ttk.LabelFrame(frame, text="Import Settings", padding=5)
            action_frame.pack(fill="x", pady=(0,10), padx=5)

            # Modalità insert / upsert
            mode_var = tk.StringVar(value="insert")
            rb_insert = ttk.Radiobutton(action_frame, text="Insert", variable=mode_var, value="insert")
            rb_upsert = ttk.Radiobutton(action_frame, text="Upsert", variable=mode_var, value="upsert")
            rb_insert.pack(side="left", padx=5)
            rb_upsert.pack(side="left", padx=5)

            # Selezione External ID field (solo se upsert)
            field_var = tk.StringVar()
            lbl_field = ttk.Label(action_frame, text="External ID Field:")
            cmb_field = ttk.Combobox(action_frame, values=list(df.columns),
                                     textvariable=field_var, state="disabled", width=20)
            lbl_field.pack(side="left", padx=(20,5))
            cmb_field.pack(side="left", padx=5)

            # Abilita/disabilita combo al cambio di mode_var
            def _on_mode_change(*args, combo=cmb_field, var=mode_var):
                combo.config(state="readonly" if var.get()=="upsert" else "disabled")
                if var.get()=="insert":
                    field_var.set("")  # reset se torni a insert
            mode_var.trace_add("write", _on_mode_change)

            # Salvo i var in un dict per dopo
            self.sheet_settings[sheet_name] = {
                "mode_var": mode_var,
                "field_var": field_var
            }

            # ─── Treeview ───────────────────────────────────────
            tree_frame = ttk.Frame(frame)
            tree_frame.pack(fill="both", expand=True, padx=5, pady=(0,5))
            tree = ttk.Treeview(tree_frame, columns=list(df.columns), show="headings")
            vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=tree.yview)
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            tree_frame.rowconfigure(0, weight=1)
            tree_frame.columnconfigure(0, weight=1)

            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=100, anchor="w")

            for _, row in df.iterrows():
                tree.insert("", "end", values=[row[c] for c in df.columns])

        # ─── Pulsante Conferma ────────────────────────────────
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill="x", pady=(0,10))
        confirm_btn = ttk.Button(
            btn_frame, text="Conferma impostazioni", command=lambda: self._confirm_import_settings(win)
        )
        confirm_btn.pack(side="right", padx=10)

    def _confirm_import_settings(self, window):
        """
        Raccolta le scelte per ogni sheet e salva in import_config.yaml
        sotto 'import_settings'.
        """
        from config import load_import_config, save_import_config, IMPORT_CONFIG_FILE

        # Carico la configurazione di import esistente
        import_cfg = load_import_config()

        # Sovrascrivo solo la sezione import_settings
        settings = {}
        for sheet, vars in self.sheet_settings.items():
            action = vars["mode_var"].get()
            ext_id = vars["field_var"].get() if action == "upsert" else None
            settings[sheet] = {
                "action": action,
                "externalIdField": ext_id
            }
        import_cfg["import_settings"] = settings

        # Salvo su import_config.yaml
        try:
            save_import_config(import_cfg)
            messagebox.showinfo(
                "Salvato",
                f"Le impostazioni di import sono salvate in:\n{IMPORT_CONFIG_FILE}"
            )
            window.destroy()
        except Exception as e:
            messagebox.showerror("Errore salvataggio", str(e))
            
    def _update_config_with_input_table_path(self, path):
        try:
            # Leggi la configurazione esistente
            with open(CONFIG_FILE, "r") as f:
                cfg = yaml.safe_load(f)
            # Imposta o sovrascrive il parametro input_tables
            cfg["input_tables"] = path
            # Scrivi nuovamente il file YAML
            with open(CONFIG_FILE, "w") as f:
                yaml.safe_dump(cfg, f, sort_keys=False, default_flow_style=False)
        except Exception as e:
            messagebox.showwarning(
                "Impossibile aggiornare config",
                f"Errore durante il salvataggio di input_tables in {CONFIG_FILE}:\n{e}"
                )


if __name__ == "__main__":
    app = MigrationToolApp()
    app.mainloop()
