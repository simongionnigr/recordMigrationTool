import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from simple_salesforce import Salesforce
import yaml
from collections import deque

import pandas as pd

from config import load_config, load_import_config, save_config, save_import_config, IMPORT_CONFIG_FILE,  CONFIG_FILE
from sf_client import SalesforceClient
from excel_utils import read_spreadsheet, write_spreadsheet

# Carica la configurazione YAML
CONFIG = load_config()
IMPORT_CONFIG = load_import_config()

class MigrationToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Salesforce Data Migrator")
        self.geometry("900x600")

        self.sf_client = None
        self.query_df = None
        self.relationship_df = None   # terrà i mapping child→parent→field


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

        # Relazioni tra SObject 
        rel_frame = ttk.LabelFrame(main, text="Relazioni SObject", padding=10)
        rel_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10,0))
        ttk.Button(rel_frame, text="Carica relazioni…", command=self._load_relationships) \
            .grid(row=0, column=0, sticky="w", padx=(0,5))
        self.relationship_label = ttk.Label(rel_frame, text="Nessun file relazioni selezionato")
        self.relationship_label.grid(row=0, column=1, sticky="w")
        rel_frame.columnconfigure(1, weight=1)


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

         # ——— Import button (grafica, senza logica per ora) ———
        self.import_button = ttk.Button(
            main,
            text="Importa Dati",
            command=self._start_run_import, 
            state="disabled"           
        )
        self.import_button.grid(
            row=4, column=0, columnspan=2,
            sticky="ew", pady=(10,0)
        )

        # Progress bar (inizialmente nascosta)
        self.progress = ttk.Progressbar(self, mode="indeterminate")

        # Menu contestuale per copia cella
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Copy", command=self._copy_cell)
        self.query_table.bind("<Button-3>", self._show_context_menu)

        # Abilitazione run_button al cambiamento credenziali
        # ogni volta che digito in un campo credenziali, aggiorno lo stato dei bottoni
        for prefix in ("source", "target"):
            for field in ("username","password","security_token"):
                ent = getattr(self, f"{prefix}_{field}")
                ent.bind("<KeyRelease>", lambda e: self._update_buttons_state())


        self._toggle_source()

    def _load_relationships(self):
        """
        Carica un CSV/XLSX a 3 colonne:
        [child_sobject, parent_sobject, child_lookup_field]
        """
        types = [("CSV file","*.csv"),("Excel file","*.xlsx")]
        path = filedialog.askopenfilename(filetypes=types)
        if not path:
            return

        try:
            if path.lower().endswith(".csv"):
                sep = CONFIG["query_csv"].get("separator", ",")
                df = pd.read_csv(path, sep=sep)
            else:
                df = pd.read_excel(path)
            # Prendo solo prime 3 colonne e le rinomino
            df = df.iloc[:, :3]
            df.columns = ["child_sobject", "parent_sobject", "child_field"]
            self.relationship_df = df
            self.relationship_label.config(text=os.path.basename(path))
        except Exception as e:
            messagebox.showerror("Errore file relazioni", str(e))
            self.relationship_df = None


    def _apply_relationship_mappings(self, sheets: dict):
        """
        Per ogni mapping child→parent→field:
        - sostituisce in sheets[child][field] i valori originali
            con i corrispondenti record_id di sheets[parent].
        """
        for _, row in self.relationship_df.iterrows():
            child = row["child_sobject"]
            parent = row["parent_sobject"]
            field = row["child_field"]

            if child not in sheets or parent not in sheets:
                continue
            df_child  = sheets[child]
            df_parent = sheets[parent]

            # Serve la colonna "Id" nel parent per mappare → record_id
            if field not in df_child.columns or "Id" not in df_parent.columns:
                continue

            # Costruisco dict: original parent Id → nuovo record_id
            mapping = df_parent.set_index("Id")["record_id"].to_dict()
            # Applico la sostituzione; se non trovo corrispondenza, lascio il valore originale
            df_child[field] = df_child[field].map(mapping).fillna(df_child[field])

            sheets[child] = df_child


    def _make_cred_fields(self, frame, prefix):
        ttk.Label(frame, text="Tipo ambiente").grid(row=0, column=0, sticky="w")
        comb = ttk.Combobox(frame, values=["Sandbox", "Production"], state="readonly")
        comb.current(0)
        comb.grid(row=0, column=1, sticky="ew")

        # ← aggiungi qui l’assegnazione per entrambi i prefissi
        if prefix == "source":
            self.source_env_type = comb
        elif prefix == "target":
            self.target_env_type = comb

        labels = ["Username", "Password", "Security Token"]
        for i, label in enumerate(labels, start=1):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w")
            show = "*" if label != "Username" else None
            entry = ttk.Entry(frame, show=show) if show else ttk.Entry(frame)
            entry.grid(row=i, column=1, sticky="ew")
            setattr(self, f"{prefix}_{label.lower().replace(' ','_')}", entry)

        for i in range(4):
            frame.rowconfigure(i, pad=5)
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
        self._update_buttons_state()


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

    def _update_buttons_state(self):
        # ── Abilita “Esegui Query e Esporta Excel” ───────────
        run_ok = False
        if self.source_mode.get() == "env":
            run_ok = (
                self.query_df is not None
                and self.source_username.get().strip()
                and self.source_password.get().strip()
                and self.source_security_token.get().strip()
            )
        else:
            run_ok = self.file_label.cget("text").endswith((".csv", ".xls", ".xlsx"))

        self.run_button.config(state="normal" if run_ok else "disabled")

        # ── Abilita “Importa Dati” ────────────────────────────
        # 1) dati pronti (export o file)
        data_ready = run_ok
        # 2) credenziali destinazione
        tgt_ok = (
            getattr(self, "target_username").get().strip() and
            getattr(self, "target_password").get().strip() and
            getattr(self, "target_security_token").get().strip()
        )
        self.import_button.config(state="normal" if data_ready and tgt_ok else "disabled")


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
            cfg = load_import_config()
            cfg["input_tables"] = path
            save_import_config(cfg)

        if self.source_mode.get()=="file" and path.lower().endswith((".xls", ".xlsx")):
            self._show_excel_tabs(path)

        self._update_buttons_state()


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
        #sheets_written = []

        # 1) Esecuzione delle query e raccolta risultati in dict
        col_sobj, col_soql = CONFIG["query_csv"]["columns"]
        sheets = {}
        for _, row in self.query_df.iterrows():
            sobject = row[col_sobj]
            soql    = row[col_soql]
            try:
                resp    = self.sf.query_all(soql)
                records = resp.get("records", [])
                data    = [{k:v for k,v in r.items() if k!="attributes"} for r in records]
                df_res  = pd.DataFrame(data)
                sheets[sobject[:31]] = df_res
            except Exception as err:
                self.after(0, lambda e=err, so=sobject:
                           messagebox.showwarning("Query fallita",
                                                  f"{so}: {e}"))
                continue

        # 2) Se non ho scritte tabelle, aggiungo un foglio di fallback
        if not sheets:
            sheets["Callback"] = pd.DataFrame(
                ["Nessuna query eseguita correttamente"], columns=["message"]
            )

        # 3) Scrivo con write_spreadsheet, passando anche le relazioni
        write_spreadsheet(save_path, sheets, self.relationship_df)

        # Informo al termine
        def finish_and_show():
            messagebox.showinfo("Completato", f"File salvato in:\n{save_path}")
            self._show_excel_tabs(save_path)
        self.after(0, finish_and_show)

    
    def _on_queries_complete(self):
        self.progress.stop()
        self.progress.pack_forget()
        self._update_run_button_state()


    def _start_run_import(self):
        """
        Disable import button, show progress bar e lancia il thread di import.
        """
        self.import_button.config(state="disabled")
        self.progress.pack(fill="x", padx=10, pady=(5,0))
        self.progress.start(10)

        thread = threading.Thread(target=self._run_import_logic, daemon=True)
        thread.start()

    def _run_import_logic(self):
        """
        1) Autentica su Salesforce destinazione
        2) Legge import_config.yaml (input_tables, import_order, import_settings)
        3) Legge i fogli da input_tables, inietta record_id se serve
        4) Esegue insert/upsert per ogni oggetto nell'ordine corretto
        5) Scrive su ciascun foglio gli sf_id e gli errori
        6) Applica mapping relazioni ai figli
        7) Salva il file di output e mostra anteprima
        """
        try:
            # --- 1) Autenticazione destinazione ---
            dest_cfg = load_import_config()
            # recupero credenziali dai widget
            dest_type = self.target_env_type.get().lower()
            domain    = "test" if dest_type=="sandbox" else "login"
            sf_dest = Salesforce(
                username=self.target_username.get().strip(),
                password=self.target_password.get().strip(),
                security_token=self.target_security_token.get().strip(),
                domain=domain
            )

            # --- 2) Lettura config di import ---
            cfg = load_import_config()
            path = cfg.get("input_tables")
            if not path:
                raise RuntimeError("Parametro 'input_tables' non trovato in import_config.yaml")

            import_order    = cfg.get("import_order", [])
            import_settings = cfg.get("import_settings", {})
            ignore_cols     = cfg.get("ignore_columns", ["record_id","to_import","sf_id","error"])

            # --- 3) Carica i fogli con record_id già iniettato ---
            sheets = read_spreadsheet(path)

            # se non c'è import_order, uso l'ordine dei fogli
            if not import_order:
                import_order = list(sheets.keys())

            # preparazione mappe
            sf_id_map    = {}   # {sobject: {internal_id: sf_id}}
            failure_map  = {}   # {sobject: set(internal_id)}

            # --- 4) Ciclo di import per ciascun oggetto ---
            for sobject in import_order:
                if sobject not in sheets:
                    continue

                df = sheets[sobject]
                internal_ids = df["record_id"].tolist()
                # --- Droppo le colonne da ignorare ---
                data_df = df.drop(columns=ignore_cols, errors="ignore")

                # scelta modalità
                setting = import_settings.get(sobject, {})
                action      = setting.get("action", "insert")
                ext_id_fld  = setting.get("externalIdField")

                results = []
                for idx, record in enumerate(data_df.to_dict(orient="records")):
                    try:
                        if action=="upsert" and ext_id_fld:
                            res = sf_dest.__getattr__(sobject).upsert(
                                f"{ext_id_fld}/{record[ext_id_fld]}", record
                            )
                            sf_id = res.get("id") or res.get("Id")
                        else:
                            res = sf_dest.__getattr__(sobject).create(record)
                            sf_id = res.get("id") or res.get("Id")
                        results.append((internal_ids[idx], sf_id, None))
                    except Exception as e:
                        results.append((internal_ids[idx], None, str(e)))

                # 5) Annotazione sf_id ed errori sul DataFrame
                df["sf_id"] = [r[1] for r in results]
                df["error"] = [r[2] for r in results]

                # aggiorno mappe per le relazioni future
                sf_id_map[sobject]   = {r[0]: r[1] for r in results if r[1]}
                failure_map[sobject] = {r[0] for r in results if not r[1]}

                sheets[sobject] = df

                # 6) Applica mapping relazioni ai figli
                if self.relationship_df is not None:
                    for _, rel in self.relationship_df.iterrows():
                        parent = rel["parent_sobject"]
                        child  = rel["child_sobject"]
                        field  = rel["child_field"]
                        if parent != sobject or child not in sheets:
                            continue
                        child_df = sheets[child]
                        # colonna di flag per non importare
                        if "to_import" not in child_df.columns:
                            child_df["to_import"] = True
                        if "error" not in child_df.columns:
                            child_df["error"] = None

                        # sostituisco id e flaggo fallimenti
                        new_vals, to_imp, errs = [], [], []
                        for rid in child_df[field]:
                            if rid in sf_id_map[parent]:
                                new_vals.append(sf_id_map[parent][rid])
                                to_imp.append(True)
                                errs.append(None)
                            else:
                                new_vals.append(rid)
                                to_imp.append(False)
                                errs.append(
                                    f"Parent {parent} load failed"
                                    if rid in failure_map[parent]
                                    else None
                                )
                        child_df[field]     = new_vals
                        child_df["to_import"]= to_imp
                        # unisco eventuali errori preesistenti
                        child_df["error"]   = child_df["error"].fillna(pd.Series(errs, index=child_df.index))

                        sheets[child] = child_df

            # --- 7) Salvataggio del nuovo Excel ---
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel file","*.xlsx")]
            )
            if save_path:
                write_spreadsheet(save_path, sheets, self.relationship_df)

            # notifica + anteprima
            self.after(0, lambda:
                messagebox.showinfo("Import completato", f"File di log salvato in:\n{save_path}")
            )
            self.after(0, lambda: self._show_excel_tabs(save_path))

        except Exception as exc:
            # catturo exc come default arg in modo che la lambda lo conservi
            self.after(0, lambda exc=exc:
                messagebox.showerror("Errore Import", str(exc))
            )
        finally:
            self.after(0, self.progress.stop)
            self.after(0, self.progress.pack_forget)
            # riabilita il pulsante per eventuali retry
            self.after(0, lambda: self.import_button.config(state="normal"))

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
        
        # ─── Applica le relazioni Lookup se caricate ─────────────────
        if self.relationship_df is not None:
            self._apply_relationship_mappings(sheets)

        # Carica config attuale (per non sovrascrivere altre sezioni)
        cfg = load_import_config()

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
        sotto 'import_settings' e 'import_order'.
        """
        # 1) Carico la configurazione di import esistente
        import_cfg = load_import_config()

        # 2) Costruisco import_settings come prima
        settings = {}
        for sheet, vars in self.sheet_settings.items():
            action = vars["mode_var"].get()
            ext_id = vars["field_var"].get() if action == "upsert" else None
            settings[sheet] = {
                "action": action,
                "externalIdField": ext_id
            }
        import_cfg["import_settings"] = settings

        # 3) Calcolo l'ordine di import (topological sort)
        if self.relationship_df is not None and not self.relationship_df.empty:
            # Nodi: tutti i parent e child unici
            rels = self.relationship_df
            nodes = set(rels["child_sobject"]) | set(rels["parent_sobject"])
            # Grafo parent -> set(children)
            graph = {nod: set() for nod in nodes}
            in_degree = {nod: 0 for nod in nodes}

            for _, row in rels.iterrows():
                parent = row["parent_sobject"]
                child  = row["child_sobject"]
                # aggiungo arco parent->child
                graph[parent].add(child)
                in_degree[child] += 1

            # coda dei nodi con in_degree 0
            q = deque([n for n,d in in_degree.items() if d == 0])
            order = []
            while q:
                n = q.popleft()
                order.append(n)
                for ch in graph[n]:
                    in_degree[ch] -= 1
                    if in_degree[ch] == 0:
                        q.append(ch)

            # se c'è un ciclo, includo comunque eventuali nodi mancanti
            if len(order) < len(nodes):
                missing = nodes - set(order)
                order.extend(sorted(missing))

            import_cfg["import_order"] = order

        # 4) Salvo su import_config.yaml
        try:
            save_import_config(import_cfg)
            messagebox.showinfo(
                "Salvato",
                f"Le impostazioni di import e l'ordine sono salvate in:\n{IMPORT_CONFIG_FILE}"
            )
            window.destroy()
        except Exception as e:
            messagebox.showerror("Errore salvataggio", str(e))

    def _update_config_with_input_table_path(self, path):
        try:
            # Leggi la configurazione esistente
            with open(IMPORT_CONFIG_FILE, "r") as f:
                cfg = yaml.safe_load(f)
            # Imposta o sovrascrive il parametro input_tables
            cfg["input_tables"] = path
            # Scrivi nuovamente il file YAML
            with open(IMPORT_CONFIG_FILE, "w") as f:
                yaml.safe_dump(cfg, f, sort_keys=False, default_flow_style=False)
        except Exception as e:
            messagebox.showwarning(
                "Impossibile aggiornare config",
                f"Errore durante il salvataggio di input_tables in {CONFIG_FILE}:\n{e}"
                )


if __name__ == "__main__":
    app = MigrationToolApp()
    app.mainloop()
