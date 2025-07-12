import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import yaml

# Caricamento configurazione YAML
CONFIG_FILE = "config.yaml"
DEFAULT_CONFIG = {
    "query_csv": {
        "columns": ["sobject_api", "soql"],
        "prompt": "Carica file query CSV",
        "filetypes": [("CSV file", "*.csv")]
    }
}
if os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE) as f:
        CONFIG = yaml.safe_load(f)
else:
    CONFIG = DEFAULT_CONFIG

class MigrationToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Salesforce Data Migrator")
        self.geometry("700x550")
        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Selettore fonte dati
        source_type_frame = ttk.LabelFrame(main_frame, text="Fonte dati", padding=10)
        source_type_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10))
        self.source_mode = tk.StringVar(value="env")
        ttk.Radiobutton(source_type_frame, text="Ambiente di partenza", variable=self.source_mode,
                        value="env", command=self._toggle_source).grid(row=0, column=0, padx=5)
        ttk.Radiobutton(source_type_frame, text="File CSV/Excel", variable=self.source_mode,
                        value="file", command=self._toggle_source).grid(row=0, column=1, padx=5)

        # Contenitori colonna sinistra e destra
        self.left_container = ttk.Frame(main_frame)
        self.left_container.grid(row=1, column=0, sticky="nsew", padx=(0,5))
        self.target_frame = ttk.LabelFrame(main_frame, text="Ambiente di destinazione", padding=10)
        self.target_frame.grid(row=1, column=1, sticky="nsew", padx=(5,0))
        main_frame.rowconfigure(1, weight=1)

        # Frame credenziali ambiente di partenza
        self.env_frame = ttk.LabelFrame(self.left_container, text="Ambiente di partenza", padding=10)
        self.env_frame.pack(fill="both", expand=True)
        self._make_cred_fields(self.env_frame, "source")

        # Frame caricamento file dati generico (CSV/Excel)
        self.file_frame = ttk.LabelFrame(self.left_container, text="Carica file dati", padding=10)
        self.file_type = tk.StringVar(value="CSV")
        ttk.Radiobutton(self.file_frame, text="CSV", variable=self.file_type, value="CSV").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(self.file_frame, text="Excel (.xlsx)", variable=self.file_type, value="XLSX").grid(row=0, column=1, padx=5)
        ttk.Button(self.file_frame, text="Seleziona file…", command=self._select_file).grid(row=0, column=2, padx=10)
        self.file_label = ttk.Label(self.file_frame, text="Nessun file selezionato")
        self.file_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(5,0))

        # Frame caricamento specifico per query CSV
        self.query_frame = ttk.LabelFrame(self.left_container, text="Query per SObject", padding=10)
        ttk.Button(self.query_frame, text=CONFIG["query_csv"]["prompt"], command=self._load_query_csv).grid(row=0, column=0, sticky="w")
        cols = CONFIG["query_csv"]["columns"]
        self.query_table = ttk.Treeview(self.query_frame, columns=cols, show="headings", height=8)
        for col in cols:
            self.query_table.heading(col, text=col)
            self.query_table.column(col, width=150)
        self.query_table.grid(row=1, column=0, sticky="nsew", pady=(5,0))
        self.query_frame.rowconfigure(1, weight=1)
        self.query_frame.columnconfigure(0, weight=1)

        # Credenziali destinazione
        self._make_cred_fields(self.target_frame, "target")

        # Inizializzazione visibilità
        self._toggle_source()

    def _make_cred_fields(self, frame, prefix):
        ttk.Label(frame, text="Tipo ambiente").grid(row=0, column=0, sticky="w")
        comb = ttk.Combobox(frame, values=["Sandbox", "Production"], state="readonly")
        comb.current(0)
        comb.grid(row=0, column=1, sticky="ew")
        ttk.Label(frame, text="Username").grid(row=1, column=0, sticky="w")
        setattr(self, f"{prefix}_username", ttk.Entry(frame))
        getattr(self, f"{prefix}_username").grid(row=1, column=1, sticky="ew")
        ttk.Label(frame, text="Password").grid(row=2, column=0, sticky="w")
        setattr(self, f"{prefix}_password", ttk.Entry(frame, show="*"))
        getattr(self, f"{prefix}_password").grid(row=2, column=1, sticky="ew")
        ttk.Label(frame, text="Security Token").grid(row=3, column=0, sticky="w")
        setattr(self, f"{prefix}_token", ttk.Entry(frame, show="*"))
        getattr(self, f"{prefix}_token").grid(row=3, column=1, sticky="ew")
        for i in range(4): frame.rowconfigure(i, pad=5)
        frame.columnconfigure(1, weight=1)

    def _toggle_source(self):
        mode = self.source_mode.get()
        if mode == "env":
            self.file_frame.pack_forget()
            self.env_frame.pack(fill="both", expand=True)
            self.query_frame.pack(fill="both", expand=True, pady=(10,0))
        else:
            self.env_frame.pack_forget()
            self.query_frame.pack_forget()
            self.file_frame.pack(fill="both", expand=True)

    def _select_file(self):
        ft = self.file_type.get()
        types = [("CSV file","*.csv")] if ft == "CSV" else [("Excel file","*.xlsx")]
        path = filedialog.askopenfilename(filetypes=types)
        self.file_label.config(text=path if path else "Nessun file selezionato")

    def _load_query_csv(self):
        types = CONFIG["query_csv"]["filetypes"]
        path = filedialog.askopenfilename(filetypes=types)
        if not path:
            return
        try:
            df = pd.read_csv(path,sep=CONFIG["query_csv"]["separator"] ,usecols=CONFIG["query_csv"]["columns"])
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile leggere il file:\n{e}")
            return
        # Pulisci tabella
        for row in self.query_table.get_children():
            self.query_table.delete(row)
        # Inserisci righe
        for _, r in df.iterrows():
            values = [r[c] for c in CONFIG["query_csv"]["columns"]]
            self.query_table.insert("", "end", values=values)

if __name__ == "__main__":
    app = MigrationToolApp()
    app.mainloop()
