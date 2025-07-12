import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class MigrationToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Salesforce Data Migrator")
        self.geometry("600x400")
        self._create_widgets()

    def _create_widgets(self):
        # Frame principale: due colonne
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill="both", expand=True)

        left = ttk.LabelFrame(main_frame, text="Ambiente di partenza", padding=10)
        right = ttk.LabelFrame(main_frame, text="Ambiente di destinazione", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        right.grid(row=0, column=1, sticky="nsew", padx=(5,0))
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Funzione per creare i campi comuni
        def make_cred_fields(frame, prefix):
            ttk.Label(frame, text="Tipo ambiente").grid(row=0, column=0, sticky="w")
            t = ttk.Combobox(frame, values=["Sandbox", "Production"], state="readonly")
            t.current(0)
            t.grid(row=0, column=1, sticky="ew")

            ttk.Label(frame, text="Username").grid(row=1, column=0, sticky="w")
            setattr(self, f"{prefix}_username", ttk.Entry(frame))
            getattr(self, f"{prefix}_username").grid(row=1, column=1, sticky="ew")

            ttk.Label(frame, text="Password").grid(row=2, column=0, sticky="w")
            setattr(self, f"{prefix}_password", ttk.Entry(frame, show="*"))
            getattr(self, f"{prefix}_password").grid(row=2, column=1, sticky="ew")

            ttk.Label(frame, text="Security Token").grid(row=3, column=0, sticky="w")
            setattr(self, f"{prefix}_token", ttk.Entry(frame, show="*"))
            getattr(self, f"{prefix}_token").grid(row=3, column=1, sticky="ew")

            for i in range(4):
                frame.rowconfigure(i, pad=5)
            frame.columnconfigure(1, weight=1)

        make_cred_fields(left, "source")
        make_cred_fields(right, "target")

        # Sezione caricamento file
        file_frame = ttk.LabelFrame(self, text="Carica file dati", padding=10)
        file_frame.pack(fill="x", padx=10, pady=(10,0))

        self.file_type = tk.StringVar(value="CSV")
        ttk.Radiobutton(file_frame, text="CSV", variable=self.file_type, value="CSV").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(file_frame, text="Excel (.xlsx)", variable=self.file_type, value="XLSX").grid(row=0, column=1, padx=5)

        ttk.Button(file_frame, text="Seleziona file…", command=self._select_file).grid(row=0, column=2, padx=10)
        self.file_label = ttk.Label(file_frame, text="Nessun file selezionato")
        self.file_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(5,0))

    def _select_file(self):
        ft = self.file_type.get()
        types = [("CSV file","*.csv")] if ft=="CSV" else [("Excel file","*.xlsx")]
        path = filedialog.askopenfilename(filetypes=types)
        if path:
            self.file_label.config(text=path)
        else:
            self.file_label.config(text="Nessun file selezionato")

if __name__ == "__main__":
    app = MigrationToolApp()
    app.mainloop()
