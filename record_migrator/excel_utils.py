import pandas as pd
import uuid


def read_spreadsheet(path: str) -> dict:
    """
    Legge un .xlsx/.xls (tutti i fogli) o un .csv (singolo foglio)
    e restituisce un dict { sheet_name: DataFrame }.
    In ogni DataFrame aggiunge in prima colonna 'record_id' con un UUID univoco per riga.
    """
    ext = path.lower().rsplit(".", 1)[-1]
    if ext in ("xls", "xlsx"):
        sheets = pd.read_excel(path, sheet_name=None)
    elif ext == "csv":
        sheets = {path: pd.read_csv(path)}
    else:
        raise ValueError(f"Formato non supportato: {path}")

    # Inietta il record_id
    for name, df in sheets.items():
        # genera un uuid4 per ogni riga
        ids = [str(uuid.uuid4()) for _ in range(len(df))]
        df.insert(0, "record_id", ids)
        sheets[name] = df

    return sheets

def write_spreadsheet(path: str, sheets: dict):
    """
    Scrive sul file Excel indicato tutti i DataFrame contenuti in sheets,
    creando un foglio per ciascuna chiave.
    Aggiunge una colonna 'record_id' con UUID per ogni riga.
    Aggiunge un foglio "Callback" se sheets è vuoto.
    """
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            # crea copia per non modificare l'originale
            df_copy = df.copy()
            ids = [str(uuid.uuid4()) for _ in range(len(df_copy))]
            df_copy.insert(0, "record_id", ids)
            df_copy.to_excel(writer, sheet_name=name[:31], index=False)

        if not sheets:
            # fallback con un solo record
            fallback = pd.DataFrame(["Nessuna query eseguita correttamente"], columns=["message"])
            fallback.insert(0, "record_id", [str(uuid.uuid4())])
            fallback.to_excel(writer, sheet_name="Callback", index=False)
