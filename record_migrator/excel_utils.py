import pandas as pd

def read_spreadsheet(path: str) -> dict:
    """
    Legge un .xlsx/.xls (tutti i fogli) o un .csv (singolo foglio)
    e restituisce un dict { sheet_name: DataFrame }.
    """
    ext = path.lower().rsplit(".", 1)[-1]
    if ext in ("xls", "xlsx"):
        return pd.read_excel(path, sheet_name=None)
    elif ext == "csv":
        return {path: pd.read_csv(path)}
    else:
        raise ValueError(f"Formato non supportato: {path}")

def write_spreadsheet(path: str, sheets: dict):
    """
    Scrive sul file Excel indicato tutti i DataFrame contenuti in sheets,
    creando un foglio per ciascuna chiave.
    Aggiunge un foglio "Callback" se sheets è vuoto.
    """
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
        if not sheets:
            pd.DataFrame(
                ["Nessuna query eseguita correttamente"]
            ).to_excel(writer, sheet_name="Callback", index=False, header=False)
