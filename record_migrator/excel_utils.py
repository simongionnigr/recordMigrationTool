import pandas as pd
import uuid

def apply_relationship_mappings(sheets: dict, relationship_df: pd.DataFrame):
    """
    Applica sui DataFrame in 'sheets' i mapping definiti in relationship_df:
      - child_sobject, parent_sobject, child_field
    sostituendo in df_child[child_field] i valori originali (Id) 
    con i record_id del df_parent.
    """
    for _, row in relationship_df.iterrows():
        child = row["child_sobject"]
        parent = row["parent_sobject"]
        field = row["child_field"]

        if child not in sheets or parent not in sheets:
            continue

        df_child  = sheets[child]
        df_parent = sheets[parent]

        # devo avere il field nel child e la colonna "Id" + "record_id" nel parent
        if field not in df_child.columns or "Id" not in df_parent.columns:
            continue

        # mappo original Id → record_id
        mapping = df_parent.set_index("Id")["record_id"].to_dict()
        df_child[field] = df_child[field].map(mapping).fillna(df_child[field])
        sheets[child] = df_child

def read_spreadsheet(path: str) -> dict:
    """
    Legge un .xlsx/.xls (tutti i fogli) o un .csv (singolo foglio)
    e restituisce un dict { sheet_name: DataFrame }.
    Inietta 'record_id' solo se non esiste già.
    """
    ext = path.lower().rsplit(".", 1)[-1]
    if ext in ("xls", "xlsx"):
        sheets = pd.read_excel(path, sheet_name=None)
    elif ext == "csv":
        sheets = {path: pd.read_csv(path)}
    else:
        raise ValueError(f"Formato non supportato: {path}")

    for name, df in sheets.items():
        # Inietta record_id solo se manca
        if "record_id" not in df.columns:
            ids = [str(uuid.uuid4()) for _ in range(len(df))]
            df.insert(0, "record_id", ids)
        sheets[name] = df

    return sheets


def write_spreadsheet(path: str, sheets: dict, relationship_df: pd.DataFrame = None):
    """
    Scrive tutti i DataFrame in 'sheets' in un file Excel:
      - inserisce 'record_id' solo se non esiste già
      - se relationship_df non è None, applica i mapping su sheets
      - salva con openpyxl
      - se sheets è vuoto, crea un foglio di fallback
    """
    # 1) Inietta record_id solo se manca
    processed = {}
    for name, df in sheets.items():
        df_copy = df.copy()
        if "record_id" not in df_copy.columns:
            ids = [str(uuid.uuid4()) for _ in range(len(df_copy))]
            df_copy.insert(0, "record_id", ids)
        processed[name] = df_copy

    # 2) Applica mapping relazioni se fornito
    if relationship_df is not None:
        apply_relationship_mappings(processed, relationship_df)

    # 3) Scrivi su Excel
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df_c in processed.items():
            df_c.to_excel(writer, sheet_name=name[:31], index=False)
        if not processed:
            # fallback
            fallback = pd.DataFrame(["Nessuna query eseguita correttamente"], columns=["message"])
            if "record_id" not in fallback.columns:
                fallback.insert(0, "record_id", [str(uuid.uuid4())])
            fallback.to_excel(writer, sheet_name="Callback", index=False)
