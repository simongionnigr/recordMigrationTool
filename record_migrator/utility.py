import pandas as pd
import numpy as np
from typing import List, Dict, Iterable, Optional

def sanitize_for_salesforce(
    df: pd.DataFrame,
    drop_empty_fields: bool = False,
    keep_keys: Optional[Iterable[str]] = None,
) -> List[Dict]:
    """
    Converte un DataFrame in una lista di dict pronti per Salesforce:
    - Sostituisce NaN/NaT/pd.NA con None
    - Converte numpy scalars in tipi Python puri
    - Converte Timestamp in ISO 8601
    - (Opzionale) rimuove le chiavi con valore None, tranne quelle in keep_keys

    Parameters
    ----------
    df : pd.DataFrame
        Dati da inviare.
    drop_empty_fields : bool
        Se True rimuove le chiavi con valore None.
    keep_keys : Iterable[str] | None
        Campi da mantenere comunque (es. externalIdField).

    Returns
    -------
    List[Dict]
        Lista di record “puliti”.
    """
    keep_keys = set(keep_keys or [])

    def _clean_value(v):
        if pd.isna(v):
            return None
        if isinstance(v, (np.integer, np.floating)):
            return v.item()
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime().isoformat()
        return v

    cleaned_records = []
    for _, row in df.iterrows():
        rec = {k: _clean_value(v) for k, v in row.items()}
        if drop_empty_fields:
            rec = {k: v for k, v in rec.items() if (v is not None or k in keep_keys)}
        cleaned_records.append(rec)
    return cleaned_records
