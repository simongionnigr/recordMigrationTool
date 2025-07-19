# Salesforce Data Migrator

**Salesforce Data Migrator** è un'applicazione desktop con interfaccia grafica (GUI) per esportare ed importare dati tra organizzazioni Salesforce (sandbox o production).

## 🧰 Funzionalità principali

- **Esportazione dati**

  - Autenticazione su sandbox o production tramite username, password e security token
  - Esecuzione batch di query SOQL definite in un file CSV
  - Generazione automatica delle relazioni (lookup o master detail) tra SObject selezionati
  - Esportazione dei risultati in file Excel con fogli separati e `record_id` unico per ogni record

- **Importazione dati**

  - Configurazione di modalità `insert` o `upsert` per ciascun SObject
  - Selezione di un campo External ID per le operazioni di upsert
  - Applicazione automatica delle relazioni parent–child per gestire i riferimenti tra record
  - Salvataggio del log con `sf_id` generati e eventuali errori in un file Excel di output
  - Calcolo automatico dell'ordine di import sulla base delle dipendenze tra oggetti

## 🖥️ Requisiti

- Python ≥ 3.8
- Librerie Python:
  ```
  pandas
  openpyxl
  pyyaml
  simple-salesforce
  ttkthemes
  ```
- `tkinter` (incluso nella maggior parte delle distribuzioni Python)

## 📦 Installazione

1. Clona il repository:

   ```bash
   git clone https://github.com/tuo-utente/salesforce-data-migrator.git
   cd salesforce-data-migrator
   ```

2. Crea un ambiente virtuale (opzionale ma consigliato):

   ```bash
   python -m venv venv
   source venv/bin/activate   # Windows: venv\Scripts\activate
   ```

3. Installa le dipendenze:

   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Avvio

Per avviare l'applicazione:

```bash
python main.py
```

## ⚙️ Configurazione

### config.yaml

Salvato automaticamente al primo avvio, contiene impostazioni per la GUI e per il caricamento delle query:

```yaml
query_csv:
  columns: ["sobject_api", "soql"]
  prompt: "Carica file query CSV"
  filetypes: [["CSV file", "*.csv"]]
  separator: ";"
ui:
  theme: "clam"
```

### import_config.yaml

Memorizza percorso del file di input, ordine e impostazioni di import:

```yaml
input_tables: "/percorso/a/file.xlsx"
import_order:
  - Account
  - Contact
import_settings:
  Account:
    action: "upsert"
    externalIdField: "External_Id__c"
  Contact:
    action: "insert"
ignore_columns: ["record_id", "to_import", "sf_id", "error"]
```

## 📁 Struttura del progetto

```
.
├── config.py           # Gestione caricamento e salvataggio config YAML
├── config.yaml         # Config GUI e query (generato automaticamente)
├── import_config.yaml  # Config import (generato automaticamente)
├── excel_utils.py      # Lettura/scrittura di file Excel e CSV, generazione record_id
├── gui.py              # Interfaccia grafica principale con Tkinter e ttkthemes
├── sf_client.py        # Wrapper per autenticazione e query SOQL con simple_salesforce
├── main.py             # Entry point dell’applicazione
├── requirements.txt    # Elenco dipendenze Python
```

## 📝 Formati file

### File query CSV

Deve contenere le seguenti due colonne:

| sobject_api | soql                          |
| ------------ | ----------------------------- |
| Account      | SELECT Id, Name FROM Account  |
| Contact      | SELECT Id, Email FROM Contact |

### File relazioni CSV (opzionale)

Tre colonne: `child_sobject`, `parent_sobject`, `child_field`, ad esempio:

| child_sobject | parent_sobject | child_field |
| -------------- | --------------- | ------------ |
| Contact        | Account         | AccountId    |

## 📌 Note

- Viene generato un `record_id` UUID per ogni riga se non presente.
- I nomi dei fogli Excel sono limitati a 31 caratteri per compatibilità.
- In caso di errore in una query o in un import, l’app mostra un warning ma prosegue con le altre operazioni.

