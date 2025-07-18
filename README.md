# Salesforce Data Migrator

**Salesforce Data Migrator** è un'applicazione desktop con interfaccia grafica (GUI) che consente di eseguire query SOQL multiple su un'org Salesforce (sandbox o production) e salvare i risultati in un file Excel. È pensata per agevolare attività di migrazione e ispezione dati tra ambienti.

## 🧰 Funzionalità principali

- Autenticazione in ambiente Salesforce tramite username, password e security token
- Esecuzione batch di query definite in un file CSV
- Output dei risultati in un file Excel con fogli separati per ogni oggetto interrogato
- Supporto alla visualizzazione dei risultati esportati direttamente nell’interfaccia
- Alternativa di input da file Excel/CSV già contenente i dati
- Salvataggio e lettura di configurazioni da file YAML

## 🖥️ Requisiti

- Python ≥ 3.8
- Librerie Python elencate in `requirements.txt`

## 📦 Installazione

1. Clona il repository:
   ```bash
   git clone https://github.com/tuo-utente/salesforce-data-migrator.git
   cd salesforce-data-migrator
   ```

2. Crea un ambiente virtuale (opzionale ma consigliato):
   ```bash
   python -m venv venv
   source venv/bin/activate  # Su Windows: venv\Scripts\activate
   ```

3. Installa le dipendenze:
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Avvio

Lancia l'applicazione con:
```bash
python main.py
```

## 📁 Struttura del progetto

```
.
├── config.py           # Gestione della configurazione YAML
├── config.yaml         # Configurazione persistente salvata automaticamente
├── excel_utils.py      # Lettura/scrittura CSV e Excel
├── gui.py              # Interfaccia grafica principale
├── main.py             # Entry point
├── sf_client.py        # Wrapper per le API di Salesforce
├── requirements.txt    # Dipendenze del progetto
```

## 🗃️ Formato file query

Il file CSV deve contenere almeno due colonne:

| sobject_api | soql                         |
|-------------|------------------------------|
| Account     | SELECT Id, Name FROM Account |
| Contact     | SELECT Id, Email FROM Contact|

Può essere configurato nel file `config.yaml`.

## 🛟 Note

- Il file `config.yaml` viene creato/modificato automaticamente e può essere rimosso in sicurezza per ripristinare le impostazioni predefinite.
- Il nome dei fogli Excel è troncato a 31 caratteri per compatibilità con Excel.
- In caso di errore in una query, l'applicazione mostra un warning ma prosegue con le altre.

