# Salesforce Data Migrator

**Salesforce Data Migrator** è uno strumento desktop con interfaccia grafica (GUI) per migrare dati tra org Salesforce. Supporta sia l'esportazione batch tramite query SOQL su più oggetti, sia l'import di record da Excel o CSV verso un ambiente di destinazione, con gestione delle relazioni tra oggetti.

## 🧰 Funzionalità principali

- Autenticazione in ambienti Sandbox o Production tramite API di Salesforce
- Esecuzione di query SOQL da file CSV e esportazione risultati in Excel
- Importazione massiva di dati da Excel in un ambiente Salesforce
- Gestione automatica delle lookup relationship tra oggetti (opzionale)
- Interfaccia utente temabile basata su `ttkthemes`
- Configurazione persistente tramite file YAML

## 📦 Installazione

Puoi installare lo strumento come pacchetto CLI Python:

```bash
pip install --upgrade .
```

oppure:

```bash
pip install .
```

Per modalità di sviluppo (modifiche live):

```bash
pip install -e .
```

## 🚀 Utilizzo

Dopo l'installazione, puoi avviare il programma con il comando:

```bash
salesforce-data-migrator
```

oppure, se preferisci, puoi usare direttamente lo script principale:

```bash
python -m record_migrator.main
```

## 🖥️ Versione Standalone

Per chi non vuole installare Python, è disponibile nella sezione **Releases** uno ZIP contenente l'eseguibile già compilato (`.exe`) pronto all'uso su Windows. Basta scaricarlo, estrarlo e lanciare l'eseguibile.

## 📁 Struttura del progetto

```
record_migrator/
├── __init__.py
├── main.py
├── gui.py
├── config.py
├── excel_utils.py
├── sf_client.py
├── config.yaml
├── import_config.yaml
setup.py
requirements.txt
```

## 🗃️ Formato file query

Il file CSV delle query deve contenere almeno le seguenti colonne:

| sobject_api | soql                         |
|-------------|------------------------------|
| Account     | SELECT Id, Name FROM Account |
| Contact     | SELECT Id, Email FROM Contact|

## 🔄 Formato file relazioni (opzionale)

Un file CSV o XLSX con le relazioni tra oggetti può essere fornito o generato automaticamente. Deve avere le seguenti colonne:

| child_sobject | parent_sobject | child_field |
|---------------|----------------|-------------|

## 🛟 Note

- I file `config.yaml` e `import_config.yaml` vengono creati e aggiornati automaticamente.
- In caso di errori durante query o import, il programma continua e li segnala nella GUI.
- I fogli Excel hanno nome troncato a 31 caratteri per compatibilità con Excel.
