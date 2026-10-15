# edi-tracciati-gdo

Generazione di tracciati a record fisso per la rifatturazione verso la Grande
Distribuzione Organizzata e riconciliazione delle fatture dei corrieri.

Richiede: database ARCA Evolution (SQL Server).

## Contesto

Strumenti sviluppati per le esigenze di un'azienda di distribuzione che usa
ARCA Evolution, pubblicati in forma anonimizzata con l'autorizzazione
dell'azienda.

## Problema

- Alcuni clienti della GDO richiedono ogni mese l'elenco di fatture e note di
  credito in un tracciato testuale a lunghezza fissa (128 caratteri per riga),
  con regole di formato diverse per ciascun cliente.
- Le fatture dei corrieri vanno controllate spedizione per spedizione rispetto
  ai documenti di trasporto registrati nel gestionale.

## Soluzione

- `tracciato-gdo/` estrae da ARCA testate (record `01`) e righe (record `02`)
  dei documenti del periodo, applica le regole di ciascun cliente, gestisce i
  riferimenti ai DDT nelle fatture riepilogative e produce un file per
  cliente, con invio opzionale tramite Gmail API.
- `carrier-invoice-reconciliation/` affianca a ogni riga della fattura del
  corriere i dati ricostruiti da ARCA (destinatario, peso, tipo di spedizione)
  in un Excel di confronto.

## Struttura

```
edi-tracciati-gdo/
├── tracciato-gdo/
│   └── tracciato_gdo_export.py          estrazione multi-cliente e invio
└── carrier-invoice-reconciliation/
    ├── reconcile_carrier_a.py           fattura CORRIERE_A (CSV)
    └── reconcile_carrier_b.py           fattura CORRIERE_B (Excel)
```

## Tecnologie

- Python 3.10+: pandas, SQLAlchemy, pyodbc, openpyxl, xlrd, python-dotenv
- Gmail API (google-api-python-client, OAuth 2.0)
- ARCA Evolution su SQL Server: `DOTes`, `DORig`, `CF`, `CFDest`, `AR`

## Configurazione

1. Installare il driver ODBC per SQL Server e le dipendenze con
   `pip install -r requirements.txt`.
2. Copiare `.env.example` in `.env` e impostare:

   | Variabile                 | Significato                                      |
   |---------------------------|--------------------------------------------------|
   | `SERVER_DB`, `DB_ARCA`    | server/istanza e database di ARCA                |
   | `DB_USER`, `DB_PASSWORD`  | credenziali SQL (vuote: autenticazione Windows)  |
   | `GMAIL_CREDENTIALS_FILE`  | client OAuth scaricato da Google Cloud Console   |
   | `GMAIL_TOKEN_FILE`        | token generato al primo accesso                  |

3. Configurare clienti e destinatari del tracciato come descritto in
   `tracciato-gdo/README.md`.

Il file `.env` non è versionato.
