# tracciato-gdo

Generazione del TRACCIATO_GDO, il file di rifatturazione a record fisso
richiesto da alcuni clienti della Grande Distribuzione Organizzata.

Richiede: database ARCA Evolution (SQL Server).

## Problema

Ogni mese alcuni clienti GDO richiedono fatture (`FTV`, `FT2`) e note di credito
(`NCC`) del periodo in un file di testo con righe di 128 caratteri. Ogni
cliente applica varianti di formato proprie, e le fatture riepilogative devono
riportare una testata per ciascun DDT di riferimento.

## Soluzione

`tracciato_gdo_export.py` genera il tracciato direttamente da ARCA.

Struttura dei record:

| Record | Contenuto |
|--------|-----------|
| `01` testata | Id documento, numero e data, numero e data DDT, codice punto vendita (`xPV` della destinazione), tipo documento, valuta |
| `02` riga | Id documento, codice articolo, descrizione (30 caratteri), unità di misura, quantità, prezzo unitario e totale, aliquota IVA |

Elaborazioni principali:

- le righe descrittive `Rif: DDT <num> del <data> Dest. <cod>` delle fatture
  riepilogative diventano testate `01` inserite prima degli articoli del DDT;
  il codice punto vendita viene recuperato dal DDT originale o, in alternativa,
  dalla destinazione indicata nella descrizione;
- le opzioni di formato sono definite per cliente in `ClientManager.clienti`:
  `tipo_doc_format` (`alpha` F/N o `numeric` 0/1), `include_valuta`,
  `prezzo_al_pezzo` e `qta_in_pezzi` (conversione da cartone a pezzo),
  `um_fissa`, `rimuovi_trattini`, `ddt_headers`;
- le righe di lunghezza diversa da 128 caratteri vengono segnalate nel log;
- il file di ogni cliente può essere inviato via Gmail API con un riepilogo
  nel corpo del messaggio.

## Utilizzo

1. Configurare il `.env` alla radice del repository (connessione ad ARCA e,
   per l'invio, `GMAIL_CREDENTIALS_FILE` e `GMAIL_TOKEN_FILE`).
2. In `ClientManager` impostare per ogni cliente codice ARCA, nome, destinatario
   (`email`), copie (`cc`) e opzioni di formato.
3. In `calculate_previous_month_period()` scegliere tra il periodo fisso
   (`data_inizio`/`data_fine`) e il calcolo automatico, commentato, dal 7 del
   mese precedente al 7 del mese corrente.
4. Eseguire:
   ```bash
   python tracciato_gdo_export.py
   ```

L'esecuzione standard elabora tutti i clienti senza inviare email; i file
vengono salvati in `output_tracciato_gdo/` e il log in
`tracciato_gdo_extraction.log`. Per l'invio si usa
`TracciatoGdoExtractorMultiClient().run()` oppure
`process_client(..., send_email=True)`. Il locale `Italian_Italy.1252`, usato
per i nomi dei mesi, è specifico di Windows.
