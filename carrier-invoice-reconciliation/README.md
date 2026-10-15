# carrier-invoice-reconciliation

Confronto tra le fatture dei corrieri e le spedizioni registrate in ARCA
Evolution.

Richiede: database ARCA Evolution (SQL Server).

## Problema

Le fatture dei corrieri elencano centinaia di spedizioni con destinatario e
peso tassato. Per verificarle occorre ritrovare ogni spedizione nel gestionale
e confrontarne i dati, un controllo che a mano richiede ore.

## Soluzione

Due script, uno per formato di fattura:

| Script | Fattura | Collegamento con ARCA |
|--------|---------|-----------------------|
| `reconcile_carrier_a.py` | CSV con separatore `;` (ISO-8859-1) | riferimento mittente = ID ordine del software di spedizione, tramite la tabella `xDOTesShippingSoftware` |
| `reconcile_carrier_b.py` | Excel `.xlsx`/`.xls` | numero e data del DDT (`DOTes`, `Cd_DO = 'DDT'`) |

Per ogni spedizione della fattura lo script cerca il documento in ARCA e
inserisce subito sotto una riga con origine `Arca` contenente:

- destinatario, dalla destinazione diversa (`CFDest`) o dall'anagrafica (`CF`);
- peso calcolato come quantità × fattore di conversione × peso lordo
  dell'articolo (`DORig`, `AR`);
- tipo di spedizione (`PACCO` o `PEDANA`), dedotto dal tipo di documento.

Se il documento non viene trovato, la riga ARCA resta vuota. Il risultato è
salvato in `output_finale.xlsx` (CORRIERE_A) o `output_finale_corriere_b.xlsx`
(CORRIERE_B), con le colonne contenenti zeri iniziali formattate come testo.

`xDOTesShippingSoftware` (colonne `Id_DoTes`, `ShippingSoftware_OrderID`,
`ShippingSoftware_TrackingNumber`) è una tabella personalizzata che collega i
documenti di ARCA agli ordini del software di spedizione: il nome della
tabella e delle colonne va adattato alla propria installazione. Anche i nomi
delle colonne della fattura, in `columns_to_read`, dipendono dal tracciato del
corriere.

## Utilizzo

Copiare la fattura nella cartella dello script ed eseguire lo script
corrispondente; viene richiesto il nome del file:

```bash
python reconcile_carrier_a.py
python reconcile_carrier_b.py   # l'estensione del file è facoltativa
```
