"""
Riconciliazione fattura CORRIERE_A (CSV) con le spedizioni registrate in ARCA:
sotto ogni riga della fattura aggiunge la riga calcolata da ARCA (destinatario, peso).
"""
import pandas as pd
import pyodbc
import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


def get_connection_string():
    """Stringa di connessione ODBC ad ARCA, letta dalle variabili d'ambiente (.env)."""
    conn_str = (
        "Driver={SQL Server};"
        f"Server={os.environ['SERVER_DB']};"
        f"Database={os.environ['DB_ARCA']};"
    )
    if os.getenv("DB_USER"):
        return conn_str + f"UID={os.environ['DB_USER']};PWD={os.getenv('DB_PASSWORD', '')};"
    return conn_str + "Trusted_Connection=yes;"


# Funzione per creare la connessione al database SQL Server
def create_db_connection():
    return pyodbc.connect(get_connection_string())

# Funzione per aggiungere una riga vuota dopo ogni riga
def add_empty_rows(dataframe):
    # Creazione di un dataframe vuoto con le stesse colonne
    empty_df = pd.DataFrame(columns=dataframe.columns)
    
    # Iterazione attraverso ogni riga del dataframe
    for index, row in dataframe.iterrows():
        # Aggiungi la riga originale al dataframe vuoto
        empty_df = pd.concat([empty_df, pd.DataFrame([row])], ignore_index=True)
        
        # Aggiungi una riga vuota
        empty_df = pd.concat([empty_df, pd.DataFrame([['']*len(dataframe.columns)], columns=dataframe.columns)], ignore_index=True)
    
    return empty_df


# Funzione per leggere i dati dal CSV selezionando solo le colonne specificate e aggiungendoli a un dataframe
def read_csv_and_add_to_dataframe(csv_file_path, dataframe):
    columns_to_read = [
        'RIFERIM.MITTENTE NUMERICO', 'DATA RITIRO (AMG)', 'RAGIONE SOCIALE DESTINAT.',
        'INDIRIZZO DESTINATARIO', 'LOCALITA DESTINATARIO', 'CAP DESTINATARIO',
        'PROVINCIA DESTINATARIO', 'PESO BOLLETTATO IN KG.', 'PESO RILEVATO IN KG.', 'IMPONIBILE IVA', 'PORTO',
        'SIGLA VARIA 1', 'VARIA 1', 'SIGLA VARIA 2', 'VARIA 2',
        'SIGLA VARIA 3', 'VARIA 3', 'SIGLA VARIA 4', 'VARIA 4',
        'SIGLA VARIA 5', 'VARIA 5', 'SIGLA VARIA 6', 'VARIA 6',
        'SIGLA VARIA 7', 'VARIA 7', 'SIGLA VARIA 8', 'VARIA 8',
        'NUMERO COLLI'
    ]
    csv_data = pd.read_csv(csv_file_path, encoding='ISO-8859-1', delimiter=';', usecols=columns_to_read, dtype={'CAP DESTINATARIO': str})
    csv_data['ORIGINE'] = 'Fattura CORRIERE_A'
    return pd.concat([dataframe, csv_data], ignore_index=True)


# Funzione per esportare il dataframe su Excel mantenendo 'CAP DESTINATARIO' come testo
def export_to_excel_with_text_format(dataframe, file_name):
    # Usa ExcelWriter per specificare un formato di scrittura
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        # Scrivi il dataframe
        dataframe.to_excel(writer, index=False)

        # Ottieni il workbook e il worksheet attuali
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Applica il formato 'testo' solo alla colonna 'CAP DESTINATARIO'
        for col in worksheet['F']:  # Supponendo che 'CAP DESTINATARIO' sia la colonna F
            col.number_format = '@'  # Imposta il formato testo per la colonna
        
# Funzione per eseguire il confronto con il database e generare una riga basata su SQL
def fetch_and_compare_data(cursor, riferimento_mittente):
    # Query per ottenere Id_DoTes da xDOTesShippingSoftware.
    # NB: xDOTesShippingSoftware è una tabella personalizzata che collega i documenti
    # agli ordini del software di spedizione: adattare tabella e colonne alla propria.
    query = f"""
        SELECT xDOTesShippingSoftware.Id_DoTes, xDOTesShippingSoftware.ShippingSoftware_OrderID, DOTes.DataDoc, DOTes.Cd_CF, DOTes.Id_DOTes, ShippingSoftware_TrackingNumber
        FROM dbo.xDOTesShippingSoftware xDOTesShippingSoftware
        JOIN dbo.DOTes DOTes ON xDOTesShippingSoftware.Id_DoTes = DOTes.Id_DoTes
        WHERE xDOTesShippingSoftware.ShippingSoftware_OrderID = '{riferimento_mittente}'
    """
    cursor.execute(query)
    row = cursor.fetchone()

    if not row:
        return None  # Se non ci sono risultati, restituisci None

    id_do_tes = row[0]
    data_doc = row[2]
    cd_cf = row[3]
    cod_verifica = row[4]
    tracking_number = row[5]
    
    # Query per ottenere i dati del destinatario (CFDest)
    query_cfdest = f"""
        SELECT CFDest.Descrizione, CFDest.Indirizzo, CFDest.Localita, CFDest.Cap, CFDest.Cd_Provincia
        FROM dbo.CFDest CFDest
        WHERE CFDest.Cd_CF = '{cd_cf}'
    """
    cursor.execute(query_cfdest)
    cfdest_row = cursor.fetchone()

    # Se non c'è corrispondenza in CFDest, cerca in CF
    if not cfdest_row:
        query_cf = f"""
            SELECT CF.Descrizione, CF.Indirizzo, CF.Localita, CF.Cap, CF.Cd_Provincia
            FROM dbo.CF CF
            WHERE CF.Cd_CF = '{cd_cf}'
        """
        cursor.execute(query_cf)
        cf_row = cursor.fetchone()

        if cf_row:
            descrizione, indirizzo, localita, cap, provincia = cf_row
        else:
            return None  # Se non trovi nulla in CF, esci
    else:
        descrizione, indirizzo, localita, cap, provincia = cfdest_row

    # Query per ottenere il peso degli articoli
    query_dorig = f"""
        SELECT DORig.Cd_AR, DORig.Cd_ARMisura, DORig.FattoreToUM1, AR.PesoLordoMks, DORig.Qta, NumeroDoc, Cd_DO
        FROM dbo.DORig DORig
        JOIN dbo.AR AR ON DORig.Cd_AR = AR.Cd_AR
        WHERE DORig.Id_DoTes = '{id_do_tes}'
    """
    cursor.execute(query_dorig)

    peso_totale = 0
    
    for rig_row in cursor.fetchall():
        # DEBUG PER CAPIRE LA QUERY CHE RIGHE ESTRAE
        # print(f'rig_row: {rig_row}')
        cd_ar, cd_armisura, fattore_um, peso_lordo, quantita, numero_doc, Cd_DO = rig_row[0], rig_row[1], rig_row[2], rig_row[3], rig_row[4], rig_row[5], rig_row[6]
        # if cd_armisura.upper() != 'PZ' and cd_armisura.upper() != 'CA' and cd_armisura.upper() != 'CF':  # Se non è PZ, moltiplica per il fattore
        #     peso_totale += fattore_um * peso_lordo
        # elif cd_armisura.upper() == 'CA' or cd_armisura.upper() == 'CF':      
        #     peso_totale += fattore_um * peso_lordo * quantita    
        # else:
        #     peso_totale += peso_lordo * quantita
        if not cd_ar:
            continue
        
        peso_lordo = (peso_lordo * 1000)
        peso_totale += fattore_um * peso_lordo * quantita

        if Cd_DO.strip() in ["RCW", "FBC", "RCA"]:
            Cd_DO = 'PACCO'
        elif Cd_DO.strip() in ["FTA", "FBB", "FTV"]:
            Cd_DO = 'PEDANA'

    # Converti il peso da grammi a kg e arrotondalo
    peso_totale_kg = round(peso_totale / 1000, 1)

    # Crea la riga con i dati estratti
    nuova_riga = {
        'ORIGINE': 'Arca',
        'RIFERIM.MITTENTE NUMERICO': riferimento_mittente,
        'DATA RITIRO (AMG)': data_doc.strftime("%d/%m/%Y"),
        'RAGIONE SOCIALE DESTINAT.': descrizione,
        'INDIRIZZO DESTINATARIO': indirizzo,
        'LOCALITA DESTINATARIO': localita,
        'CAP DESTINATARIO': cap,
        'PROVINCIA DESTINATARIO': provincia,
        'PESO BOLLETTATO IN KG.': peso_totale_kg,
        'PESO RILEVATO IN KG.': peso_totale_kg,
        'NUMERO DOC': numero_doc,
        'COD VERIFICA': cod_verifica,
        'TIPO ORDINE': Cd_DO,
        'IMPONIBILE IVA': '',
        'PORTO': '',
        'SIGLA VARIA 1': '', 'VARIA 1': '',
        'SIGLA VARIA 2': '', 'VARIA 2': '',
        'SIGLA VARIA 3': '', 'VARIA 3': '',
        'SIGLA VARIA 4': '', 'VARIA 4': '',
        'SIGLA VARIA 5': '', 'VARIA 5': '',
        'SIGLA VARIA 6': '', 'VARIA 6': '',
        'SIGLA VARIA 7': '', 'VARIA 7': '',
        'SIGLA VARIA 8': '', 'VARIA 8': '',
        'NUMERO COLLI': '',
        'TRACKING NUMBER': tracking_number

    }

    return nuova_riga

# Esempio di utilizzo
def main():
    # Creazione di un dataframe vuoto con le colonne che ti interessano, incluso "ORIGINE"
    df = pd.DataFrame(columns=[
        'ORIGINE', 'TIPO ORDINE','RIFERIM.MITTENTE NUMERICO', 'DATA RITIRO (AMG)', 'TRACKING NUMBER', 'RAGIONE SOCIALE DESTINAT.',
        'INDIRIZZO DESTINATARIO', 'LOCALITA DESTINATARIO', 'CAP DESTINATARIO',
        'PROVINCIA DESTINATARIO', 'PESO BOLLETTATO IN KG.', 'PESO RILEVATO IN KG.', 'IMPONIBILE IVA', 'PORTO',
        'SIGLA VARIA 1', 'VARIA 1', 'SIGLA VARIA 2', 'VARIA 2',
        'SIGLA VARIA 3', 'VARIA 3', 'SIGLA VARIA 4', 'VARIA 4',
        'SIGLA VARIA 5', 'VARIA 5', 'SIGLA VARIA 6', 'VARIA 6',
        'SIGLA VARIA 7', 'VARIA 7', 'SIGLA VARIA 8', 'VARIA 8',
        'NUMERO COLLI', 'NUMERO DOC', 'COD VERIFICA'
    ])

    # Leggi i dati dal CSV e aggiungili al dataframe
    folder = os.path.dirname(os.path.abspath(__file__))
    nome_file = input("Inserisci il nome del file CSV: ").strip().strip('"')
    if not nome_file.lower().endswith('.csv'):
        nome_file += '.CSV'
    csv_file = os.path.join(folder, nome_file)
    df = read_csv_and_add_to_dataframe(csv_file, df)

    # Stabilisci la connessione al database
    conn = create_db_connection()
    cursor = conn.cursor()

    # Iterazione su ogni riga del dataframe
    for index, row in df.iterrows():
        riferimento_mittente = row['RIFERIM.MITTENTE NUMERICO']
        
        # Prendi i dati dal database per quel riferimento
        nuova_riga = fetch_and_compare_data(cursor, riferimento_mittente)

        if nuova_riga:
            # Aggiungi la nuova riga sotto la riga corrente
            df = pd.concat([df.iloc[:index*2+1], pd.DataFrame([nuova_riga]), df.iloc[index*2+1:]], ignore_index=True)
        else:
            # Aggiungi una riga vuota se non c'è corrispondenza nel database
            empty_row = pd.DataFrame([[''] * (len(df.columns) - 1) + ['Arca']], columns=df.columns)  # ORIGINE sarà "Arca"
            df = pd.concat([df.iloc[:index*2+1], empty_row, df.iloc[index*2+1:]], ignore_index=True)

    # Esporta il dataframe con i dati aggiornati in un file Excel
    export_to_excel_with_text_format(df, "output_finale.xlsx")
    
    print("Dati esportati con successo in 'output_finale.xlsx'")

if __name__ == "__main__":
    main()