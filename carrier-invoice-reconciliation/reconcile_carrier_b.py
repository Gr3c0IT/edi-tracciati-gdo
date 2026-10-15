"""
Riconciliazione fattura CORRIERE_B (Excel) con i DDT registrati in ARCA:
sotto ogni riga della fattura aggiunge la riga calcolata da ARCA (destinatario, peso).
"""
import pandas as pd
import pyodbc
import os
from datetime import datetime

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
    empty_df = pd.DataFrame(columns=dataframe.columns)
    for index, row in dataframe.iterrows():
        empty_df = pd.concat([empty_df, pd.DataFrame([row])], ignore_index=True)
        empty_df = pd.concat([empty_df, pd.DataFrame([['']*len(dataframe.columns)], columns=dataframe.columns)], ignore_index=True)
    return empty_df


# Funzione per leggere i dati dal file CORRIERE_B (Excel) selezionando solo le colonne che ci servono
def read_corriere_b_file_and_add_to_dataframe(corriere_b_file_path, dataframe):
    columns_to_read = [
        'Data Sped.', 'Numero', 'Fil.', 'Porto', 'Mittente', 'Partenza',
        'Destinatario', 'Indirizzo Destinatario', 'Destinazione', 'ProvDes',
        'Numero DDT', 'Data DDT', 'Colli', 'Peso', 'Bancali', 'Epal',
        'Metri Cubi', 'Tassato', 'Tot Imponibile', 'Tipo Pallet'
    ]

    # NOTA: il file CORRIERE_B è normalmente un Excel (.xlsx o .xls).
    # Se invece il file è un .csv, sostituisci la riga sotto con:
    # corriere_b_data = pd.read_csv(corriere_b_file_path, encoding='ISO-8859-1', delimiter=';', usecols=columns_to_read, dtype={'Numero DDT': str})
    # Scelgo l'engine in base all'estensione reale del file:
    # .xlsx/.xlsm -> openpyxl, .xls (vecchio formato binario) -> xlrd (richiede 'pip install xlrd')
    ext = os.path.splitext(corriere_b_file_path)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

    # Leggo TUTTE le colonne (senza usecols) e ripulisco gli header da spazi spuri,
    # perché nel file CORRIERE_B alcune intestazioni hanno spazi nascosti (es. ' Indirizzo Destinatario').
    # Così evitiamo che un piccolo refuso nel file sorgente rompa lo script.
    corriere_b_data_full = pd.read_excel(corriere_b_file_path, dtype={'Numero DDT': str}, engine=engine)
    corriere_b_data_full.columns = [str(c).strip() for c in corriere_b_data_full.columns]

    colonne_indispensabili = [
        'Numero DDT', 'Data DDT', 'Destinatario', 'Indirizzo Destinatario',
        'Destinazione', 'ProvDes', 'Peso'
    ]
    mancanti_gravi = [c for c in colonne_indispensabili if c not in corriere_b_data_full.columns]
    if mancanti_gravi:
        raise ValueError(
            f"Colonne indispensabili non trovate nel file CORRIERE_B: {mancanti_gravi}\n"
            f"Colonne disponibili nel file: {list(corriere_b_data_full.columns)}"
        )

    # Le colonne accessorie possono mancare da un file all'altro (CORRIERE_B non è sempre coerente):
    # se assenti, le creo vuote invece di bloccare lo script.
    for colonna in columns_to_read:
        if colonna not in corriere_b_data_full.columns:
            corriere_b_data_full[colonna] = ''

    corriere_b_data = corriere_b_data_full[columns_to_read].copy()

    # Normalizzo le date da stringa/Timestamp a oggetto date puro, utile per la query parametrizzata
    corriere_b_data['Data DDT'] = pd.to_datetime(corriere_b_data['Data DDT'], dayfirst=True).dt.date

    # Rinomino le colonne CORRIERE_B che hanno un equivalente diretto nel dato Arca,
    # così finiscono nella STESSA colonna del foglio finale (confronto verticale)
    rename_map = {
        'Numero DDT': 'NUMERO DDT',
        'Data DDT': 'DATA DDT',
        'Destinatario': 'RAGIONE SOCIALE DESTINAT.',
        'Indirizzo Destinatario': 'INDIRIZZO DESTINATARIO',
        'Destinazione': 'LOCALITA DESTINATARIO',
        'ProvDes': 'PROVINCIA DESTINATARIO',
        'Peso': 'PESO (kg)',
    }
    corriere_b_data = corriere_b_data.rename(columns=rename_map)

    corriere_b_data['ORIGINE'] = 'Fattura CORRIERE_B'
    return pd.concat([dataframe, corriere_b_data], ignore_index=True)


# Funzione per esportare il dataframe su Excel mantenendo colonne "sensibili" come testo
def export_to_excel_with_text_format(dataframe, file_name):
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        # Colonna 'Numero DDT' come testo, per non perdere eventuali zeri iniziali
        for col in worksheet['B']:  # verificare l'indice colonna corretto in base all'ordine finale
            col.number_format = '@'


# Funzione per eseguire il confronto con il database e generare una riga basata su SQL
def fetch_and_compare_data(cursor, numero_ddt, data_ddt):
    numero_ddt_pulito = str(numero_ddt).strip()

    # Aggancio: Cd_DO = 'DDT' + Data documento (oggetto date, gestita dal driver) + NumeroDoc "ripulito" da spazi
    query = """
        SELECT Id_DoTes, DataDoc, Cd_CF, NumeroDoc
        FROM dbo.DOTes
        WHERE Cd_DO = 'DDT'
          AND CAST(DataDoc AS DATE) = ?
          AND LTRIM(RTRIM(NumeroDoc)) = ?
    """
    data_ddt_param = data_ddt.strftime('%Y-%m-%d') if hasattr(data_ddt, 'strftime') else str(data_ddt)
    cursor.execute(query, (data_ddt_param, numero_ddt_pulito))
    row = cursor.fetchone()

    if not row:
        return None  # Nessuna corrispondenza trovata in DOTes

    id_do_tes = row[0]
    data_doc = row[1]
    cd_cf = row[2]
    numero_doc_db = row[3]

    # Dati destinatario: prima CFDest, poi fallback su CF (stessa logica dello script CORRIERE_A)
    query_cfdest = """
        SELECT Descrizione, Indirizzo, Localita, Cap, Cd_Provincia
        FROM dbo.CFDest
        WHERE Cd_CF = ?
    """
    cursor.execute(query_cfdest, (cd_cf,))
    cfdest_row = cursor.fetchone()

    if not cfdest_row:
        query_cf = """
            SELECT Descrizione, Indirizzo, Localita, Cap, Cd_Provincia
            FROM dbo.CF
            WHERE Cd_CF = ?
        """
        cursor.execute(query_cf, (cd_cf,))
        cf_row = cursor.fetchone()
        if cf_row:
            descrizione, indirizzo, localita, cap, provincia = cf_row
        else:
            return None
    else:
        descrizione, indirizzo, localita, cap, provincia = cfdest_row

    # Peso ricalcolato da DORig/AR, stessa identica logica dello script CORRIERE_A
    query_dorig = """
        SELECT DORig.Cd_AR, DORig.Cd_ARMisura, DORig.FattoreToUM1, AR.PesoLordoMks, DORig.Qta, Cd_DO
        FROM dbo.DORig DORig
        JOIN dbo.AR AR ON DORig.Cd_AR = AR.Cd_AR
        WHERE DORig.Id_DoTes = ?
    """
    cursor.execute(query_dorig, (id_do_tes,))

    peso_totale = 0
    tipo_ordine = ''

    for rig_row in cursor.fetchall():
        cd_ar, cd_armisura, fattore_um, peso_lordo, quantita, Cd_DO = rig_row[0], rig_row[1], rig_row[2], rig_row[3], rig_row[4], rig_row[5]

        if not cd_ar:
            continue

        peso_lordo = peso_lordo * 1000
        peso_totale += fattore_um * peso_lordo * quantita

        if Cd_DO and Cd_DO.strip() in ["RCW", "FBC", "RCA"]:
            tipo_ordine = 'PACCO'
        elif Cd_DO and Cd_DO.strip() in ["FTA", "FBB", "FTV"]:
            tipo_ordine = 'PEDANA'

    peso_totale_kg = round(peso_totale / 1000, 1)

    nuova_riga = {
        'ORIGINE': 'Arca',
        'TIPO ORDINE': tipo_ordine,
        'NUMERO DDT': str(numero_doc_db).strip(),
        'DATA DDT': data_doc.date() if hasattr(data_doc, 'date') else data_doc,
        'RAGIONE SOCIALE DESTINAT.': descrizione,
        'INDIRIZZO DESTINATARIO': indirizzo,
        'LOCALITA DESTINATARIO': localita,
        'CAP DESTINATARIO': cap,
        'PROVINCIA DESTINATARIO': provincia,
        'PESO (kg)': peso_totale_kg,
        'COD VERIFICA': id_do_tes,
    }

    return nuova_riga


def main():
    df = pd.DataFrame(columns=[
        'ORIGINE', 'TIPO ORDINE', 'NUMERO DDT', 'DATA DDT',
        'RAGIONE SOCIALE DESTINAT.', 'INDIRIZZO DESTINATARIO', 'LOCALITA DESTINATARIO',
        'CAP DESTINATARIO', 'PROVINCIA DESTINATARIO', 'PESO (kg)', 'COD VERIFICA',
        'Data Sped.', 'Numero', 'Fil.', 'Porto', 'Mittente', 'Partenza',
        'Colli', 'Bancali', 'Epal', 'Metri Cubi', 'Tassato', 'Tot Imponibile', 'Tipo Pallet'
    ])

    folder = os.path.dirname(os.path.abspath(__file__))
    nome_file = input("Inserisci il nome del file CORRIERE_B: ").strip().strip('"')

    if nome_file.lower().endswith(('.xlsx', '.xls', '.xlsm')):
        corriere_b_file = os.path.join(folder, nome_file)
    else:
        # Nessuna estensione indicata: provo in ordine .xlsx, poi .xls, poi .xlsm
        corriere_b_file = None
        for possibile_ext in ('.xlsx', '.xls', '.xlsm'):
            candidato = os.path.join(folder, nome_file + possibile_ext)
            if os.path.isfile(candidato):
                corriere_b_file = candidato
                break
        if corriere_b_file is None:
            raise FileNotFoundError(
                f"Nessun file '{nome_file}' con estensione .xlsx/.xls/.xlsm trovato in {folder}"
            )

    df = read_corriere_b_file_and_add_to_dataframe(corriere_b_file, df)

    conn = create_db_connection()
    cursor = conn.cursor()

    for index, row in df.iterrows():
        numero_ddt = row['NUMERO DDT']
        data_ddt = row['DATA DDT']

        if pd.isna(numero_ddt) or pd.isna(data_ddt):
            nuova_riga = None
        else:
            nuova_riga = fetch_and_compare_data(cursor, numero_ddt, data_ddt)

        if nuova_riga:
            df = pd.concat([df.iloc[:index*2+1], pd.DataFrame([nuova_riga]), df.iloc[index*2+1:]], ignore_index=True)
        else:
            empty_row = pd.DataFrame([[''] * (len(df.columns) - 1) + ['Arca']], columns=df.columns)
            df = pd.concat([df.iloc[:index*2+1], empty_row, df.iloc[index*2+1:]], ignore_index=True)

    # Formatto la colonna DATA DDT (oggetti date, sia CORRIERE_B che Arca) come stringa dd/mm/yyyy
    # solo ora, dopo aver finito i confronti, così il confronto verticale è leggibile
    df['DATA DDT'] = df['DATA DDT'].apply(
        lambda d: d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') and not pd.isna(d) else d
    )

    export_to_excel_with_text_format(df, "output_finale_corriere_b.xlsx")
    print("Dati esportati con successo in 'output_finale_corriere_b.xlsx'")


if __name__ == "__main__":
    main()