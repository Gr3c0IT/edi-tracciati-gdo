#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script per l'estrazione del tracciato TRACCIATO_GDO multi-cliente con invio automatico via Gmail
Supporta 5 clienti: CLIENTE_A, CLIENTE_B, CLIENTE_D, CLIENTE_C, CLIENTE_E
"""

import pandas as pd
import urllib.parse
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta
import calendar
import os
import logging
import pickle
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import locale

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


locale.setlocale(locale.LC_TIME, 'Italian_Italy.1252')

# Configurazione logging (per Windows)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tracciato_gdo_extraction.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class ClientManager:
    """Gestione informazioni clienti TRACCIATO_GDO"""
   
    def __init__(self):
        self.clienti = {
            'C000001': {
                'codice': 'C000001',
                'nome': 'CLIENTE_A',
                'email': 'collega_a@azienda.example.com',  # in produzione: referente_a@cliente-a.example.com
                'cc': ['collega_a@azienda.example.com'],
                'nome_breve': 'CLIENTE_A'
            },
            'C000002': {
                'codice': 'C000002',
                'nome': 'CLIENTE_B',
                'email': 'collega_a@azienda.example.com',  # in produzione: referente_b@cliente-b.example.com, referente_c@cliente-b.example.com
                'cc': ['collega_a@azienda.example.com'],
                'nome_breve': 'CLIENTE_B'
            },
            'C000003': {
                'codice': 'C000003',
                'nome': 'CLIENTE_C',
                'email': 'collega_a@azienda.example.com',  # in produzione: referente_d@cliente-d.example.com
                'cc': ['collega_a@azienda.example.com'],
                'nome_breve': 'CLIENTE_C',
                # Configurazioni tracciato
                'ddt_headers': True,  # Gestione testate DDT da righe "Rif: DDT"
                'tipo_doc_format': 'alpha',  # 'F', 'N'
                'include_valuta': True,
                'prezzo_al_pezzo': True,  # Divisione prezzi per FattoreToUM1
                'qta_in_pezzi': True,  # Moltiplicazione Qta per FattoreToUM1
                'um_fissa': 'PZ',  # Unità di misura fissa
                'rimuovi_trattini': True  # Rimuove trattini dalla descrizione
            },
            'C000004': {
                'codice': 'C000004',
                'nome': 'CLIENTE_D',
                'email': 'collega_a@azienda.example.com',  # in produzione: referente_e@cliente-d.example.com
                'cc': ['collega_a@azienda.example.com'],
                'nome_breve': 'CLIENTE_D',
                # Configurazioni tracciato
                'ddt_headers': True,  # Gestione testate DDT
                'tipo_doc_format': 'numeric',  # '0', '1'
                'include_valuta': False,
                'prezzo_al_pezzo': False,
                'qta_in_pezzi': False,
                'um_fissa': None,
                'rimuovi_trattini': False
            },
            'C000005': {
                'codice': 'C000005',
                'nome': 'CLIENTE_E',
                'email': 'collega_a@azienda.example.com',
                'cc': ['collega_a@azienda.example.com'],
                'nome_breve': 'CLIENTE_E',
                # Configurazioni tracciato (standard TRACCIATO_GDO Euro, come CLIENTE_A)
                'ddt_headers': True,  # Gestione testate DDT da righe "Rif: DDT" (come CLIENTE_D)
                'tipo_doc_format': 'alpha',  # 'F', 'N'
                'include_valuta': True,  # EUR in pos. 95-97
                'prezzo_al_pezzo': False,
                'qta_in_pezzi': False,
                'um_fissa': None,
                'rimuovi_trattini': False
            }
        }
    
    def get_cliente(self, codice):
        """Restituisce informazioni cliente per codice"""
        return self.clienti.get(codice)
    
    def get_all_clienti(self):
        """Restituisce tutti i clienti"""
        return self.clienti
    
    def list_clienti(self):
        """Stampa lista clienti disponibili con indirizzi CC"""
        print("\n" + "="*60)
        print("CLIENTI DISPONIBILI:")
        print("="*60)
        for codice, info in self.clienti.items():
            print(f"{codice} - {info['nome']}")
            print(f"    Email: {info['email']}")
            if info.get('cc'):
                print(f"    CC: {', '.join(info['cc'])}")
            else:
                print(f"    CC: Nessuno")
            print("-" * 60)
    
    def get_config(self, codice_cliente, key, default=None):
        """Recupera una configurazione specifica per un cliente"""
        return self.clienti.get(codice_cliente, {}).get(key, default)
    
    def requires_ddt_headers(self, codice_cliente):
        """Verifica se il cliente richiede gestione testate DDT"""
        return self.get_config(codice_cliente, 'ddt_headers', False)

class GmailSender:
    """Gestione invio email via Gmail API"""
    
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.send']
        self.token_file = os.getenv('GMAIL_TOKEN_FILE', 'token.pickle')
        self.credentials_file = os.getenv('GMAIL_CREDENTIALS_FILE', 'credentials.json')
        self.service = None
        
    def authenticate(self):
        """Autentica con Gmail API usando credentials.json e token.pickle"""
        try:
            creds = None
            
            # Carica token esistente
            if os.path.exists(self.token_file):
                with open(self.token_file, 'rb') as token:
                    creds = pickle.load(token)
            
            # Se non ci sono credenziali valide, richiedi autorizzazione
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    if not os.path.exists(self.credentials_file):
                        logging.error("[ERROR] File credentials.json non trovato!")
                        return False
                    
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.credentials_file, self.SCOPES)
                    creds = flow.run_local_server(port=0)
                
                # Salva token per la prossima volta
                with open(self.token_file, 'wb') as token:
                    pickle.dump(creds, token)
            
            # Costruisci servizio Gmail
            self.service = build('gmail', 'v1', credentials=creds)
            logging.info("[OK] Autenticazione Gmail completata")
            return True
            
        except Exception as e:
            logging.error(f"[ERROR] Errore nell'autenticazione Gmail: {str(e)}")
            return False
    
    def create_message_with_attachment(self, to, subject, message_text, attachment_path, cc=None):
        """Crea messaggio email con allegato"""
        try:
            message = MIMEMultipart()
            message['to'] = to
            message['subject'] = subject
            
            if cc and len(cc) > 0:
                message['cc'] = ', '.join(cc)
                logging.info(f"[EMAIL] CC aggiunti: {', '.join(cc)}")

            # Corpo del messaggio
            message.attach(MIMEText(message_text, 'html'))
            
            # Aggiungi allegato
            if os.path.exists(attachment_path):
                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                filename = os.path.basename(attachment_path)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {filename}',
                )
                message.attach(part)
            
            # Codifica messaggio
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
            return {'raw': raw_message}
            
        except Exception as e:
            logging.error(f"[ERROR] Errore nella creazione messaggio: {str(e)}")
            return None
    
    def send_message(self, message):
        """Invia messaggio via Gmail API"""
        try:
            result = self.service.users().messages().send(userId="me", body=message).execute()
            logging.info(f"[OK] Email inviata. ID messaggio: {result['id']}")
            return True
        except Exception as e:
            logging.error(f"[ERROR] Errore nell'invio email: {str(e)}")
            return False

class TracciatoGdoExtractorMultiClient:
    """Classe principale per l'estrazione TRACCIATO_GDO multi-cliente"""
    
    def __init__(self):
        """Inizializza il sistema"""
        self.engine = None
        self.output_folder = "output_tracciato_gdo"
        self.client_manager = ClientManager()
        self.gmail_sender = GmailSender()
        self.create_output_folder()

    def _insert_ddt_headers_df(self, df, codice_cliente):
        """
        Inserisce testate DDT nei punti corretti lavorando su DataFrame
        """
        import re
        
        processed_rows = []
        pending_headers = []
        last_header_by_doc = {}
        first_movement_seen = {}  # Traccia se abbiamo già visto un movimento con Cd_AR per questo documento
        
        for idx, row in df.iterrows():
            id_dotes = row['Id_DOTes']
            tipo_rec = row['TipoRec']
            riga_tracciato = row['RigaTracciato']
            
            if tipo_rec == '01':
                # Testata - memorizza e aggiungi
                last_header_by_doc[id_dotes] = row.to_dict()
                processed_rows.append(row.to_dict())
                
            elif tipo_rec == '02':
                # Movimento - verifica se è "Rif: DDT" o articolo normale
                cd_ar = riga_tracciato[7:22].strip()  # Posizione Cd_AR nel tracciato
                
                if not cd_ar:  # Nessun codice articolo
                    descrizione_completa = riga_tracciato[22:52].strip()  # Posizione Descrizione
                    
                    # Check case-insensitive per "Rif: DDT" o "rif: ddt"
                    if descrizione_completa.lower().startswith('rif: ddt'):
                        # ========== LOGICA ANTI-DUPLICATO ==========
                        # Se questo è il PRIMO movimento del documento (non abbiamo ancora visto articoli con Cd_AR)
                        # allora NON creare la testata (c'è già quella originale dalla query)
                        if id_dotes not in first_movement_seen:
                            # Primo movimento è "Rif: DDT" → SALTA creazione testata
                            continue  # Skip sia la testata che la riga
                        # ==========================================
                        
                        # Se arriviamo qui, abbiamo già visto articoli con Cd_AR
                        # quindi possiamo creare la nuova testata DDT
                        if id_dotes in last_header_by_doc:
                            new_header = self._create_ddt_header_from_dict(
                                last_header_by_doc[id_dotes],
                                descrizione_completa,
                                codice_cliente
                            )
                            pending_headers.append(new_header)
                        # NON includere questa riga nei risultati
                        continue
                
                # ========== TRACCIAMENTO PRIMO MOVIMENTO CON Cd_AR ==========
                # Marca che abbiamo visto almeno un movimento con Cd_AR per questo documento
                if cd_ar and id_dotes not in first_movement_seen:
                    first_movement_seen[id_dotes] = True
                # ============================================================
                
                # Articolo normale con Cd_AR
                # Prima aggiungi eventuali testate DDT in sospeso
                if pending_headers:
                    processed_rows.extend(pending_headers)
                    pending_headers = []
                
                # Poi aggiungi l'articolo
                processed_rows.append(row.to_dict())
        
        # NON aggiungere testate rimaste in coda - non hanno articoli dopo
        if pending_headers:
            import logging
            for header in pending_headers:
                num_ddt = header['RigaTracciato'][19:25].strip()
                logging.warning(f"[DDT] Testata DDT {num_ddt} scartata: nessun articolo dopo questa riga 'Rif: DDT'")
        
        # Converti lista di dict in DataFrame
        result_df = pd.DataFrame(processed_rows)
        result_df.reset_index(drop=True, inplace=True)
        
        return result_df

    def _create_ddt_header_from_dict(self, original_header, descrizione_ddt, codice_cliente):
        """
        Crea una nuova testata basata su quella originale,
        modificando DDT e recuperando xPV dal DDT originale
        """
        import re
        import logging
        import pandas as pd
        
        id_dotes = original_header['Id_DOTes']
        riga_originale = original_header['RigaTracciato']
        
        # ========== ESTRAZIONE NUMERO E DATA DDT ==========
        match_num = re.search(
            r'(?:rif|Rif)[\s:]*(?:ddt|DDT)[\s:]*(\d+)',
            descrizione_ddt,
            re.IGNORECASE
        )
        num_ddt = match_num.group(1) if match_num else ''
        
        match_data = re.search(
            r'(?:del|DEL)[\s:]*(\d{1,2})[-/]?(\d{1,2})[-/]?(\d{4})',
            descrizione_ddt,
            re.IGNORECASE
        )
        
        if match_data:
            gg, mm, aaaa = match_data.groups()
            gg = gg.zfill(2)
            mm = mm.zfill(2)
            data_ddt_formatted = f"{aaaa[-2:]}{mm}{gg}"
            data_ddt_completa = f"{aaaa}-{mm}-{gg}"
        else:
            data_ddt_formatted = '      '
            data_ddt_completa = None
        
# ========== RECUPERO xPV DAL DDT ORIGINALE (TENTATIVO 1: con TRIM) ==========
        xpv_value = ''
        
        if num_ddt and data_ddt_completa:
            try:
                # TENTATIVO 1: Query con TRIM per rimuovere spazi da NumeroDoc
                query_ddt = f"""
                SELECT TOP 1 t.Cd_CFDest
                FROM [dbo].DOTes t
                WHERE LTRIM(RTRIM(t.NumeroDoc)) = '{num_ddt}'
                AND t.Cd_DO = 'DDT'
                AND t.Cd_CF = '{codice_cliente}'
                AND CAST(t.DataDoc AS DATE) = '{data_ddt_completa}'
                """
                
                df_ddt = pd.read_sql_query(query_ddt, self.engine)
                
                if not df_ddt.empty and df_ddt.iloc[0]['Cd_CFDest'] is not None:
                    cd_cfdest = df_ddt.iloc[0]['Cd_CFDest']
                    
                    query_xpv = f"""
                    SELECT TOP 1 dest.xPV
                    FROM [dbo].CFDest dest
                    WHERE dest.Cd_CFDest = '{cd_cfdest}'
                    AND dest.Cd_CF = '{codice_cliente}'
                    AND dest.xPV IS NOT NULL
                    """
                    
                    df_xpv = pd.read_sql_query(query_xpv, self.engine)
                    
                    if not df_xpv.empty and df_xpv.iloc[0]['xPV'] is not None:
                        xpv_value = str(df_xpv.iloc[0]['xPV']).strip()
                        #logging.info(f"[DDT] ✓ xPV recuperato per DDT {num_ddt} (via NumeroDoc+TRIM): '{xpv_value}'")
                    else:
                        logging.warning(f"[DDT] xPV non trovato per Cd_CFDest={cd_cfdest}")
                else:
                    logging.warning(f"[DDT] DDT originale non trovato con TRIM: NumDoc={num_ddt}, Data={data_ddt_completa}")
                    
            except Exception as e:
                logging.error(f"[DDT] Errore recupero xPV (tentativo 1 - TRIM): {e}")
        
        # ========== FALLBACK: ESTRAI Dest. DALLA DESCRIZIONE ==========
# ========== FALLBACK: ESTRAI Dest. DALLA DESCRIZIONE ==========
        if not xpv_value:  # Se xPV ancora vuoto, prova con Dest.
            try:
                logging.info(f"[DDT] → Tentativo fallback: estrazione Dest. dalla descrizione")
                
                # RECUPERA DESCRIZIONE ORIGINALE COMPLETA dal database (non troncata)
                try:
                    query_desc = f"""
                    SELECT TOP 1 r.Descrizione
                    FROM [dbo].DORig r
                    WHERE r.Id_DOTes = {id_dotes}
                    AND r.Descrizione LIKE 'Rif: DDT {num_ddt} %'
                    """
                    df_desc = pd.read_sql_query(query_desc, self.engine)
                    
                    if not df_desc.empty and df_desc.iloc[0]['Descrizione'] is not None:
                        descrizione_ddt = df_desc.iloc[0]['Descrizione']
                        logging.info(f"[DDT] Descrizione completa recuperata: '{descrizione_ddt}'")
                except Exception as e:
                    logging.warning(f"[DDT] Errore recupero descrizione completa: {e}")
                
                # Cerca pattern "Dest. XXX" nella descrizione
                match_dest = re.search(
                    r'(?:Dest|dest)[\s.:]*(\d+)',
                    descrizione_ddt,
                    re.IGNORECASE
                )
                
                if match_dest:
                    cd_cfdest = match_dest.group(1).zfill(3)  # es. "1" -> "001"
                    logging.info(f"[DDT] Dest. estratto dalla descrizione: '{cd_cfdest}'")
                    
                    query_xpv_fallback = f"""
                    SELECT TOP 1 dest.xPV
                    FROM [dbo].CFDest dest
                    WHERE dest.Cd_CFDest = '{cd_cfdest}'
                    AND dest.Cd_CF = '{codice_cliente}'
                    AND dest.xPV IS NOT NULL
                    """
                    
                    df_xpv_fallback = pd.read_sql_query(query_xpv_fallback, self.engine)
                    
                    if not df_xpv_fallback.empty and df_xpv_fallback.iloc[0]['xPV'] is not None:
                        xpv_value = str(df_xpv_fallback.iloc[0]['xPV']).strip()
                        logging.info(f"[DDT] ✓ xPV recuperato via Dest. (FALLBACK): '{xpv_value}'")
                    else:
                        logging.warning(f"[DDT] xPV non trovato per Dest={cd_cfdest} (fallback)")
                else:
                    logging.warning(f"[DDT] Dest. non trovato nella descrizione: '{descrizione_ddt}'")
                    
            except Exception as e:
                logging.error(f"[DDT] Errore recupero xPV (fallback Dest.): {e}")
        
        # ========== RICOSTRUISCI TESTATA ==========
        # NON copiare xPV dalla testata originale, usa SOLO quello recuperato
        nuova_riga = (
            riga_originale[:13] +      # 01 + Id + NumDoc  
            riga_originale[13:19] +    # DataDoc
            num_ddt.rjust(6) +         # NumDDT
            data_ddt_formatted +       # DataDDT
            riga_originale[31:77] +    # Tutto tranne xPV (spazi, codici, ecc)
            xpv_value.rjust(15) +      # xPV NUOVO recuperato dal DDT specifico
            riga_originale[92:]        # Resto (tipo doc, valuta, ecc)
        )
        
        # ========== DEBUG ==========
        logging.info(f"[DDT] Testata creata - DDT={num_ddt}, xPV='{xpv_value}'")
        
        return {
            'Id_DOTes': id_dotes,
            'TipoRec': '01',
            'RigaTracciato': nuova_riga
        }

    def create_output_folder(self):
        """Crea la cartella di output se non esiste"""
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
            logging.info(f"Creata cartella: {self.output_folder}")
    
    def calculate_previous_month_period(self):
        """Calcola il periodo dal 7 del mese precedente al 7 del mese corrente"""
        oggi = datetime.now()
        
        # DEBUG: verifica data corrente del sistema
        logging.info(f"[DEBUG] Data corrente sistema: {oggi.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 7 del mese corrente
        settimo_mese_corrente = oggi.replace(day=7)
        logging.info(f"[DEBUG] 7 del mese corrente: {settimo_mese_corrente.strftime('%Y-%m-%d')}")
        
        # 7 del mese precedente
        if oggi.month == 1:  # Se siamo a gennaio, mese precedente è dicembre dell'anno prima
            settimo_mese_precedente = oggi.replace(year=oggi.year-1, month=12, day=7)
        else:
            settimo_mese_precedente = oggi.replace(month=oggi.month-1, day=7)
        
        logging.info(f"[DEBUG] 7 del mese precedente: {settimo_mese_precedente.strftime('%Y-%m-%d')}")
        
        # Formattiamo per SQL Server (YYYY-MM-DD)

        # # Parte di codice per scaricare tracciato_gdo in base alla data impostata
        data_inizio = '2026-09-01'
        data_fine = '2026-09-30'
        nome_periodo = f'dal {data_inizio} al {data_fine}'
        #Dopo aver attivato le variabili data_inizio, data_fine, nome_periodo, ricordati di commentare queste sotto.
        

        #data_inizio = settimo_mese_precedente.strftime('%Y-%m-%d')
        #data_fine = settimo_mese_corrente.strftime('%Y-%m-%d')
        #nome_periodo = f"dal {settimo_mese_precedente.strftime('%d %B %Y')} al {settimo_mese_corrente.strftime('%d %B %Y')}"
        
        logging.info(f"[PERIOD] Periodo estrazione: {data_inizio} - {data_fine}")
        logging.info(f"[PERIOD] Periodo: {nome_periodo}")
        
        return data_inizio, data_fine, nome_periodo
    
    def connect_database(self):
        """Stabilisce la connessione al database SQL Server"""
        try:
            params = urllib.parse.quote_plus(get_connection_string())
            self.engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % params)
            
            # Test connessione
            with self.engine.connect() as conn:
                result = conn.execute(text("SELECT 1"))
                logging.info("[OK] Connessione al database stabilita con successo")
                return True
                
        except Exception as e:
            logging.error(f"[ERROR] Errore nella connessione al database: {str(e)}")
            return False
    
    def get_sql_query(self, codice_cliente, data_inizio, data_fine, id_dotes_filter=None):
            
            """Restituisce la query SQL per il tracciato TRACCIATO_GDO di un cliente specifico nel periodo specificato"""

            if id_dotes_filter:
                if isinstance(id_dotes_filter, (list, tuple, set)):
                    ids_str = ",".join(str(x) for x in id_dotes_filter)
                    filtro_dotes = f"AND t.Id_DOTes IN ({ids_str})"
                else:
                    filtro_dotes = f"AND t.Id_DOTes = {id_dotes_filter}"
            else:
                filtro_dotes = ""
            
            # Aggiungi 'DDT' se c'è filtro id_dotes oppure per C000006
            ddt_condition = ", 'DDT'" if (id_dotes_filter or codice_cliente == 'C000006') else ""
            return f"""
            -- QUERY PER TRACCIATO TRACCIATO_GDO - Cliente: {codice_cliente}
            -- Periodo: {data_inizio} - {data_fine}
            -- Database: [DB_ARCA].[dbo]
            -- Filtri: Cd_Do IN ('FTV', 'FT2', 'NCC') AND Cd_CF = '{codice_cliente}' AND DataDoc nel mese precedente

            -- =====================================================
            -- RECORD TESTATA (Tipo 01) - 128 caratteri fissi
            -- =====================================================
            SELECT 
                t.Id_DOTes,
                CAST(0 AS INT) AS Id_DORig,
                '01' AS TipoRec,
                '01' +                                                                            
                RIGHT('00000' + CAST(t.Id_DOTes AS VARCHAR(5)), 5) +                          
                RIGHT('      ' + ISNULL(CAST(t.NumeroDoc AS VARCHAR(6)), ''), 6) +            
                RIGHT('0' + CAST(YEAR(t.DataDoc) % 100 AS VARCHAR(2)), 2) + 
                RIGHT('0' + CAST(MONTH(t.DataDoc) AS VARCHAR(2)), 2) + 
                RIGHT('0' + CAST(DAY(t.DataDoc) AS VARCHAR(2)), 2) +                          
                RIGHT('      ' + ISNULL(bolla.XDDT, ''), 6) +                                
                CASE 
                    WHEN bolla.XDATADDT IS NOT NULL THEN 
                        RIGHT('0' + CAST(YEAR(bolla.XDATADDT) % 100 AS VARCHAR(2)), 2) + 
                        RIGHT('0' + CAST(MONTH(bolla.XDATADDT) AS VARCHAR(2)), 2) + 
                        RIGHT('0' + CAST(DAY(bolla.XDATADDT) AS VARCHAR(2)), 2)
                    ELSE '      '
                END +                                                                                                                                           
                '               ' +                                                             
                ' ' +                                                                           
                '               ' +                                                             
                '               ' +                                                             
                RIGHT('               ' + ISNULL(LTRIM(RTRIM(xpv.xPV)), ''), 15) +  -- xPV nella testata (trim spazi)                     
                ' ' +                                                                           
                CASE 
                    WHEN '{codice_cliente}' IN ('C000004','C000003') THEN  
                        CASE 
                            WHEN t.Cd_Do = 'NCC' THEN '1'  
                            WHEN t.Cd_Do IN ('FTV', 'FT2') THEN '0'  
                            ELSE '0'
                        END
                    ELSE  
                        CASE 
                            WHEN t.Cd_Do = 'NCC' THEN 'N'  
                            WHEN t.Cd_Do IN ('FTV', 'FT2') THEN 'F'  
                            ELSE 'F'
                        END
                END + 
                CASE 
                    WHEN '{codice_cliente}' IN ('C000004','C000003') THEN '   '  
                    ELSE RIGHT('   ' + ISNULL(t.Cd_VL, 'EUR'), 3)   
                END + 
                '                               '                                               
                AS RigaTracciato

            FROM [dbo].DOTes t
            LEFT JOIN (
                SELECT Id_DOTes, xPV
                FROM (
                    SELECT 
                        t2.Id_DOTes,
                        dest.xPV,
                        ROW_NUMBER() OVER (
                            PARTITION BY t2.Id_DOTes 
                            ORDER BY ddt_orig.Id_DOTes DESC
                        ) AS rn
                    FROM [dbo].DOTes t2
                    -- Trova numero DDT dalla subquery bolla
                    LEFT JOIN (
                        SELECT 
                            r.Id_DOTes,
                            COALESCE(
                                MIN(NULLIF(r.XDDT, '')),
                                (
                                    SELECT TOP 1 
                                        SUBSTRING(r2.Descrizione,
                                            PATINDEX('%ddt%', LOWER(r2.Descrizione)) + 4,
                                            CHARINDEX(' ', r2.Descrizione + ' ', PATINDEX('%ddt%', LOWER(r2.Descrizione)) + 4) 
                                                - (PATINDEX('%ddt%', LOWER(r2.Descrizione)) + 4)
                                        )
                                    FROM [dbo].DORig r2
                                    WHERE r2.Id_DOTes = r.Id_DOTes
                                    AND LOWER(r2.Descrizione) LIKE '%ddt%'
                                )
                            ) AS NumDDT,
                            COALESCE(
                                MIN(r.XDATADDT),
                                (
                                    SELECT TOP 1 
                                        TRY_CAST(
                                            SUBSTRING(r2.Descrizione,
                                                PATINDEX('%del%', LOWER(r2.Descrizione)) + 4,
                                                10
                                            ) AS DATE
                                        )
                                    FROM [dbo].DORig r2
                                    WHERE r2.Id_DOTes = r.Id_DOTes
                                    AND LOWER(r2.Descrizione) LIKE '%ddt%del%'
                                )
                            ) AS DataDDT
                        FROM [dbo].DORig r
                        WHERE r.Cd_CF = '{codice_cliente}'
                        GROUP BY r.Id_DOTes
                    ) ddt_info ON t2.Id_DOTes = ddt_info.Id_DOTes
                    -- Cerca DDT originale per ottenere Cd_CFDest (MATCH ESATTO, non più LIKE)
                    LEFT JOIN [dbo].DOTes ddt_orig ON 
                        LTRIM(RTRIM(ddt_orig.NumeroDoc)) = LTRIM(RTRIM(ISNULL(ddt_info.NumDDT, '')))
                        AND ddt_info.NumDDT IS NOT NULL
                        AND ddt_info.NumDDT != ''
                        AND ddt_orig.Cd_DO = 'DDT'
                        AND ddt_orig.Cd_CF = '{codice_cliente}'
                        AND CAST(ddt_orig.DataDoc AS DATE) = CAST(ddt_info.DataDDT AS DATE)
                    -- Ottieni xPV da CFDest
                    LEFT JOIN [dbo].CFDest dest ON 
                        COALESCE(t2.Cd_CFDest, ddt_orig.Cd_CFDest) = dest.Cd_CFDest
                        AND dest.Cd_CF = '{codice_cliente}'
                        AND dest.xPV IS NOT NULL
                    WHERE t2.Cd_CF = '{codice_cliente}'
                ) sub
                WHERE rn = 1
            ) xpv ON t.Id_DOTes = xpv.Id_DOTes
            LEFT JOIN (
                SELECT 
                    r.Id_DOTes,
                    COALESCE(
                        MIN(NULLIF(r.XDDT, '')),
                        (
                            SELECT TOP 1 
                                SUBSTRING(r2.Descrizione,
                                    PATINDEX('%ddt%', LOWER(r2.Descrizione)) + 4,
                                    CHARINDEX(' ', r2.Descrizione + ' ', PATINDEX('%ddt%', LOWER(r2.Descrizione)) + 4) 
                                        - (PATINDEX('%ddt%', LOWER(r2.Descrizione)) + 4)
                                )
                            FROM [dbo].DORig r2
                            WHERE r2.Id_DOTes = r.Id_DOTes
                            AND LOWER(r2.Descrizione) LIKE '%ddt%'
                        )
                    ) AS XDDT,
                    COALESCE(
                        MIN(r.XDATADDT),
                        (
                            SELECT TOP 1 
                                TRY_CAST(
                                    SUBSTRING(r2.Descrizione,
                                        PATINDEX('%del%', LOWER(r2.Descrizione)) + 4,
                                        10
                                    ) AS DATE
                                )
                            FROM [dbo].DORig r2
                            WHERE r2.Id_DOTes = r.Id_DOTes
                            AND LOWER(r2.Descrizione) LIKE '%ddt%del%'
                        )
                    ) AS XDATADDT
                FROM [dbo].DORig r
                WHERE r.Cd_Do IN ('FTV', 'FT2', 'NCC')
                AND r.Cd_CF = '{codice_cliente}'
                GROUP BY r.Id_DOTes
            ) bolla ON t.Id_DOTes = bolla.Id_DOTes

            WHERE t.Cd_Do IN ('FTV', 'FT2', 'NCC'{ddt_condition})
            AND t.Cd_CF = '{codice_cliente}'
            AND CAST(t.DataDoc AS DATE) >= CAST('{data_inizio}' AS DATE)
            AND CAST(t.DataDoc AS DATE) <= CAST('{data_fine}' AS DATE)
            {filtro_dotes}

            UNION ALL

            -- =====================================================
            -- RECORD MOVIMENTI (Tipo 02) - 128 caratteri fissi
            -- =====================================================
            SELECT
                r.Id_DOTes,
                r.Id_DORig,
                '02' AS TipoRec,
                '02' +                                                                            
                RIGHT('00000' + CAST(r.Id_DOTes AS VARCHAR(5)), 5) +                          
                RIGHT('               ' + ISNULL(r.Cd_AR, ''), 15) +                          
                LEFT(
                    ISNULL(
                        CASE 
                            WHEN '{codice_cliente}' = 'C000003' THEN 
                                REPLACE(
                                    REPLACE(
                                        REPLACE(r.Descrizione, '-', ''),  -- ASCII hyphen U+002D
                                    '–', ''),  -- en dash U+2013
                                '—', '')  -- em dash U+2014
                            ELSE r.Descrizione
                        END, 
                        ''
                    ) + '                              ', 
                    30
                ) +    
                RIGHT('  ' + 
                    CASE 
                        WHEN '{codice_cliente}' = 'C000003' THEN 'PZ'
                        ELSE ISNULL(r.Cd_ARMisura, '')
                    END, 
                    2
                ) +                                
                RIGHT('0000000' + CAST(CAST(
                    ISNULL(
                        CASE 
                            WHEN '{codice_cliente}' = 'C000003' THEN 
                                r.Qta * ISNULL(r.FattoreToUM1, 1)
                            ELSE 
                                r.Qta
                        END, 
                        0
                    ) * 100 AS BIGINT) AS VARCHAR(7)), 7) +      
                RIGHT('000000000' + CAST(CAST(
                    ISNULL(
                        CASE 
                            WHEN '{codice_cliente}' = 'C000003' THEN 
                                r.PrezzoUnitarioV / NULLIF(r.FattoreToUM1, 0)
                            ELSE 
                                r.PrezzoUnitarioV
                        END, 
                        0
                    ) * 1000 AS BIGINT) AS VARCHAR(9)), 9) +
                RIGHT('000000000' + CAST(CAST(ISNULL(r.PrezzoTotaleV, 0) * 1000 AS BIGINT) AS VARCHAR(9)), 9) +    
                '    ' +                                                                        
                ' ' +                                                                          
                CASE
                    WHEN r.Cd_Aliquota = 'W10' THEN '10'
                    WHEN r.Cd_Aliquota = 'W22' THEN '22'
                    ELSE RIGHT('  ' + ISNULL(REPLACE(r.Cd_Aliquota, 'W', ''), ''), 2)
                END +                                                                          
                ' ' +                                                                          
                '1' +                                                                          
                '                 ' +                                                          
                ' ' +                                                                          
                '                ' +                                                            
                '      '                                                                        
                AS RigaTracciato
            FROM [dbo].DORig r
            WHERE r.Cd_Do IN ('FTV', 'FT2', 'NCC'{ddt_condition})
            AND r.Cd_CF = '{codice_cliente}'
            AND (
                -- Righe normali con codice articolo
                (r.Cd_AR IS NOT NULL AND r.Cd_AR != '')
                OR
                -- Righe "Rif: DDT" solo per clienti specifici
                (
                    '{codice_cliente}' IN ('C000004', 'C000003', 'C000005')
                    AND (r.Cd_AR IS NULL OR r.Cd_AR = '')
                    AND r.Descrizione LIKE 'Rif: DDT %'
                )
            )
            AND CAST(r.DataDoc AS DATE) >= CAST('{data_inizio}' AS DATE)
            AND CAST(r.DataDoc AS DATE) <= CAST('{data_fine}' AS DATE)
            {filtro_dotes.replace("t.Id_DOTes", "r.Id_DOTes")}

            ORDER BY Id_DOTes, TipoRec, Id_DORig
            """

    def process_results(self, results, codice_cliente):
        """Post-processing risultati query"""
        if self.client_manager.requires_ddt_headers(codice_cliente):
            return self._insert_ddt_headers(results, codice_cliente)
        else:
            return results

    def _insert_ddt_headers(self, results, codice_cliente):
        """
        Inserisce testate DDT nei punti corretti
        
        Logica:
        1. Trova righe movimento (tipo 02) con "Rif: DDT" 
        2. Crea nuova testata basata sull'ultima testata dello stesso documento
        3. Inserisce testata PRIMA del prossimo articolo con Cd_AR
        4. Rimuove la riga "Rif: DDT" dai risultati finali
        """
        processed_results = []
        pending_headers = []
        last_header_by_doc = {}  # Dizionario {Id_DOTes: ultima_testata}
        
        for row in results:
            id_dotes = row[0]  # Prima colonna: Id_DOTes
            tipo_rec = row[1]  # Seconda colonna: TipoRec
            riga_tracciato = row[2]  # Terza colonna: RigaTracciato
            
            if tipo_rec == '01':
                # Testata - memorizza e aggiungi
                last_header_by_doc[id_dotes] = row
                processed_results.append(row)
                
            elif tipo_rec == '02':
                # Movimento - verifica se è "Rif: DDT" o articolo normale
                cd_ar = riga_tracciato[7:22].strip()  # Posizione Cd_AR nel tracciato
                
                if not cd_ar:  # Nessun codice articolo
                    descrizione_completa = riga_tracciato[22:52].strip()  # Posizione Descrizione
                    
                    if descrizione_completa.startswith('Rif: DDT'):
                        # Questa è una riga DDT - crea testata
                        if id_dotes in last_header_by_doc:
                            new_header = self._create_ddt_header_from_row(
                                last_header_by_doc[id_dotes],
                                descrizione_completa,
                                codice_cliente
                            )
                            pending_headers.append(new_header)
                        # NON includere questa riga nei risultati
                        continue
                
                # Articolo normale con Cd_AR
                # Prima aggiungi eventuali testate DDT in sospeso
                if pending_headers:
                    processed_results.extend(pending_headers)
                    pending_headers = []
                
                # Poi aggiungi l'articolo
                processed_results.append(row)
        
        # Aggiungi eventuali testate rimaste in coda
        if pending_headers:
            processed_results.extend(pending_headers)
        
        return processed_results

    def _create_ddt_header_from_row(self, original_header, descrizione_ddt, codice_cliente):
        """
        Crea una nuova testata basata su quella originale,
        modificando solo numero e data DDT
        
        Args:
            original_header: tuple (Id_DOTes, TipoRec, RigaTracciato)
            descrizione_ddt: str - es. "Rif: DDT 123 del 01-12-2024 Dest. 001"
            codice_cliente: str
        
        Returns:
            tuple: nuova testata modificata
        """
        import re
        
        id_dotes = original_header[0]
        riga_originale = original_header[2]
        
        # Estrai numero DDT (es. "123")
        match_num = re.search(r'Rif: DDT\s+(\d+)', descrizione_ddt)
        num_ddt = match_num.group(1) if match_num else ''
        
        # Estrai data DDT (es. "01-12-2024")
        match_data = re.search(r'del\s+(\d{2})-(\d{2})-(\d{4})', descrizione_ddt)
        if match_data:
            gg, mm, aaaa = match_data.groups()
            data_ddt_formatted = f"{aaaa[-2:]}{mm}{gg}"  # Formato AAMMGG
        else:
            data_ddt_formatted = '      '  # 6 spazi se non trovata
        
        # Ricostruisci la riga tracciato modificando solo DDT
        # Struttura: 01 + Id(5) + NumDoc(6) + DataDoc(6) + NumDDT(6) + DataDDT(6) + ...
        nuova_riga = (
            riga_originale[:13] +  # Mantieni: 01 + Id + NumDoc
            riga_originale[13:19] +  # Mantieni: DataDoc
            num_ddt.rjust(6) +  # NUOVO: NumDDT
            data_ddt_formatted +  # NUOVO: DataDDT
            riga_originale[31:]  # Mantieni: resto della riga
        )
        
        return (id_dotes, '01', nuova_riga)

    def _extract_ddt_info(self, descrizione):
        """
        Estrae numero e data DDT dalla descrizione
        
        Args:
            descrizione: str - es. "Rif: DDT 123 del 01-12-2024 Dest. 001"
        
        Returns:
            tuple: (num_ddt, data_ddt_formatted) - es. ("123", "241201")
        """
        import re
        
        # Estrai numero DDT
        match_num = re.search(r'Rif: DDT\s+(\d+)', descrizione)
        num_ddt = match_num.group(1) if match_num else ''
        
        # Estrai data DDT (formato gg-mm-aaaa)
        match_data = re.search(r'del\s+(\d{2})-(\d{2})-(\d{4})', descrizione)
        if match_data:
            gg, mm, aaaa = match_data.groups()
            data_ddt_formatted = f"{aaaa[-2:]}{mm}{gg}"  # AAMMGG
        else:
            data_ddt_formatted = '      '
        
        return num_ddt, data_ddt_formatted

    def extract_data_for_client(self, codice_cliente, id_dotes_filter=None):
        try:
            cliente = self.client_manager.get_cliente(codice_cliente)
            if not cliente:
                logging.error(f"[ERROR] Cliente {codice_cliente} non trovato")
                return None, None, None
            
            data_inizio, data_fine, nome_mese = self.calculate_previous_month_period()
            logging.info(f"[PROGRESS] Estrazione {cliente['nome']} per {nome_mese}")
            logging.info(f"[PROGRESS] Periodo: {data_inizio} - {data_fine}")
            
            query = self.get_sql_query(codice_cliente, data_inizio, data_fine, id_dotes_filter)
            df = pd.read_sql_query(query, self.engine)
            
            if self.client_manager.requires_ddt_headers(codice_cliente):
                df = self._insert_ddt_headers_df(df, codice_cliente)
                logging.info(f"[PROGRESS] Processing testate DDT completato per {cliente['nome']}")
            
            logging.info(f"[OK] Estratte {len(df)} righe per {cliente['nome']} - {nome_mese}")
            
            # ← AGGIUNTO: salta validazione se DataFrame vuoto
            if not df.empty:
                invalid_rows = df[df['RigaTracciato'].str.len() != 128]
                if len(invalid_rows) > 0:
                    logging.warning(f"[WARNING] {len(invalid_rows)} righe con lunghezza errata per {cliente['nome']}")
            
            return df, cliente, nome_mese
            
        except Exception as e:
            logging.error(f"[ERROR] Errore nell'estrazione per {codice_cliente}: {str(e)}")
            return None, None, None
    
    def save_to_file(self, df, cliente, nome_mese):
        """Salva i dati in un file per il cliente specifico"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # Formatta nome mese per filename (es: "January 2025" -> "Jan2025")
            mese_filename = nome_mese.replace(" ", "").replace("January", "Jan").replace("February", "Feb").replace("March", "Mar").replace("April", "Apr").replace("May", "May").replace("June", "Jun").replace("July", "Jul").replace("August", "Aug").replace("September", "Sep").replace("October", "Oct").replace("November", "Nov").replace("December", "Dec")
            
            filename = f"tracciato_tracciato_gdo_{cliente['nome_breve']}_{mese_filename}_{timestamp}.txt"
            filepath = os.path.join(self.output_folder, filename)
            
            # Scrive il tracciato
            with open(filepath, 'w', encoding='utf-8') as f:
                for riga in df['RigaTracciato']:
                    f.write(riga + '\n')
            
            logging.info(f"[OK] File salvato per {cliente['nome']} ({nome_mese}): {filepath}")
            logging.info(f"[STATS] {len(df)} righe totali - {len(df[df['TipoRec'] == '01'])} testate - {len(df[df['TipoRec'] == '02'])} movimenti")
            
            return filepath
            
        except Exception as e:
            logging.error(f"[ERROR] Errore nel salvataggio per {cliente['nome']}: {str(e)}")
            return None
    
    def create_email_body(self, cliente, stats, nome_mese):
            """Crea il corpo dell'email per il cliente con il mese specificato"""
            return f"""
            <html>
            <body>
                <h2>Tracciato TRACCIATO_GDO - {cliente['nome']}</h2>
                
                <p>Gentile Cliente,</p>
                
                <p>In allegato il tracciato TRACCIATO_GDO per il periodo <strong>{nome_mese}</strong> con i seguenti dati:</p>
                
                <ul>
                    <li><strong>Data estrazione:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</li>
                    <li><strong>Cliente:</strong> {cliente['nome']} ({cliente['codice']})</li>
                    <li><strong>Periodo:</strong> {nome_mese} (tutto il mese)</li>
                    <li><strong>Documenti elaborati:</strong> {stats['documenti']}</li>
                    <li><strong>Righe movimento:</strong> {stats['righe_movimento']}</li>
                    <li><strong>Righe totali:</strong> {stats['righe_totali']}</li>
                </ul>
                                
                <p><strong>Nota:</strong> Questo tracciato contiene tutti i documenti (fatture FTV, FT2 e note credito NCC) del mese di <strong>{nome_mese}</strong>.</p>
                
                <p>Cordiali saluti,<br>
                <strong>Sistema di Rifatturazione Automatica</strong></p>
                
                <hr>
                <small><i>Questo messaggio è stato generato automaticamente dal sistema di estrazione TRACCIATO_GDO.<br>
                Esecuzione programmata: ogni 7 del mese per il mese precedente.</i></small>
            </body>
            </html>
            """

    def send_email_to_client(self, cliente, filepath, stats, nome_mese):
        """Invia email al cliente con il tracciato per il mese specificato"""
        try:
            logging.info(f"[EMAIL] Preparazione invio email a {cliente['nome']} per {nome_mese}")
            logging.info(f"[EMAIL] Destinatario principale: {cliente['email']}")
            
            # Controlla se ci sono indirizzi CC
            cc_list = cliente.get('cc', [])
            if cc_list:
                logging.info(f"[EMAIL] Destinatari CC: {', '.join(cc_list)}")
            else:
                logging.info(f"[EMAIL] Nessun destinatario CC configurato")
            
            # Autenticazione Gmail
            if not self.gmail_sender.authenticate():
                logging.error(f"[EMAIL] Autenticazione Gmail fallita")
                return False
            
            # Crea messaggio
            subject = f"Tracciato TRACCIATO_GDO - {cliente['nome']} - {nome_mese} - {datetime.now().strftime('%d/%m/%Y')}"
            
            # Crea corpo email - NOTA: DEVE avere 3 parametri oltre a self
            try:
                body = self.create_email_body(cliente, stats, nome_mese)
                logging.info(f"[EMAIL] Corpo email creato per {cliente['nome']}")
            except Exception as e:
                logging.error(f"[EMAIL] Errore nella creazione corpo email: {str(e)}")
                return False
            
            # Crea messaggio con allegato e CC
            try:
                message = self.gmail_sender.create_message_with_attachment(
                    to=cliente['email'],
                    subject=subject,
                    message_text=body,
                    attachment_path=filepath,
                    cc=cc_list  # ← AGGIUNTO: Passa la lista CC
                )
                
                if not message:
                    logging.error(f"[EMAIL] Errore nella creazione messaggio")
                    return False
                    
                logging.info(f"[EMAIL] Messaggio creato per {cliente['email']}")
                if cc_list:
                    logging.info(f"[EMAIL] Messaggio include CC: {', '.join(cc_list)}")
            except Exception as e:
                logging.error(f"[EMAIL] Errore nella creazione messaggio: {str(e)}")
                return False
            
            # Invio email
            try:
                result = self.gmail_sender.send_message(message)
                if result:
                    logging.info(f"[EMAIL] Email inviata con successo a {cliente['nome']}")
                    if cc_list:
                        logging.info(f"[EMAIL] Email inviata anche in CC a: {', '.join(cc_list)}")
                return result
            except Exception as e:
                logging.error(f"[EMAIL] Errore nell'invio: {str(e)}")
                return False
            
        except Exception as e:
            logging.error(f"[ERROR] Errore generale nell'invio email a {cliente['nome']}: {str(e)}")
            return False
    
    def process_client(self, codice_cliente, send_email=False, id_dotes_filter=None):
        """Processa un singolo cliente per il mese precedente"""
        logging.info(f"[START] Elaborazione cliente {codice_cliente}")
        
        # Estrazione dati
        result = self.extract_data_for_client(codice_cliente, id_dotes_filter)
        if result[0] is None:
            return False
            
        df, cliente, nome_mese = result
        
        if len(df) == 0:
            logging.warning(f"[WARNING] Nessun dato trovato per {cliente['nome']} nel mese {nome_mese}")
            return False
        
        # Salvataggio file
        filepath = self.save_to_file(df, cliente, nome_mese)
        if not filepath:
            return False
        
        # Invio email se richiesto
        if send_email:
            stats = {
                'documenti': df['Id_DOTes'].nunique(),
                'righe_movimento': len(df[df['TipoRec'] == '02']),
                'righe_totali': len(df)
            }
            
            if self.send_email_to_client(cliente, filepath, stats, nome_mese):
                logging.info(f"[SUCCESS] Cliente {cliente['nome']} processato e email inviata ({nome_mese})")
            else:
                logging.error(f"[ERROR] Cliente {cliente['nome']} processato ma email non inviata")
                return False
        else:
            logging.info(f"[SUCCESS] Cliente {cliente['nome']} processato per {nome_mese} (senza invio email)")
        
        return True
    
    
    def run(self):
        """Esecuzione automatica completa - nessuna interazione utente"""
        logging.info("[AUTOMATIC] Avvio sistema estrazione TRACCIATO_GDO automatico")
        
        # Connessione database
        if not self.connect_database():
            logging.error("[AUTOMATIC] Connessione database fallita - terminazione")
            return False
        
        # Processamento automatico tutti i clienti con invio email
        success = self.process_all_clients()
        
        if success:
            logging.info("[AUTOMATIC] Sistema completato con successo")
        else:
            logging.error("[AUTOMATIC] Sistema completato con errori")
        
        return success
    
    def process_all_clients(self):
        """Processa automaticamente tutti i clienti con invio email"""
        data_inizio, data_fine, nome_mese = self.calculate_previous_month_period()
        logging.info(f"[AUTO] Elaborazione automatica tutti i clienti per {nome_mese}")
        
        clienti = self.client_manager.get_all_clienti()
        successi = 0
        
        for codice_cliente in clienti.keys():
            logging.info("-" * 80)
            if self.process_client_automatic(codice_cliente):
                successi += 1
        
        logging.info("-" * 80)
        logging.info(f"[AUTO] Elaborazione completata: {successi}/{len(clienti)} clienti processati")
        
        if successi == len(clienti):
            logging.info(f"[SUCCESS] Tutti i clienti elaborati e email inviate per {nome_mese}")
            return True
        else:
            logging.error(f"[ERROR] Solo {successi} clienti su {len(clienti)} elaborati correttamente")
            return False
    
    def process_client_automatic(self, codice_cliente):
        """Processa automaticamente un singolo cliente con invio email obbligatorio"""
        logging.info(f"[AUTO] Elaborazione automatica cliente {codice_cliente}")
        
        # Estrazione dati
        result = self.extract_data_for_client(codice_cliente)
        if result[0] is None:
            return False
            
        df, cliente, nome_mese = result
        
        if len(df) == 0:
            logging.warning(f"[WARNING] Nessun dato trovato per {cliente['nome']} nel mese {nome_mese}")
            logging.info(f"[SUCCESS] Cliente {cliente['nome']} processato (nessun dato per {nome_mese}) - OK")
            return True  # ← CAMBIATO: Nessun dato è un risultato valido, non un errore
        
        # Salvataggio file
        filepath = self.save_to_file(df, cliente, nome_mese)
        if not filepath:
            return False
        
        # Invio email automatico
        stats = {
            'documenti': df['Id_DOTes'].nunique(),
            'righe_movimento': len(df[df['TipoRec'] == '02']),
            'righe_totali': len(df)
        }
        
        if self.send_email_to_client(cliente, filepath, stats, nome_mese):
            logging.info(f"[SUCCESS] Cliente {cliente['nome']} processato e email inviata ({nome_mese})")
            return True
        else:
            logging.error(f"[ERROR] Cliente {cliente['nome']} processato ma email non inviata")
            return False

def main():
    import sys
    
    logging.info("[START] Avvio sistema TRACCIATO_GDO automatico")
    
    extractor = TracciatoGdoExtractorMultiClient()  # prima l'istanza
    
    if not extractor.connect_database():
        logging.error("Connessione fallita!")
        sys.exit(1)
    
    # --- MODALITÀ singolo cliente + ordine/i specifico/i: singolo cliente + filtro Id_DOTes ---
    # id_dotes_list = [1001, 1002, 1003]
    # extractor.process_client('C000001', send_email=False, id_dotes_filter= id_dotes_list)
    
    # --- MODALITÀ NORMALE: tutti i clienti (commenta il blocco sopra e decommenta questo) ---
    for codice in extractor.client_manager.get_all_clienti():
        extractor.process_client(codice, send_email=False)

    # --- PROVA CLIENTE_E ---
    # extractor.process_client('C000005', send_email=False)
    
    # logging.info("[FINAL] Sistema terminato")
    
#    if success:
#        logging.info("[FINAL] Sistema terminato con successo")
#        sys.exit(0)  # Exit code 0 = successo
#     else:
#        logging.error("[FINAL] Sistema terminato con errori")
#        sys.exit(1)  # Exit code 1 = errore

if __name__ == "__main__":
    main()