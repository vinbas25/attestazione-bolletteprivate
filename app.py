import streamlit as st
import fitz  # PyMuPDF
import re
import datetime
import pandas as pd
import logging
from typing import Optional, Dict, List, Tuple
from io import BytesIO

# CONFIGURAZIONE LAYOUT E STILE STREAMLIT
st.set_page_config(layout="wide")

st.markdown("""
    <style>
        /* Nasconde menu, header e footer */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        
        /* Riduce padding per sfruttare tutto lo spazio */
        .main .block-container {
            padding-top: 1rem;
            padding-right: 1rem;
            padding-left: 1rem;
            padding-bottom: 1rem;
        }
    </style>
""", unsafe_allow_html=True)

# Configurazione del logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Mappa mesi in italiano
MESI_MAP = {
    "gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4, "maggio": 5, "giugno": 6,
    "luglio": 7, "agosto": 8, "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12
}

# Elenco esteso di società conosciute con regex specifiche
SOCIETA_CONOSCIUTE = {
    "AGSM AIM ENERGIA": r"AGSM\s*AIM",
    "A2A ENERGIA": r"A2A\s*ENERGIA",
    "ACQUE VERONA": r"ACQUE\s*VERONA",
    "ACQUE SPA": r"ACQUE\s*SPA",
    "AQUEDOTTO DEL FIORA": r"AQUEDOTTO\s*DEL\s*FIORA",
    "ASA LIVORNO": r"ASA\s*LIVORNO",
    "ENEL ENERGIA": r"ENEL\s*ENERGIA",
    "NUOVE ACQUE": r"NUOVE\s*ACQUE",
    "GAIA SPA": r"GAIA\s*SPA",
    "PUBLIACQUA": r"PUBLIACQUA",
    "EDISON ENERGIA": r"EDISON\s*ENERGIA"
}

def estrai_testo_da_pdf(file) -> str:
    """Estrae il testo da un file PDF con gestione errori migliorata."""
    try:
        doc = fitz.open(stream=file.read(), filetype="pdf")
        testo = ""
        for page in doc:
            testo += page.get_text()
        return testo
    except fitz.FileDataError:
        logger.error(f"File {file.name} non valido o corrotto")
        return ""
    except Exception as e:
        logger.error(f"Errore durante l'estrazione del testo dal PDF {file.name}: {str(e)}")
        return ""

def estrai_societa(testo: str) -> str:
    """Estrae la società con precisione migliorata."""
    try:
        for societa, pattern in SOCIETA_CONOSCIUTE.items():
            if re.search(pattern, testo, re.IGNORECASE):
                return societa
        patterns = [
            r'\b([A-Z]{2,}\s*(?:AIM|ENERGIA|GAS|ACQUA|SPA))\b',
            r'\b(SPA|S\.P\.A\.|SRL|S\.R\.L\.)\b'
        ]
        for pattern in patterns:
            match = re.search(pattern, testo)
            if match:
                return match.group(0).strip()
    except Exception as e:
        logger.error(f"Errore durante l'estrazione della società: {str(e)}")
    return "N/D"

def estrai_periodo(testo: str) -> str:
    """Estrae il periodo di riferimento con più pattern."""
    try:
        patterns = [
            r'dal\s+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\s+al\s+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
            r'periodo\s+di\s+riferimento\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\s*-\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
            r'Periodo di riferimento\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\s*-\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
            r'rif\.\s*periodo\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\s*al\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
            # Nuovi pattern aggiunti
            r'dal\s+(\d{1,2}/\d{1,2}/\d{4})\s+al\s+(\d{1,2}/\d{1,2}/\d{4})',  # Formato con slash e anno a 4 cifre
            r'Periodo di riferimento\s+(\d{1,2}/\d{1,2}/\d{4}\s*-\s*\d{1,2}/\d{1,2}/\d{4})',  # Formato con trattino
            r'Periodo\s*:\s*(\d{1,2}/\d{1,2}/\d{4})\s*-\s*(\d{1,2}/\d{1,2}/\d{4})',  # Versione abbreviata
            r'Periodo fatturazione\s*:\s*(\d{1,2}/\d{1,2}/\d{4})\s*-\s*(\d{1,2}/\d{1,2}/\d{4})',  # Alternativa con "fatturazione"
            r'dal\s+(\d{2}/\d{2}/\d{4})\s+al\s+(\d{2}/\d{2}/\d{4})',
            r'Periodo di riferimento (\d{2}/\d{2}/\d{4}) - (\d{2}/\d{2}/\d{4})',
            r'Periodo di riferimento (\d{2}/\d{2}/\d{4}) - (\d{2}/\d{2}/\d{4})',
            r'(\d{2}/\d{2}/\d{4}) - (\d{2}/\d{2}/\d{4})',
        ]
        for pattern in patterns:
            matches = re.finditer(pattern, testo, re.IGNORECASE)
            for match in matches:
                if len(match.groups()) == 2:
                    return f"{match.group(1)} - {match.group(2)}"
    except Exception as e:
        logger.error(f"Errore durante l'estrazione del periodo: {str(e)}")
    return "N/D"

def parse_date(g: str, m: str, y: str) -> Optional[datetime.date]:
    """Parsing data migliorato con più formati supportati."""
    try:
        giorno = int(g)
        if m.isdigit():
            mese = int(m)
        else:
            mese = MESI_MAP.get(m.lower().strip(), 0)
        if len(y) == 2:
            anno = 2000 + int(y)
        else:
            anno = int(y)
        if 1 <= mese <= 12 and 1 <= giorno <= 31:
            return datetime.date(anno, mese, giorno)
    except (ValueError, TypeError) as e:
        logger.error(f"Errore durante il parsing della data: {str(e)}")
    return None

def estrai_data_fattura(testo: str) -> str:
    """Estrae la data della fattura con più pattern e fallback."""
    try:
        patterns = [
            r'(?:data\s*fattura|fattura\s*del|emissione)\s*[:\-]?\s*(\d{1,2})[\/\-\.\s](\d{1,2}|\w+)[\/\-\.\s](\d{2,4})',
            r'(?:data\s*emissione|emesso\s*il)\s*[:\-]?\s*(\d{1,2})[\/\-\.\s](\d{1,2}|\w+)[\/\-\.\s](\d{2,4})',
            r'\b(\d{2})[\/\-\.](\d{2})[\/\-\.](\d{4})\b',
            r'\b(\d{4})[\/\-\.](\d{2})[\/\-\.](\d{2})\b',
            r'\b(\d{1,2})\s+(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)\s+(\d{4})\b',
            r'\b(?:al|il)\s+(\d{1,2})\s+(\w+)\s+(\d{4})\b'
        ]
        for pattern in patterns:
            matches = re.finditer(pattern, testo, re.IGNORECASE)
            for match in matches:
                if len(match.groups()) == 3:
                    data = parse_date(match.group(1), match.group(2), match.group(3))
                    if data:
                        return data.strftime("%d/%m/%Y")
    except Exception as e:
        logger.error(f"Errore durante l'estrazione della data: {str(e)}")
    return "N/D"

def estrai_pod_pdr(testo: str) -> str:
    """Estrae POD o PDR unificato con pattern specifici."""
    try:
        pod_patterns = [
            r'POD\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Punto\s*di\s*Prelievo\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Codice\s*POD\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'(?:matricola\s*contatore|matr\.?\s*cont\.?|numero\s*contatore)\s*[:=\-]?\s*([A-Z0-9]{8,12})(?:\s|$)',
            r'(?:matricola\s*contatore|matr\.?\s*cont\.?|numero\s*contatore)\s*[:=\-]?\s*([A-Z0-9\-]{8,12})(?:\s|$)',
            r'(?:matricola\s*contatore|matr\.?\s*cont\.?|numero\s*contatore)\s*[:=\-]?\s*([A-Z0-9]{8,14})(?:\s|$)',
            r'Contatore\s*n\.\s*(\d{6,})',
        ]
        for pattern in pod_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        pdr_patterns = [
            r'PDR\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Punto\s*di\s*Ricerca\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Codice\s*PDR\s*[:\-]?\s*([A-Z0-9]{14,16})'
        ]
        for pattern in pdr_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                return match.group(1).strip()
    except Exception as e:
        logger.error(f"Errore durante l'estrazione del POD/PDR: {str(e)}")
    return "N/D"

def estrai_indirizzo(testo: str) -> str:
    """
    Tenta di estrarre l'indirizzo del cliente da un testo utilizzando regex.
    
    Args:
        testo: Stringa contenente il testo da analizzare
        
    Returns:
        Stringa con l'indirizzo estratto o "N/D" se non trovato
    """
    try:
        patterns = [
            r'Indirizzo\s*[:\-]?\s*((?:Via|Viale|Piazza|Corso|C\.so|V\.le|Str\.).+?\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'Servizio\s*erogato\s*in\s*((?:Via|Viale|Piazza|Corso|C\.so|V\.le|Str\.).+?\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'Luogo\s*di\s*fornitura\s*[:\-]?\s*((?:Via|Viale|Piazza|Corso|C\.so|V\.le|Str\.).+?\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'Indirizzo\s*di\s*fornitura\s*[:\-]?\s*((?:Via|Viale|Piazza|Corso|C\.so|V\.le|Str\.).+?\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'Indirizzo\s*fornitura\s*((?:Via|Viale|Piazza|Corso|C\.so|V\.le|Str\.).+?\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'(?:Indirizzo|Servizio erogato in|Luogo di fornitura|Indirizzo di fornitura|Indirizzo fornitura)\s*[:\-]?\s*((?:Via|Viale|Piazza|Corso|C\.so|V\.le|Str\.)\s+[A-Za-zÀ-ÿ\s]+?\s*\d{1,5}(?:\s*[A-Za-z]?)?)',
            r'DATI FORNITURA.*?VIA\s(.*?\d{5}\s\w{2})',
            r'(?:DATI FORNITURA|Indirizzo|Luogo di fornitura|Servizio erogato in|Ubicazione).*?VIA\s(.*?\d{5}\s\w{2})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                indirizzo = match.group(1).strip()
                # Pulizia aggiuntiva dell'indirizzo
                indirizzo = re.sub(r'^\W+|\W+$', '', indirizzo)  # Rimuove punteggiatura all'inizio/fine
                return indirizzo
                
        return "N/D"
        
    except Exception as e:
        logger.error(f"Errore durante l'estrazione dell'indirizzo: {str(e)}", exc_info=True)
        return "N/D"

def estrai_numero_fattura(testo: str) -> str:
    """Estrae il numero della fattura con più pattern e validazione."""
    try:
        patterns = [
            r'Numero fattura elettronica valida ai fini fiscali\s*[:]?\s*([A-Z]{0,4}\s*[0-9\/\-]+\s*[0-9]+)',
            r'(?:numero\s*fattura|n°\s*fattura|fattura\s*n\.?)\s*[:\-]?\s*([A-Z]{0,4}\s*[0-9\/\-]+\s*[0-9]+)',
            r'(?:doc\.|documento)\s*[:\-]?\s*([A-Z]{0,4}\s*[0-9\/\-]+\s*[0-9]+)',
            r'[Ff]attura\s+(?:elektronica\s+)?[nN]°?\s*[:\-]?\s*([A-Z]{0,4}\s*[0-9\/\-]+\s*[0-9]+)',
            r'Numero Fattura\s*[:]?\s*([A-Z]{0,4}\s*[0-9\/\-]+\s*[0-9]+)', # Aggiunto il pattern per "Numero Fattura"
            r'\b\d{2,4}[\/\-]\d{3,8}\b',
            r'\b[A-Z]{2,5}\s*\d{4,}\/\d{2,}\b'
        ]
        for pattern in patterns:
            matches = re.finditer(pattern, testo, re.IGNORECASE)
            for match in matches:
                num = match.group(1) if match.groups() else match.group(0)
                num = num.strip()
                if len(num) >= 5 and any(c.isdigit() for c in num):
                    return num
    except Exception as e:
        logger.error(f"Errore durante l'estrazione del numero della fattura: {str(e)}")
    return "N/D"

def estrai_totale_bolletta(testo: str) -> Tuple[str, str]:
    """Estrae il totale e la valuta con più pattern."""
    try:
        patterns = [
            r'totale\s*(?:fattura|bolletta)\s*[:\-]?\s*[€]?\s*([\d\.,]+)\s*([€]?)',
            r'importo\s*totale\s*[:\-]?\s*[€]?\s*([\d\.,]+)\s*([€]?)',
            r'pagare\s*[:\-]?\s*[€]?\s*([\d\.,]+)\s*([€]?)',
            r'totale\s*dovuto\s*[:\-]?\s*[€]?\s*([\d\.,]+)\s*([€]?)'
        ]
        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match and len(match.groups()) >= 1:
                importo = match.group(1).replace('.', '').replace(',', '.')
                try:
                    float(importo)
                    valuta = match.group(2) if len(match.groups()) >= 2 and match.group(2) else "€"
                    return importo, valuta
                except ValueError:
                    continue
    except Exception as e:
        logger.error(f"Errore durante l'estrazione del totale della bolletta: {str(e)}")
    return "N/D", "€"

import re
import logging

logger = logging.getLogger(__name__)

def estrai_consumi(testo: str) -> str:
    """Estrae i consumi fatturati da testo OCR o PDF, gestendo vari formati e fallback."""
    try:
        # Blocchi prioritari basati su intestazioni note (come RIEPILOGO CONSUMI FATTURATI)
        testo_upper = testo.upper()
        idx = testo_upper.find("RIEPILOGO CONSUMI FATTURATI")
        if idx != -1:
            snippet = testo_upper[idx:idx+600]

            match = re.search(r'TOTALE COMPLESSIVO DI[:\-]?\s*([\d\.,]+)', snippet)
            if not match:
                match = re.search(r'TOTALE\s+QUANTITÀ[:\-]?\s*([\d\.,]+)', snippet)

            if match:
                try:
                    valore = float(match.group(1).replace('.', '').replace(',', '.'))
                    return f"{valore:.2f} Smc"
                except:
                    pass  # Continua con i pattern generali

        # Pattern generali e multi-bolletta (aggiunti i nuovi pattern per casi come "329 mc")
        patterns = [
            # Pattern specifico per bollette tipo Nuove Acque (es. "Consumo\n329 mc")
            r'Consumo\s*\n\s*(\d+)\s*mc',  # Cattura "329" dopo "Consumo" e a capo
            r'Consumo\s+nel\s+periodo\s+di\s+\d+\s+giorni:\s*([\d\.,]+)\s*mc',  # Es: "Consumo nel periodo di 141 giorni: 299 mc"
            r'Letture e Consumi.*?Contatore n\.\s*\d+.*?(\d+)\s*mc',  # Cattura consumo da tabelle
            r'Consumo\s+stimato\s*[:\-]?\s*([\d\.,]+)\s*mc',  # Fallback per stime
            r'Consumo\s+fatturato\s*[:\-]?\s*([\d\.,]+)\s*mc',  # Fallback esplicito
            r'totale\s+smc\s+fatturati\s*[:\-]?\s*([\d]{1,3}(?:[\.,][\d]{3})*(?:[\.,]\d+)?)',
            r'Totale\s+quantità\s*[:\-]?\s*([\d.]+,\d+)\s*Smc',
            r'totale\s+consumo\s+fatturato\s+per\s+il\s+periodo\s+di\s+riferimento\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi)',
            r'(?:consumo\s*fatturato|consumo\s*stimato\s*fatturato|consumo\s*totale)\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi)',
            r'(?:riepilogo\s*consumi[^\n]*\n.*\n.*?)([\d\.,]+)\s*(mc|m³|metri\s*cubi)',
            r'(?:prospetto\s*letture\s*e\s*consumi[^\n]*\n.*\n.*?\d+)\s+([\d\.,]+)\s*$',
            r'(?:dettaglio\s*consumi[^\n]*\n.*\n.*?\d+\s+)([\d\.,]+)\s*$',
        ]

        for pattern in patterns:
            matches = re.finditer(pattern, testo, re.IGNORECASE | re.MULTILINE)
            for match in matches:
                try:
                    valore_raw = match.group(1)
                    valore_normalizzato = valore_raw.replace('.', '').replace(',', '.')
                    consumo = float(valore_normalizzato)

                    if len(match.groups()) > 1 and match.group(2):
                        unita = match.group(2).lower()
                    else:
                        unita = "mc"  # Default per i nuovi pattern

                    return f"{consumo:.2f} {unita}"
                except (ValueError, IndexError):
                    logger.debug(f"Errore nel processare il match: {match.group() if match else 'N/A'}")
                    continue

        # Fallback per casi estremi (es. testo con formattazione irregolare)
        fallback = re.search(r'(\d+)\s*mc\s+Importo\s+da\s+pagare', testo)  # Cattura "329 mc" prima di "Importo da pagare"
        if fallback:
            return f"{float(fallback.group(1)):.2f} mc"

    except Exception as e:
        logger.error(f"Errore durante l'estrazione dei consumi: {str(e)}", exc_info=True)

    return "N/D"

def estrai_dati_cliente(testo: str) -> str:
    """Estrae i dati del cliente (codice cliente, partita IVA, ecc.)."""
    try:
        patterns = [
            r'(?:Numero\s*Contatore|Contatore)[\s:]*([0-9]{8,9})',
            r'(?:Matricola|Contatore|S/N)[\s:]*([A-Z0-9]{14,15})'
        ]
        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return "N/D"
    except Exception as e:
        logger.error(f"Errore durante l'estrazione dei dati cliente: {str(e)}")
        return "N/D"

def estrai_dati(file) -> Dict[str, str]:
    """Estrae tutti i dati da un singolo file PDF."""
    testo = estrai_testo_da_pdf(file)
    if not testo:
        return None
    societa = estrai_societa(testo)
    pod = estrai_pod_pdr(testo)
    totale, valuta = estrai_totale_bolletta(testo)
    consumi = estrai_consumi(testo)
    indirizzo = estrai_indirizzo(testo)
    dati_cliente = estrai_dati_cliente(testo)
    return {
        "Società": societa,
        "Periodo di Riferimento": estrai_periodo(testo),
        "Data Fattura": estrai_data_fattura(testo),
        "POD": pod,
        "Dati Cliente": dati_cliente,
        "Indirizzo": indirizzo,
        "Numero Fattura": estrai_numero_fattura(testo),
        f"Totale ({valuta})": totale,
        "File": file.name,
        "Consumi": consumi
    }

def crea_excel(dati_lista: List[Dict[str, str]]) -> Optional[BytesIO]:
    """Crea un file Excel in memoria con i dati estratti."""
    try:
        colonne_ordinate = [
            "Società",
            "Periodo di Riferimento",
            "Data Fattura",
            "POD",
            "Dati Cliente",
            "Indirizzo",
            "Numero Fattura",
            "Totale (€)",
            "Consumi",
            "File"
        ]
        df = pd.DataFrame([d for d in dati_lista if d is not None])
        if len(df) == 0:
            st.warning("Nessun dato valido da esportare")
            return None
        colonne_presenti = [col for col in colonne_ordinate if col in df.columns]
        df = df[colonne_presenti]
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Report')
            workbook = writer.book
            worksheet = writer.sheets['Report']
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })
            data_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'top',
                'border': 1
            })
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            for row in range(1, len(df)+1):
                for col in range(len(df.columns)):
                    worksheet.write(row, col, df.iloc[row-1, col], data_format)
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
        output.seek(0)
        return output
    except Exception as e:
        logger.error(f"Errore durante la creazione del file Excel: {str(e)}")
        return None

def mostra_grafico_consumi(dati_lista: List[Dict[str, str]]):
    """Mostra un grafico comparativo dei consumi se disponibili."""
    try:
        df = pd.DataFrame([d for d in dati_lista if d is not None])
        if len(df) == 0:
            return
        if "Consumi" not in df.columns:
            return
        df['Consumo_val'] = df['Consumi'].str.extract(r'([\d\.]+)')[0].astype(float)
        df = df.dropna(subset=['Consumo_val'])
        if len(df) < 2:
            return
        st.subheader("📈 Confronto Consumi")
        unita = df['Consumi'].iloc[0].split()[-1] if len(df['Consumi'].iloc[0].split()) > 1 else ""
        chart_data = df[['File', 'Consumo_val']].rename(columns={'Consumo_val': 'Consumo'})
        chart_data = chart_data.set_index('File')
        st.bar_chart(chart_data)
        if unita:
            st.caption(f"Unità di misura: {unita}")
    except Exception as e:
        st.warning(f"Impossibile generare il grafico: {str(e)}")

def main():
    st.title("📊 Analizzatore Bollette Migliorato")
    st.markdown("""
    **Carica una o più bollette PDF** per estrarre automaticamente i dati principali.
    """)

    # Aggiungi CSS personalizzato per allargare la visualizzazione
    st.markdown("""
    <style>
    div[data-baseweb="base-input"] {
        width: 100%;
    }
    div[data-testid="stDataFrame"] {
        width: 100%;
    }
    </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.header("Impostazioni")
        mostra_grafici = st.checkbox("Mostra grafici comparativi", value=True)
        raggruppa_societa = st.checkbox("Raggruppa per società", value=True)

    file_pdf_list = st.file_uploader(
        "Seleziona i file PDF delle bollette",
        type=["pdf"],
        accept_multiple_files=True,
        help="Puoi selezionare più file contemporaneamente"
    )

    if file_pdf_list:
        risultati = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        for i, file in enumerate(file_pdf_list):
            status_text.text(f"Elaborazione {i+1}/{len(file_pdf_list)}: {file.name[:30]}...")
            progress_bar.progress((i + 1) / len(file_pdf_list))
            try:
                dati = estrai_dati(file)
                if dati:
                    risultati.append(dati)
            except Exception as e:
                logger.error(f"Errore durante l'elaborazione di {file.name}: {str(e)}")
                continue
        progress_bar.empty()
        if risultati:
            status_text.success(f"✅ Elaborazione completata! {len(risultati)} file processati con successo.")
            st.subheader("📋 Dati Estratti")
            if raggruppa_societa:
                societa_disponibili = sorted(list(set(d['Società'] for d in risultati if pd.notna(d['Società']) and (d['Società'] != "N/D"))))
                if societa_disponibili:
                    societa = st.selectbox(
                        "Filtra per società",
                        options=["Tutte"] + societa_disponibili,
                        index=0
                    )
                    if societa != "Tutte":
                        risultati_filtrati = [d for d in risultati if d['Società'] == societa]
                    else:
                        risultati_filtrati = risultati
                else:
                    risultati_filtrati = risultati
                    st.warning("Nessuna società riconosciuta nei documenti")
            else:
                risultati_filtrati = risultati

            # Utilizza st.data_editor per una migliore interazione
            df = pd.DataFrame(risultati_filtrati)
            st.data_editor(
                df,
                use_container_width=True,
                hide_index=True,
                disabled=True,
                key="data_editor"
            )

            if mostra_grafici and risultati_filtrati:
                mostra_grafico_consumi(risultati_filtrati)

            st.subheader("📤 Esporta Dati")
            col1, col2 = st.columns(2)
            with col1:
                excel_data = crea_excel(risultati_filtrati)
                if excel_data:
                    st.download_button(
                        label="Scarica Excel",
                        data=excel_data,
                        file_name="report_consumi.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Scarica i dati in formato Excel"
                    )
            with col2:
                if risultati_filtrati:
                    csv = pd.DataFrame(risultati_filtrati).to_csv(index=False, sep=';').encode('utf-8')
                    st.download_button(
                        label="Scarica CSV",
                        data=csv,
                        file_name="report_consumi.csv",
                        mime="text/csv",
                        help="Scarica i dati in formato CSV (delimitato da punto e virgola)"
                    )
        else:
            status_text.warning("⚠️ Nessun dato valido estratto dai file caricati")

    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; font-size: 14px; color: gray;">
        Strumento sviluppato dal Mar. Vincenzo Basile<br>
        Supporta i principali fornitori italiani di luce, gas e acqua
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
