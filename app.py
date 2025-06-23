import streamlit as st
import fitz  # PyMuPDF
import re
import datetime
import pandas as pd
from typing import Optional, Dict, List, Tuple
from io import BytesIO

# Configurazione pagina Streamlit
st.set_page_config(
    page_title="📊 Report Consumi Migliorato",
    layout="wide",
    page_icon="📈",
    initial_sidebar_state="expanded"
)

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
    "ENI GAS E LUCE": r"ENI\s*GAS\s*E\s*LUCE",
    "HERA COMM": r"HERA\s*COMM",
    "IREN": r"IREN",
    "PUBLIACQUA": r"PUBLIACQUA",
    "SORGENIA": r"SORGENIA",
    "EDISON ENERGIA": r"EDISON\s*ENERGIA"
}

# Unità di misura per i consumi
UNITA_MISURA = {
    "luce": "kWh",
    "elettricità": "kWh",
    "energia": "kWh", 
    "gas": "mc",
    "acqua": "mc",
    "idrico": "mc"
}

def estrai_testo_da_pdf(file) -> str:
    """Estrae il testo da un file PDF con gestione errori migliorata."""
    try:
        doc = fitz.open(stream=file.read(), filetype="pdf")
        testo = ""
        for page in doc:
            testo += page.get_text()
        return testo
    except Exception as e:
        st.error(f"Errore durante l'estrazione del testo dal PDF {file.name}: {str(e)}")
        return ""

def estrai_societa(testo: str) -> Tuple[str, str]:
    """Estrae la società e il tipo di fornitura con precisione migliorata."""
    try:
        # Cerca corrispondenza esatta con società conosciute
        for societa, pattern in SOCIETA_CONOSCIUTE.items():
            if re.search(pattern, testo, re.IGNORECASE):
                # Determina tipo fornitura
                if "ACQU" in societa.upper():
                    return societa, "Acqua"
                elif "GAS" in societa.upper():
                    return societa, "Gas"
                elif "ENERG" in societa.upper():
                    return societa, "Luce"
                return societa, "N/D"

        # Pattern generici di fallback
        patterns = [
            (r'\b([A-Z]{2,}\s*(?:AIM|ENERGIA|GAS|ACQUA))\b', "N/D"),
            (r'\b(SPA|S\.P\.A\.|SRL|S\.R\.L\.)\b', "N/D")
        ]

        for pattern, tipo in patterns:
            match = re.search(pattern, testo)
            if match:
                return match.group(0).strip(), tipo

    except Exception as e:
        st.error(f"Errore durante l'estrazione della società: {str(e)}")

    return "N/D", "N/D"

def estrai_periodo(testo: str) -> str:
    """Estrae il periodo di riferimento con più pattern."""
    try:
        patterns = [
            r'dal\s+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\s+al\s+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
            r'periodo\s+di\s+riferimento\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\s*-\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
            r'rif\.\s*periodo\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\s*al\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})'
        ]

        for pattern in patterns:
            matches = re.finditer(pattern, testo, re.IGNORECASE)
            for match in matches:
                if len(match.groups()) == 2:
                    return f"{match.group(1)} - {match.group(2)}"

    except Exception as e:
        st.error(f"Errore durante l'estrazione del periodo: {str(e)}")

    return "N/D"

def parse_date(g: str, m: str, y: str) -> Optional[datetime.date]:
    """Parsing data migliorato con più formati supportati."""
    try:
        giorno = int(g)
        
        # Gestione mese come numero o nome
        if m.isdigit():
            mese = int(m)
        else:
            mese = MESI_MAP.get(m.lower().strip(), 0)
        
        # Gestione anno a 2 o 4 cifre
        if len(y) == 2:
            anno = 2000 + int(y)
        else:
            anno = int(y)
        
        # Validazione data
        if 1 <= mese <= 12 and 1 <= giorno <= 31:
            return datetime.date(anno, mese, giorno)
            
    except (ValueError, TypeError) as e:
        st.error(f"Errore durante il parsing della data: {str(e)}")
    
    return None

def estrai_data_fattura(testo: str) -> str:
    """Estrae la data della fattura con più pattern e fallback."""
    try:
        patterns = [
            r'data\s*fattura\s*[:\-]?\s*(\d{1,2})[\/\-\.\s](\d{1,2}|\w+)[\/\-\.\s](\d{2,4})',
            r'fattura\s*del\s*(\d{1,2})[\/\-\.\s](\d{1,2}|\w+)[\/\-\.\s](\d{2,4})',
            r'emissione\s*:\s*(\d{1,2})[\/\-\.\s](\d{1,2}|\w+)[\/\-\.\s](\d{2,4})',
            r'documento\s*emesso\s*in\s*data\s*(\d{1,2})[\/\-\.\s](\d{1,2}|\w+)[\/\-\.\s](\d{2,4})'
        ]

        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match and len(match.groups()) == 3:
                data = parse_date(match.group(1), match.group(2), match.group(3))
                if data:
                    return data.strftime("%d/%m/%Y")

        # Fallback per date in formato ISO
        iso_match = re.search(r'(\d{4}-\d{2}-\d{2})', testo)
        if iso_match:
            try:
                return datetime.datetime.strptime(iso_match.group(1), "%Y-%m-%d").strftime("%d/%m/%Y")
            except:
                pass

    except Exception as e:
        st.error(f"Errore durante l'estrazione della data della fattura: {str(e)}")

    return "N/D"

def estrai_pod_pdr(testo: str) -> Tuple[str, str]:
    """Estrae POD (luce) e PDR (gas) con pattern specifici."""
    try:
        # Cerca POD per elettricità
        pod_patterns = [
            r'POD\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Punto\s*di\s*Prelievo\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Codice\s*POD\s*[:\-]?\s*([A-Z0-9]{14,16})'
        ]
        
        # Cerca PDR per gas
        pdr_patterns = [
            r'PDR\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Punto\s*di\s*Ricerca\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Codice\s*PDR\s*[:\-]?\s*([A-Z0-9]{14,16})'
        ]
        
        pod = ""
        pdr = ""
        
        for pattern in pod_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                pod = match.group(1).strip()
                break
                
        for pattern in pdr_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                pdr = match.group(1).strip()
                break
                
        return pod, pdr
        
    except Exception as e:
        st.error(f"Errore durante l'estrazione del POD/PDR: {str(e)}")
        return "N/D", "N/D"

def estrai_indirizzo(testo: str) -> str:
    """Tenta di estrarre l'indirizzo del cliente."""
    try:
        patterns = [
            r'Indirizzo\s*[:\-]?\s*((?:Via|Viale|Piazza|Corso).+?\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'Servizio\s*erogato\s*in\s*((?:Via|Viale|Piazza|Corso).+?\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'Luogo\s*di\s*fornitura\s*[:\-]?\s*((?:Via|Viale|Piazza|Corso).+?\d{1,5}(?:\s*[A-Za-z]?)?)\b'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                return match.group(1).strip()
                
        return "N/D"
    except Exception as e:
        st.error(f"Errore durante l'estrazione dell'indirizzo: {str(e)}")
        return "N/D"

def estrai_numero_fattura(testo: str) -> str:
    """Estrae il numero della fattura con più pattern e validazione."""
    try:
        patterns = [
            r'numero\s*fattura\s*elettronica\s*[:\-]?\s*([A-Z0-9]{6,20})',
            r'n°\s*fattura\s*[:\-]?\s*([A-Z0-9]{6,20})',
            r'fattura\s*n\.?\s*([A-Z0-9]{6,20})',
            r'documento\s*[:\-]?\s*([A-Z0-9]{6,20})',
            r'rif\.\s*fattura\s*[:\-]?\s*([A-Z0-9]{6,20})'
        ]

        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                num = match.group(1).strip()
                # Validazione base del numero fattura
                if len(num) >= 6 and any(c.isdigit() for c in num):
                    return num

        # Fallback per numeri fattura complessi
        complex_pattern = r'\b(?:[A-Z]{2,5})[-/]?\d{4,}[-/]?\d*\b'
        match = re.search(complex_pattern, testo)
        if match:
            return match.group(0).strip()

    except Exception as e:
        st.error(f"Errore durante l'estrazione del numero della fattura: {str(e)}")

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
                valuta = match.group(2) if len(match.groups()) >= 2 and match.group(2) else "€"
                return importo, valuta

    except Exception as e:
        st.error(f"Errore durante l'estrazione del totale della bolletta: {str(e)}")

    return "N/D", "€"

def estrai_consumi(testo: str) -> Tuple[str, str]:
    """Estrae i consumi e l'unità di misura con più precisione."""
    try:
        # Cerca prima i consumi energetici (kWh)
        energy_patterns = [
            r'(?:consumo|energia)\s*(?:fatturato|complessivo)\s*[:\-]?\s*([\d\.,]+)\s*(kWh)?',
            r'totale\s*energia\s*(?:attiva|fatturata)\s*[:\-]?\s*([\d\.,]+)\s*(kWh)?',
            r'consumo\s*periodo\s*[:\-]?\s*([\d\.,]+)\s*(kWh)?'
        ]

        for pattern in energy_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match and match.group(1):
                consumo = match.group(1).replace('.', '').replace(',', '.')
                unita = match.group(2).lower() if match.group(2) else "kWh"
                return consumo, unita

        # Cerca consumi gas (mc)
        gas_patterns = [
            r'(?:consumo|gas)\s*(?:fatturato|complessivo)\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi)?',
            r'totale\s*gas\s*(?:naturale|fatturato)\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi)?',
            r'consumo\s*gas\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi)?'
        ]

        for pattern in gas_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match and match.group(1):
                consumo = match.group(1).replace('.', '').replace(',', '.')
                unita = "mc"
                if match.group(2):
                    unita = match.group(2).lower()
                    if "metri cubi" in unita:
                        unita = "mc"
                return consumo, unita

        # Cerca consumi acqua (mc)
        water_patterns = [
            r'(?:consumo|acqua)\s*(?:fatturato|complessivo)\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi|l|litri)?',
            r'totale\s*acqua\s*(?:fatturata|consumata)\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi|l|litri)?',
            r'volume\s*acqua\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi|l|litri)?'
        ]

        for pattern in water_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match and match.group(1):
                consumo = match.group(1).replace('.', '').replace(',', '.')
                unita = "mc"
                if match.group(2):
                    unita = match.group(2).lower()
                    if "litri" in unita:
                        unita = "l"
                return consumo, unita

    except Exception as e:
        st.error(f"Errore durante l'estrazione dei consumi: {str(e)}")

    return "N/D", "N/D"

def estrai_dati(file) -> Dict:
    """Estrae tutti i dati da un singolo file PDF."""
    testo = estrai_testo_da_pdf(file)
    societa, tipo_fornitura = estrai_societa(testo)
    pod, pdr = estrai_pod_pdr(testo)
    totale, valuta = estrai_totale_bolletta(testo)
    consumo, unita_misura = estrai_consumi(testo)
    
    return {
        "File": file.name,
        "Società": societa,
        "Tipo Fornitura": tipo_fornitura,
        "Periodo di Riferimento": estrai_periodo(testo),
        "Data Fattura": estrai_data_fattura(testo),
        "POD": pod,
        "PDR": pdr,
        "Indirizzo": estrai_indirizzo(testo),
        "Numero Fattura": estrai_numero_fattura(testo),
        f"Totale ({valuta})": totale,
        f"Consumo ({unita_misura})": consumo,
        "Note": ""
    }

def crea_excel(dati_lista: List[Dict]) -> BytesIO:
    """Crea un file Excel in memoria con i dati estratti."""
    output = BytesIO()
    df = pd.DataFrame(dati_lista)
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
        
        # Formattazione foglio Excel
        workbook = writer.book
        worksheet = writer.sheets['Report']
        
        # Formato header
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Formato dati
        data_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1
        })
        
        # Applica formati
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        for row in range(1, len(df)+1):
            for col in range(len(df.columns)):
                worksheet.write(row, col, df.iloc[row-1, col], data_format)
        
        # Auto-adjust column widths
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
    
    output.seek(0)
    return output

def mostra_grafico_consumi(dati_lista: List[Dict]):
    """Mostra un grafico comparativo dei consumi se disponibili."""
    try:
        df = pd.DataFrame(dati_lista)
        
        # Filtra solo righe con consumi numerici
        df['Consumo (val)'] = pd.to_numeric(df.iloc[:, -2], errors='coerce')
        df = df.dropna(subset=['Consumo (val)'])
        
        if len(df) > 1:
            st.subheader("📈 Confronto Consumi")
            
            # Determina unità di misura (assumiamo sia la stessa per tutte)
            unita = df.iloc[0, -1]  # Ultima colonna è l'unità
            
            # Crea grafico
            chart_data = df[['File', 'Consumo (val)']].set_index('File')
            st.bar_chart(chart_data)
            
            st.caption(f"Unità di misura: {unita}")
    except Exception as e:
        st.warning(f"Impossibile generare il grafico dei consumi: {str(e)}")

def main():
    st.title("📊 Report Consumi Migliorato")
    st.markdown("""
    **Carica una o più bollette PDF** per estrarre automaticamente i dati principali.
    """)
    
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
                risultati.append(dati)
            except Exception as e:
                st.error(f"Errore durante l'elaborazione di {file.name}: {str(e)}")
                continue
        
        progress_bar.empty()
        status_text.success(f"✅ Elaborazione completata! {len(risultati)} file processati.")
        
        if risultati:
            # Mostra tabella risultati
            st.subheader("📋 Dati Estratti")
            
            if raggrupa_societa:
                societa = st.selectbox(
                    "Filtra per società",
                    options=["Tutte"] + sorted(list(set(d['Società'] for d in risultati if d['Società'] != "N/D")),
                    index=0
                )
                
                if societa != "Tutte":
                    risultati_filtrati = [d for d in risultati if d['Società'] == societa]
                else:
                    risultati_filtrati = risultati
            else:
                risultati_filtrati = risultati
            
            st.dataframe(
                pd.DataFrame(risultati_filtrati),
                use_container_width=True,
                hide_index=True
            )
            
            # Mostra grafici se richiesto
            if mostra_grafici:
                mostra_grafico_consumi(risultati_filtrati)
            
            # Pulsanti esportazione
            st.subheader("📤 Esporta Dati")
            col1, col2 = st.columns(2)
            
            with col1:
                # Esporta Excel
                excel_data = crea_excel(risultati_filtrati)
                st.download_button(
                    label="Scarica Excel",
                    data=excel_data,
                    file_name="report_consumi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                # Esporta CSV
                csv = pd.DataFrame(risultati_filtrati).to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Scarica CSV",
                    data=csv,
                    file_name="report_consumi.csv",
                    mime="text/csv"
                )
    
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; font-size: 14px; color: gray;">
        Strumento sviluppato per l'estrazione automatica di dati da bollette PDF<br>
        Supporta i principali fornitori italiani di luce, gas e acqua
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
