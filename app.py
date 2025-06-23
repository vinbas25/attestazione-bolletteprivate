import streamlit as st
import fitz  # PyMuPDF
import re
import datetime
import pandas as pd
from typing import Optional, Dict, List, Tuple
from io import BytesIO
import logging

# Configurazione logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configurazione pagina Streamlit
st.set_page_config(
    page_title="📊 Analizzatore Bollette Avanzato",
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
    "GAIA SPA": r"GAIA\s*SPA",
    "HERA COMM": r"HERA\s*COMM",
    "IREN": r"IREN",
    "PUBLIACQUA": r"PUBLIACQUA",
    "SORGENIA": r"SORGENIA",
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
    except fitz.FileDataError as e:
        logger.error(f"FileDataError in {file.name}: {str(e)}")
        st.error(f"File {file.name} non valido o corrotto")
        return ""
    except Exception as e:
        logger.error(f"Errore durante l'estrazione da PDF {file.name}: {str(e)}", exc_info=True)
        st.error(f"Errore durante l'estrazione del testo dal PDF {file.name}")
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
        return "N/D"
    except Exception as e:
        logger.error(f"Errore estrazione società: {str(e)}", exc_info=True)
        return "N/D"

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
        return "N/D"
    except Exception as e:
        logger.error(f"Errore estrazione periodo: {str(e)}", exc_info=True)
        return "N/D"

def estrai_indirizzo(testo: str) -> str:
    """Estrae l'indirizzo con maggiore accuratezza."""
    try:
        patterns = [
            r'Indirizzo\s*[:\-]?\s*((?:Via|Viale|Piazza|Corso|Contrada|Borgo|Frazione)[\s\w]+?\s*\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'Luogo\s*di\s*fornitura\s*[:\-]?\s*((?:[\w\s]+)\s*\d{1,5}(?:\s*[A-Za-z]?)?)',
            r'Servizio\s*erogato\s*in\s*([\w\s]+?\s*\d{1,5}(?:\s*[A-Za-z]?)?)\b',
            r'Fornitura\s*presso\s*:\s*([\w\s]+?\s*\d{1,5}(?:\s*[A-Za-z]?)?)\b'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                indirizzo = match.group(1).strip()
                # Pulizia dell'indirizzo
                indirizzo = re.sub(r'\s+', ' ', indirizzo)  # Rimuovi spazi multipli
                indirizzo = re.sub(r'[\n\t]', ' ', indirizzo)  # Rimuovi caratteri speciali
                return indirizzo
        return "N/D"
    except Exception as e:
        logger.error(f"Errore estrazione indirizzo: {str(e)}", exc_info=True)
        return "N/D"

def estrai_consumi(testo: str) -> str:
    """Estrae i consumi con pattern più completi e validazione."""
    try:
        patterns = [
            # Pattern per energia elettrica
            r'(?:consumo\s*energia\s*attiva|energia\s*fatturata)\s*[:\-]?\s*([\d\.,]+)\s*(kWh)',
            # Pattern per gas
            r'(?:consumo\s*gas|volume\s*gas)\s*[:\-]?\s*([\d\.,]+)\s*(m³|mc|metri\s*cubi)',
            # Pattern per acqua
            r'(?:consumo\s*acqua|volume\s*acqua)\s*[:\-]?\s*([\d\.,]+)\s*(m³|mc|litri|l)',
            # Pattern generici
            r'(?:consumo\s*totale|totale\s*consumi)\s*[:\-]?\s*([\d\.,]+)\s*(kWh|m³|mc|litri|l)',
            r'(?:RIEPILOGO\s*CONSUMI\s*FATTURATI)\s*[\s\S]*?([\d\.,]+)\s*(kWh|m³|mc|litri|l)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match and match.group(1):
                try:
                    consumo = match.group(1).replace('.', '').replace(',', '.')
                    consumo_float = float(consumo)
                    unita = match.group(2).lower() if match.group(2) else ""
                    
                    # Normalizza unità di misura
                    if not unita:
                        if "energia" in pattern.lower() or "kwh" in pattern.lower():
                            unita = "kWh"
                        elif "gas" in pattern.lower() or "mc" in pattern.lower() or "m³" in pattern.lower():
                            unita = "mc"
                        elif "acqua" in pattern.lower():
                            unita = "mc"  # Default per acqua
                    
                    return f"{consumo_float:.3f} {unita}"
                except ValueError:
                    continue
        return "N/D"
    except Exception as e:
        logger.error(f"Errore estrazione consumi: {str(e)}", exc_info=True)
        return "N/D"

# [Altre funzioni rimangono uguali...]

def main():
    st.title("📊 Analizzatore Bollette Avanzato")
    st.markdown("""
    **Carica una o più bollette PDF** per estrarre automaticamente i dati principali.
    """)
    
    with st.sidebar:
        st.header("Impostazioni Avanzate")
        mostra_grafici = st.checkbox("Mostra grafici comparativi", value=True)
        raggruppa_societa = st.checkbox("Raggruppa per società", value=True)
        mostra_dettagli = st.checkbox("Mostra dettagli tecnici", value=False)
    
    # Sezione di caricamento file con drag & drop migliorato
    with st.expander("📤 Carica i tuoi file PDF", expanded=True):
        file_pdf_list = st.file_uploader(
            "Trascina i file qui o fai click per selezionare", 
            type=["pdf"], 
            accept_multiple_files=True,
            help="Supporta più file contemporaneamente. Dimensioni massime: 200MB per file",
            label_visibility="collapsed"
        )
    
    if file_pdf_list:
        # Mostra anteprima file selezionati
        st.subheader("📋 File Selezionati")
        cols = st.columns(4)
        for i, file in enumerate(file_pdf_list):
            cols[i%4].info(f"**{i+1}.** {file.name[:25]}...")
        
        # Elaborazione file
        risultati = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        errori = 0
        
        for i, file in enumerate(file_pdf_list):
            status_text.text(f"🔍 Analisi {i+1}/{len(file_pdf_list)}: {file.name[:30]}...")
            progress_bar.progress((i + 1) / len(file_pdf_list))
            
            try:
                dati = estrai_dati(file)
                if dati:
                    risultati.append(dati)
                else:
                    errori += 1
            except Exception as e:
                logger.error(f"Errore elaborazione {file.name}: {str(e)}", exc_info=True)
                errori += 1
                continue
        
        progress_bar.empty()
        
        if risultati:
            status_text.success(f"✅ Elaborazione completata! {len(risultati)} file processati ({errori} errori)")
            
            # Mostra statistiche rapide
            mostra_statistiche_riepilogo(risultati)
            
            # Mostra tabella risultati
            st.subheader("📋 Dati Estratti")
            
            if raggruppa_societa:
                societa_disponibili = sorted(list(set(d['Società'] for d in risultati if d['Società'] != "N/D")))
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
            else:
                risultati_filtrati = risultati
            
            # Tabella interattiva
            df = pd.DataFrame(risultati_filtrati)
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
                column_order=["Società", "Periodo di Riferimento", "Data Fattura", "Totale (€)", "Consumi", "File"],
                column_config={
                    "Totale (€)": st.column_config.NumberColumn(format="%.2f €"),
                    "Consumi": st.column_config.TextColumn(width="medium")
                }
            )
            
            # Mostra grafici se richiesto
            if mostra_grafici and risultati_filtrati:
                mostra_grafico_consumi(risultati_filtrati)
            
            # Sezione esportazione migliorata
            st.subheader("📤 Esporta Dati")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                excel_data = crea_excel(risultati_filtrati)
                if excel_data:
                    st.download_button(
                        label="💾 Scarica Excel",
                        data=excel_data,
                        file_name="report_consumi.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Scarica tutti i dati in formato Excel",
                        use_container_width=True
                    )
            
            with col2:
                if risultati_filtrati:
                    csv = pd.DataFrame(risultati_filtrati).to_csv(index=False, sep=';').encode('utf-8')
                    st.download_button(
                        label="📝 Scarica CSV",
                        data=csv,
                        file_name="report_consumi.csv",
                        mime="text/csv",
                        help="Scarica i dati in formato CSV (delimitato da punto e virgola)",
                        use_container_width=True
                    )
            
            with col3:
                if mostra_dettagli:
                    with st.expander("Dettagli Tecnici"):
                        st.json(risultati_filtrati[:2])  # Mostra solo i primi 2 per esempio
        else:
            status_text.warning("⚠️ Nessun dato valido estratto dai file caricati")
    
    # Footer migliorato
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; font-size: 14px; color: gray;">
        <strong>Analizzatore Bollette Avanzato</strong><br>
        Supporta i principali fornitori italiani di luce, gas e acqua<br>
        <em>Versione 2.0 - Ottimizzato per precisione e usabilità</em>
    </div>
    """, unsafe_allow_html=True)

def mostra_statistiche_riepilogo(dati: List[Dict[str, str]]):
    """Mostra un riepilogo statistico dei dati estratti."""
    if not dati:
        return
    
    df = pd.DataFrame(dati)
    
    try:
        # Calcola statistiche
        totale = df['Totale (€)'].replace('N/D', '0').astype(float).sum()
        num_bollette = len(df)
        num_societa = df['Società'].nunique()
        
        # Estrai consumi se presenti
        consumi_totali = ""
        if 'Consumi' in df.columns:
            try:
                consumi_df = df[df['Consumi'] != 'N/D'].copy()
                if not consumi_df.empty:
                    consumi_df['Valore'] = consumi_df['Consumi'].str.extract(r'([\d\.]+)')[0].astype(float)
                    consumi_df['Unità'] = consumi_df['Consumi'].str.extract(r'([^\d\s]+)$')[0]
                    
                    if not consumi_df['Unità'].empty:
                        unita = consumi_df['Unità'].mode()[0]
                        somma_consumi = consumi_df['Valore'].sum()
                        consumi_totali = f"{somma_consumi:.2f} {unita}"
            except Exception as e:
                logger.error(f"Errore calcolo consumi: {str(e)}", exc_info=True)
        
        # Mostra metriche
        cols = st.columns(4)
        cols[0].metric("Bollette analizzate", num_bollette)
        cols[1].metric("Società diverse", num_societa)
        cols[2].metric("Importo totale", f"{totale:.2f} €")
        if consumi_totali:
            cols[3].metric("Consumo totale", consumi_totali)
    
    except Exception as e:
        logger.error(f"Errore generazione statistiche: {str(e)}", exc_info=True)

# [Altre funzioni rimangono uguali...]

if __name__ == "__main__":
    main()
