import streamlit as st
import pdfplumber
import re
import pandas as pd
from typing import List, Dict, Tuple

# ----------------------- FUNZIONI DI ESTRAZIONE -----------------------

def estrai_testo_da_pdf(file) -> str:
    """Estrae il testo da un file PDF."""
    try:
        with pdfplumber.open(file) as pdf:
            testo = ""
            for pagina in pdf.pages:
                testo += pagina.extract_text() + "\n"
        return testo
    except Exception as e:
        st.error(f"Errore durante l'apertura del PDF: {e}")
        return ""

def estrai_societa(testo: str) -> Tuple[str, str]:
    """Estrae la società e il tipo di fornitura (luce/gas)."""
    testo = testo.lower()
    societa = ""
    tipo_fornitura = ""

    if "enel" in testo:
        societa = "Enel"
    elif "hera" in testo:
        societa = "Hera"
    elif "iren" in testo:
        societa = "Iren"
    elif "acea" in testo:
        societa = "Acea"
    elif "e-on" in testo or "eon" in testo:
        societa = "E.ON"
    elif "edison" in testo:
        societa = "Edison"
    else:
        societa = "Sconosciuta"

    if "energia elettrica" in testo or "luce" in testo:
        tipo_fornitura = "Luce"
    elif "gas naturale" in testo or "gas" in testo:
        tipo_fornitura = "Gas"
    else:
        tipo_fornitura = "Sconosciuto"

    return societa, tipo_fornitura

def estrai_pod_pdr(testo: str) -> Tuple[str, str]:
    """Estrae il codice POD e PDR."""
    pod_match = re.search(r'POD[\s:]*([A-Z0-9]{14,})', testo, re.IGNORECASE)
    pdr_match = re.search(r'PDR[\s:]*([0-9]{10,})', testo, re.IGNORECASE)

    pod = pod_match.group(1).strip() if pod_match else ""
    pdr = pdr_match.group(1).strip() if pdr_match else ""
    return pod, pdr

def estrai_totale_bolletta(testo: str) -> Tuple[str, str]:
    """Estrae l'importo totale e la valuta."""
    match = re.search(r'Totale.*?([\d.,]+)[ ]*(€|EUR)?', testo, re.IGNORECASE)
    if match:
        valore = match.group(1).replace(".", "").replace(",", ".")
        valuta = match.group(2) or "€"
        return valore, valuta
    return "", ""

def estrai_consumi(testo: str) -> str:
    """Estrae il consumo energetico (kWh o Smc)."""
    match = re.search(r'Consumo.*?([\d.,]+)[ ]*(kWh|Smc)', testo, re.IGNORECASE)
    if match:
        return f"{match.group(1)} {match.group(2)}"
    return ""

def estrai_indirizzo(testo: str) -> str:
    """Estrae l'indirizzo di fornitura."""
    match = re.search(r'(?:Fornitura|Fornitura presso|Indirizzo):?\s*(.*?)\n', testo, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def estrai_dati_cliente(testo: str) -> str:
    """Estrae il nome del cliente."""
    match = re.search(r'(?:Intestatario|Cliente|Titolare):?\s*(.*?)\n', testo, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def estrai_data_fattura(testo: str) -> str:
    """Estrae la data della fattura."""
    match = re.search(r'Data fattura[:\s]*([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE)
    return match.group(1) if match else ""

def estrai_numero_fattura(testo: str) -> str:
    """Estrae il numero della fattura."""
    match = re.search(r'Numero fattura[:\s]*([\w/-]+)', testo, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def estrai_periodo(testo: str) -> str:
    """Estrae il periodo di riferimento della bolletta."""
    match = re.search(r'Periodo.*?:?\s*([0-9]{2}/[0-9]{2}/[0-9]{4})\s*-\s*([0-9]{2}/[0-9]{2}/[0-9]{4})', testo)
    return f"{match.group(1)} - {match.group(2)}" if match else ""

# ----------------------- LOGICA DI PARSING FILE -----------------------

def estrai_dati(file) -> Dict:
    """Estrae tutti i dati da un singolo file PDF."""
    testo = estrai_testo_da_pdf(file)
    if not testo:
        return None

    societa, tipo_fornitura = estrai_societa(testo)
    pod, pdr = estrai_pod_pdr(testo)
    totale, valuta = estrai_totale_bolletta(testo)
    consumi = estrai_consumi(testo)
    indirizzo = estrai_indirizzo(testo)
    dati_cliente = estrai_dati_cliente(testo)
    data_fattura = estrai_data_fattura(testo)
    numero_fattura = estrai_numero_fattura(testo)
    periodo_riferimento = estrai_periodo(testo)

    return {
        "Società": societa,
        "Tipo Fornitura": tipo_fornitura,
        "POD": pod,
        "PDR": pdr,
        "Totale": totale,
        "Valuta": valuta,
        "Consumi": consumi,
        "Indirizzo": indirizzo,
        "Dati Cliente": dati_cliente,
        "Data Fattura": data_fattura,
        "Numero Fattura": numero_fattura,
        "Periodo di Riferimento": periodo_riferimento
    }

# ----------------------- INTERFACCIA STREAMLIT -----------------------

st.set_page_config(page_title="Analizzatore Bollette", layout="wide")
st.title("🔍 Analizzatore Bollette PDF")

uploaded_files = st.file_uploader("Carica uno o più PDF di bollette", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.info(f"Hai caricato {len(uploaded_files)} file.")
    risultati = []

    for file in uploaded_files:
        with st.spinner(f"Estrazione dati da: {file.name}"):
            dati = estrai_dati(file)
            if dati:
                dati["File"] = file.name
                risultati.append(dati)
            else:
                st.warning(f"⚠️ Nessun dato estratto da: {file.name}")

    if risultati:
        df = pd.DataFrame(risultati)
        st.success("✅ Estrazione completata!")
        st.dataframe(df)

        # Esporta in Excel o CSV
        col1, col2 = st.columns(2)
        with col1:
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Scarica CSV", data=csv, file_name="dati_bollette.csv", mime="text/csv")

        with col2:
            excel = df.to_excel(index=False, engine='openpyxl')
            st.download_button("📥 Scarica Excel", data=excel, file_name="dati_bollette.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("❌ Nessun dato valido estratto dai PDF caricati.")
else:
    st.info("Carica almeno un file PDF per iniziare.")
