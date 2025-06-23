import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from docx import Document
from io import BytesIO

# --- Estrazione dati dal PDF ---
def estrai_dati_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    testo = "".join(pagina.get_text() for pagina in doc)

    # --- Società: cerca nome preciso (es. AGSM AIM Energia) nel testo iniziale ---
    match_societa = re.search(r'\b(AGSM\s*AIM\s*ENERGIA|AGSM\s*ENERGIA|AIM\s*ENERGIA)\b', testo, re.IGNORECASE)
    nome_societa = match_societa.group(1).upper() if match_societa else "N/D"

    # --- Numero fattura ---
    numero_fattura = re.search(
        r'Numero\s+fattura\s+elettronica\s+valida\s+ai\s+fini\s+fiscali\s*:\s*([A-Z0-9/-]+)', testo, re.IGNORECASE
    )

    # --- Data di chiusura documento ---
    data_chiusura = re.search(
        r'Documento\s+di\s+chiusura.*?([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE
    )

    # --- Periodo di riferimento ---
    periodo = re.search(
        r'dal\s+([0-9]{2}/[0-9]{2}/[0-9]{4})\s+al\s+([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE
    )
    periodo_rif = f"{periodo.group(1)} - {periodo.group(2)}" if periodo else "N/D"

    # --- Totale bolletta ---
    totale_bolletta = re.search(
        r'Totale\s+bolletta\s*[:\-]?\s*€?\s*([\d.,]+)', testo, re.IGNORECASE
    )

    return {
        "societa": nome_societa,
        "periodo_riferimento": periodo_rif,
        "data": data_chiusura.group(1) if data_chiusura else "N/D",
        "pod": "",               # Vuoto
        "dati_cliente": "",      # Vuoto
        "via": "",               # Vuoto
        "numero_fattura": numero_fattura.group(1) if numero_fattura else "N/D",
        "totale_bolletta": totale_bolletta.group(1) if totale_bolletta else "N/D"
    }

# --- Word opzionale ---
def crea_attestazione(dati):
    doc = Document()
    doc.add_heading("Attestazione di Consumo", level=1)
    for k, v in dati.items():
        doc.add_paragraph(f"{k.replace('_', ' ').capitalize()}: {v}")
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- App Streamlit ---
st.set_page_config(page_title="Report Bolletta PDF", layout="centered")
st.title("📄 Report Estratto da Bolletta PDF")

file_pdf = st.file_uploader("Carica la bolletta in PDF", type=["pdf"])

if file_pdf:
    with st.spinner("Estrazione dati in corso..."):
        dati = estrai_dati_da_pdf(file_pdf)

    st.success("✅ Dati estratti!")

    # --- Ordine delle colonne richiesto ---
    colonne_finali = [
        "societa",
        "periodo_riferimento",
        "data",
        "pod",
        "dati_cliente",
        "via",
        "numero_fattura",
        "totale_bolletta"
    ]

    df = pd.DataFrame([dati])[colonne_finali]
    df.columns = [
        "Società",
        "Periodo di Riferimento",
        "Data",
        "POD",
        "Dati Cliente",
        "Via",
        "Numero Fattura",
        "Totale Bolletta (€)"
    ]

    st.subheader("📊 Report Finale (copiabile in Excel)")
    st.dataframe(df, use_container_width=True)

    # Download DOCX
    buffer = crea_attestazione(dati)
    st.download_button(
        label="📥 Scarica Attestazione Word",
        data=buffer,
        file_name="attestazione.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
