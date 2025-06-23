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

    # Estrazioni con regex
    nome_societa = re.search(
        r'(?:Intestatario|Ragione\s+sociale|Cliente|Società)\s*[:\-]?\s*(.+)', testo, re.IGNORECASE
    )
    numero_fattura = re.search(
        r'Numero\s+fattura\s+elettronica\s+valida\s+ai\s+fini\s+fiscali\s*:\s*([A-Z0-9/-]+)', testo, re.IGNORECASE
    )
    data_chiusura = re.search(
        r'Documento\s+di\s+chiusura.*?([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE
    )

    # Totale per Word (non mostrato nel report finale)
    totale_bolletta = re.search(
        r'Totale\s+(?:bolletta|da\s+pagare).*?:?\s*€?\s*([\d.,]+)', testo, re.IGNORECASE
    )

    return {
        "nome_societa": nome_societa.group(1).strip() if nome_societa else "N/D",
        "data_chiusura": data_chiusura.group(1) if data_chiusura else "N/D",
        "numero_fattura": numero_fattura.group(1) if numero_fattura else "N/D",
        "totale": totale_bolletta.group(1) if totale_bolletta else "N/D"
    }

# --- Creazione attestazione Word (facoltativa) ---
def crea_attestazione(dati):
    doc = Document()
    doc.add_heading("Attestazione di Consumo", level=1)
    doc.add_paragraph(f"Numero Fattura: {dati['numero_fattura']}")
    doc.add_paragraph(f"Documento di Chiusura del: {dati['data_chiusura']}")
    doc.add_paragraph(f"Totale Bolletta: € {dati['totale']}")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Interfaccia Streamlit ---
st.set_page_config(page_title="Report Bolletta PDF", layout="centered")
st.title("📄 Generatore Report Bolletta")

file_pdf = st.file_uploader("Carica una bolletta in PDF", type=["pdf"])

if file_pdf:
    with st.spinner("Estrazione dati in corso..."):
        dati = estrai_dati_da_pdf(file_pdf)

    st.success("✅ Dati estratti con successo!")

    # --- Mostra report in tabella copiabile in Excel ---
    df = pd.DataFrame([{
        "Nome Società": dati["nome_societa"],
        "Documento di Chiusura del (Data Fattura)": dati["data_chiusura"],
        "Numero Fattura": dati["numero_fattura"]
    }])

    st.subheader("📊 Report Finale (puoi copiarlo in Excel)")
    st.dataframe(df, use_container_width=True)

    # --- Download DOCX opzionale ---
    buffer = crea_attestazione(dati)
    st.download_button(
        label="📥 Scarica Attestazione Word",
        data=buffer,
        file_name="attestazione.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
