import streamlit as st
import fitz  # PyMuPDF
import re
from docx import Document
from io import BytesIO

def estrai_dati_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    testo = ""
    for pagina in doc:
        testo += pagina.get_text()

    numero = re.search(r'Bolletta\s+n°\s*(\d+)', testo)
    data = re.search(r'Data\s+emissione[:\s]*(\d{2}/\d{2}/\d{4})', testo)
    importo = re.search(r'Totale\s+da\s+pagare[:\s]*€?\s*([\d.,]+)', testo)

    return {
        "numero": numero.group(1) if numero else "N/D",
        "data": data.group(1) if data else "N/D",
        "importo": importo.group(1) if importo else "N/D"
    }

def crea_attestazione(dati):
    doc = Document()
    doc.add_heading("Attestazione di Consumo", level=1)
    doc.add_paragraph(f"Numero Bolletta: {dati['numero']}")
    doc.add_paragraph(f"Data Emissione: {dati['data']}")
    doc.add_paragraph(f"Importo: € {dati['importo']}")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="Attestazione Bolletta", layout="centered")

st.title("📄 Generatore di Attestazioni da Bollette PDF")

file_pdf = st.file_uploader("Carica una bolletta in PDF", type=["pdf"])

if file_pdf:
    with st.spinner("Estrazione dati dalla bolletta..."):
        dati = estrai_dati_da_pdf(file_pdf)

    st.success("✅ Dati estratti correttamente!")
    st.write(f"**Numero Bolletta:** {dati['numero']}")
    st.write(f"**Data Emissione:** {dati['data']}")
    st.write(f"**Importo:** € {dati['importo']}")

    buffer = crea_attestazione(dati)

    st.download_button(
        label="📥 Scarica Attestazione",
        data=buffer,
        file_name="attestazione.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )