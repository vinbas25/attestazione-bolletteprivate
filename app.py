import streamlit as st
import fitz  # PyMuPDF
import re
from docx import Document
from io import BytesIO

# Funzione per estrarre i dati dalla bolletta PDF
def estrai_dati_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    testo = ""
    for pagina in doc:
        testo += pagina.get_text()
    
    # Visualizza il testo per debugging (opzionale)
    # st.text(testo)

    # Regex migliorate e flessibili
    numero_fattura = re.search(r'Numero\s+fattura\s+elettronica\s+valida\s+ai\s+fini\s+fiscali\s*:\s*([A-Z0-9/-]+)', testo, re.IGNORECASE)
    data_chiusura = re.search(r'Documento\s+di\s+chiusura\s+del\s*:?[\s\n]*([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE)
    totale_bolletta = re.search(r'Totale\s+(?:bolletta|da\s+pagare).*?:?\s*€?\s*([\d.,]+)', testo, re.IGNORECASE)

    return {
        "numero_fattura": numero_fattura.group(1) if numero_fattura else "N/D",
        "data_chiusura": data_chiusura.group(1) if data_chiusura else "N/D",
        "totale": totale_bolletta.group(1) if totale_bolletta else "N/D"
    }

# Funzione per creare l'attestazione Word
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

# Configurazione Streamlit
st.set_page_config(page_title="Attestazione Bolletta", layout="centered")
st.title("📄 Generatore di Attestazioni da Bolletta PDF")

# Upload PDF
file_pdf = st.file_uploader("Carica la bolletta in PDF", type=["pdf"])

# Elaborazione
if file_pdf:
    with st.spinner("Estrazione dati dalla bolletta..."):
        dati = estrai_dati_da_pdf(file_pdf)

    st.success("✅ Dati estratti correttamente!")
    st.write(f"**Numero Fattura Elettronica:** {dati['numero_fattura']}")
    st.write(f"**Documento di Chiusura del:** {dati['data_chiusura']}")
    st.write(f"**Totale Bolletta:** € {dati['totale']}")

    buffer = crea_attestazione(dati)

    st.download_button(
        label="📥 Scarica Attestazione",
        data=buffer,
        file_name="attestazione.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
