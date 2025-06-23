import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd

# --- Estrazione dati dal PDF ---
def estrai_dati_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    testo = "".join(pagina.get_text() for pagina in doc)

    # Società
    match_societa = re.search(r'\b(AGSM\s*AIM\s*ENERGIA|AGSM\s*ENERGIA|AIM\s*ENERGIA)\b', testo, re.IGNORECASE)
    nome_societa = match_societa.group(1).upper() if match_societa else "N/D"

    # Numero fattura
    numero_fattura = re.search(
        r'Numero\s+fattura\s+elettronica\s+valida\s+ai\s+fini\s+fiscali\s*:\s*([A-Z0-9/-]+)', testo, re.IGNORECASE
    )

    # Data di chiusura
    data_chiusura = re.search(
        r'Documento\s+di\s+chiusura.*?([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE
    )

    # Periodo di riferimento
    periodo = re.search(
        r'dal\s+([0-9]{2}/[0-9]{2}/[0-9]{4})\s+al\s+([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE
    )
    periodo_rif = f"{periodo.group(1)} - {periodo.group(2)}" if periodo else "N/D"

    # Totale bolletta
    totale_bolletta = re.search(
        r'Totale\s+bolletta\s*[:\-]?\s*€?\s*([\d.,]+)', testo, re.IGNORECASE
    )

    # Consumi fatturati (sommati)
    consumi_trovati = re.findall(
        r'consumi\s+fatturati\s*[:\-]?\s*([\d.,]+)', testo, re.IGNORECASE
    )
    consumi_valori = []
    for c in consumi_trovati:
        try:
            valore = float(c.replace(".", "").replace(",", "."))
            consumi_valori.append(valore)
        except:
            continue
    totale_consumi = round(sum(consumi_valori), 2) if consumi_valori else "N/D"

    return {
        "File": "",
        "Società": nome_societa,
        "Periodo di Riferimento": periodo_rif,
        "Data": data_chiusura.group(1) if data_chiusura else "N/D",
        "POD": "",
        "Dati Cliente": "",
        "Via": "",
        "Numero Fattura": numero_fattura.group(1) if numero_fattura else "N/D",
        "Totale Bolletta (€)": totale_bolletta.group(1) if totale_bolletta else "N/D",
        "Consumi": totale_consumi
    }

# --- Visualizza tabella HTML copiabile ---
def mostra_tabella_html(dati):
    html = "<table style='border-collapse: collapse; width: 100%;'>"
    html += "<tr>" + "".join(f"<th style='border: 1px solid black; padding: 4px;'>{col}</th>" for col in dati.keys()) + "</tr>"
    html += "<tr>" + "".join(f"<td style='border: 1px solid black; padding: 4px;'>{val}</td>" for val in dati.values()) + "</tr>"
    html += "</table>"
    st.markdown("### 📋 Copia la tabella qui sotto e incolla in Excel")
    st.markdown(html, unsafe_allow_html=True)

# --- Streamlit App ---
st.set_page_config(page_title="Report Bolletta", layout="centered")
st.title("📄 Report Estratto da Bolletta PDF")

file_pdf = st.file_uploader("Carica la bolletta in PDF", type=["pdf"])

if file_pdf:
    with st.spinner("Estrazione dati in corso..."):
        dati = estrai_dati_da_pdf(file_pdf)

    st.success("✅ Dati estratti correttamente!")

    # ✅ Mostra tabella HTML per copia/incolla
    mostra_tabella_html(dati)
