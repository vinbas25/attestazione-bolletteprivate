import streamlit as st
import fitz  # PyMuPDF
import re

def estrai_testo_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    testo = "".join(pagina.get_text() for pagina in doc)
    return testo

def estrai_societa(testo):
    # Provo a trovare società AGSM AIM ENERGIA o varianti comuni
    patterns = [
        r'\bAGSM\s*AIM\s*ENERGIA\b',
        r'\bAGSM\s*ENERGIA\b',
        r'\bAIM\s*ENERGIA\b',
        # Aggiungi altre varianti se serve
    ]
    for p in patterns:
        m = re.search(p, testo, re.IGNORECASE)
        if m:
            return m.group(0).upper()
    return "N/D"

def estrai_periodo(testo):
    # Cerco pattern "dal dd/mm/yyyy al dd/mm/yyyy"
    m = re.search(r'dal\s+([0-9]{2}/[0-9]{2}/[0-9]{4})\s+al\s+([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE)
    if m:
        return f"{m.group(1)} - {m.group(2)}"
    return "N/D"

def estrai_data_chiusura(testo):
    # Data chiusura con pattern "Documento di chiusura ... dd/mm/yyyy"
    m = re.search(r'Documento\s+di\s+chiusura.*?([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE)
    return m.group(1) if m else "N/D"

def estrai_numero_fattura(testo):
    m = re.search(r'Numero\s+fattura\s+elettronica\s+valida\s+ai\s+fini\s+fiscali\s*:\s*([A-Z0-9/-]+)', testo, re.IGNORECASE)
    return m.group(1) if m else "N/D"

def estrai_totale_bolletta(testo):
    # Cerca "Totale bolletta" seguito da numero con o senza simbolo €
    m = re.search(r'Totale\s+bolletta\s*[:\-]?\s*€?\s*([\d.,]+)', testo, re.IGNORECASE)
    if m:
        return m.group(1)
    # fallback: cerca solo "Totale" seguito da numero, ma da usare con cautela
    m2 = re.search(r'Totale\s*[:\-]?\s*€?\s*([\d.,]+)', testo, re.IGNORECASE)
    return m2.group(1) if m2 else "N/D"

def estrai_consumi_intelligente(testo):
    # Cerco tutte le occorrenze di "consumo/i fatturato/i"
    pattern = re.compile(r'consumo\s+fatturato|consumi\s+fatturati', re.IGNORECASE)
    lines = testo.split('\n')
    consumi_valori = []

    for i, line in enumerate(lines):
        if pattern.search(line):
            # Prendo linee successive e precedenti (max 3 prima e dopo)
            start = max(0, i - 3)
            end = min(len(lines), i + 4)
            blocco = lines[start:end]
            testo_blocco = " ".join(blocco)

            # Estraggo numeri italiani con separatore migliaia e decimali (es: 1.234,56 o 1234,56)
            numeri = re.findall(r'\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?', testo_blocco)

            for num in numeri:
                try:
                    valore = float(num.replace('.', '').replace(',', '.'))
                    consumi_valori.append(valore)
                except:
                    pass

    return round(sum(consumi_valori), 2) if consumi_valori else "N/D"

def estrai_dati_da_pdf(file):
    testo = estrai_testo_da_pdf(file)

    dati = {
        "Società": estrai_societa(testo),
        "Periodo di Riferimento": estrai_periodo(testo),
        "Data": estrai_data_chiusura(testo),
        "POD": "",
        "Dati Cliente": "",
        "Via": "",
        "Numero Fattura": estrai_numero_fattura(testo),
        "Totale Bolletta (€)": estrai_totale_bolletta(testo),
        "File": "",
        "Consumi": estrai_consumi_intelligente(testo)
    }

    return dati

def mostra_tabella_html(dati):
    html = "<table style='border-collapse: collapse; width: 100%;'>"
    html += "<tr>" + "".join(f"<th style='border: 1px solid black; padding: 6px;'>{col}</th>" for col in dati.keys()) + "</tr>"
    html += "<tr>" + "".join(f"<td style='border: 1px solid black; padding: 6px;'>{val}</td>" for val in dati.values()) + "</tr>"
    html += "</table>"

    st.markdown("### 📋 Copia la tabella qui sotto e incolla direttamente in Excel")
    st.markdown(html, unsafe_allow_html=True)

# -- STREAMLIT UI --

st.set_page_config(page_title="Report Bolletta Intelligente", layout="centered")
st.title("📄 Report Estratto da Bolletta PDF (Estrazione Intelligente)")

file_pdf = st.file_uploader("Carica la bolletta in PDF", type=["pdf"])

if file_pdf:
    with st.spinner("Estrazione dati in corso..."):
        dati = estrai_dati_da_pdf(file_pdf)
    st.success("✅ Dati estratti correttamente!")
    mostra_tabella_html(dati)
