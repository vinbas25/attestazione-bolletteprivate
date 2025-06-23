import streamlit as st
import fitz  # PyMuPDF
import re
import datetime

def estrai_testo_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    testo = "".join(pagina.get_text() for pagina in doc)
    return testo

def estrai_societa(testo):
    patterns = [
        r'\bAGSM\s*AIM\s*ENERGIA\b',
        r'\bAGSM\s*ENERGIA\b',
        r'\bAIM\s*ENERGIA\b',
    ]
    for p in patterns:
        m = re.search(p, testo, re.IGNORECASE)
        if m:
            return m.group(0).upper()
    return "N/D"

def estrai_periodo(testo):
    m = re.search(r'dal\s+([0-9]{2}/[0-9]{2}/[0-9]{4})\s+al\s+([0-9]{2}/[0-9]{2}/[0-9]{4})', testo, re.IGNORECASE)
    if m:
        return f"{m.group(1)} - {m.group(2)}"
    return "N/D"

def estrai_data_fattura_intelligente(testo):
    testo_lower = testo.lower()

    keywords = [
        "data fattura",
        "data emissione",
        "documento di chiusura",
        "data chiusura",
        "emissione fattura",
        "emissione",
        "fattura",
        "data"
    ]

    pattern_date = r'([0-3]?\d)[/\-\s]([01]?\d|gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)[/\-\s]?(\d{2,4})'

    mesi = {
        "gennaio":1, "febbraio":2, "marzo":3, "aprile":4, "maggio":5, "giugno":6,
        "luglio":7, "agosto":8, "settembre":9, "ottobre":10, "novembre":11, "dicembre":12
    }

    for keyword in keywords:
        for match in re.finditer(keyword, testo_lower):
            start = match.start()
            blocco = testo_lower[start:start+100]

            for m in re.finditer(pattern_date, blocco):
                g, mese, anno = m.groups()

                try:
                    giorno = int(g)
                    if mese.isdigit():
                        mese_num = int(mese)
                    else:
                        mese_num = mesi.get(mese, 0)

                    if len(anno) == 2:
                        anno_num = 2000 + int(anno)
                    else:
                        anno_num = int(anno)

                    if mese_num == 0:
                        continue

                    dt = datetime.date(anno_num, mese_num, giorno)
                    return dt.strftime("%d/%m/%Y")

                except:
                    continue

    m_generale = re.search(r'([0-3]?\d)[/\-]([01]?\d)[/\-](\d{4})', testo_lower)
    if m_generale:
        try:
            giorno = int(m_generale.group(1))
            mese = int(m_generale.group(2))
            anno = int(m_generale.group(3))
            dt = datetime.date(anno, mese, giorno)
            return dt.strftime("%d/%m/%Y")
        except:
            pass

    return "N/D"

def estrai_data_chiusura(testo):
    return estrai_data_fattura_intelligente(testo)

def estrai_numero_fattura(testo):
    m = re.search(r'Numero\s+fattura\s+elettronica\s+valida\s+ai\s+fini\s+fiscali\s*:\s*([A-Z0-9/-]+)', testo, re.IGNORECASE)
    return m.group(1) if m else "N/D"

def estrai_totale_bolletta(testo):
    m = re.search(r'Totale\s+bolletta\s*[:\-]?\s*€?\s*([\d.,]+)', testo, re.IGNORECASE)
    if m:
        return m.group(1)
    m2 = re.search(r'Totale\s*[:\-]?\s*€?\s*([\d.,]+)', testo, re.IGNORECASE)
    return m2.group(1) if m2 else "N/D"

def estrai_consumi_da_riquadro(testo):
    testo_upper = testo.upper()
    consumi_valore = "N/D"

    idx = testo_upper.find("RIEPILOGO CONSUMI FATTURATI")
    if idx == -1:
        return consumi_valore

    snippet = testo_upper[idx:idx+500]

    m = re.search(r'TOTALE COMPLESSIVO DI\s*[:\-]?\s*([\d\.,]+)', snippet, re.IGNORECASE)
    if m:
        numero_str = m.group(1)
        try:
            consumi_valore = float(numero_str.replace('.', '').replace(',', '.'))
        except:
            consumi_valore = "N/D"

    return consumi_valore

def estrai_consumi_intelligente(testo):
    return estrai_consumi_da_riquadro(testo)

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
