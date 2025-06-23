import streamlit as st
import fitz  # PyMuPDF
import re
import datetime

# Estrae il testo intero dal PDF
def estrai_testo_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    return "".join(p.get_text() for p in doc)

# Trova società
def estrai_societa(testo):
    m = re.search(r'\bAGSM\s*AIM\s*ENERGIA\b', testo, re.IGNORECASE)
    return m.group(0).upper() if m else "N/D"

# Estrae periodo
def estrai_periodo(testo):
    m = re.search(r'dal\s+(\d{2}/\d{2}/\d{4})\s+al\s+(\d{2}/\d{2}/\d{4})', testo, re.IGNORECASE)
    return f"{m.group(1)} - {m.group(2)}" if m else "N/D"

# Parsing di date (formati italiano e mese in lettere)
mesi_map = {
    "gennaio":1, "febbraio":2, "marzo":3, "aprile":4, "maggio":5, "giugno":6,
    "luglio":7, "agosto":8, "settembre":9, "ottobre":10, "novembre":11, "dicembre":12
}
def parse_date(g, m, y):
    try:
        day = int(g)
        month = int(m) if m.isdigit() else mesi_map.get(m.lower(), 0)
        year = int(y) if len(y)==4 else 2000+int(y)
        return datetime.date(year, month, day)
    except:
        return None

# Ricerca data fattura
def estrai_data_fattura(testo):
    patterns = [
        r'fattura del\s*(\d{1,2}[\/\-\.\s](?:\d{1,2}|gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)[\/\-\.\s]\d{2,4})',
        r'data\s+fattura[:\-]?\s*(\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4})',
        r'data\s+emissione[:\-]?\s*(\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4})',
        r'emissione[:\-]?\s*(\d{1,2}[\/\-\.\s](?:\d{1,2}|gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)[\/\-\.\s]\d{2,4})',
        r'documento\s+di\s+chiusura.*?(\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4})'
    ]
    for pat in patterns:
        m = re.search(pat, testo, re.IGNORECASE)
        if m:
            # standardizza
            parts = re.split(r'[\/\-\.\s]+', m.group(1))
            dt = parse_date(parts[0], parts[1], parts[2])
            if dt: return dt.strftime("%d/%m/%Y")
    # fallback: prendi la prima data in formato dd/mm/yyyy o dd-mm-yyyy
    m2 = re.search(r'(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})', testo)
    if m2:
        dt = parse_date(m2.group(1), m2.group(2), m2.group(3))
        if dt: return dt.strftime("%d/%m/%Y")
    return "N/D"

# Estrai numero fattura
def estrai_numero_fattura(testo):
    m = re.search(r'Numero\s+fattura[^0-9A-Z]*(\w+)', testo, re.IGNORECASE)
    return m.group(1) if m else "N/D"

# Estrai totale bolletta
def estrai_totale_bolletta(testo):
    m = re.search(r'Totale\s+bolletta[:\-]?\s*€?\s*([\d\.,]+)', testo, re.IGNORECASE)
    return m.group(1) if m else "N/D"

# Estrai consumi da blocco specifico
def estrai_consumi(testo):
    t = testo.upper()
    idx = t.find("RIEPILOGO CONSUMI FATTURATI")
    if idx==-1: return "N/D"
    snip = t[idx:idx+500]
    m = re.search(r'TOTALE COMPLESSIVO DI[:\-]?\s*([\d\.,]+)', snip)
    if m:
        try:
            return float(m.group(1).replace('.', '').replace(',', '.'))
        except:
            return "N/D"
    return "N/D"

# Main extraction
def estrai_dati(file):
    txt = estrai_testo_da_pdf(file)
    return {
        "Società": estrai_societa(txt),
        "Periodo di Riferimento": estrai_periodo(txt),
        "Data": estrai_data_fattura(txt),
        "POD": "",
        "Dati Cliente": "",
        "Via": "",
        "Numero Fattura": estrai_numero_fattura(txt),
        "Totale Bolletta (€)": estrai_totale_bolletta(txt),
        "File": "",
        "Consumi": estrai_consumi(txt)
    }

def mostra_tabella(d):
    cols, vals = list(d.keys()), list(d.values())
    html = "<table style='border-collapse:collapse;width:100%'>"
    html += "<tr>" + "".join(f"<th style='border:1px solid #888;padding:4px'>{c}</th>" for c in cols) + "</tr>"
    html += "<tr>" + "".join(f"<td style='border:1px solid #888;padding:4px'>{v}</td>" for v in vals) + "</tr>"
    html += "</table>"
    st.markdown("### 📋 Copia/Incolla in Excel")
    st.markdown(html, unsafe_allow_html=True)

# Streamlit UI
st.set_page_config(page_title="Estrazione Bolletta", layout="centered")
st.title("📄 Lettore Bolletta PDF")

f = st.file_uploader("Carica bolletta PDF", type="pdf")
if f:
    with st.spinner("Elaborazione..."):
        dati = estrai_dati(f)
    st.success("✅ Dati pronti")
    mostra_tabella(dati)
