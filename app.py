import streamlit as st
import fitz  # PyMuPDF
import re
import datetime

# Estrae il testo intero dal PDF
def estrai_testo_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    return "".join(p.get_text() for p in doc)

# Trova società (modificare se necessario per altri fornitori)
def estrai_societa(testo):
    m = re.search(r'\bAGSM\s*AIM\s*ENERGIA\b', testo, re.IGNORECASE)
    return m.group(0).upper() if m else "N/D"

# Estrai periodo
def estrai_periodo(testo):
    m = re.search(r'dal\s+(\d{2}/\d{2}/\d{4})\s+al\s+(\d{2}/\d{2}/\d{4})', testo, re.IGNORECASE)
    return f"{m.group(1)} - {m.group(2)}" if m else "N/D"

# Parsing date
mesi_map = {
    "gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4, "maggio": 5, "giugno": 6,
    "luglio": 7, "agosto": 8, "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12
}
def parse_date(g, m, y):
    try:
        giorno = int(g)
        mese = int(m) if m.isdigit() else mesi_map.get(m.lower(), 0)
        anno = int(y) if len(y) == 4 else 2000 + int(y)
        if mese == 0: return None
        return datetime.date(anno, mese, giorno)
    except:
        return None

# Estrai data fattura (ricerca intelligente)
def estrai_data_fattura(testo):
    patterns = [
        r'fattura del\s*(\d{1,2}[\/\-\.\s](?:\d{1,2}|gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)[\/\-\.\s]\d{2,4})',
        r'data\s+fattura[:\-]?\s*(\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4})',
        r'data\s+emissione[:\-]?\s*(\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4})',
        r'emissione[:\-]?\s*(\d{1,2}[\/\-\.\s](?:\d{1,2}|gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)[\/\-\.\s]\d{2,4})',
        r'documento\s+di\s+chiusura.*?(\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4})'
    ]
    for pattern in patterns:
        match = re.search(pattern, testo, re.IGNORECASE)
        if match:
            parts = re.split(r'[\/\-\.\s]+', match.group(1))
            if len(parts) >= 3:
                data = parse_date(parts[0], parts[1], parts[2])
                if data:
                    return data.strftime("%d/%m/%Y")
    # Fallback: prima data numerica
    fallback = re.search(r'(\d{2})[\/\-](\d{2})[\/\-](\d{4})', testo)
    if fallback:
        data = parse_date(fallback.group(1), fallback.group(2), fallback.group(3))
        if data:
            return data.strftime("%d/%m/%Y")
    return "N/D"

# Estrai numero fattura (intelligente)
def estrai_numero_fattura(testo):
    patterns = [
        r'numero\s+fattura\s+elettronica[^:]*[:\s]\s*([A-Z0-9\-\/]+)',
        r'numero\s+fattura[:\s\-]*([A-Z0-9\-\/]+)',
        r'n°\s*fattura[:\s\-]*([A-Z0-9\-\/]+)',
        r'fattura\s+n\.?\s*([A-Z0-9\-\/]+)',
        r'fattura\s+([A-Z]{1,3}[0-9]{3,})'
    ]
    for pattern in patterns:
        match = re.search(pattern, testo, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    # fallback: qualsiasi codice simile a una fattura
    match = re.search(r'\b[A-Z]{1,4}[-/]?[0-9]{3,}[-/]?[A-Z0-9]*\b', testo)
    return match.group(0).strip() if match else "N/D"

# Estrai totale bolletta
def estrai_totale_bolletta(testo):
    m = re.search(r'Totale\s+bolletta[:\-]?\s*€?\s*([\d\.,]+)', testo, re.IGNORECASE)
    return m.group(1) if m else "N/D"

# Estrai consumi fatturati
def estrai_consumi(testo):
    testo_upper = testo.upper()
    idx = testo_upper.find("RIEPILOGO CONSUMI FATTURATI")
    if idx == -1:
        return "N/D"
    snippet = testo_upper[idx:idx+500]
    match = re.search(r'TOTALE COMPLESSIVO DI[:\-]?\s*([\d\.,]+)', snippet)
    if match:
        try:
            return float(match.group(1).replace('.', '').replace(',', '.'))
        except:
            return "N/D"
    return "N/D"

# Estrazione dati dal file
def estrai_dati(file):
    testo = estrai_testo_da_pdf(file)
    return {
        "Società": estrai_societa(testo),
        "Periodo di Riferimento": estrai_periodo(testo),
        "Data": estrai_data_fattura(testo),
        "POD": "",
        "Dati Cliente": "",
        "Via": "",
        "Numero Fattura": estrai_numero_fattura(testo),
        "Totale Bolletta (€)": estrai_totale_bolletta(testo),
        "File": "",
        "Consumi": estrai_consumi(testo)
    }

# Mostra in tabella HTML copiabile
def mostra_tabella(dati):
    cols = list(dati.keys())
    vals = list(dati.values())
    html = "<table style='border-collapse:collapse;width:100%'>"
    html += "<tr>" + "".join(f"<th style='border:1px solid #888;padding:6px'>{c}</th>" for c in cols) + "</tr>"
    html += "<tr>" + "".join(f"<td style='border:1px solid #888;padding:6px'>{v}</td>" for v in vals) + "</tr>"
    html += "</table>"
    st.markdown("### 📋 Copia la tabella e incolla in Excel")
    st.markdown(html, unsafe_allow_html=True)

# Streamlit UI
st.set_page_config(page_title="Estrazione Bolletta", layout="centered")
st.title("📄 Estrazione Intelligente da Bolletta PDF")

file_pdf = st.file_uploader("Carica una bolletta PDF", type=["pdf"])

if file_pdf:
    with st.spinner("Elaborazione in corso..."):
        dati = estrai_dati(file_pdf)
    st.success("✅ Dati estratti con successo!")
    mostra_tabella(dati)
