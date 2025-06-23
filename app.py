import streamlit as st
import fitz  # PyMuPDF
import re
import datetime

# Estrai testo da PDF
def estrai_testo_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    return "".join(p.get_text() for p in doc)

# Società
def estrai_societa(testo):
    match = re.search(r'\bAGSM\s*AIM\s*ENERGIA\b', testo, re.IGNORECASE)
    return match.group(0).upper() if match else "N/D"

# Periodo
def estrai_periodo(testo):
    match = re.search(r'dal\s+(\d{2}/\d{2}/\d{4})\s+al\s+(\d{2}/\d{2}/\d{4})', testo, re.IGNORECASE)
    return f"{match.group(1)} - {match.group(2)}" if match else "N/D"

# Mesi
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

# Data fattura
def estrai_data_fattura(testo):
    patterns = [
        r'fattura del\s*(\d{1,2}[\/\-\.\s](\w+)[\/\-\.\s](\d{2,4}))',
        r'data\s+fattura[:\-]?\s*(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{4})',
        r'data\s+emissione[:\-]?\s*(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{4})',
        r'documento\s+di\s+chiusura.*?(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{4})'
    ]
    for pat in patterns:
        match = re.search(pat, testo, re.IGNORECASE)
        if match:
            groups = match.groups()
            data = parse_date(groups[0], groups[1], groups[2])
            if data:
                return data.strftime("%d/%m/%Y")
    fallback = re.search(r'(\d{2})[\/\-](\d{2})[\/\-](\d{4})', testo)
    if fallback:
        data = parse_date(fallback.group(1), fallback.group(2), fallback.group(3))
        if data:
            return data.strftime("%d/%m/%Y")
    return "N/D"

# Numero fattura
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
    match = re.search(r'\b[A-Z]{1,4}[-/]?[0-9]{3,}[-/]?[A-Z0-9]*\b', testo)
    return match.group(0).strip() if match else "N/D"

# Totale bolletta
def estrai_totale_bolletta(testo):
    match = re.search(r'Totale\s+bolletta[:\-]?\s*€?\s*([\d\.,]+)', testo, re.IGNORECASE)
    return match.group(1) if match else "N/D"

# Consumi
def estrai_consumi(testo):
    testo_upper = testo.upper()
    idx = testo_upper.find("RIEPILOGO CONSUMI FATTURATI")
    if idx == -1:
        return "N/D"
    snippet = testo_upper[idx:idx+600]
    match = re.search(r'TOTALE COMPLESSIVO DI[:\-]?\s*([\d\.,]+)', snippet)
    if match:
        try:
            return float(match.group(1).replace('.', '').replace(',', '.'))
        except:
            return "N/D"
    return "N/D"

# Dati da singolo file
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
        "File": file.name,
        "Consumi": estrai_consumi(testo)
    }

# Mostra tabella finale
def mostra_tabella(dati_lista):
    if not dati_lista:
        return
    colonne = list(dati_lista[0].keys())
    html = "<table style='border-collapse:collapse;width:100%'>"
    html += "<tr>" + "".join(f"<th style='border:1px solid #888;padding:6px;background:#eee'>{col}</th>" for col in colonne) + "</tr>"
    for dati in dati_lista:
        html += "<tr>" + "".join(f"<td style='border:1px solid #888;padding:6px'>{dati[col]}</td>" for col in colonne) + "</tr>"
    html += "</table>"
    st.markdown("### 📋 Report finale (puoi copiarlo e incollarlo in Excel):")
    st.markdown(html, unsafe_allow_html=True)

# Streamlit UI
st.set_page_config(page_title="Report Consumi", layout="wide")
st.title("📊 Report Consumi")

file_pdf_list = st.file_uploader("Carica una o più bollette PDF", type=["pdf"], accept_multiple_files=True)

if file_pdf_list:
    risultati = []
    with st.spinner("Estrazione in corso..."):
        for file in file_pdf_list:
            dati = estrai_dati(file)
            risultati.append(dati)
    st.success(f"✅ Elaborati {len(risultati)} file.")
    mostra_tabella(risultati)

# Footer
st.markdown("""---""")
st.markdown("<p style='text-align:center;font-size:14px;color:gray;'>Creato dal Mar. Vincenzo Basile</p>", unsafe_allow_html=True)
