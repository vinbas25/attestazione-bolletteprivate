import streamlit as st
import fitz  # PyMuPDF
import re
import datetime

# Estrai testo da PDF
def estrai_testo_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    return "".join(page.get_text() for page in doc)

# Estrai società
def estrai_societa(testo):
    match = re.search(r'\bAGSM\s*AIM\s*ENERGIA\b', testo, re.IGNORECASE)
    return match.group(0).upper() if match else "N/D"

# Estrai periodo di riferimento
def estrai_periodo(testo):
    match = re.search(r'dal\s+(\d{2}/\d{2}/\d{4})\s+al\s+(\d{2}/\d{2}/\d{4})', testo, re.IGNORECASE)
    return f"{match.group(1)} - {match.group(2)}" if match else "N/D"

# Mappa mesi in italiano
mesi_map = {
    "gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4, "maggio": 5, "giugno": 6,
    "luglio": 7, "agosto": 8, "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12
}

# Parsing data: accetta giorno, mese (numero o nome), anno
def parse_date(g, m, y):
    try:
        giorno = int(g)
        if m.isdigit():
            mese = int(m)
        else:
            mese = mesi_map.get(m.lower(), 0)
        anno = int(y) if len(y) == 4 else 2000 + int(y)
        if mese == 0:
            return None
        return datetime.date(anno, mese, giorno)
    except:
        return None

# Estrai data fattura con pattern multipli e fallback
def estrai_data_fattura(testo):
    patterns = [
        r'fattura del\s*(\d{1,2})[\/\-\.\s](\w+)[\/\-\.\s](\d{2,4})',
        r'data\s+fattura[:\-]?\s*(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{4})',
        r'data\s+emissione[:\-]?\s*(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{4})',
        r'documento\s+di\s+chiusura.*?(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{4})'
    ]
    for pat in patterns:
        match = re.search(pat, testo, re.IGNORECASE)
        if match:
            groups = match.groups()
            # Qui differenzio tra pattern con mese come nome o numero
            if len(groups) == 3:
                data = parse_date(groups[0], groups[1], groups[2])
                if data:
                    return data.strftime("%d/%m/%Y")
    # Fallback generico per data in formato dd/mm/yyyy o dd-mm-yyyy
    fallback = re.search(r'(\d{2})[\/\-](\d{2})[\/\-](\d{4})', testo)
    if fallback:
        data = parse_date(fallback.group(1), fallback.group(2), fallback.group(3))
        if data:
            return data.strftime("%d/%m/%Y")
    return "N/D"

# Estrai numero fattura con vari pattern e fallback
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
    # Fallback: cerca un codice fattura tipico (es. ABC-1234 o simili)
    match = re.search(r'\b[A-Z]{1,4}[-/]?[0-9]{3,}[-/]?[A-Z0-9]*\b', testo)
    return match.group(0).strip() if match else "N/D"

# Estrai totale bolletta
def estrai_totale_bolletta(testo):
    match = re.search(r'Totale\s+bolletta[:\-]?\s*€?\s*([\d\.,]+)', testo, re.IGNORECASE)
    return match.group(1).replace('.', '').replace(',', '.') if match else "N/D"

# Estrai consumi
def estrai_consumi(testo):
    testo_upper = testo.upper()
    idx = testo_upper.find("RIEPILOGO CONSUMI FATTURATI")
    if idx == -1:
        return "N/D"
    snippet = testo_upper[idx:idx+600]
    match = re.search(r'TOTALE COMPLESSIVO DI[:\-]?\s*([\d\.,]+)', snippet)
    if match:
        try:
            valore = match.group(1).replace('.', '').replace(',', '.')
            return float(valore)
        except:
            return "N/D"
    return "N/D"

# Estrai dati da singolo file PDF
def estrai_dati(file):
    testo = estrai_testo_da_pdf(file)
    return {
        "Società": estrai_societa(testo),
        "Periodo di Riferimento": estrai_periodo(testo),
        "Data": estrai_data_fattura(testo),
        "POD": "",  # da implementare se necessario
        "Dati Cliente": "",  # da implementare se necessario
        "Via": "",  # da implementare se necessario
        "Numero Fattura": estrai_numero_fattura(testo),
        "Totale Bolletta (€)": estrai_totale_bolletta(testo),
        "File": file.name,
        "Consumi": estrai_consumi(testo)
    }

# Mostra tabella finale in Streamlit
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

# Footer semplice
st.markdown("---")
st.markdown("<p style='text-align:center;font-size:14px;color:gray;'>Powered Mar. Vincenzo Basile</p>", unsafe_allow_html=True)
