import streamlit as st
import fitz  # PyMuPDF
import re
import datetime
from typing import Optional, Dict, List

# Mappa mesi in italiano
MESI_MAP = {
    "gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4, "maggio": 5, "giugno": 6,
    "luglio": 7, "agosto": 8, "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12
}

# Elenco di società conosciute
SOCIETA_CONOSCIUTE = [
    "AGSM AIM ENERGIA",
    "A2A",
    "ACQUE SPA",
    "AQUEDOTTO DEL FIORA",
    "ASA",
    "FIRENZE ACQUE",
    "GEAL",
    "GAIA SPA",
    "PUBLIACQUA SPA"
]

def estrai_testo_da_pdf(file) -> str:
    """Estrae il testo da un file PDF."""
    try:
        doc = fitz.open(stream=file.read(), filetype="pdf")
        return "".join(page.get_text() for page in doc)
    except Exception as e:
        st.error(f"Errore durante l'estrazione del testo dal PDF: {e}")
        return ""

def estrai_societa(testo: str) -> str:
    """Estrae la società dal testo utilizzando tecniche avanzate."""
    try:
        for societa in SOCIETA_CONOSCIUTE:
            if societa.lower() in testo.lower():
                return societa

        patterns = [
            r'\b([A-Z]{2,}\s*(?:AIM|ENERGIA|S\.?P\.?A\.?|SRL|GREEN|COMM|ACQUE))\b',
            r'\b([A-Z]{2,}\s*(?:ENERGIA|POWER|LIGHT|GAS|ACQUA))\b',
            r'\b(AGSM|A2A|ACQUE|AQUEDOTTO|ASA|FIRENZE|GEAL|GAIA|PUBLIACQUA)\b'
        ]

        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                return match.group(0).upper()
    except Exception as e:
        st.error(f"Errore durante l'estrazione della società: {e}")

    return "N/D"

def estrai_periodo(testo: str) -> str:
    """Estrae il periodo di riferimento dal testo."""
    try:
        match = re.search(r'dal\s+(\d{2}/\d{2}/\d{4})\s+al\s+(\d{2}/\d{2}/\d{4})', testo, re.IGNORECASE)
        if match:
            return f"{match.group(1)} - {match.group(2)}"
    except Exception as e:
        st.error(f"Errore durante l'estrazione del periodo: {e}")

    return "N/D"

def parse_date(g: str, m: str, y: str) -> Optional[datetime.date]:
    """Parsing data: accetta giorno, mese (numero o nome), anno."""
    try:
        giorno = int(g)
        mese = int(m) if m.isdigit() else MESI_MAP.get(m.lower(), 0)
        anno = int(y) if len(y) == 4 else 2000 + int(y)
        if 1 <= mese <= 12:
            return datetime.date(anno, mese, giorno)
    except ValueError as e:
        st.error(f"Errore durante il parsing della data: {e}")

    return None

def estrai_data_fattura(testo: str) -> str:
    """Estrae la data della fattura dal testo."""
    try:
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
                if len(groups) == 3:
                    data = parse_date(groups[0], groups[1], groups[2])
                    if data:
                        return data.strftime("%d/%m/%Y")

        fallback = re.search(r'(\d{2})[\/\-](\d{2})[\/\-](\d{4})', testo)
        if fallback:
            data = parse_date(fallback.group(1), fallback.group(2), fallback.group(3))
            if data:
                return data.strftime("%d/%m/%Y")
    except Exception as e:
        st.error(f"Errore durante l'estrazione della data della fattura: {e}")

    return "N/D"

def estrai_numero_fattura(testo: str) -> str:
    """Estrae il numero della fattura dal testo."""
    try:
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
        if match:
            return match.group(0).strip()
    except Exception as e:
        st.error(f"Errore durante l'estrazione del numero della fattura: {e}")

    return "N/D"

def estrai_totale_bolletta(testo: str) -> str:
    """Estrae il totale della bolletta dal testo."""
    try:
        match = re.search(r'Totale\s+bolletta[:\-]?\s*€?\s*([\d\.,]+)', testo, re.IGNORECASE)
        if match:
            return match.group(1).replace('.', '').replace(',', '.')
    except Exception as e:
        st.error(f"Errore durante l'estrazione del totale della bolletta: {e}")

    return "N/D"

def estrai_consumi(testo: str) -> str:
    """Estrae i consumi dal testo."""
    try:
        testo_upper = testo.upper()
        idx = testo_upper.find("RIEPILOGO CONSUMI FATTURATI")
        if idx == -1:
            return "N/D"

        snippet = testo_upper[idx:idx+600]
        match = re.search(r'TOTALE COMPLESSIVO DI[:\-]?\s*([\d\.,]+)', snippet)
        if match:
            valore = match.group(1).replace('.', '').replace(',', '.')
            return float(valore)
    except Exception as e:
        st.error(f"Errore durante l'estrazione dei consumi: {e}")

    return "N/D"

def estrai_dati(file) -> Dict:
    """Estrae i dati da un singolo file PDF."""
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

def mostra_tabella(dati_lista: List[Dict]) -> None:
    """Mostra la tabella finale in Streamlit."""
    if not dati_lista:
        st.warning("Nessun dato da visualizzare.")
        return

    colonne = list(dati_lista[0].keys())
    html = "<table style='border-collapse:collapse;width:100%'>"
    html += "<tr>" + "".join(f"<th style='border:1px solid #888;padding:6px;background:#eee'>{col}</th>" for col in colonne) + "</tr>"
    for dati in dati_lista:
        html += "<tr>" + "".join(f"<td style='border:1px solid #888;padding:6px'>{dati[col]}</td>" for col in colonne) + "</tr>"
    html += "</table>"

    st.markdown("### 📋 Report finale (puoi copiarlo e incollarlo in Excel):")
    st.markdown(html, unsafe_allow_html=True)

def main():
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

    st.markdown("---")
    st.markdown("<p style='text-align:center;font-size:14px;color:gray;'>Powered by ChatGPT</p>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
