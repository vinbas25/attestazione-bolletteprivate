import os
import re
import fitz  # PyMuPDF
import pandas as pd

# Elenco dei fornitori riconosciuti, inclusa Gaia Spa
FORNITORI = [
    "Enel Energia", "Edison", "A2A", "Hera", "Iren", "Acea",
    "E.ON", "Illumia", "Green Network", "Pulsee", "Wekiwi", "Gaia Spa"
]

# Estrae tutto il testo da un file PDF
def estrai_testo_pdf(percorso_file):
    with fitz.open(percorso_file) as doc:
        return "\n".join([pagina.get_text() for pagina in doc])

# Identifica il fornitore della bolletta
def identifica_fornitore(testo):
    for fornitore in FORNITORI:
        if fornitore.lower() in testo.lower():
            return fornitore
    return "Sconosciuto"

# Estrae POD o PDR e li unifica sotto la stessa colonna
def trova_pod_o_pdr(testo):
    match = re.search(r'(IT\d{14}|IT[A-Z0-9]{14,})', testo)
    if match:
        return match.group()
    match = re.search(r'\b\d{14}\b', testo)
    if match:
        return match.group()
    return "Non trovato"

# Estrae il periodo di fatturazione
def trova_periodo(testo):
    match = re.search(r'(\d{2}/\d{2}/\d{4}).{1,40}?(\d{2}/\d{2}/\d{4})', testo)
    if match:
        return f"{match.group(1)} - {match.group(2)}"
    return "Non trovato"

# Estrae il consumo, anche con pattern alternativi
def trova_consumo(testo):
    pattern_labels = [
        r'Totale consumo fatturato.*?(\d+[.,]?\d*)\s*(kWh|mc)',
        r'RIEPILOGO CONSUMI FATTURATI.*?(\d+[.,]?\d*)\s*(kWh|mc)',
        r'Consumo.*?(\d+[.,]?\d*)\s*(kWh|mc)',
        r'Fatturato.*?(\d+[.,]?\d*)\s*(kWh|mc)'
    ]
    for pattern in pattern_labels:
        match = re.search(pattern, testo, re.IGNORECASE | re.DOTALL)
        if match:
            return f"{match.group(1).replace(',', '.')}".strip()
    return "Non trovato"

# Estrae l’importo totale da pagare
def trova_importo(testo):
    match = re.search(r'Totale da pagare.*?(\d+[.,]\d{2})', testo)
    if match:
        return match.group(1).replace(',', '.')
    return "Non trovato"

# Analizza un singolo file PDF
def analizza_pdf(percorso_file):
    testo = estrai_testo_pdf(percorso_file)
    return {
        "Società": identifica_fornitore(testo),
        "POD": trova_pod_o_pdr(testo),
        "Indirizzo Fornitura": "Non disponibile",  # Placeholder migliorabile
        "Periodo Fatturazione": trova_periodo(testo),
        "Consumo": trova_consumo(testo),
        "Importo Totale (€)": trova_importo(testo)
    }

# Analizza tutti i file PDF presenti in una cartella
def analizza_cartella(cartella):
    risultati = []
    for nome_file in os.listdir(cartella):
        if nome_file.lower().endswith(".pdf"):
            percorso_file = os.path.join(cartella, nome_file)
            dati = analizza_pdf(percorso_file)
            risultati.append(dati)
    return pd.DataFrame(risultati)

# Punto di ingresso principale
if __name__ == "__main__":
    cartella_pdf = "bollette"  # Nome della cartella contenente i PDF
    df = analizza_cartella(cartella_pdf)
    df.to_excel("riepilogo_bollette.xlsx", index=False)
    print("Riepilogo completato. File salvato come 'riepilogo_bollette.xlsx'.")
