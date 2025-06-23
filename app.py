import streamlit as st
import fitz
import openai
import os

# CONFIGURA QUI LA TUA API KEY OPENAI
openai.api_key = st.secrets["openai_api_key"] if "openai_api_key" in st.secrets else os.getenv("OPENAI_API_KEY")

def estrai_testo_da_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    return "".join(p.get_text() for p in doc)

def chiedi_all_ai(testo_pdf):
    prompt = f"""
Hai davanti a te il testo estratto da una bolletta PDF. Estrai con attenzione i seguenti dati:
- Società (solo il nome esatto della società che ha emesso la bolletta)
- Periodo di riferimento (indicato con 'dal' e 'al')
- Data fattura (data di emissione o di chiusura documento)
- Numero fattura (codice alfanumerico, se presente)
- Totale bolletta (importo finale da pagare)
- Consumi fatturati complessivi (valore in kWh o metri cubi)

Rispondi solo con i dati, in questo formato:

Società: ...
Periodo di Riferimento: ...
Data: ...
Numero Fattura: ...
Totale Bolletta (€): ...
Consumi: ...
    """

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "Sei un assistente che estrae dati da bollette PDF."},
                {"role": "user", "content": prompt + "\n\n" + testo_pdf}
            ],
            temperature=0
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        return f"Errore durante l'analisi AI: {e}"

# Streamlit UI
st.set_page_config(page_title="Report Consumi AI", layout="wide")
st.title("📊 Report Consumi")

file_pdf_list = st.file_uploader("Carica una o più bollette PDF", type=["pdf"], accept_multiple_files=True)

if file_pdf_list:
    st.markdown("### 🔍 Risultati AI")
    for file in file_pdf_list:
        st.subheader(f"📄 {file.name}")
        testo = estrai_testo_da_pdf(file)
        risposta_ai = chiedi_all_ai(testo)
        st.code(risposta_ai, language="markdown")

st.markdown("---")
st.markdown("<p style='text-align:center;font-size:14px;color:gray;'>Creato dal Mar. Vincenzo Basile</p>", unsafe_allow_html=True)
