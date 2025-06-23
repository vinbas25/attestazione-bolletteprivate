# Modifica alla mappa delle società
SOCIETA_CONOSCIUTE = {
    "AGSM AIM ENERGIA": r"AGSM\s*AIM",
    "A2A ENERGIA": r"A2A\s*ENERGIA",
    "ACQUE VERONA": r"ACQUE\s*VERONA",
    "ACQUE SPA": r"ACQUE\s*SPA",
    "AQUEDOTTO DEL FIORA": r"AQUEDOTTO\s*DEL\s*FIORA",
    "ASA LIVORNO": r"ASA\s*LIVORNO",
    "ENEL ENERGIA": r"ENEL\s*ENERGIA",
    "ENI GAS E LUCE": r"ENI\s*GAS\s*E\s*LUCE",
    "GAIA SPA": r"GAIA\s*SPA",  # Aggiunta GAIA SPA
    "HERA COMM": r"HERA\s*COMM",
    "IREN": r"IREN",
    "PUBLIACQUA": r"PUBLIACQUA",
    "SORGENIA": r"SORGENIA",
    "EDISON ENERGIA": r"EDISON\s*ENERGIA"
}

# Modifica alla funzione estrai_societa (rimossa la parte del tipo fornitura)
def estrai_societa(testo: str) -> str:
    """Estrae la società con precisione migliorata."""
    try:
        # Cerca corrispondenza esatta con società conosciute
        for societa, pattern in SOCIETA_CONOSCIUTE.items():
            if re.search(pattern, testo, re.IGNORECASE):
                return societa

        # Pattern generici di fallback
        patterns = [
            r'\b([A-Z]{2,}\s*(?:AIM|ENERGIA|GAS|ACQUA|SPA))\b',
            r'\b(SPA|S\.P\.A\.|SRL|S\.R\.L\.)\b'
        ]

        for pattern in patterns:
            match = re.search(pattern, testo)
            if match:
                return match.group(0).strip()

    except Exception as e:
        st.error(f"Errore durante l'estrazione della società: {str(e)}")

    return "N/D"

# Modifica alla funzione estrai_pod_pdr (ora restituisce un unico codice)
def estrai_pod_pdr(testo: str) -> str:
    """Estrae POD o PDR unificato con pattern specifici."""
    try:
        # Cerca prima POD (ha priorità)
        pod_patterns = [
            r'POD\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Punto\s*di\s*Prelievo\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Codice\s*POD\s*[:\-]?\s*([A-Z0-9]{14,16})'
        ]
        
        for pattern in pod_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                return match.group(1).strip()
                
        # Se non trova POD, cerca PDR
        pdr_patterns = [
            r'PDR\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Punto\s*di\s*Ricerca\s*[:\-]?\s*([A-Z0-9]{14,16})',
            r'Codice\s*PDR\s*[:\-]?\s*([A-Z0-9]{14,16})'
        ]
        
        for pattern in pdr_patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match:
                return match.group(1).strip()
                
    except Exception as e:
        st.error(f"Errore durante l'estrazione del POD/PDR: {str(e)}")
    
    return "N/D"

# Modifica alla funzione estrai_consumi con nuovi pattern
def estrai_consumi(testo: str) -> str:
    """Estrae i consumi con pattern più completi."""
    try:
        # Nuovi pattern aggiunti
        patterns = [
            r'(?:Totale\s*consumo\s*fatturato|RIEPILOGO\s*CONSUMI\s*FATTURATI|consumo\s*periodo)\s*(?:[:\-]?\s*)?([\d\.,]+)\s*(kWh|mc|m³|metri\s*cubi|l|litri)?',
            r'(?:consumo\s*fatturato\s*per\s*il\s*periodo\s*di\s*riferimento)\s*[:\-]?\s*([\d\.,]+)\s*(kWh|mc|m³|metri\s*cubi|l|litri)?',
            r'(?:energia\s*(?:attiva|fatturata)\s*complessiva)\s*[:\-]?\s*([\d\.,]+)\s*(kWh)?',
            r'(?:gas\s*naturale\s*fatturato)\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi)?',
            r'(?:volume\s*acqua\s*fatturato)\s*[:\-]?\s*([\d\.,]+)\s*(mc|m³|metri\s*cubi|l|litri)?'
        ]

        for pattern in patterns:
            match = re.search(pattern, testo, re.IGNORECASE)
            if match and match.group(1):
                try:
                    consumo = float(match.group(1).replace('.', '').replace(',', '.'))
                    unita = match.group(2).lower() if match.group(2) else ""
                    
                    # Normalizza unità di misura
                    if not unita:
                        if "energia" in pattern.lower() or "kwh" in pattern.lower():
                            unita = "kWh"
                        elif "gas" in pattern.lower() or "mc" in pattern.lower() or "m³" in pattern.lower():
                            unita = "mc"
                        elif "acqua" in pattern.lower():
                            unita = "mc"  # Default per acqua
                    
                    return f"{consumo:.2f} {unita}"
                except ValueError:
                    continue

    except Exception as e:
        st.error(f"Errore durante l'estrazione dei consumi: {str(e)}")

    return "N/D"

# Modifica alla funzione estrai_dati per riflettere i cambiamenti
def estrai_dati(file) -> Dict:
    """Estrae tutti i dati da un singolo file PDF."""
    testo = estrai_testo_da_pdf(file)
    if not testo:
        return None
        
    societa = estrai_societa(testo)
    pod = estrai_pod_pdr(testo)
    totale, valuta = estrai_totale_bolletta(testo)
    consumi = estrai_consumi(testo)
    indirizzo = estrai_indirizzo(testo)
    dati_cliente = estrai_dati_cliente(testo)
    
    return {
        "Società": societa,
        "Periodo di Riferimento": estrai_periodo(testo),
        "Data Fattura": estrai_data_fattura(testo),
        "POD": pod,
        "Dati Cliente": dati_cliente,
        "Indirizzo": indirizzo,
        "Numero Fattura": estrai_numero_fattura(testo),
        f"Totale ({valuta})": totale,
        "File": file.name,
        "Consumi": consumi
    }

# Modifica alla funzione crea_excel per aggiornare le colonne
def crea_excel(dati_lista: List[Dict]) -> Optional[BytesIO]:
    """Crea un file Excel con le colonne aggiornate."""
    try:
        # Nuovo ordine colonne senza "Tipo Fornitura" e con POD unificato
        colonne_ordinate = [
            "Società",
            "Periodo di Riferimento",
            "Data Fattura",
            "POD",
            "Dati Cliente",
            "Indirizzo",
            "Numero Fattura",
            "Totale (€)",
            "Consumi",
            "File"
        ]
        
        df = pd.DataFrame([d for d in dati_lista if d is not None])
        
        if len(df) == 0:
            st.warning("Nessun dato valido da esportare")
            return None
            
        # Riordina le colonne
        colonne_presenti = [col for col in colonne_ordinate if col in df.columns]
        df = df[colonne_presenti]
        
        # Resto del codice rimane uguale...
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Report')
            
            workbook = writer.book
            worksheet = writer.sheets['Report']
            
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })
            
            # ... (resto del codice)
            
        output.seek(0)
        return output
        
    except Exception as e:
        st.error(f"Errore durante la creazione del file Excel: {str(e)}")
        return None
