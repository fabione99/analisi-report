import streamlit as st
import fitz
import re
import matplotlib.pyplot as plt

def extract_text_from_pdf(pdf_file):
    """Estrae il testo da un PDF con gestione degli errori."""
    try:
        with fitz.open("pdf", pdf_file.read()) as doc:
            return "\n".join(page.get_text() for page in doc)
    except Exception as e:
        st.error(f"Errore durante la lettura del PDF: {e}")
        return None

def find_value(labels, text):
    """Trova un valore numerico dopo una o più label, gestendo formati diversi."""
    if not text:
        return None
    for label in labels:
        match = re.search(fr"{label}[:\s]*([\d.,]+)", text, re.IGNORECASE)
        if match:
            value_str = match.group(1).replace(".", "").replace(",", ".")
            try:
                return float(value_str)
            except ValueError:
                st.warning(f"Valore non numerico trovato per '{label}': {match.group(1)}")
            return None
    return None

def extract_financial_data(text):
    """Estrae i dati finanziari chiave dal testo."""
    if not text:
        return None
    data = {
        "Crediti verso clienti": find_value(["Crediti verso clienti"], text),
        "Rimanenze": find_value(["Rimanenze"], text),
        "Disponibilità liquide": find_value(["Disponibilità liquide"], text),
        "Altre attività correnti": find_value(["Altre attività correnti"], text),
        "Acconti / anticipi": find_value(["Acconti / anticipi"], text),
        "Debiti verso fornitori": find_value(["Debiti verso fornitori"], text),
        "Altri": find_value(["Altri"], text),
        "Ricavi": find_value(["Ricavi"], text),
        "Capitale circolante netto": find_value(["Capitale circolante netto"], text),
    }
    return data

def evaluate_company(financial_data):
    """Genera un'analisi finanziaria e calcola le percentuali."""
    if not financial_data or not any(financial_data.values()):
        return "Dati finanziari non trovati o incompleti nel PDF.", {}

    report = " **Analisi Finanziaria** \n\n"
    ricavi = financial_data.get("Ricavi")
    percentages = {}

    for key, value in financial_data.items():
        if key != "Ricavi" and value is not None and ricavi is not None and ricavi != 0: #rimossa la verifica della Partita IVA
            percent = (value / ricavi) * 100
            percentages[key] = percent
            report += f" **{key}**: € {value:,.2f} ({percent:.2f}% dei ricavi)\n"
        elif key != "Ricavi" and value is not None: #rimossa la verifica della Partita IVA
            report += f" **{key}**: € {value:,.2f} (Ricavi non disponibili per il calcolo della percentuale)\n"

    # Calcola la percentuale di "Crediti verso clienti" sul fatturato
    crediti_clienti = financial_data.get("Crediti verso clienti")
    if crediti_clienti is not None and ricavi is not None and ricavi != 0:
        percent_crediti_fatturato = (crediti_clienti / ricavi) * 100
        report += f" **Crediti verso clienti / Fatturato**: {percent_crediti_fatturato:.2f}%\n"

    # Separa le sezioni attivo e passivo nella descrizione
    report_attivo = "\n **ATTIVO:** \n"
    for key in ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"]:
        if key in financial_data and financial_data[key] is not None:
            report_attivo += f" **{key}**: € {financial_data[key]:,.2f} ({percentages.get(key, 0):.2f}% dei ricavi)\n"

    report_passivo = "\n **PASSIVO:** \n"
    for key in ["Acconti / anticipi", "Debiti verso fornitori", "Altri"]:
        if key in financial_data and financial_data[key] is not None:
            report_passivo += f" **{key}**: € {financial_data[key]:,.2f} ({percentages.get(key, 0):.2f}% dei ricavi)\n"

    report += report_attivo + report_passivo
    return report, percentages

def plot_chart(percentages, chart_type="pie", title=""):
    """Crea un grafico a torta o a barre con estetica migliorata."""
    labels = [k for k, v in percentages.items() if v is not None]
    values = [v for v in percentages.values() if v is not None]

    if not labels:
        return None

    fig, ax = plt.subplots(figsize=(8, 6))
    if chart_type == "pie":
        ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=140, colors=plt.cm.Paired.colors)
        ax.set_title(title, fontsize=14)
    else:
        ax.bar(labels, values, color=plt.cm.viridis.colors)
        ax.set_ylabel("Percentuale rispetto ai ricavi", fontsize=12)
        ax.set_title(title, fontsize=14)
        plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    return fig

def plot_aggregated_bar(percentages, financial_data):
    """Crea un grafico a barre aggregato per confrontare attivo, passivo e capitale circolante netto con valori numerici e spaziatura."""
    attivo_sum = sum(percentages.get(k, 0) for k in ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"])
    passivo_sum = sum(percentages.get(k, 0) for k in ["Acconti / anticipi", "Debiti verso fornitori", "Altri"])
    capitale_circolante_netto = financial_data.get("Capitale circolante netto")
    capitale_circolante_netto_percent = (capitale_circolante_netto / financial_data.get("Ricavi")) * 100 if capitale_circolante_netto is not None and financial_data.get("Ricavi") is not None and financial_data.get("Ricavi") != 0 else 0

    labels = ['Attivo', 'Passivo', 'Capitale Circolante Netto']
    values = [attivo_sum, passivo_sum, capitale_circolante_netto_percent]

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.bar(labels, values, color=['#1f77b4', '#ff7f0e', '#2ca02c'], width=0.6)  # Aggiunto width per spaziatura

    # Aggiunta dei valori numerici sopra le barre
    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, yval + 1, round(yval, 2), ha='center', va='bottom')

    ax.set_ylabel("Percentuale rispetto ai ricavi", fontsize=12)
    ax.set_title("Confronto Attivo, Passivo e Capitale Circolante Netto", fontsize=14)
    plt.xticks(labels)  # Utilizzo delle etichette definite
    plt.tight_layout()
    return fig

def extract_report_header(text):
    """Estrae le prime 50 parole del testo, escludendo le prime 14 e fermandosi dopo la Partita IVA."""
    if not text:
        return None
    words = text.split()
    header_words = words[14:64]  # Prende le parole dalla 15a alla 64a
    header = " ".join(header_words)
    match = re.search(r"Partita IVA[:\s]*([\d]+)", header, re.IGNORECASE)
    if match:
        partita_iva_index = header.find(match.group(0))
        return header[:partita_iva_index + len(match.group(0))]  # Restituisce l'intestazione fino alla Partita IVA
    return header

def extract_partita_iva(text):
    """Estrae la Partita IVA dal testo."""
    if not text:
        return None
    match = re.search(r"Partita IVA[:\s]*([\d]+)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None

st.title(" Analisi Finanziaria Prospect")
pdf_file = st.file_uploader("Carica il PDF", type=["pdf"])

if st.button("Analizza"):
    if pdf_file:
        text = extract_text_from_pdf(pdf_file)
        if text:
            report_header = extract_report_header(text)
            financial_data = extract_financial_data(text)
            analysis, percentages = evaluate_company(financial_data)

            st.subheader(f"Intestazione Report: \n{report_header}")  # visualizzazione intestazione report con a capo
            st.subheader("Risultati Analisi Finanziaria")

            st.write(analysis)
            if percentages:
                st.subheader("Grafici Finanziari")
                attivo_percentages = {k: v for k, v in percentages.items() if k in ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"]}
                passivo_percentages = {k: v for k, v in percentages.items() if k in ["Acconti / anticipi", "Debiti verso fornitori", "Altri"]}

                if attivo_percentages:
                    st.pyplot(plot_chart(attivo_percentages, "pie", "Percentuale attivo rispetto ai ricavi"))
                    st.pyplot(plot_chart(attivo_percentages, "bar", "Analisi attivo finanziario"))

                if passivo_percentages:
                    st.pyplot(plot_chart(passivo_percentages, "pie", "Percentuale passivo rispetto ai ricavi"))
                    st.pyplot(plot_chart(passivo_percentages, "bar", "Analisi passivo finanziario"))

                st.pyplot(plot_aggregated_bar(percentages, financial_data))

        else:
            st.warning("Dati insufficienti per i grafici.")
    else:
        st.error("Carica un PDF.")