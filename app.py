import streamlit as st
import fitz
import re
import matplotlib.pyplot as plt
import io
from weasyprint import HTML

def extract_text_from_pdf(pdf_file):
    try:
        with fitz.open("pdf", pdf_file.read()) as doc:
            return "\n".join(page.get_text() for page in doc)
    except Exception as e:
        st.error(f"Errore durante la lettura del PDF: {e}")
        return None

def find_value(labels, text):
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
        "Ebitda Coface": find_value(["Ebitda Coface"], text),
    }
    return data

def evaluate_company(financial_data):
    if not financial_data or not any(financial_data.values()):
        return "Dati finanziari non trovati o incompleti nel PDF.", {}

    ricavi = financial_data.get("Ricavi")
    ebitda_coface = financial_data.get("Ebitda Coface")
    rapporto_fatturato_ebitda = ricavi / ebitda_coface if ricavi and ebitda_coface else None

    percentages = {}
    for key, value in financial_data.items():
        if key != "Ricavi" and value and ricavi:
            percentages[key] = (value / ricavi) * 100

    output = "Attivo:\n"
    for key in ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"]:
        val = financial_data.get(key)
        if val is not None:
            output += f"- {key}: € {val:,.2f} ({percentages.get(key, 0):.2f}% dei ricavi)\n"

    output += "\nPassivo:\n"
    for key in ["Acconti / anticipi", "Debiti verso fornitori", "Altri"]:
        val = financial_data.get(key)
        if val is not None:
            output += f"- {key}: € {val:,.2f} ({percentages.get(key, 0):.2f}% dei ricavi)\n"

    output += "\nAltri dati:\n"
    for key in ["Capitale circolante netto", "Ebitda Coface"]:
        val = financial_data.get(key)
        if val is not None:
            output += f"- {key}: € {val:,.2f} ({percentages.get(key, 0):.2f}% dei ricavi)\n"

    output += "\nRapporti:\n"
    cvc = financial_data.get("Crediti verso clienti")
    if cvc and ricavi:
        output += f"- Crediti verso clienti / Fatturato: {(cvc / ricavi) * 100:.2f}%\n"
    if rapporto_fatturato_ebitda:
        output += f"- Fatturato / Ebitda Coface: {rapporto_fatturato_ebitda:.2f}\n"
        output += f"--> Ogni 1000€ non incassati, servono {1000 * rapporto_fatturato_ebitda:.2f}€ di nuovo fatturato per compensare\n"

    return output, percentages, rapporto_fatturato_ebitda

def annotate_bars(ax, bars, value_type="percent"):
    for bar in bars:
        yval = bar.get_height()
        if value_type == "euro":
            label = f'€ {yval:,.2f}'
        elif value_type == "percent":
            label = f'{yval:.2f}%'
        else:
            label = f'{yval:.2f}'
        ax.text(bar.get_x() + bar.get_width() / 2, yval, label, ha='center', va='bottom')

def plot_percent_bars(title, data):
    labels = data.keys()
    values = data.values()
    if values:
        fig, ax = plt.subplots()
        bars = ax.bar(labels, values, color=plt.cm.viridis.colors)
        ax.set_title(title)
        ax.set_ylabel("Percentuale rispetto ai ricavi")
        plt.xticks(rotation=45, ha="right")
        annotate_bars(ax, bars, "percent")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

def plot_aggregated_bar(percentages, financial_data):
    attivo = sum(percentages.get(k, 0) for k in ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"])
    passivo = sum(percentages.get(k, 0) for k in ["Acconti / anticipi", "Debiti verso fornitori", "Altri"])
    ccn = financial_data.get("Capitale circolante netto", 0)
    ricavi = financial_data.get("Ricavi", 1)
    ccn_percent = (ccn / ricavi) * 100 if ricavi else 0

    labels = ['Attivo', 'Passivo', 'Capitale Circolante Netto']
    values = [attivo, passivo, ccn_percent]

    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.bar(labels, values, color=['#1f77b4', '#ff7f0e', '#2ca02c'])
    ax.set_ylabel("Percentuale rispetto ai ricavi")
    ax.set_title("Confronto Attivo, Passivo e CCN")
    annotate_bars(ax, bars, "percent")
    plt.tight_layout()
    return fig

def plot_fatturato_ebitda_histogram(rapporto):
    fig, ax = plt.subplots()
    labels = ['Crediti non incassati', 'Fatturato necessario per compensare']
    values = [1000, rapporto * 1000 if rapporto else 0]
    bars = ax.bar(labels, values, color=['skyblue', 'lightcoral'])
    ax.set_title("Impatto Economico dei Crediti non Incassati")
    ax.set_ylabel("Importi (€)")
    annotate_bars(ax, bars, "euro")
    plt.tight_layout()
    return fig

def extract_report_header(text):
    if not text:
        return None
    words = text.split()
    header_words = words[14:64]
    header = " ".join(header_words)
    match = re.search(r"Partita IVA[:\s]*([\d]+)", header, re.IGNORECASE)
    if match:
        idx = header.find(match.group(0))
        return header[:idx + len(match.group(0))]
    return header

def extract_partita_iva(text):
    if not text:
        return None
    match = re.search(r"Partita IVA[:\s]*([\d]+)", text, re.IGNORECASE)
    return match.group(1).strip() if match else None

def create_pdf_report_weasyprint(report_header, analysis, rapporto_fatturato_ebitda):
    html_content = f"""
    <html><head><style>
        body {{ font-family: 'DejaVu Sans', sans-serif; font-size: 12pt; }}
        h1 {{ font-size: 1.5em; }} h2 {{ font-size: 1.2em; }}
        pre {{ white-space: pre-wrap; word-break: break-word; }}
    </style></head><body>
    <h1>Intestazione Report: {report_header}</h1>
    <h2>Risultati Analisi Finanziaria</h2>
    <pre>{analysis}</pre>
    """
    if rapporto_fatturato_ebitda:
        html_content += f"""
        <h2>Impatto Economico dei Crediti non Incassati</h2>
        <p>{rapporto_fatturato_ebitda:.2f}</p>
        <img src="fatturato_ebitda_hist.png" alt="Grafico">
        """
    html_content += "</body></html>"
    pdf_buffer = HTML(string=html_content).write_pdf()
    return io.BytesIO(pdf_buffer)

# STREAMLIT APP
st.title("Analisi Finanziaria Prospect")
pdf_file = st.file_uploader("Carica il PDF", type=["pdf"])

if st.button("Analizza"):
    if pdf_file:
        text = extract_text_from_pdf(pdf_file)
        if text:
            report_header = extract_report_header(text)
            financial_data = extract_financial_data(text)
            analysis, percentages, rapporto = evaluate_company(financial_data)

            st.subheader(f"Intestazione Report:\n{report_header}")
            st.subheader("Risultati Analisi")
            st.write(analysis)

            if rapporto:
                fig_fatt_ebitda = plot_fatturato_ebitda_histogram(rapporto)
                st.pyplot(fig_fatt_ebitda)
                fig_fatt_ebitda.savefig("fatturato_ebitda_hist.png")

            if percentages:
                st.subheader("Grafici Finanziari")
                keys_attivo = ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"]
                keys_passivo = ["Acconti / anticipi", "Debiti verso fornitori", "Altri"]

                attivo_percentages = {k: percentages[k] for k in keys_attivo if k in percentages}
                passivo_percentages = {k: percentages[k] for k in keys_passivo if k in percentages}

                plot_percent_bars("Analisi Attivo Finanziario", attivo_percentages)

                plot_percent_bars("Analisi Passivo Finanziario", passivo_percentages)

                fig_agg = plot_aggregated_bar(percentages, financial_data)
                st.pyplot(fig_agg)
                plt.close(fig_agg)

                if all(financial_data.get(k) is not None for k in ["Ricavi", "Crediti verso clienti", "Capitale circolante netto"]):
                    fig_confronto, ax_confronto = plt.subplots()
                    labels_confronto = ["Ricavi", "Crediti verso clienti", "Capitale circolante netto"]
                    values_confronto = [financial_data[k] for k in labels_confronto]
                    bars_confronto = ax_confronto.bar(labels_confronto, values_confronto, color=plt.cm.viridis.colors)
                    ax_confronto.set_title("Confronto Fatturato / Crediti Verso Clienti / CCN")
                    ax_confronto.set_ylabel("Importi (€)")
                    annotate_bars(ax_confronto, bars_confronto, "euro")
                    plt.xticks(labels_confronto)
                    plt.tight_layout()
                    st.pyplot(fig_confronto)
                    plt.close(fig_confronto)
        else:
            st.warning("Dati insufficienti per l'analisi.")
    else:
        st.error("Carica un PDF.")
