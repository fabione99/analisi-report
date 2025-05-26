import streamlit as st
import fitz
import re
import matplotlib.pyplot as plt
import io
import pandas as pd

# --- FUNZIONI PDF ---

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
        label = f'{yval:.2f}%' if value_type == "percent" else f'€ {yval:,.2f}'
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

# --- STREAMLIT PDF ---

st.title("📄 Analisi Finanziaria da PDF")
pdf_file = st.file_uploader("Carica il PDF", type=["pdf"], key="pdf_uploader")

if st.button("Analizza PDF"):
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

                if all(financial_data.get(k) is not None for k in ["Ricavi", "Crediti verso clienti", "Capitale circolante netto"]):
                    fig_confronto, ax_confronto = plt.subplots()
                    labels_confronto = ["Ricavi", "Crediti verso clienti", "Capitale circolante netto"]
                    values_confronto = [financial_data[k] for k in labels_confronto]
                    bars_confronto = ax_confronto.bar(labels_confronto, values_confronto, color=plt.cm.viridis.colors)
                    ax_confronto.set_title("Confronto Fatturato / Crediti / CCN")
                    ax_confronto.set_ylabel("Importi (€)")
                    annotate_bars(ax_confronto, bars_confronto, "euro")
                    plt.tight_layout()
                    st.pyplot(fig_confronto)
        else:
            st.warning("Dati insufficienti per l'analisi.")
    else:
        st.error("Carica un PDF.")

# --- STREAMLIT EXCEL ---

st.title("📊 Analisi Portafoglio da Excel")
excel_file = st.file_uploader("Carica il file Excel", type=["xlsx"], key="excel_uploader")

if st.button("Analizza Excel"):
    if excel_file:
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()
        df = df[df['Livello Rischio'].str.strip() != '0-Non disponibile']

        pivot1 = df.pivot_table(index='Livello Rischio', values='Ragione Sociale', aggfunc='count', margins=True, margins_name='Totale')
        pivot1.rename(columns={'Ragione Sociale': 'Numero Aziende'}, inplace=True)

        pivot2 = df.pivot_table(index='Livello Rischio', columns='Late Payment Index', values='Ragione Sociale', aggfunc='count', fill_value=0, margins=True, margins_name='Totale')
        pivot3 = df.pivot_table(index='Livello Rischio', columns='Tipo Valutazione', values='Ragione Sociale', aggfunc='count', fill_value=0, margins=True, margins_name='Totale')
        pivot4 = df.pivot_table(index='Livello Rischio', values=['Ragione Sociale', 'Advanced Opinion'], aggfunc={'Ragione Sociale': 'count', 'Advanced Opinion': 'sum'}, margins=True, margins_name='Totale')
        pivot4.rename(columns={'Ragione Sociale': 'Numero Aziende'}, inplace=True)

        st.subheader("1️⃣ Numero aziende per Livello di Rischio")
        st.dataframe(pivot1)

        st.subheader("2️⃣ Aziende per Rischio e Late Payment Index")
        st.dataframe(pivot2)

        st.subheader("3️⃣ Aziende per Rischio e Tipo di Valutazione")
        st.dataframe(pivot3)

        st.subheader("4️⃣ Numero aziende e Totale Advanced Opinion per Rischio")
        st.dataframe(pivot4)

        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            pivot1.to_excel(writer, sheet_name="Aziende per Rischio")
            pivot2.to_excel(writer, sheet_name="Rischio vs LPI")
            pivot3.to_excel(writer, sheet_name="Rischio vs Valutazione")
            pivot4.to_excel(writer, sheet_name="Rischio + Adv Opinion")
        output_excel.seek(0)

        st.download_button(
            label="📥 Scarica report Excel",
            data=output_excel,
            file_name="analisi_portafoglio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Carica prima un file Excel valido.")
