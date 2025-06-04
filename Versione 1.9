import streamlit as st
import fitz
import re
import matplotlib.pyplot as plt
import pandas as pd
import io

# --- SELEZIONE LINGUA ---

lang = st.selectbox(
    "🌐 Seleziona la lingua / Select language / Selecciona idioma",
    ["Italiano", "English", "Español"]
)

# --- TRADUZIONI ---

translations = {
    "Italiano": {
        "title": "📄 Analisi Finanziaria da PDF",
        "upload_pdf": "Carica il PDF",
        "analyze": "Analizza PDF",
        "error_pdf": "Errore durante la lettura del PDF",
        "missing_data": "Dati insufficienti per l'analisi.",
        "upload_prompt": "Carica un PDF.",
        "header": "Intestazione Report",
        "results": "Risultati Analisi",
        "assets": "Attivo",
        "liabilities": "Passivo",
        "other_data": "Altri dati",
        "ratios": "Rapporti",
        "chart_assets": "Analisi Attivo Finanziario",
        "chart_liabilities": "Analisi Passivo Finanziario",
        "chart_summary": "Confronto Attivo, Passivo e CCN",
        "chart_impact": "Impatto Economico dei Crediti non Incassati",
        "chart_comparison": "Confronto Fatturato / Crediti / CCN",
        "percent_revenue": "Percentuale rispetto ai ricavi",
        "amounts_eur": "Importi (€)",
        "per_1000_msg": "Ogni 1000€ non incassati, servono {amount:.2f}€ di nuovo fatturato per compensare",
        "uncollected_label": "Crediti non incassati",
        "revenue_needed_label": "Fatturato necessario per compensare",
        "ccn_summary": "Confronto Attivo, Passivo e CCN",
        "turnover_comparison": "Confronto Fatturato / Crediti / CCN",
        "excel_title": "📊 Analisi Portafoglio da Excel",
        "upload_excel": "Carica il file Excel",
        "analyze_excel": "Analizza Excel",
        "warning_upload_excel": "Carica prima un file Excel valido.",
        "section_1": "1️⃣ Numero aziende per Livello di Rischio",
        "section_2": "2️⃣ Aziende per Rischio e Late Payment Index",
        "section_3": "3️⃣ Aziende per Rischio e Tipo di Valutazione",
        "section_4": "4️⃣ Numero aziende e Totale Advanced Opinion per Rischio",
        "section_critical": "🚨 Possibili Casi Critici",
        "download_excel": "📥 Scarica report Excel",
        "company_count": "Numero Aziende",
        "percent_total": "% sul Totale Aziende",
        "total": "Totale"

    },
    "English": {
        "title": "📄 Financial Analysis from PDF",
        "upload_pdf": "Upload PDF",
        "analyze": "Analyze PDF",
        "error_pdf": "Error while reading the PDF",
        "missing_data": "Insufficient data for analysis.",
        "upload_prompt": "Please upload a PDF.",
        "header": "Report Header",
        "results": "Analysis Results",
        "assets": "Assets",
        "liabilities": "Liabilities",
        "other_data": "Other Data",
        "ratios": "Ratios",
        "chart_assets": "Asset Analysis",
        "chart_liabilities": "Liability Analysis",
        "chart_summary": "Comparison: Assets, Liabilities & NWC",
        "chart_impact": "Economic Impact of Uncollected Receivables",
        "chart_comparison": "Comparison: Revenue / Receivables / NWC",
        "percent_revenue": "Percentage of Revenue",
        "amounts_eur": "Amounts (€)",
        "per_1000_msg": "For every €1000 uncollected, you need €{amount:.2f} in new revenue to compensate",
        "uncollected_label": "Uncollected Receivables",
        "revenue_needed_label": "Revenue Needed to Compensate",
        "ccn_summary": "Comparison: Assets, Liabilities & Working Capital",
        "turnover_comparison": "Comparison: Revenue / Receivables / NWC",
        "excel_title": "📊 Portfolio Analysis from Excel",
        "upload_excel": "Upload Excel file",
        "analyze_excel": "Analyze Excel",
        "warning_upload_excel": "Please upload a valid Excel file first.",
        "section_1": "1️⃣ Number of Companies by Risk Level",
        "section_2": "2️⃣ Companies by Risk and Late Payment Index",
        "section_3": "3️⃣ Companies by Risk and Evaluation Type",
        "section_4": "4️⃣ Number of Companies and Total Advanced Opinion by Risk",
        "section_critical": "🚨 Potential Critical Cases",
        "download_excel": "📥 Download Excel Report",
        "company_count": "Number of Companies",
        "percent_total": "% of Total Companies",
        "total": "Total"


    },
    "Español": {
        "title": "📄 Análisis Financiero desde PDF",
        "upload_pdf": "Sube el PDF",
        "analyze": "Analizar PDF",
        "error_pdf": "Error al leer el PDF",
        "missing_data": "Datos insuficientes para el análisis.",
        "upload_prompt": "Por favor, sube un PDF.",
        "header": "Encabezado del Informe",
        "results": "Resultados del Análisis",
        "assets": "Activos",
        "liabilities": "Pasivos",
        "other_data": "Otros Datos",
        "ratios": "Ratios",
        "chart_assets": "Análisis de Activos",
        "chart_liabilities": "Análisis de Pasivos",
        "chart_summary": "Resumen: Activos, Pasivos y CTN",
        "chart_impact": "Impacto Económico de Créditos No Cobrados",
        "chart_comparison": "Comparación: Ingresos / Créditos / CTN",
        "percent_revenue": "Porcentaje sobre ingresos",
        "amounts_eur": "Importes (€)",
        "per_1000_msg": "Por cada €1000 no cobrados, se necesitan €{amount:.2f} de nuevos ingresos para compensar",
        "uncollected_label": "Créditos no cobrados",
        "revenue_needed_label": "Ingresos necesarios para compensar",
        "ccn_summary": "Resumen: Activos, Pasivos y CTN",
        "turnover_comparison": "Comparación: Ingresos / Créditos / CTN",
        "excel_title": "📊 Análisis de Cartera desde Excel",
        "upload_excel": "Sube el archivo Excel",
        "analyze_excel": "Analizar Excel",
        "warning_upload_excel": "Por favor, sube primero un archivo Excel válido.",
        "section_1": "1️⃣ Número de empresas por Nivel de Riesgo",
        "section_2": "2️⃣ Empresas por Riesgo e Índice de Morosidad",
        "section_3": "3️⃣ Empresas por Riesgo y Tipo de Evaluación",
        "section_4": "4️⃣ Número de empresas y Evaluación Avanzada total por Riesgo",
        "section_critical": "🚨 Posibles Casos Críticos",
        "download_excel": "📥 Descargar informe Excel",
        "company_count": "Número de empresas",
        "percent_total": "% del total de empresas",
        "total": "Total"
    }
}



excel_filename_by_lang = {
    "Italiano": "analisi_portafoglio.xlsx",
    "English": "portfolio_analysis.xlsx",
    "Español": "análisis_cartera.xlsx"
}

field_labels = {
    "Italiano": {
        "Crediti verso clienti": "Crediti verso clienti",
        "Rimanenze": "Rimanenze",
        "Disponibilità liquide": "Disponibilità liquide",
        "Altre attività correnti": "Altre attività correnti",
        "Acconti / anticipi": "Acconti / anticipi",
        "Debiti verso fornitori": "Debiti verso fornitori",
        "Altri": "Altri",
        "Ricavi": "Ricavi",
        "Capitale circolante netto": "Capitale circolante netto",
        "Ebitda Coface": "Ebitda Coface",
    },
    "English": {
        "Crediti verso clienti": "Accounts Receivable",
        "Rimanenze": "Stock",
        "Disponibilità liquide": "Cash And Equivalent",
        "Altre attività correnti": "Other Current Assets",
        "Acconti / anticipi": "Advances Received For WIP",
        "Debiti verso fornitori": "Accounts Payable",
        "Altri": "Others",
        "Ricavi": "Revenue",
        "Capitale circolante netto": "Working Capital",
        "Ebitda Coface": "Coface EBITDA",
    },
    "Español": {
        "Crediti verso clienti": "Accounts Receivable",
        "Rimanenze": "Stock",
        "Disponibilità liquide": "Cash And Equivalent",
        "Altre attività correnti": "Otros activos corrientes",
        "Acconti / anticipi": "Advances Received For WIP",
        "Debiti verso fornitori": "Accounts Payable",
        "Altri": "Otros",
        "Ricavi": "Ingresos",
        "Capitale circolante netto": "Capital de trabajo neto",
        "Ebitda Coface": "Ebitda Coface",
    }
}
excel_column_names = {
    "Italiano": {
        "Easynumber": "Easynumber",
        "Rif. Cliente": "Rif. Cliente",
        "Ragione Sociale Validata": "Ragione Sociale Validata",
        "Livello Rischio": "Livello Rischio",
        "Late Payment Index": "Late Payment Index",
        "Tipo Valutazione": "Tipo Valutazione",
        "Segnalazioni Negative": "Segnalazioni Negative",
        "Ragione Sociale": "Ragione Sociale",
        "Advanced Opinion": "Advanced Opinion"
    },
    "English": {
        "Easynumber": "Easynumber",
        "Rif. Cliente": "Customer Ref.",
        "Ragione Sociale Validata": "Verified Name",
        "Livello Rischio": "Risk Level",
        "Late Payment Index": "Late payment index",
        "Tipo Valutazione": "Evaluation Type",
        "Segnalazioni Negative": "Negative Events",
        "Ragione Sociale": "Name",
        "Advanced Opinion": "Advanced Opinion"
    },
    "Español": {
        "Easynumber": "Easynumber",
        "Rif. Cliente": "Referencia del Comprador",
        "Ragione Sociale Validata": "Razón Social Verificada",
        "Livello Rischio": "Nivel de Riesgo",
        "Late Payment Index": "Late Payment Index",
        "Tipo Valutazione": "Tipo de Evaluación",
        "Segnalazioni Negative": "Información negativa",
        "Ragione Sociale": "Razón Social",
        "Advanced Opinion": "Evaluación Avanzada"
    }
}
sheet_names = {
    "Italiano": {
        "pivot1": "Aziende per Rischio",
        "pivot2": "Rischio vs LPI",
        "pivot3": "Rischio vs Valutazione",
        "pivot4": "Rischio + Adv Opinion",
        "casi_critici": "Possibili Casi Critici"
    },
    "English": {
        "pivot1": "Companies by Risk",
        "pivot2": "Risk vs LPI",
        "pivot3": "Risk vs Evaluation",
        "pivot4": "Risk + Adv Opinion",
        "casi_critici": "Potential Critical Cases"
    },
    "Español": {
        "pivot1": "Empresas por Riesgo",
        "pivot2": "Riesgo vs LPI",
        "pivot3": "Riesgo vs Evaluación",
        "pivot4": "Riesgo + Evaluación Avanzada",
        "casi_critici": "Casos Críticos Potenciales"
    }
}
excel_value_map = {
    "Livello Rischio": {
        "1-Molto Basso": {
            "Italiano": "1-Molto Basso",
            "English": "1-Very Low Risk",
            "Español": "1-Riesgo Muy Bajo"
        },
        "2-Basso": {
            "Italiano": "2-Basso",
            "English": "2-Low Risk",
            "Español": "2-Riesgo Bajo"
        },
        "3-Medio Basso": {
            "Italiano": "3-Medio Basso",
            "English": "3-Medium to Low Risk",
            "Español": "3-Riesgo Medio Bajo"
        },
        "4-Medio Alto": {
            "Italiano": "4-Medio Alto",
            "English": "4-Medium to High Risk",
            "Español": "4-Riesgo Medio Alto"
        },
        "5-Alto": {
            "Italiano": "5-Alto",
            "English": "5-High Risk",
            "Español": "5-Riesgo Alto"
        },
        "6-Molto Alto": {
            "Italiano": "6-Molto Alto",
            "English": "6-Very High Risk",
            "Español": "6-Riesgo Muy Alto"
        }
    },
    "Late Payment Index": {
    "Moderata evidenza": {
        "Italiano": "Moderata evidenza",
        "English": "Keep track",
        "Español": "Cierta información negativa"
    },
    "Elevata evidenza": {
        "Italiano": "Elevata evidenza",
        "English": "Warning",
        "Español": "Información negativa considerable"
    },
    "Nessuna evidenza": {
        "Italiano": "Nessuna evidenza",
        "English": "No negative payment",
        "Español": "Sin información negativa en pagos"
    }
    },
    "Tipo Valutazione": {
        "2-Medio Intervento Umano": {
            "Italiano": "2-Medio Intervento Umano",
            "English": "2-Medium Human Activity",
            "Español": "2-Intervención Manual Media"
        },
        "3-Alto Intervento Umano": {
            "Italiano": "3-Alto Intervento Umano",
            "English": "3-High Human Activity",
            "Español": "3-Intervención Manual Alta"
        }
    }
}

t = translations[lang]
f = field_labels[lang]

# --- FUNZIONI ---

def extract_text_from_pdf(pdf_file):
    try:
        with fitz.open("pdf", pdf_file.read()) as doc:
            return "\n".join(page.get_text() for page in doc)
    except Exception as e:
        st.error(f"{t['error_pdf']}: {e}")
        return None

def find_value(labels, text):
    for label in labels:
        match = re.search(fr"{label}[:\s]*([\d.,]+)", text, re.IGNORECASE)
        if match:
            value_str = match.group(1).replace(".", "").replace(",", ".")
            try:
                return float(value_str)
            except ValueError:
                return None
    return None

def extract_financial_data(text, language):
    label_map = {
        "Italiano": {
            "Crediti verso clienti": ["Crediti verso clienti"],
            "Rimanenze": ["Rimanenze"],
            "Disponibilità liquide": ["Disponibilità liquide"],
            "Altre attività correnti": ["Altre attività correnti"],
            "Acconti / anticipi": ["Acconti / anticipi"],
            "Debiti verso fornitori": ["Debiti verso fornitori"],
            "Altri": ["Altri"],
            "Ricavi": ["Ricavi"],
            "Capitale circolante netto": ["Capitale circolante netto"],
            "Ebitda Coface": ["Ebitda Coface"],
        },
        "English": {
            "Crediti verso clienti": ["Accounts Receivable"],
            "Rimanenze": ["Stock"],
            "Disponibilità liquide": ["Cash And Equivalent"],
            "Altre attività correnti": ["Other current assets"],
            "Acconti / anticipi": ["Advances Received For WIP"],
            "Debiti verso fornitori": ["Accounts Payable"],
            "Altri": ["Others"],
            "Ricavi": ["Turnover"],
            "Capitale circolante netto": ["Working Capital"],
            "Ebitda Coface": ["Ebitda Coface"],
        },
        "Español": {
            "Crediti verso clienti": ["Accounts Receivable"],
            "Rimanenze": ["Stock"],
            "Disponibilità liquide": ["Cash And Equivalent"],
            "Altre attività correnti": ["Otros activos corrientes"],
            "Acconti / anticipi": ["Advances Received For WIP"],
            "Debiti verso fornitori": ["Accounts Payable"],
            "Altri": ["Otros"],
            "Ricavi": ["Turnover"],
            "Capitale circolante netto": ["Working Capital"],
            "Ebitda Coface": ["Ebitda Coface"],
        },
    }

    return {field: find_value(labels, text) for field, labels in label_map[language].items()}

def evaluate_company(data):
    output = f"{t['assets']}:\n"
    ricavi = data.get("Ricavi")
    ebitda = data.get("Ebitda Coface")
    rapporto = ricavi / ebitda if ricavi and ebitda else None
    perc = {k: (v / ricavi * 100) for k, v in data.items() if k != "Ricavi" and v and ricavi}

    for k in ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"]:
        if data.get(k) is not None:
            output += f"- {f[k]}: € {data[k]:,.2f} ({perc.get(k, 0):.2f}% {t['percent_revenue'].lower()})\n"

    output += f"\n{t['liabilities']}:\n"
    for k in ["Acconti / anticipi", "Debiti verso fornitori", "Altri"]:
        if data.get(k) is not None:
            output += f"- {f[k]}: € {data[k]:,.2f} ({perc.get(k, 0):.2f}% {t['percent_revenue'].lower()})\n"

    output += f"\n{t['other_data']}:\n"
    for k in ["Capitale circolante netto", "Ebitda Coface"]:
        if data.get(k) is not None:
            output += f"- {f[k]}: € {data[k]:,.2f} ({perc.get(k, 0):.2f}% {t['percent_revenue'].lower()})\n"

    output += f"\n{t['ratios']}:\n"
    if data.get("Crediti verso clienti") and ricavi:
        output += f"- {f['Crediti verso clienti']} / {f['Ricavi']}: {(data['Crediti verso clienti'] / ricavi) * 100:.2f}%\n"
    if rapporto:
        output += f"- {f['Ricavi']} / {f['Ebitda Coface']}: {rapporto:.2f}\n"
        output += f"--> {t['per_1000_msg'].format(amount=rapporto * 1000)}\n"

    return output, perc, rapporto

def plot_percent_bars(title, data):
    fig, ax = plt.subplots()
    labels = [f[k] for k in data.keys()]
    bars = ax.bar(labels, data.values())
    ax.set_title(title)
    ax.set_ylabel(t["percent_revenue"])
    plt.xticks(rotation=45)
    for bar in bars:
        y = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, y, f"{y:.2f}%", ha='center', va='bottom')
    st.pyplot(fig)

def plot_impact_chart(rapporto):
    fig, ax = plt.subplots()
    labels = [t["uncollected_label"], t["revenue_needed_label"]]
    values = [1000, rapporto * 1000]
    bars = ax.bar(labels, values)
    ax.set_title(t["chart_impact"])
    ax.set_ylabel(t["amounts_eur"])
    for bar in bars:
        y = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, y, f"€ {y:,.2f}", ha='center', va='bottom')
    plt.tight_layout()
    st.pyplot(fig)

def plot_summary_chart(percentages, data):
    attivo = sum(percentages.get(k, 0) for k in ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"])
    passivo = sum(percentages.get(k, 0) for k in ["Acconti / anticipi", "Debiti verso fornitori", "Altri"])
    ricavi = data.get("Ricavi", 1)
    ccn = data.get("Capitale circolante netto", 0)
    ccn_percent = (ccn / ricavi) * 100 if ricavi else 0

    labels = [t["assets"], t["liabilities"], f["Capitale circolante netto"]]
    values = [attivo, passivo, ccn_percent]

    fig, ax = plt.subplots()
    bars = ax.bar(labels, values)
    ax.set_title(t["ccn_summary"])
    ax.set_ylabel(t["percent_revenue"])
    for bar in bars:
        y = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, y, f"{y:.2f}%", ha='center', va='bottom')
    plt.tight_layout()
    st.pyplot(fig)

def plot_turnover_comparison_chart(data):
    fig, ax = plt.subplots()
    labels = [f["Ricavi"], f["Crediti verso clienti"], f["Capitale circolante netto"]]
    values = [data.get("Ricavi", 0), data.get("Crediti verso clienti", 0), data.get("Capitale circolante netto", 0)]
    bars = ax.bar(labels, values)
    ax.set_title(t["chart_comparison"])
    ax.set_ylabel(t["amounts_eur"])
    for bar in bars:
        y = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, y, f"€ {y:,.2f}", ha='center', va='bottom')
    plt.tight_layout()
    st.pyplot(fig)

def extract_report_header(text):
    words = text.split()
    header = " ".join(words[14:64])
    match = re.search(r"Partita IVA[:\s]*([\d]+)", header, re.IGNORECASE)
    return header[:header.find(match.group(0)) + len(match.group(0))] if match else header

def translate_excel_values(df, lang):
    # Traduce i valori per Livello Rischio, Late Payment Index e Tipo Valutazione
    for col in ["Livello Rischio", "Late Payment Index", "Tipo Valutazione"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda val: excel_value_map[col].get(val, {}).get(lang, val))
    return df

# --- UI ---

st.title(t["title"])

pdf_file = st.file_uploader(t['upload_pdf'], type=["pdf"], key="pdf_uploader")

if st.button(t['analyze']):
    if pdf_file:
        text = extract_text_from_pdf(pdf_file)
        if text:
            header = extract_report_header(text)
            data = extract_financial_data(text, lang)
            analysis, perc, rapporto = evaluate_company(data)

            st.subheader(f"{t['header']}:\n{header}")
            st.subheader(t["results"])
            st.write(analysis)

            if rapporto:
                plot_impact_chart(rapporto)
            if perc:
                attivo_keys = ["Crediti verso clienti", "Rimanenze", "Disponibilità liquide", "Altre attività correnti"]
                passivo_keys = ["Acconti / anticipi", "Debiti verso fornitori", "Altri"]

                attivo_data = {k: perc[k] for k in attivo_keys if k in perc}
                passivo_data = {k: perc[k] for k in passivo_keys if k in perc}

                plot_percent_bars(t["chart_assets"], attivo_data)
                plot_percent_bars(t["chart_liabilities"], passivo_data)
                plot_summary_chart(perc, data)
                plot_turnover_comparison_chart(data)
        else:
            st.warning(t['missing_data'])
    else:
        st.error(t['upload_prompt'])

# --- STREAMLIT EXCEL ---

st.title(t["excel_title"])
excel_file = st.file_uploader(t["upload_excel"], type=["xlsx"], key="excel_uploader")

if st.button(t["analyze_excel"]):
    if excel_file:
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()

        col_map = excel_column_names[lang]
        sheet_map = sheet_names[lang]
        rischio_col = col_map["Livello Rischio"]

        if rischio_col not in df.columns:
            st.error(f"Colonna '{rischio_col}' non trovata nel file.")
            st.stop()

        # Filtra righe senza rischio valido
        df = df[df[rischio_col].fillna("").astype(str).str.strip() != ""]
        df = df[df[rischio_col].astype(str).str.strip() != "0-Non disponibile"]

        totale_numero_aziende = len(df)

        # Pivot 1
        pivot1 = df.pivot_table(
            index=rischio_col,
            values=col_map["Ragione Sociale"],
            aggfunc='count',
            margins=True,
            margins_name=t["total"]
        )
        pivot1.rename(columns={col_map["Ragione Sociale"]: t["company_count"]}, inplace=True)
        pivot1[t["percent_total"]] = (pivot1[t["company_count"]] / totale_numero_aziende * 100).round(2).apply(lambda x: f"{x:.2f}".replace(".", ",") + '%')
        if t["total"] in pivot1.index:
            pivot1.loc[t["total"], t["percent_total"]] = "100,00%"

        # Pivot 2
        pivot2_count = df.pivot_table(
            index=rischio_col,
            columns=col_map["Late Payment Index"],
            values=col_map["Ragione Sociale"],
            aggfunc='count',
            fill_value=0,
            margins=True,
            margins_name=t["total"]
        )
        pivot2_percent = pivot2_count.applymap(
            lambda x: f"{(x / totale_numero_aziende * 100):.2f}".replace(".", ",") + "%")

        # Ordine desiderato delle colonne (tradotte)
        lpi_order_keys = ["Elevata evidenza", "Moderata evidenza", "Nessuna evidenza"]
        lpi_order_translated = [excel_value_map["Late Payment Index"].get(k, {}).get(lang, k) for k in lpi_order_keys]

        # Ricostruzione pivot2 con colonne ordinate e Totale/%
        pivot2_combined = pd.DataFrame()

        for col in lpi_order_translated:
            if col in pivot2_count.columns:
                pivot2_combined[col] = pivot2_count[col]
                pivot2_combined[f"{col} (%)"] = pivot2_percent[col]

        # ➕ Aggiungi Totale (colonna) come penultima e % come ultima
        if t["total"] in pivot2_count.columns:
            pivot2_combined[t["total"]] = pivot2_count[t["total"]]
            pivot2_combined[f"{t['total']} (%)"] = pivot2_percent[t["total"]]

        # ➕ Se presente la riga Totale, forza 100,00%
        if t["total"] in pivot2_combined.index:
            pivot2_combined.loc[t["total"], f"{t['total']} (%)"] = "100,00%"

        # Pivot 3
        pivot3_count = df.pivot_table(
            index=rischio_col,
            columns=col_map["Tipo Valutazione"],
            values=col_map["Ragione Sociale"],
            aggfunc='count',
            fill_value=0,
            margins=True,
            margins_name=t["total"]
        )
        pivot3_percent = pivot3_count.applymap(lambda x: f"{(x / totale_numero_aziende * 100):.2f}".replace(".", ",") + "%")
        pivot3_combined = pd.DataFrame()
        for col in pivot3_count.columns:
            pivot3_combined[col] = pivot3_count[col]
            pivot3_combined[f"{col} (%)"] = pivot3_percent[col]
        if t["total"] in pivot3_combined.index:
            for col in pivot3_count.columns:
                valore = pivot3_combined.loc[t["total"], col]
                percentuale = (valore / totale_numero_aziende * 100)
                pivot3_combined.loc[t["total"], f"{col} (%)"] = f"{percentuale:.2f}".replace(".", ",") + "%"

        # Pivot 4
        pivot4 = df.pivot_table(
            index=rischio_col,
            values=[col_map["Ragione Sociale"], col_map["Advanced Opinion"]],
            aggfunc={col_map["Ragione Sociale"]: "count", col_map["Advanced Opinion"]: "sum"},
            margins=True,
            margins_name=t["total"]
        )
        pivot4.rename(columns={col_map["Ragione Sociale"]: t["company_count"]}, inplace=True)
        pivot4[col_map["Advanced Opinion"]] = pivot4[col_map["Advanced Opinion"]].apply(
            lambda x: f"€ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        # Visualizzazione
        st.subheader(t["section_1"])
        st.dataframe(pivot1)

        st.subheader(t["section_2"])
        st.dataframe(pivot2_combined)

        st.subheader(t["section_3"])
        st.dataframe(pivot3_combined)

        st.subheader(t["section_4"])
        st.dataframe(pivot4)

        st.subheader(t["section_critical"])

        livelli_critici = [excel_value_map["Livello Rischio"][k][lang] for k in ["5-Alto", "6-Molto Alto"]]
        lpi_critici = [excel_value_map["Late Payment Index"][k][lang] for k in
                       ["Elevata evidenza", "Moderata evidenza"]]
        tipo_valutazione_critica = [excel_value_map["Tipo Valutazione"][k][lang] for k in
                                    ["2-Medio Intervento Umano", "3-Alto Intervento Umano"]]

        casi_critici = df[
            df[rischio_col].isin(livelli_critici) &
            df[col_map["Late Payment Index"]].isin(lpi_critici) &
            df[col_map["Tipo Valutazione"]].isin(tipo_valutazione_critica)
            ]

        # Ordina per Late Payment Index
        lpi_order = [excel_value_map["Late Payment Index"][k][lang] for k in ["Elevata evidenza", "Moderata evidenza"]]
        casi_critici["_lpi_order"] = casi_critici[col_map["Late Payment Index"]].apply(
            lambda x: lpi_order.index(x) if x in lpi_order else 99)
        casi_critici = casi_critici.sort_values("_lpi_order").drop(columns="_lpi_order")

        colonne_output = [
            col_map["Easynumber"],
            col_map["Rif. Cliente"],
            col_map["Ragione Sociale Validata"],
            col_map["Livello Rischio"],
            col_map["Late Payment Index"],
            col_map["Tipo Valutazione"],
            col_map["Segnalazioni Negative"]
        ]

        st.dataframe(casi_critici[colonne_output])
        from xlsxwriter.utility import xl_col_to_name

        # Accesso al workbook e worksheet
        # Output Excel con fogli tradotti
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            pivot1.to_excel(writer, sheet_name=sheet_map["pivot1"])
            pivot2_combined.to_excel(writer, sheet_name=sheet_map["pivot2"])
            pivot3_combined.to_excel(writer, sheet_name=sheet_map["pivot3"])
            pivot4.to_excel(writer, sheet_name=sheet_map["pivot4"])
            casi_critici[colonne_output].to_excel(writer, sheet_name=sheet_map["casi_critici"], index=False)

            # 💡 Inserisci la formattazione qui dentro!
            workbook = writer.book

            # Mappa dei colori rischio
            rischio_colori = {
                "1-Molto Basso": "#479D72",
                "2-Basso": "#99CC40",
                "3-Medio Basso": "#81EC85",
                "4-Medio Alto": "#D4F1A5",
                "5-Alto": "#FF9900",
                "6-Molto Alto": "#FF0000",
            }

            from xlsxwriter.utility import xl_col_to_name


            # Funzione che colora la colonna "Livello Rischio" se presente
            def colora_colonna_rischio(df_sorgente, nome_foglio):
                if col_map["Livello Rischio"] not in df_sorgente.columns:
                    return
                col_index = list(df_sorgente.columns).index(col_map["Livello Rischio"])
                col_letter = xl_col_to_name(col_index)
                worksheet = writer.sheets[nome_foglio]

                for rischio_codice, colore in rischio_colori.items():
                    localizzato = excel_value_map["Livello Rischio"].get(rischio_codice, {}).get(lang)
                    if localizzato:
                        formato = workbook.add_format({"bg_color": colore})
                        worksheet.conditional_format(
                            f"{col_letter}2:{col_letter}1000",
                            {
                                "type": "cell",
                                "criteria": "==",
                                "value": f'"{localizzato}"',
                                "format": formato,
                            }
                        )


            # Applica la colorazione a tutti i fogli
            colora_colonna_rischio(pivot1.reset_index(), sheet_map["pivot1"])
            colora_colonna_rischio(pivot2_combined.reset_index(), sheet_map["pivot2"])
            colora_colonna_rischio(pivot3_combined.reset_index(), sheet_map["pivot3"])
            colora_colonna_rischio(pivot4.reset_index(), sheet_map["pivot4"])
            colora_colonna_rischio(casi_critici[colonne_output], sheet_map["casi_critici"])

        output_excel.seek(0)

        st.download_button(
            label=t["download_excel"],
            data=output_excel,
            file_name=excel_filename_by_lang[lang],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning(t["warning_upload_excel"])
