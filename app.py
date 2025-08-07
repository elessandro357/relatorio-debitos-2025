import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import io

st.set_page_config(layout="wide", page_title="Relatório de Débitos")

st.title("📊 Sistema de Relatórios de Débitos por Secretaria")

uploaded_file = st.file_uploader("📁 Envie a planilha Excel", type="xlsx")

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file)
    
    header_row = 2
    dados_unificados = []

    for i in range(1, df_raw.shape[1], 3):
        secretaria = df_raw.iloc[1, i]
        if pd.isna(secretaria):
            continue

        bloco = df_raw.iloc[header_row + 1:, i:i+3].copy()
        bloco.columns = ["Data", "Fornecedor", "Valor"]
        bloco["Secretaria"] = secretaria
        bloco["CNPJ"] = ""
        dados_unificados.append(bloco)

    df = pd.concat(dados_unificados, ignore_index=True)
    df.dropna(subset=["Data", "Fornecedor", "Valor"], inplace=True)
    df["Data"] = pd.to_datetime(df["Data"], errors='coerce')
    df["Valor"] = pd.to_numeric(df["Valor"], errors='coerce')
    df.dropna(subset=["Data", "Valor"], inplace=True)
    df["Valor"] = df["Valor"].round(2)
    df = df[["Data", "Fornecedor", "CNPJ", "Valor", "Secretaria"]]

    st.subheader("📌 Tabela de Débitos")
    st.dataframe(df, use_container_width=True)

    st.subheader("📈 Gráficos")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**🔹 Débitos por Secretaria**")
        fig1, ax1 = plt.subplots()
        df.groupby("Secretaria")["Valor"].sum().sort_values().plot(kind="barh", ax=ax1)
        ax1.set_xlabel("Valor (R$)")
        st.pyplot(fig1)

    with col2:
        st.markdown("**🔹 Top 10 Fornecedores com Maior Débito**")
        fig2, ax2 = plt.subplots()
        df.groupby("Fornecedor")["Valor"].sum().sort_values(ascending=False).head(10).plot(kind="bar", ax=ax2)
        ax2.set_ylabel("Valor (R$)")
        plt.xticks(rotation=45, ha='right')
        st.pyplot(fig2)

    st.subheader("📥 Downloads")

    # Download PDF
    def gerar_pdf(dataframe):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(200, 10, txt="Relatório de Débitos por Secretaria", ln=True, align="C")
        pdf.set_font("Arial", size=10)
        pdf.ln(10)
        for index, row in dataframe.iterrows():
            linha = f"{row['Data'].strftime('%d/%m/%Y')} | {row['Fornecedor']} | {row['CNPJ']} | R$ {row['Valor']:,.2f} | {row['Secretaria']}"
            pdf.multi_cell(0, 8, linha)
        buffer = io.BytesIO()
        pdf.output(buffer)
        buffer.seek(0)
        return buffer

    pdf_bytes = gerar_pdf(df)
    st.download_button("📄 Baixar Relatório em PDF", data=pdf_bytes, file_name="relatorio_debitos.pdf", mime="application/pdf")

    # Download Excel
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    st.download_button("📊 Baixar Planilha Tratada (.xlsx)", data=excel_buffer, file_name="planilha_tratada.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Download Gráfico 1
    grafico1_buffer = io.BytesIO()
    fig1.savefig(grafico1_buffer, format='png')
    grafico1_buffer.seek(0)
    st.download_button("📉 Baixar Gráfico por Secretaria", data=grafico1_buffer, file_name="grafico_por_secretaria.png", mime="image/png")

    # Download Gráfico 2
    grafico2_buffer = io.BytesIO()
    fig2.savefig(grafico2_buffer, format='png')
    grafico2_buffer.seek(0)
    st.download_button("📊 Baixar Gráfico por Fornecedor", data=grafico2_buffer, file_name="grafico_top_fornecedores.png", mime="image/png")