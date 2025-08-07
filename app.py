import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

st.set_page_config(layout="wide", page_title="Dashboard de Débitos 2025")

st.title("📊 Dashboard de Débitos por Secretaria - 2025")

uploaded_file = st.file_uploader("📁 Envie a planilha Excel (com colunas: DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA)", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Normaliza nomes de colunas
    df.columns = df.columns.str.strip().str.upper()

    # Verifica colunas obrigatórias
    required_cols = ["DATA", "FORNECEDOR", "CNPJ", "VALOR", "SECRETARIA"]
    if not all(col in df.columns for col in required_cols):
        st.error(f"❌ A planilha precisa conter as colunas: {', '.join(required_cols)}")
    else:
        # Conversões
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
        df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce")
        df.dropna(subset=["DATA", "FORNECEDOR", "VALOR", "SECRETARIA"], inplace=True)

        # === FILTROS ===
        st.sidebar.header("Filtros")
        secretarias = st.sidebar.multiselect("Selecione a Secretaria", sorted(df["SECRETARIA"].unique()))
        fornecedores = st.sidebar.multiselect("Selecione o Fornecedor", sorted(df["FORNECEDOR"].unique()))
        data_min = st.sidebar.date_input("Data inicial", df["DATA"].min())
        data_max = st.sidebar.date_input("Data final", df["DATA"].max())

        # Aplica filtros
        df_filtrado = df.copy()
        if secretarias:
            df_filtrado = df_filtrado[df_filtrado["SECRETARIA"].isin(secretarias)]
        if fornecedores:
            df_filtrado = df_filtrado[df_filtrado["FORNECEDOR"].isin(fornecedores)]
        df_filtrado = df_filtrado[(df_filtrado["DATA"] >= pd.to_datetime(data_min)) & (df_filtrado["DATA"] <= pd.to_datetime(data_max))]

        # === GRÁFICOS ===
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Débitos por Secretaria")
            graf1 = df_filtrado.groupby("SECRETARIA")["VALOR"].sum().reset_index()
            fig1 = px.bar(graf1, x="VALOR", y="SECRETARIA", orientation="h", text="VALOR",
                          color="SECRETARIA", color_discrete_sequence=px.colors.qualitative.Set2)
            fig1.update_traces(texttemplate='R$ %{text:,.2f}', textposition='outside')
            st.plotly_chart(fig1, use_container_width=True)

        with col2:
            st.subheader("Top 10 Fornecedores")
            graf2 = df_filtrado.groupby("FORNECEDOR")["VALOR"].sum().reset_index().sort_values(by="VALOR", ascending=False).head(10)
            fig2 = px.bar(graf2, x="FORNECEDOR", y="VALOR", text="VALOR",
                          color="FORNECEDOR", color_discrete_sequence=px.colors.qualitative.Set3)
            fig2.update_traces(texttemplate='R$ %{text:,.2f}', textposition='outside')
            fig2.update_xaxes(tickangle=45)
            st.plotly_chart(fig2, use_container_width=True)

        # === TABELA ===
        st.subheader("📋 Dados Filtrados")
        st.dataframe(df_filtrado, use_container_width=True)

        # === DOWNLOADS ===
        st.subheader("📥 Exportar Dados Filtrados")

        # Excel
        excel_buffer = io.BytesIO()
        df_filtrado.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        st.download_button("📊 Baixar Excel", data=excel_buffer, file_name="dados_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF
        def gerar_pdf(dataframe):
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(200, 10, txt="Relatório de Débitos Filtrados", ln=True, align="C")
            pdf.set_font("Arial", size=10)
            pdf.ln(10)
            for _, row in dataframe.iterrows():
                linha = f"{row['DATA'].strftime('%d/%m/%Y')} | {row['FORNECEDOR']} | {row['CNPJ']} | R$ {row['VALOR']:,.2f} | {row['SECRETARIA']}"
                pdf.multi_cell(0, 8, linha)
            buffer = io.BytesIO()
            pdf.output(buffer)
            buffer.seek(0)
            return buffer

        pdf_bytes = gerar_pdf(df_filtrado)
        st.download_button("📄 Baixar PDF", data=pdf_bytes, file_name="relatorio_filtrado.pdf", mime="application/pdf")

else:
    st.info("Envie uma planilha para começar.")
