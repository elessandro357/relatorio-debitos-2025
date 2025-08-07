import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import io

st.set_page_config(layout="wide", page_title="Relatório de Débitos 2025")
st.title("📊 Sistema de Relatórios de Débitos por Secretaria")

uploaded_file = st.file_uploader("📁 Envie a planilha Excel no formato correto", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # Verifica se todas as colunas necessárias existem
        expected_columns = ["DATA", "FORNECEDOR", "CNPJ", "VALOR", "SECRETARIA"]
        if not all(col in df.columns for col in expected_columns):
            st.error(f"A planilha precisa conter as colunas: {', '.join(expected_columns)}")
        else:
            # Conversões
            df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
            df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce")
            df = df.dropna(subset=["DATA", "FORNECEDOR", "VALOR", "SECRETARIA"])
            df["VALOR"] = df["VALOR"].round(2)

            st.subheader("📌 Tabela de Débitos")
            st.dataframe(df, use_container_width=True)

            st.subheader("📈 Gráficos")
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**🔹 Débitos por Secretaria**")
                fig1, ax1 = plt.subplots()
                df.groupby("SECRETARIA")["VALOR"].sum().sort_values().plot(kind="barh", ax=ax1)
                ax1.set_xlabel("Valor (R$)")
                st.pyplot(fig1)

            with col2:
                st.markdown("**🔹 Top 10 Fornecedores com Maior Débito**")
                fig2, ax2 = plt.subplots()
                df.groupby("FORNECEDOR")["VALOR"].sum().sort_values(ascending=False).head(10).plot(kind="bar", ax=ax2)
                ax2.set_ylabel("Valor (R$)")
                plt.xticks(rotation=45, ha='right')
                st.pyplot(fig2)

            st.subheader("📥 Downloads")

            def gerar_pdf(dataframe):
                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                pdf.set_font("Arial", 'B', 14)
                pdf.cell(200, 10, txt="Relatório de Débitos 2025", ln=True, align="C")
                pdf.set_font("Arial", size=10)
                pdf.ln(10)
                for index, row in dataframe.iterrows():
                    linha = f"{row['DATA'].strftime('%d/%m/%Y')} | {row['FORNECEDOR']} | {row['CNPJ']} | R$ {row['VALOR']:,.2f} | {row['SECRETARIA']}"
                    pdf.multi_cell(0, 8, linha)
                buffer = io.BytesIO()
                pdf.output(buffer)
                buffer.seek(0)
                return buffer

            # PDF
            pdf_bytes = gerar_pdf(df)
            st.download_button("📄 Baixar Relatório em PDF", data=pdf_bytes, file_name="relatorio_debitos_2025.pdf", mime="application/pdf")

            # Excel
            excel_buffer = io.BytesIO()
            df.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            st.download_button("📊 Baixar Planilha Tratada (.xlsx)", data=excel_buffer, file_name="planilha_tratada.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # Gráfico por Secretaria
            grafico1_buffer = io.BytesIO()
            fig1.savefig(grafico1_buffer, format='png')
            grafico1_buffer.seek(0)
            st.download_button("📉 Baixar Gráfico por Secretaria", data=grafico1_buffer, file_name="grafico_por_secretaria.png", mime="image/png")

            # Gráfico por Fornecedor
            grafico2_buffer = io.BytesIO()
            fig2.savefig(grafico2_buffer, format='png')
            grafico2_buffer.seek(0)
            st.download_button("📊 Baixar Gráfico por Fornecedor", data=grafico2_buffer, file_name="grafico_top_fornecedores.png", mime="image/png")

    except Exception as e:
        st.error(f"Erro ao processar a planilha: {str(e)}")

else:
    st.info("Envie uma planilha Excel com as colunas: DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA.")
