import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io
from datetime import date

# =========================
# Configuração do app
# =========================
st.set_page_config(layout="wide", page_title="Dashboard de Débitos 2025")
st.title("📊 Dashboard Interativo de Débitos por Secretaria - 2025")
st.caption("Envie a planilha e use os filtros para explorar os dados. Downloads respeitam os filtros aplicados.")

# =========================
# Funções auxiliares
# =========================
@st.cache_data(show_spinner=False)
def load_excel(uploaded_file: io.BytesIO) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file)
    # Normaliza nomes de colunas (remove espaços, caixa alta)
    df.columns = df.columns.str.strip().str.upper()
    return df

def validate_columns(df: pd.DataFrame) -> tuple[bool, list]:
    required = ["DATA", "FORNECEDOR", "CNPJ", "VALOR", "SECRETARIA"]
    missing = [c for c in required if c not in df.columns]
    return (len(missing) == 0, missing)

def cast_types(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce")
    df = df.dropna(subset=["DATA", "FORNECEDOR", "VALOR", "SECRETARIA"])
    df["VALOR"] = df["VALOR"].round(2)
    return df

def gerar_pdf(dataframe: pd.DataFrame) -> io.BytesIO:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="Relatório de Débitos - Dados Filtrados", ln=True, align="C")
    pdf.set_font("Arial", size=10)
    pdf.ln(8)

    if dataframe.empty:
        pdf.multi_cell(0, 8, "Nenhum registro para os filtros selecionados.")
    else:
        # Cabeçalho
        pdf.set_font("Arial", 'B', 10)
        pdf.multi_cell(0, 7, "DATA | FORNECEDOR | CNPJ | VALOR | SECRETARIA")
        pdf.set_font("Arial", size=10)
        pdf.ln(2)

        for _, row in dataframe.iterrows():
            data_txt = row["DATA"].strftime("%d/%m/%Y") if pd.notna(row["DATA"]) else ""
            cnpj_txt = "" if pd.isna(row.get("CNPJ", "")) else str(row.get("CNPJ", ""))
            linha = f"{data_txt} | {row['FORNECEDOR']} | {cnpj_txt} | R$ {row['VALOR']:,.2f} | {row['SECRETARIA']}"
            pdf.multi_cell(0, 7, linha)

    # Converte PDF em bytes (compatível com Streamlit Cloud)
    pdf_bytes = pdf.output(dest="S").encode("latin-1")
    return io.BytesIO(pdf_bytes)

def filtro_multiselect_opcoes(label: str, serie: pd.Series):
    opcoes = sorted(serie.dropna().unique().tolist())
    return st.sidebar.multiselect(label, opcoes)

# =========================
# Upload do arquivo
# =========================
uploaded_file = st.file_uploader(
    "📁 Envie a planilha Excel (colunas necessárias: DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA)",
    type=["xlsx"],
)

if not uploaded_file:
    st.info("Envie uma planilha para começar.")
    st.stop()

# =========================
# Leitura e validação
# =========================
try:
    df_raw = load_excel(uploaded_file)
except Exception as e:
    st.error(f"Não foi possível ler o arquivo. Detalhes: {e}")
    st.stop()

ok, missing = validate_columns(df_raw)
if not ok:
    st.error(f"Planilha inválida. Faltam as colunas: {', '.join(missing)}")
    st.stop()

df = cast_types(df_raw)

if df.empty:
    st.warning("A planilha foi carregada, mas não há linhas válidas após conversão de tipos. Revise os dados.")
    st.stop()

# =========================
# Filtros (barra lateral)
# =========================
st.sidebar.header("🔎 Filtros")

secretarias_sel = filtro_multiselect_opcoes("Secretaria", df["SECRETARIA"])
fornecedores_sel = filtro_multiselect_opcoes("Fornecedor", df["FORNECEDOR"])

data_min_default = pd.to_datetime(df["DATA"].min()).date()
data_max_default = pd.to_datetime(df["DATA"].max()).date()

col_data1, col_data2 = st.sidebar.columns(2)
with col_data1:
    data_ini = st.date_input("Data inicial", data_min_default)
with col_data2:
    data_fim = st.date_input("Data final", data_max_default)

# Garante ordem correta
if data_ini > data_fim:
    st.sidebar.error("A data inicial não pode ser maior que a data final.")
    st.stop()

# Aplica filtros
df_filtrado = df[
    (df["DATA"] >= pd.to_datetime(data_ini)) &
    (df["DATA"] <= pd.to_datetime(data_fim))
].copy()

if secretarias_sel:
    df_filtrado = df_filtrado[df_filtrado["SECRETARIA"].isin(secretarias_sel)]

if fornecedores_sel:
    df_filtrado = df_filtrado[df_filtrado["FORNECEDOR"].isin(fornecedores_sel)]

# =========================
# KPIs
# =========================
col_k1, col_k2, col_k3 = st.columns(3)
total_valor = df_filtrado["VALOR"].sum() if not df_filtrado.empty else 0.0
qtd_linhas = len(df_filtrado)
qtd_fornec = df_filtrado["FORNECEDOR"].nunique()

col_k1.metric("Valor total filtrado", f"R$ {total_valor:,.2f}")
col_k2.metric("Registros", f"{qtd_linhas}")
col_k3.metric("Fornecedores", f"{qtd_fornec}")

st.divider()

# =========================
# Gráficos (Plotly)
# =========================
col_g1, col_g2 = st.columns(2)

with col_g1:
    st.subheader("Débitos por Secretaria")
    if df_filtrado.empty:
        st.info("Sem dados para os filtros selecionados.")
    else:
        graf1 = df_filtrado.groupby("SECRETARIA", as_index=False)["VALOR"].sum().sort_values("VALOR")
        fig1 = px.bar(
            graf1,
            x="VALOR",
            y="SECRETARIA",
            orientation="h",
            text="VALOR",
            color="SECRETARIA",
        )
        fig1.update_traces(texttemplate="R$ %{text:,.2f}", textposition="outside")
        fig1.update_layout(showlegend=False, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig1, use_container_width=True)

with col_g2:
    st.subheader("Top 10 Fornecedores")
    if df_filtrado.empty:
        st.info("Sem dados para os filtros selecionados.")
    else:
        graf2 = (
            df_filtrado.groupby("FORNECEDOR", as_index=False)["VALOR"]
            .sum()
            .sort_values(by="VALOR", ascending=False)
            .head(10)
        )
        fig2 = px.bar(
            graf2,
            x="FORNECEDOR",
            y="VALOR",
            text="VALOR",
            color="FORNECEDOR",
        )
        fig2.update_traces(texttemplate="R$ %{text:,.2f}", textposition="outside")
        fig2.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10, r=10, t=30, b=80))
        st.plotly_chart(fig2, use_container_width=True)

st.divider()

# =========================
# Tabela
# =========================
st.subheader("📋 Dados Filtrados")
st.dataframe(df_filtrado, use_container_width=True)

# =========================
# Downloads (respeitam os filtros)
# =========================
st.subheader("📥 Exportar")

# Excel
excel_buffer = io.BytesIO()
df_filtrado.to_excel(excel_buffer, index=False)
excel_buffer.seek(0)
st.download_button(
    "📊 Baixar Excel",
    data=excel_buffer,
    file_name="dados_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# PDF
pdf_bytes = gerar_pdf(df_filtrado)
st.download_button(
    "📄 Baixar PDF",
    data=pdf_bytes,
    file_name="relatorio_filtrado.pdf",
    mime="application/pdf",
)
