import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

# =========================================
# Configuração do app
# =========================================
st.set_page_config(layout="wide", page_title="Dashboard de Débitos 2025")
st.title("📊 Dashboard Interativo de Débitos por Secretaria - 2025")
st.caption("Envie a planilha e use os filtros para explorar os dados. Downloads respeitam os filtros aplicados.")

# =========================================
# Utilidades
# =========================================
def format_brl(valor):
    """Formata número como Real brasileiro (R$ 1.234,56) sem depender de locale do SO."""
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return valor

@st.cache_data(show_spinner=False)
def load_excel(uploaded_file: io.BytesIO) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file)
    # Normaliza nomes de colunas (remove espaços, caixa alta)
    df.columns = df.columns.str.strip().str.upper()
    return df

def validate_columns(df: pd.DataFrame):
    required = ["DATA", "FORNECEDOR", "CNPJ", "VALOR", "SECRETARIA"]
    missing = [c for c in required if c not in df.columns]
    return (len(missing) == 0, missing)

def cast_types(df: pd.DataFrame) -> pd.DataFrame:
    """Conversões robustas: DATA tenta padrão e dayfirst; VALOR aceita '1.234,56'."""
    df = df.copy()

    # DATA — tenta padrão; se falhar, tenta dayfirst=True
    d1 = pd.to_datetime(df["DATA"], errors="coerce")
    d2 = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
    df["DATA"] = d1.fillna(d2)

    # VALOR — tenta direto; se falhar, tenta parse BR (1.234,56)
    v1 = pd.to_numeric(df["VALOR"], errors="coerce")
    precisa_parse_brl = v1.isna() & df["VALOR"].astype(str).str.contains(r"[.,]", na=False)
    v2 = pd.to_numeric(
        df.loc[precisa_parse_brl, "VALOR"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    v1.loc[precisa_parse_brl] = v2
    df["VALOR"] = v1

    # Limpeza mínima de texto
    df["FORNECEDOR"] = df["FORNECEDOR"].astype(str).str.strip()
    df["SECRETARIA"] = df["SECRETARIA"].astype(str).str.strip()

    # Remove só o que é realmente inviável
    df = df.dropna(subset=["DATA", "VALOR", "FORNECEDOR", "SECRETARIA"]).copy()
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
            valor_txt = format_brl(row["VALOR"])
            linha = f"{data_txt} | {row['FORNECEDOR']} | {cnpj_txt} | {valor_txt} | {row['SECRETARIA']}"
            pdf.multi_cell(0, 7, linha)

    # Converte PDF em bytes (compatível com Streamlit Cloud)
    pdf_bytes = pdf.output(dest="S").encode("latin-1")
    return io.BytesIO(pdf_bytes)

def filtro_multiselect_opcoes(label: str, serie: pd.Series):
    opcoes = sorted(serie.dropna().unique().tolist())
    return st.sidebar.multiselect(label, opcoes)

# =========================================
# Upload do arquivo
# =========================================
uploaded_file = st.file_uploader(
    "📁 Envie a planilha Excel (colunas necessárias: DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA)",
    type=["xlsx"],
)

if not uploaded_file:
    st.info("Envie uma planilha para começar.")
    st.stop()

# =========================================
# Leitura e validação
# =========================================
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

# =========================================
# Diagnóstico: por que linhas foram ignoradas
# =========================================
with st.expander("🔎 Diagnóstico: linhas ignoradas na importação"):
    raw = df_raw.copy()
    raw.columns = raw.columns.str.strip().str.upper()

    data_conv1 = pd.to_datetime(raw["DATA"], errors="coerce")
    data_conv2 = pd.to_datetime(raw["DATA"], errors="coerce", dayfirst=True)
    data_ok = data_conv1.notna() | data_conv2.notna()

    valor_num = pd.to_numeric(raw["VALOR"], errors="coerce")
    valor_brl_try = pd.to_numeric(
        raw["VALOR"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce"
    )
    valor_ok = valor_num.notna() | valor_brl_try.notna()

    fornecedor_ok = raw["FORNECEDOR"].astype(str).str.strip().ne("")
    secretaria_ok = raw["SECRETARIA"].astype(str).str.strip().ne("")

    problema = (~data_ok) | (~valor_ok) | (~fornecedor_ok) | (~secretaria_ok)
    diag = raw.loc[problema, ["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]].copy()

    if not diag.empty:
        # estimativa do número da linha no Excel (linha 1 = cabeçalho)
        diag.insert(0, "LINHA_APROX_EXCEL", diag.index + 2)
        diag["MOTIVO"] = (
            (~data_ok).map({True:"DATA inválida", False:""}) + " " +
            (~valor_ok).map({True:"VALOR inválido", False:""}) + " " +
            (~fornecedor_ok).map({True:"FORNECEDOR vazio", False:""}) + " " +
            (~secretaria_ok).map({True:"SECRETARIA vazia", False:""})
        ).str.strip()
        st.warning("Algumas linhas foram ignoradas por problemas de dados:")
        st.dataframe(diag, use_container_width=True)
    else:
        st.info("Nenhuma linha foi ignorada por problemas de dados.")

st.divider()

# =========================================
# Filtros (barra lateral)
# =========================================
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

# =========================================
# KPIs
# =========================================
col_k1, col_k2, col_k3 = st.columns(3)
total_valor = df_filtrado["VALOR"].sum() if not df_filtrado.empty else 0.0
qtd_linhas = len(df_filtrado)
qtd_fornec = df_filtrado["FORNECEDOR"].nunique()

col_k1.metric("Valor total filtrado", format_brl(total_valor))
col_k2.metric("Registros", f"{qtd_linhas}")
col_k3.metric("Fornecedores", f"{qtd_fornec}")

st.divider()

# =========================================
# Gráficos (Plotly)
# =========================================
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
            text=[format_brl(v) for v in graf1["VALOR"]],
            color="SECRETARIA",
        )
        # Hover numérico com separador padrão; o texto já está em BRL
        fig1.update_traces(hovertemplate="<b>%{y}</b><br>Valor: %{x:,.2f}")
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
            text=[format_brl(v) for v in graf2["VALOR"]],
            color="FORNECEDOR",
        )
        fig2.update_traces(hovertemplate="<b>%{x}</b><br>Valor: %{y:,.2f}")
        fig2.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10, r=10, t=30, b=80))
        st.plotly_chart(fig2, use_container_width=True)

st.divider()

# =========================================
# Tabela (formatação BRL apenas para exibição)
# =========================================
st.subheader("📋 Dados Filtrados")
df_display = df_filtrado.copy()
df_display["VALOR"] = df_display["VALOR"].apply(format_brl)
st.dataframe(df_display, use_container_width=True)

# =========================================
# Downloads (respeitam os filtros)
# =========================================
st.subheader("📥 Exportar")

# Excel (mantém numérico para permitir somas)
excel_buffer = io.BytesIO()
df_filtrado.to_excel(excel_buffer, index=False)
excel_buffer.seek(0)
st.download_button(
    "📊 Baixar Excel",
    data=excel_buffer,
    file_name="dados_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# PDF (com BRL)
pdf_bytes = gerar_pdf(df_filtrado)
st.download_button(
    "📄 Baixar PDF",
    data=pdf_bytes,
    file_name="relatorio_filtrado.pdf",
    mime="application/pdf",
)
