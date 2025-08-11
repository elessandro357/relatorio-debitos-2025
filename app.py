# app.py — Comparativo Geral (estilo Power BI)
# -----------------------------------------------------------------------------
# • Lê "Comparativo geral.xlsx" (ou qualquer .xlsx enviado)
# • Detecta automaticamente abas "Débitos" e "Saldos" (ou usa mapeador de colunas)
# • Filtros (slicers) na sidebar + KPIs + gráficos interativos (Plotly)
# • Exports: Excel/CSV dos dados filtrados; PNG dos gráficos (opcional com kaleido)
# -----------------------------------------------------------------------------

import io
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.io as pio  # necessário quando exportar PNG com kaleido

# ============ Config ============ #
st.set_page_config(page_title="Comparativo Geral — Painel Interativo", layout="wide")
st.title("📊 Comparativo Geral — Painel Interativo")
st.caption("Estilo Power BI — filtros na lateral, KPIs e gráficos interativos (Plotly).")

# ============ Helpers ============ #
def format_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

@st.cache_data(show_spinner=False, ttl=300)
def load_excel_all_sheets(file) -> dict:
    """Lê todas as abas do Excel em um dicionário {nome: DataFrame}, normalizando cabeçalhos."""
    xls = pd.ExcelFile(file)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        df.columns = df.columns.str.strip().str.upper()
        sheets[name.upper()] = df
    return sheets

def validar_debitos_cols(cols):
    req = ["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]
    miss = [c for c in req if c not in cols]
    return len(miss)==0, miss, req

def validar_saldos_cols(cols):
    req = ["CONTA","NOME DA CONTA","SECRETARIA","BANCO","TIPO DE RECURSO","SALDO BANCARIO"]
    miss = [c for c in req if c not in cols]
    return len(miss)==0, miss, req

def coluna_mapper_ui(cols_atual, req_cols, key_prefix):
    st.info("Mapeie suas colunas para o modelo esperado.")
    mapeamento = {}
    for alvo in req_cols:
        opts = ["(não existe)"] + list(cols_atual)
        mapeamento[alvo] = st.selectbox(
            f"Coluna do arquivo para **{alvo}**",
            options=opts,
            index=opts.index(alvo) if alvo in cols_atual else 0,
            key=f"{key_prefix}_{alvo}"
        )
    return mapeamento

def aplicar_mapeamento(df, mapa):
    cols_novas = {}
    for alvo, origem in mapa.items():
        if origem != "(não existe)" and origem in df.columns:
            cols_novas[alvo] = df[origem]
        else:
            cols_novas[alvo] = pd.Series([None]*len(df))
    return pd.DataFrame(cols_novas)

def preparar_debitos(df):
    df = df.copy()
    # DATA
    d1 = pd.to_datetime(df["DATA"], errors="coerce")
    d2 = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
    df["DATA"] = d1.fillna(d2)
    # VALOR (aceita 1.234,56)
    v1 = pd.to_numeric(df["VALOR"], errors="coerce")
    precisa_brl = v1.isna() & df["VALOR"].astype(str).str.contains(r"[.,]", na=False)
    v2 = pd.to_numeric(
        df.loc[precisa_brl, "VALOR"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce"
    )
    v1.loc[precisa_brl] = v2
    df["VALOR"] = v1.clip(lower=0).round(2)
    # Texto
    df["FORNECEDOR"] = df["FORNECEDOR"].astype(str).str.strip()
    df["SECRETARIA"] = df["SECRETARIA"].astype(str).str.strip()
    if "CNPJ" in df.columns:
        df["CNPJ"] = df["CNPJ"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(14)
    # Limpeza
    df = df.dropna(subset=["DATA","VALOR","FORNECEDOR","SECRETARIA"]).copy()
    return df

def preparar_saldos(df):
    df = df.copy()
    df["SALDO BANCARIO"] = pd.to_numeric(df["SALDO BANCARIO"], errors="coerce").fillna(0.0)
    for c in ["SECRETARIA","BANCO","TIPO DE RECURSO","NOME DA CONTA","CONTA"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df

def exp_export_tabela(nome, df):
    st.subheader(f"📥 Exportar — {nome}")
    c1, c2 = st.columns(2)
    with c1:
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)
        st.download_button("⬇️ Excel", data=buf, file_name=f"{nome.lower().replace(' ','_')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ CSV", data=csv, file_name=f"{nome.lower().replace(' ','_')}.csv", mime="text/csv")

# ============ Upload ============ #
st.markdown("### 📤 Envie o arquivo Excel")
up = st.file_uploader("Selecione o arquivo (ex.: Comparativo geral.xlsx)", type=["xlsx"])

if not up:
    st.info("Envie o Excel para começar.")
    st.stop()

sheets = load_excel_all_sheets(up)
sheet_names = list(sheets.keys())

# Tentativa automática de achar abas
deb_sheet_guess = next((n for n in sheet_names if "DEBIT" in n), sheet_names[0])
sald_sheet_guess = next((n for n in sheet_names if "SALD" in n), None)

col_sel1, col_sel2 = st.columns(2)
with col_sel1:
    deb_tab = st.selectbox("Aba de Débitos", options=sheet_names, index=sheet_names.index(deb_sheet_guess))
with col_sel2:
    sald_tab = st.selectbox("Aba de Saldos (opcional)", options=["(nenhuma)"]+sheet_names,
                             index=(["(nenhuma)"]+sheet_names).index(sald_sheet_guess) if sald_sheet_guess else 0)

df_deb_raw = sheets[deb_tab].copy()
okd, missd, reqd = validar_debitos_cols(df_deb_raw.columns)
if not okd:
    st.warning("A aba escolhida como **Débitos** não tem todas as colunas. Faça o mapeamento:")
    mapa = coluna_mapper_ui(df_deb_raw.columns, reqd, key_prefix="deb")
    df_deb = aplicar_mapeamento(df_deb_raw, mapa)
else:
    df_deb = df_deb_raw[reqd].copy()

df_deb = preparar_debitos(df_deb)

df_sald = None
if sald_tab != "(nenhuma)":
    df_sald_raw = sheets[sald_tab].copy()
    oks, misss, reqs = validar_saldos_cols(df_sald_raw.columns)
    if not oks:
        st.warning("A aba escolhida como **Saldos** não tem todas as colunas. Faça o mapeamento:")
        mapa = coluna_mapper_ui(df_sald_raw.columns, reqs, key_prefix="sald")
        df_sald = aplicar_mapeamento(df_sald_raw, mapa)
    else:
        df_sald = df_sald_raw[reqs].copy()
    df_sald = preparar_saldos(df_sald)

st.success("Dados carregados com sucesso!")

# ============ Sidebar (filtros) ============ #
st.sidebar.header("🔎 Filtros (estilo slicer)")
# Período (somente débitos)
dmin = pd.to_datetime(df_deb["DATA"]).min().date()
dmax = pd.to_datetime(df_deb["DATA"]).max().date()
c_dt1, c_dt2 = st.sidebar.columns(2)
with c_dt1:
    dt_ini = st.date_input("Data inicial", dmin, key="dtini")
with c_dt2:
    dt_fim = st.date_input("Data final", dmax, key="dtfim")
if dt_ini > dt_fim:
    st.sidebar.error("Data inicial > Data final.")
    st.stop()

secs = sorted(df_deb["SECRETARIA"].unique().tolist())
forns = sorted(df_deb["FORNECEDOR"].unique().tolist())
f_secs = st.sidebar.multiselect("Secretaria", secs)
f_forns = st.sidebar.multiselect("Fornecedor", forns)

# Filtros de saldos, se existir
if df_sald is not None:
    bancos = sorted(df_sald["BANCO"].dropna().unique().tolist())
    tipos_rec = sorted(df_sald["TIPO DE RECURSO"].dropna().unique().tolist())
    f_bancos = st.sidebar.multiselect("Banco (saldos)", bancos)
    f_tipos = st.sidebar.multiselect("Tipo de Recurso (saldos)", tipos_rec)
else:
    f_bancos, f_tipos = [], []

# Aplicar filtros aos débitos
deb_f = df_deb[(df_deb["DATA"] >= pd.to_datetime(dt_ini)) & (df_deb["DATA"] <= pd.to_datetime(dt_fim))].copy()
if f_secs:
    deb_f = deb_f[deb_f["SECRETARIA"].isin(f_secs)]
if f_forns:
    deb_f = deb_f[deb_f["FORNECEDOR"].isin(f_forns)]

# Aplicar filtros aos saldos
if df_sald is not None:
    sald_f = df_sald.copy()
    if f_secs:
        sald_f = sald_f[sald_f["SECRETARIA"].isin(f_secs)]
    if f_bancos:
        sald_f = sald_f[sald_f["BANCO"].isin(f_bancos)]
    if f_tipos:
        sald_f = sald_f[sald_f["TIPO DE RECURSO"].isin(f_tipos)]
else:
    sald_f = None

# ============ KPIs topo ============ #
k1, k2, k3, k4 = st.columns(4)
total_debitos = deb_f["VALOR"].sum() if not deb_f.empty else 0.0
k1.metric("Total de Débitos (filtrado)", format_brl(total_debitos))
k2.metric("Registros de Débito", f"{len(deb_f)}")
k3.metric("Fornecedores", f"{deb_f['FORNECEDOR'].nunique()}")
if sald_f is not None:
    k4.metric("Saldo Bancário (filtrado)", format_brl(sald_f["SALDO BANCARIO"].sum()))
else:
    k4.metric("Saldo Bancário (filtrado)", "—")

st.divider()

# ============ Gráficos estilo Power BI ============ #
g1c, g2c = st.columns(2)

# 1) Barras horizontais: débito por secretaria
with g1c:
    st.subheader("🔹 Débitos por Secretaria")
    if deb_f.empty:
        st.info("Sem dados.")
        fig_sec = None
    else:
        g1 = deb_f.groupby("SECRETARIA", as_index=False)["VALOR"].sum().sort_values("VALOR")
        fig_sec = px.bar(g1, x="VALOR", y="SECRETARIA", orientation="h",
                         text=[format_brl(v) for v in g1["VALOR"]],
                         color="SECRETARIA")
        fig_sec.update_traces(hovertemplate="<b>%{y}</b><br>Valor: %{x:,.2f}")
        fig_sec.update_layout(showlegend=False, margin=dict(l=10,r=10,t=30,b=10))
        st.plotly_chart(fig_sec, use_container_width=True)

# 2) Top fornecedores (barras)
with g2c:
    st.subheader("🔹 Top 10 Fornecedores")
    if deb_f.empty:
        st.info("Sem dados.")
        fig_forn = None
    else:
        g2 = (deb_f.groupby("FORNECEDOR", as_index=False)["VALOR"]
              .sum().sort_values("VALOR", ascending=False).head(10))
        fig_forn = px.bar(g2, x="FORNECEDOR", y="VALOR",
                          text=[format_brl(v) for v in g2["VALOR"]],
                          color="FORNECEDOR")
        fig_forn.update_traces(hovertemplate="<b>%{x}</b><br>Valor: %{y:,.2f}")
        fig_forn.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80))
        st.plotly_chart(fig_forn, use_container_width=True)

st.divider()

# 3) Linha do tempo: soma de débitos por mês
st.subheader("📈 Série Temporal — Débitos por Mês")
if deb_f.empty:
    st.info("Sem dados.")
    fig_mensal = None
else:
    tmp = deb_f.copy()
    tmp["MES"] = pd.to_datetime(tmp["DATA"]).dt.to_period("M").dt.to_timestamp()
    g3 = tmp.groupby("MES", as_index=False)["VALOR"].sum()
    fig_mensal = px.line(g3, x="MES", y="VALOR", markers=True)
    fig_mensal.update_traces(hovertemplate="<b>%{x|%d/%m/%Y}</b><br>Total: %{y:,.2f}")
    fig_mensal.update_layout(margin=dict(l=10,r=10,t=30,b=10))
    st.plotly_chart(fig_mensal, use_container_width=True)

st.divider()

# 4) Treemap por Secretaria > Fornecedor
st.subheader("🧩 Treemap — Distribuição por Secretaria > Fornecedor")
if deb_f.empty:
    st.info("Sem dados.")
else:
    g4 = deb_f.groupby(["SECRETARIA","FORNECEDOR"], as_index=False)["VALOR"].sum()
    fig_tree = px.treemap(g4, path=["SECRETARIA","FORNECEDOR"], values="VALOR")
    fig_tree.update_traces(hovertemplate="<b>%{label}</b><br>Valor: %{value:,.2f}")
    st.plotly_chart(fig_tree, use_container_width=True)

st.divider()

# ============ Seção de Saldos (se existir) ============ #
st.header("🏦 Saldos (opcional)")
if df_sald is None:
    st.info("Nenhuma aba de Saldos selecionada.")
else:
    ksa, ksb, ksc = st.columns(3)
    ksa.metric("Saldo total (filtrado)", format_brl(sald_f["SALDO BANCARIO"].sum()))
    ksb.metric("Contas", f"{len(sald_f)}")
    ksc.metric("Secretarias", f"{sald_f['SECRETARIA'].nunique()}")

    st.subheader("🔹 Saldos por Secretaria")
    gsec = sald_f.groupby("SECRETARIA", as_index=False)["SALDO BANCARIO"].sum().sort_values("SALDO BANCARIO", ascending=False)
    if gsec.empty:
        st.info("Sem dados.")
        fig_sald = None
    else:
        fig_sald = px.bar(gsec, x="SECRETARIA", y="SALDO BANCARIO",
                          text=[format_brl(v) for v in gsec["SALDO BANCARIO"]],
                          color="SECRETARIA")
        fig_sald.update_traces(hovertemplate="<b>%{x}</b><br>Saldo: %{y:,.2f}")
        fig_sald.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80))
        st.plotly_chart(fig_sald, use_container_width=True)

    st.subheader("📋 Saldos — Dados Filtrados")
    st.dataframe(sald_f, use_container_width=True)

st.divider()

# ============ Tabelas de dados filtrados (Débitos) ============ #
st.header("📋 Débitos — Dados Filtrados")
st.dataframe(deb_f.assign(VALOR_FORMATADO=deb_f["VALOR"].apply(format_brl)), use_container_width=True)

# Exports de dados filtrados
st.subheader("📥 Exportar dados filtrados")
c1, c2 = st.columns(2)
with c1:
    buf = io.BytesIO()
    deb_f.to_excel(buf, index=False)
    buf.seek(0)
    st.download_button("⬇️ Excel — Débitos", data=buf, file_name="debitos_filtrados.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with c2:
    csv = deb_f.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ CSV — Débitos", data=csv, file_name="debitos_filtrados.csv", mime="text/csv")
