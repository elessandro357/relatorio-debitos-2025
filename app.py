# app.py — Análise de Gastos por Fornecedor (Streamlit)
# Requisitos: streamlit, pandas, plotly, fpdf, (opcional p/ PNG: kaleido==0.2.1)
# Executar: streamlit run app.py

import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

# ================================
# Config geral
# ================================
st.set_page_config(layout="wide", page_title="Análise de Gastos por Fornecedor")
st.title("📊 Análise de Gastos por Fornecedor")
st.caption("Dashboards de Débitos e Saldos • Filtros avançados • Exporta Excel/PDF • Gráficos (PNG/HTML).")

# ================================
# Utilidades / Helpers
# ================================
def format_brl(v):
    """R$ 1.234,56 sem depender de locale."""
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

@st.cache_data(show_spinner=False)
def load_table(upload) -> pd.DataFrame:
    """Carrega .xlsx ou .csv; padroniza cabeçalhos em maiúsculas sem espaços extras."""
    if upload is None:
        return pd.DataFrame()
    name = upload.name.lower()
    if name.endswith(".xlsx"):
        df = pd.read_excel(upload)
    elif name.endswith(".csv"):
        df = pd.read_csv(upload, sep=None, engine="python")
    else:
        st.error("Formato não suportado. Envie .xlsx ou .csv.")
        return pd.DataFrame()
    df.columns = df.columns.str.strip().str.upper()
    return df

def cast_types_debitos(df: pd.DataFrame) -> pd.DataFrame:
    """DATA robusta (tenta dayfirst) + VALOR aceita '1.234,56'."""
    df = df.copy()
    # DATA
    d1 = pd.to_datetime(df["DATA"], errors="coerce")
    d2 = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
    df["DATA"] = d1.fillna(d2)

    # VALOR
    v1 = pd.to_numeric(df["VALOR"], errors="coerce")
    precisa_brl = v1.isna() & df["VALOR"].astype(str).str.contains(r"[.,]", na=False)
    v2 = pd.to_numeric(
        df.loc[precisa_brl, "VALOR"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce"
    )
    v1.loc[precisa_brl] = v2
    df["VALOR"] = v1

    # Texto
    for col in ["FORNECEDOR", "SECRETARIA", "CNPJ"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Limpeza
    df = df.dropna(subset=["DATA", "VALOR", "FORNECEDOR", "SECRETARIA"]).copy()
    df["VALOR"] = df["VALOR"].round(2)
    df["ANO"] = df["DATA"].dt.year
    df["MES"] = df["DATA"].dt.month
    df["YM"] = df["DATA"].dt.to_period("M").astype(str)  # ex: 2025-07
    return df

def validar_debitos_cols(df):
    req = ["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]
    miss = [c for c in req if c not in df.columns]
    return len(miss)==0, miss

def validar_saldos_cols(df):
    req = ["CONTA","NOME DA CONTA","SECRETARIA","BANCO","TIPO DE RECURSO","SALDO BANCARIO"]
    miss = [c for c in req if c not in df.columns]
    return len(miss)==0, miss

def preparar_saldos(df_raw, apenas_livre=True):
    df = df_raw.copy()
    df.columns = df.columns.str.strip().str.upper()
    if "TIPO DE RECURSO" in df.columns:
        if apenas_livre:
            df = df[df["TIPO DE RECURSO"].astype(str).str.upper()=="LIVRE"]
    df["SALDO BANCARIO"] = pd.to_numeric(df["SALDO BANCARIO"], errors="coerce").fillna(0.0)
    for c in ["SECRETARIA","BANCO","TIPO DE RECURSO","NOME DA CONTA","CONTA"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df

def saldo_por_secretaria(df_saldos):
    return (df_saldos.groupby("SECRETARIA", as_index=False)["SALDO BANCARIO"]
            .sum().rename(columns={"SALDO BANCARIO":"SALDO_LIVRE"}))

# ===== PDF (tabela simples) =====
def _pdf_to_bytesio(pdf_obj):
    out = pdf_obj.output(dest="S")
    pdf_bytes = out if isinstance(out, (bytes, bytearray)) else out.encode("latin-1", "ignore")
    return io.BytesIO(pdf_bytes)

def _chunk_long_words(text, maxlen=30):
    s = "" if pd.isna(text) else str(text)
    parts = []
    for w in s.split():
        if len(w) > maxlen:
            parts.extend([w[i:i+maxlen] for i in range(0, len(w), maxlen)])
        else:
            parts.append(w)
    return " ".join(parts)

def gerar_pdf_listagem(df: pd.DataFrame, titulo="Relatório"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, txt=titulo, ln=True, align="C")
    pdf.ln(2)

    if df.empty:
        pdf.set_font("Helvetica", size=10)
        pdf.multi_cell(0, 7, "Nenhum registro.")
        return _pdf_to_bytesio(pdf)

    pdf.set_font("Helvetica", size=10)
    cols = list(df.columns)
    epw = pdf.w - 2 * pdf.l_margin

    if set(["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]).issubset(set(cols)):
        order = ["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]
        cols = [c for c in order if c in cols]
        w_data, w_forn, w_cnpj, w_val = 22, 70, 35, 28
        w_sec = max(epw - (w_data + w_forn + w_cnpj + w_val), 30)
        widths = [w_data, w_forn, w_cnpj, w_val, w_sec]
    else:
        widths = [epw / len(cols)] * len(cols)

    # Cabeçalho
    pdf.set_font("Helvetica", 'B', 10)
    for c, w in zip(cols, widths):
        pdf.multi_cell(w, 7, _chunk_long_words(c, 20), border=0, new_x="RIGHT", new_y="TOP")
    pdf.multi_cell(0, 2, "", border=0, new_x="LMARGIN", new_y="NEXT")

    # Linhas
    pdf.set_font("Helvetica", size=10)
    for _, row in df.iterrows():
        for c, w in zip(cols, widths):
            txt = row[c]
            if isinstance(txt, (int, float)) and c.upper().startswith("VALOR"):
                txt = format_brl(txt)
            txt = _chunk_long_words(txt, 30)
            pdf.multi_cell(w, 6, txt, border=0, new_x="RIGHT", new_y="TOP")
        pdf.multi_cell(0, 2, "", border=0, new_x="LMARGIN", new_y="NEXT")

    return _pdf_to_bytesio(pdf)

# ====== Downloads de gráficos (PNG/HTML) ======
def _fig_png_bytes(fig):
    try:
        return fig.to_image(format="png", scale=2)  # requer kaleido
    except Exception:
        return None

def _fig_html_bytes(fig):
    html = fig.to_html(full_html=False, include_plotlyjs="cdn")
    return html.encode("utf-8")

def download_fig_buttons(fig, base_filename, key_prefix):
    c1, c2 = st.columns(2)
    with c1:
        png = _fig_png_bytes(fig)
        if png:
            st.download_button("⬇️ Baixar PNG", data=png,
                               file_name=f"{base_filename}.png", mime="image/png",
                               key=f"{key_prefix}_png")
        else:
            st.caption("Para baixar PNG, adicione `kaleido==0.2.1` ao requirements.")
    with c2:
        html = _fig_html_bytes(fig)
        st.download_button("⬇️ Baixar HTML (interativo)", data=html,
                           file_name=f"{base_filename}.html", mime="text/html",
                           key=f"{key_prefix}_html")

def limpar_filtros(keys):
    changed = False
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]
            changed = True
    if changed:
        st.rerun()

# ================================
# ABAS (sem Plano de Pagamento)
# ================================
tab_dash, tab_saldos = st.tabs(["📈 Dashboard de Gastos (Débitos)", "🏦 Dashboard de Saldos"])

# --------- Aba Débitos (Gastos) ---------
with tab_dash:
    up_deb = st.file_uploader(
        "📁 Envie a planilha de **Débitos** (DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA) — aceita .xlsx ou .csv",
        type=["xlsx","csv"], key="deb_dashboard"
    )

    if not up_deb:
        st.info("Envie a planilha de Débitos para ver o dashboard.")
        st.stop()

    # Carrega e valida
    df_raw = load_table(up_deb)
    ok, miss = validar_debitos_cols(df_raw)
    if not ok:
        st.error(f"Faltam colunas em Débitos: {', '.join(miss)}")
        st.stop()

    df = cast_types_debitos(df_raw)

    # ---------------- Sidebar de filtros PRO ----------------
    st.sidebar.header("🔎 Filtros — Gastos (Débitos)")

    # Período
    dmin = pd.to_datetime(df["DATA"].min()).date()
    dmax = pd.to_datetime(df["DATA"].max()).date()
    din = st.sidebar.date_input("Data inicial", dmin, key="deb_d1")
    dfi = st.sidebar.date_input("Data final", dmax, key="deb_d2")
    if din > dfi:
        st.sidebar.error("Data inicial > Data final."); st.stop()

    # Multiseleções
    secs = st.sidebar.multiselect("Secretaria", sorted(df["SECRETARIA"].unique()), key="deb_secs")
    forn = st.sidebar.multiselect("Fornecedor", sorted(df["FORNECEDOR"].unique()), key="deb_forn")
    cnpjs = st.sidebar.multiselect("CNPJ", sorted(df["CNPJ"].astype(str).unique()), key="deb_cnpjs")

    # Busca textual e faixa de valores
    forn_q = st.sidebar.text_input("Busca por texto em Fornecedor", key="deb_forn_q")
    vmin, vmax = float(df["VALOR"].min()), float(df["VALOR"].max())
    vsel = st.sidebar.slider("Faixa de valores (R$)", min_value=0.0, max_value=max(vmax, 0.0),
                             value=(max(0.0, vmin), vmax), step=0.01, key="deb_vrange")

    # Top N para ranking
    topn = st.sidebar.number_input("Top N fornecedores (ranking)", min_value=3, max_value=50, value=10, step=1, key="deb_topn")

    # Botão Limpar
    if st.sidebar.button("🧹 Limpar filtros"):
        limpar_filtros(["deb_d1","deb_d2","deb_secs","deb_forn","deb_cnpjs","deb_forn_q","deb_vrange","deb_topn"])

    # ---------------- Aplica filtros ----------------
    df_f = df[(df["DATA"]>=pd.to_datetime(din)) & (df["DATA"]<=pd.to_datetime(dfi))].copy()
    if secs:   df_f = df_f[df_f["SECRETARIA"].isin(secs)]
    if forn:   df_f = df_f[df_f["FORNECEDOR"].isin(forn)]
    if cnpjs:  df_f = df_f[df_f["CNPJ"].astype(str).isin(cnpjs)]
    if forn_q:
        df_f = df_f[df_f["FORNECEDOR"].str.contains(forn_q, case=False, na=False)]
    if vsel:
        df_f = df_f[(df_f["VALOR"]>=vsel[0]) & (df_f["VALOR"]<=vsel[1])]

    # ---------------- KPIs ----------------
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Valor total filtrado", format_brl(df_f["VALOR"].sum() if not df_f.empty else 0))
    k2.metric("Registros", f"{len(df_f)}")
    k3.metric("Fornecedores", f"{df_f['FORNECEDOR'].nunique()}")
    k4.metric("Secretarias", f"{df_f['SECRETARIA'].nunique()}")

    st.divider()

    # ---------------- Gráficos ----------------
    g1c,g2c = st.columns(2)
    with g1c:
        st.subheader("Débitos por Secretaria")
        if df_f.empty:
            st.info("Sem dados.")
            fig1 = None
        else:
            g1 = df_f.groupby("SECRETARIA", as_index=False)["VALOR"].sum().sort_values("VALOR")
            fig1 = px.bar(g1, x="VALOR", y="SECRETARIA", orientation="h",
                          text=[format_brl(v) for v in g1["VALOR"]], color="SECRETARIA")
            fig1.update_traces(hovertemplate="<b>%{y}</b><br>Valor: %{x:,.2f}")
            fig1.update_layout(showlegend=False, margin=dict(l=10,r=10,t=30,b=10))
            st.plotly_chart(fig1, use_container_width=True)
            download_fig_buttons(fig1, "debitos_por_secretaria", "deb1")

    with g2c:
        st.subheader(f"Top {topn} Fornecedores (por valor)")
        if df_f.empty:
            st.info("Sem dados.")
            fig2 = None
        else:
            g2 = (df_f.groupby(["FORNECEDOR","CNPJ"], as_index=False)["VALOR"]
                        .sum().sort_values("VALOR", ascending=False).head(int(topn)))
            g2["FORNEC"] = g2["FORNECEDOR"] + " • " + g2["CNPJ"].astype(str)
            fig2 = px.bar(g2, x="FORNEC", y="VALOR",
                          text=[format_brl(v) for v in g2["VALOR"]], color="FORNECEDOR")
            fig2.update_traces(hovertemplate="<b>%{x}</b><br>Valor: %{y:,.2f}")
            fig2.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80))
            st.plotly_chart(fig2, use_container_width=True)
            download_fig_buttons(fig2, "top_fornecedores", "deb2")

    st.divider()

    # Série temporal por mês
    st.subheader("Série Temporal — Gastos por Mês")
    if df_f.empty:
        st.info("Sem dados.")
        fig3 = None
    else:
        g3 = df_f.groupby("YM", as_index=False)["VALOR"].sum().sort_values("YM")
        fig3 = px.line(g3, x="YM", y="VALOR", markers=True)
        fig3.update_traces(hovertemplate="<b>%{x}</b><br>Valor: %{y:,.2f}")
        fig3.update_layout(xaxis_title="Ano-Mês", yaxis_title="Valor", margin=dict(l=10,r=10,t=30,b=10))
        st.plotly_chart(fig3, use_container_width=True)
        download_fig_buttons(fig3, "serie_mensal_gastos", "deb3")

    # Heatmap Secretaria x Mês
    st.subheader("Mapa de Calor — Secretaria × Mês")
    if df_f.empty:
        st.info("Sem dados.")
        fig4 = None
    else:
        piv = (df_f.groupby(["SECRETARIA","YM"])["VALOR"].sum()
                     .unstack(fill_value=0).sort_index())
        if piv.shape[0] == 0 or piv.shape[1] == 0:
            st.info("Sem dados.")
            fig4 = None
        else:
            fig4 = px.imshow(
                piv.values,
                labels=dict(x="Ano-Mês", y="Secretaria", color="Valor"),
                x=list(piv.columns), y=list(piv.index), aspect="auto"
            )
            st.plotly_chart(fig4, use_container_width=True)
            download_fig_buttons(fig4, "heatmap_secretaria_mes", "deb4")

    st.divider()
    st.subheader("📋 Dados Filtrados")
    df_disp = df_f.copy()
    df_disp["VALOR"] = df_disp["VALOR"].apply(format_brl)
    # DATA para DD/MM/AAAA
    df_disp["DATA"] = pd.to_datetime(df_disp["DATA"]).dt.strftime("%d/%m/%Y")
    st.dataframe(df_disp[["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]], use_container_width=True)

    st.subheader("📥 Exportar (Débitos)")
    xbuf = io.BytesIO(); df_f.to_excel(xbuf, index=False); xbuf.seek(0)
    st.download_button("📊 Excel (dados filtrados)", data=xbuf,
                       file_name="debitos_filtrados.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    pdf_df = df_disp.rename(columns={"VALOR":"VALOR (BRL)"})
    pdf = gerar_pdf_listagem(pdf_df, "Débitos — Dados Filtrados")
    st.download_button("📄 PDF (dados filtrados)", data=pdf,
                       file_name="debitos_filtrados.pdf", mime="application/pdf")

# --------- Aba Saldos ---------
with tab_saldos:
    up_saldos = st.file_uploader(
        "🏦 Envie a planilha de **Saldos** (CONTA, NOME DA CONTA, SECRETARIA, BANCO, TIPO DE RECURSO, SALDO BANCARIO) — .xlsx ou .csv",
        type=["xlsx","csv"], key="saldos_tab")
    apenas_livre = st.checkbox("Considerar apenas Recurso LIVRE", value=True)

    if not up_saldos:
        st.info("Envie a planilha de Saldos para ver o dashboard.")
    else:
        sal_raw = load_table(up_saldos)
        ok_s, miss_s = validar_saldos_cols(sal_raw)
        if not ok_s:
            st.error(f"Saldos inválidos. Faltam: {', '.join(miss_s)}"); st.stop()
        sal = preparar_saldos(sal_raw, apenas_livre=apenas_livre)

        st.sidebar.header("🔎 Filtros — Saldos")
        secs_sal = st.sidebar.multiselect("Secretaria (saldos)", sorted(sal["SECRETARIA"].dropna().unique()), key="sal_secs")
        bancos   = st.sidebar.multiselect("Banco", sorted(sal["BANCO"].dropna().unique()), key="sal_bancos")
        tipos    = st.sidebar.multiselect("Tipo de Recurso", sorted(sal["TIPO DE RECURSO"].dropna().unique()), key="sal_tipos")

        if st.sidebar.button("🧹 Limpar filtros (Saldos)"):
            limpar_filtros(["sal_secs","sal_bancos","sal_tipos"])

        sal_f = sal.copy()
        if secs_sal: sal_f = sal_f[sal_f["SECRETARIA"].isin(secs_sal)]
        if bancos:   sal_f = sal_f[sal_f["BANCO"].isin(bancos)]
        if tipos:    sal_f = sal_f[sal_f["TIPO DE RECURSO"].isin(tipos)]

        k1,k2,k3 = st.columns(3)
        k1.metric("Saldo total", format_brl(sal_f["SALDO BANCARIO"].sum()))
        k2.metric("Contas", f"{len(sal_f)}")
        k3.metric("Secretarias", f"{sal_f['SECRETARIA'].nunique()}")

        st.divider()
        st.subheader("Saldos por Secretaria")
        gsec = saldo_por_secretaria(sal_f).sort_values("SALDO_LIVRE", ascending=False)
        if gsec.empty:
            st.info("Sem dados.")
        else:
            fig = px.bar(gsec, x="SECRETARIA", y="SALDO_LIVRE",
                         text=[format_brl(v) for v in gsec["SALDO_LIVRE"]], color="SECRETARIA")
            fig.update_traces(hovertemplate="<b>%{x}</b><br>Saldo: %{y:,.2f}")
            fig.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80))
            st.plotly_chart(fig, use_container_width=True)
            download_fig_buttons(fig, "saldos_por_secretaria", "sald1")

        st.divider()
        st.subheader("📋 Contas (filtradas)")
        sal_display = sal_f.copy()
        sal_display["SALDO BANCARIO"] = sal_display["SALDO BANCARIO"].apply(format_brl)
        st.dataframe(sal_display, use_container_width=True)

        st.subheader("📥 Exportar (Saldos)")
        bsal = io.BytesIO(); sal_f.to_excel(bsal, index=False); bsal.seek(0)
        st.download_button("📊 Excel (saldos filtrados)", data=bsal,
                           file_name="saldos_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        pdf_sal = sal_display.rename(columns={"SALDO BANCARIO":"SALDO (BRL)"})
        pdf2 = gerar_pdf_listagem(pdf_sal, "Saldos — Contas Filtradas")
        st.download_button("📄 PDF (saldos filtrados)", data=pdf2,
                           file_name="saldos_filtrados.pdf", mime="application/pdf")
