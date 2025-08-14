# app.py — Análise de Gastos por Fornecedor (Streamlit)
# Requisitos: streamlit, pandas, plotly, fpdf
# (opcional para incluir gráficos no PDF: kaleido==0.2.1)
# Executar: streamlit run app.py

import io
import re
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
from fpdf import FPDF

# ================================
# Config + CSS (métricas menores)
# ================================
st.set_page_config(layout="wide", page_title="Análise de Gastos por Fornecedor")
st.markdown(
    """
    <style>
      div[data-testid="stMetricValue"] { font-size: 1.4rem !important; }
      div[data-testid="stMetricLabel"] { font-size: 0.9rem !important; }
      .block-container { padding-top: 0.8rem; padding-bottom: 0.8rem; }
    </style>
    """,
    unsafe_allow_html=True
)
st.title("📊 Análise de Gastos por Fornecedor")
st.caption("Dashboards de Débitos e Saldos • Filtros avançados • Exporta Excel/PDF • Botão de imprimir.")

PLOTLY_FONT_SIZE = 12  # fonte menor em todos os gráficos

# ================================
# Helpers
# ================================
def format_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

@st.cache_data(show_spinner=False)
def load_table(upload) -> pd.DataFrame:
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
    df = df.copy()
    d1 = pd.to_datetime(df["DATA"], errors="coerce")
    d2 = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
    df["DATA"] = d1.fillna(d2)

    v1 = pd.to_numeric(df["VALOR"], errors="coerce")
    precisa_brl = v1.isna() & df["VALOR"].astype(str).str.contains(r"[.,]", na=False)
    v2 = pd.to_numeric(
        df.loc[precisa_brl, "VALOR"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce"
    )
    v1.loc[precisa_brl] = v2
    df["VALOR"] = v1

    for col in ["FORNECEDOR", "SECRETARIA", "CNPJ"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    df = df.dropna(subset=["DATA", "VALOR", "FORNECEDOR", "SECRETARIA"]).copy()
    df["VALOR"] = df["VALOR"].round(2)
    df["ANO"] = df["DATA"].dt.year
    df["YM"] = df["DATA"].dt.to_period("M").astype(str)
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
    if "TIPO DE RECURSO" in df.columns and apenas_livre:
        df = df[df["TIPO DE RECURSO"].astype(str).str.upper()=="LIVRE"]
    df["SALDO BANCARIO"] = pd.to_numeric(df["SALDO BANCARIO"], errors="coerce").fillna(0.0)
    for c in ["SECRETARIA","BANCO","TIPO DE RECURSO","NOME DA CONTA","CONTA"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df

def saldo_por_secretaria(df_saldos):
    return (df_saldos.groupby("SECRETARIA", as_index=False)["SALDO BANCARIO"]
            .sum().rename(columns={"SALDO BANCARIO":"SALDO_LIVRE"}))

# ---------- PDF seguro (sanitização Latin-1) ----------
SMART_MAP = {
    "—": "-", "–": "-", "‒": "-", "―": "-",
    "“": '"', "”": '"', "‘": "'", "’": "'",
    "•": "-", "\u00A0": " "
}
def to_pdf_text(s: str) -> str:
    s = "" if s is None else str(s)
    for k, v in SMART_MAP.items():
        s = s.replace(k, v)
    s = re.sub(r'[\u200b-\u200f\u202a-\u202e]', '', s)  # zero-width/biDi
    try:
        s.encode("latin-1")
    except UnicodeEncodeError:
        s = s.encode("latin-1", "ignore").decode("latin-1")
    return s

def _pdf_to_bytesio(pdf_obj):
    out = pdf_obj.output(dest="S")
    pdf_bytes = out if isinstance(out, (bytes, bytearray)) else out.encode("latin-1", "ignore")
    return io.BytesIO(pdf_bytes)

def gerar_pdf_listagem(df: pd.DataFrame, titulo="Relatorio"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, txt=to_pdf_text(titulo), ln=True, align="C")
    pdf.ln(2)

    if df.empty:
        pdf.set_font("Helvetica", size=10)
        pdf.multi_cell(0, 7, to_pdf_text("Nenhum registro."))
        return _pdf_to_bytesio(pdf)

    cols = list(df.columns)
    epw = pdf.w - 2 * pdf.l_margin

    if set(["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]).issubset(set(df.columns)):
        order = ["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]
        cols = [c for c in order if c in df.columns]
        w_data, w_forn, w_cnpj, w_val = 22, 70, 35, 28
        w_sec = max(epw - (w_data + w_forn + w_cnpj + w_val), 30)
        widths = [w_data, w_forn, w_cnpj, w_val, w_sec]
    else:
        widths = [epw / len(cols)] * len(cols)

    pdf.set_font("Helvetica", 'B', 10)
    for c, w in zip(cols, widths):
        pdf.multi_cell(w, 7, to_pdf_text(c), border=0, new_x="RIGHT", new_y="TOP")
    pdf.multi_cell(0, 2, "", border=0, new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Helvetica", size=10)
    for _, row in df.iterrows():
        for c, w in zip(cols, widths):
            txt = row[c] if c in row else ""
            if isinstance(txt, (int, float)) and str(c).upper().startswith("VALOR"):
                txt = format_brl(txt)
            pdf.multi_cell(w, 6, to_pdf_text(txt), border=0, new_x="RIGHT", new_y="TOP")
        pdf.multi_cell(0, 2, "", border=0, new_x="LMARGIN", new_y="NEXT")

    return _pdf_to_bytesio(pdf)

# ---- Captura PNG do Plotly (para PDF do dashboard) ----
def _fig_png_bytes(fig):
    try:
        return fig.to_image(format="png", scale=2)  # requer kaleido
    except Exception:
        return None

def gerar_pdf_dashboard(titulo, metrics: dict, figs: list):
    """Gera um PDF (título+métricas+gráficos). 'figs' = [(subtitulo, fig), ...]."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, to_pdf_text(titulo), ln=True, align="C")
    pdf.ln(2)

    pdf.set_font("Helvetica", size=10)
    for k, v in metrics.items():
        pdf.cell(0, 6, to_pdf_text(f"{k}: {v}"), ln=True)

    epw = pdf.w - 2 * pdf.l_margin
    for subtitulo, fig in figs:
        if fig is None:
            continue
        img_bytes = _fig_png_bytes(fig)
        if not img_bytes:
            pdf.ln(4)
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(0, 7, to_pdf_text(subtitulo + " (gráfico indisponível sem 'kaleido')"), ln=True)
            pdf.set_font("Helvetica", size=10)
            continue
        # Converte bytes -> stream e informa o tipo
        stream = io.BytesIO(img_bytes)
        stream.seek(0)
        pdf.ln(4)
        pdf.set_font("Helvetica", 'B', 11)
        pdf.cell(0, 7, to_pdf_text(subtitulo), ln=True)
        pdf.image(stream, w=epw, type="PNG")
    return _pdf_to_bytesio(pdf)

def limpar_filtros(keys):
    changed = False
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]
            changed = True
    if changed:
        st.rerun()

# ================================
# ABAS
# ================================
tab_dash, tab_saldos = st.tabs(["📈 Dashboard de Gastos (Débitos)", "🏦 Dashboard de Saldos"])

# --------- Aba Débitos (sem série temporal e sem heatmap) ---------
with tab_dash:
    up_deb = st.file_uploader(
        "📁 Envie a planilha de **Débitos** (DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA) — .xlsx ou .csv",
        type=["xlsx","csv"], key="deb_dashboard"
    )
    if not up_deb:
        st.info("Envie a planilha de Débitos para ver o dashboard.")
        st.stop()

    df_raw = load_table(up_deb)
    ok, miss = validar_debitos_cols(df_raw)
    if not ok:
        st.error(f"Faltam colunas em Débitos: {', '.join(miss)}")
        st.stop()
    df = cast_types_debitos(df_raw)

    # Sidebar de filtros
    st.sidebar.header("🔎 Filtros — Gastos (Débitos)")
    dmin = pd.to_datetime(df["DATA"].min()).date()
    dmax = pd.to_datetime(df["DATA"].max()).date()
    din = st.sidebar.date_input("Data inicial", dmin, key="deb_d1")
    dfi = st.sidebar.date_input("Data final", dmax, key="deb_d2")
    if din > dfi:
        st.sidebar.error("Data inicial > Data final."); st.stop()
    secs = st.sidebar.multiselect("Secretaria", sorted(df["SECRETARIA"].unique()), key="deb_secs")
    forn = st.sidebar.multiselect("Fornecedor", sorted(df["FORNECEDOR"].unique()), key="deb_forn")
    cnpjs = st.sidebar.multiselect("CNPJ", sorted(df["CNPJ"].astype(str).unique()), key="deb_cnpjs")
    forn_q = st.sidebar.text_input("Busca por texto em Fornecedor", key="deb_forn_q")
    vmin, vmax = float(df["VALOR"].min()), float(df["VALOR"].max())
    vsel = st.sidebar.slider("Faixa de valores (R$)", min_value=0.0, max_value=max(vmax, 0.0),
                             value=(max(0.0, vmin), vmax), step=0.01, key="deb_vrange")
    topn = st.sidebar.number_input("Top N fornecedores (ranking)", min_value=3, max_value=50, value=10, step=1, key="deb_topn")

    if st.sidebar.button("🧹 Limpar filtros"):
        limpar_filtros(["deb_d1","deb_d2","deb_secs","deb_forn","deb_cnpjs","deb_forn_q","deb_vrange","deb_topn"])

    # Aplica filtros
    df_f = df[(df["DATA"]>=pd.to_datetime(din)) & (df["DATA"]<=pd.to_datetime(dfi))].copy()
    if secs:   df_f = df_f[df_f["SECRETARIA"].isin(secs)]
    if forn:   df_f = df_f[df_f["FORNECEDOR"].isin(forn)]
    if cnpjs:  df_f = df_f[df_f["CNPJ"].astype(str).isin(cnpjs)]
    if forn_q: df_f = df_f[df_f["FORNECEDOR"].str.contains(forn_q, case=False, na=False)]
    if vsel:   df_f = df_f[(df_f["VALOR"]>=vsel[0]) & (df_f["VALOR"]<=vsel[1])]

    # KPIs
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Valor total filtrado", format_brl(df_f["VALOR"].sum() if not df_f.empty else 0))
    k2.metric("Registros", f"{len(df_f)}")
    k3.metric("Fornecedores", f"{df_f['FORNECEDOR'].nunique()}")
    k4.metric("Secretarias", f"{df_f['SECRETARIA'].nunique()}")

    st.divider()

    # Gráficos mantidos
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
            fig1.update_layout(showlegend=False, margin=dict(l=10,r=10,t=30,b=10), font=dict(size=PLOTLY_FONT_SIZE))
            st.plotly_chart(fig1, use_container_width=True)
    with g2c:
        st.subheader(f"Top {int(topn)} Fornecedores (por valor)")
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
            fig2.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80), font=dict(size=PLOTLY_FONT_SIZE))
            st.plotly_chart(fig2, use_container_width=True)

    st.divider()
    st.subheader("📋 Dados Filtrados")
    df_disp = df_f.copy()
    df_disp["VALOR"] = df_disp["VALOR"].apply(format_brl)
    df_disp["DATA"] = pd.to_datetime(df_disp["DATA"]).dt.strftime("%d/%m/%Y")
    st.dataframe(df_disp[["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]], use_container_width=True)

    st.subheader("📥 Exportar / Imprimir")
    # Excel
    xbuf = io.BytesIO(); df_f.to_excel(xbuf, index=False); xbuf.seek(0)
    st.download_button("📊 Excel (dados filtrados)", data=xbuf,
                       file_name="debitos_filtrados.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    # PDF (tabela)
    pdf_df = df_disp.rename(columns={"VALOR":"VALOR (BRL)"})
    pdf = gerar_pdf_listagem(pdf_df, "Debitos - Dados Filtrados")
    st.download_button("📄 PDF (dados filtrados - tabela)", data=pdf,
                       file_name="debitos_filtrados.pdf", mime="application/pdf")
    # PDF do painel (só gráficos mantidos)
    deb_metrics = {
        "Valor total filtrado": format_brl(df_f["VALOR"].sum() if not df_f.empty else 0),
        "Registros": str(len(df_f)),
        "Fornecedores": str(df_f["FORNECEDOR"].nunique()),
        "Secretarias": str(df_f["SECRETARIA"].nunique())
    }
    pdf_dash = gerar_pdf_dashboard(
        "Dashboard - Gastos (Débitos)",
        deb_metrics,
        [
            ("Débitos por Secretaria", fig1),
            (f"Top {int(topn)} Fornecedores", fig2),
        ]
    )
    st.download_button("📄 Baixar PDF do Dashboard (imprimir)", data=pdf_dash,
                       file_name="dashboard_debitos.pdf", mime="application/pdf")
    components.html(
        """
        <button onclick="window.print()" style="padding:8px 12px;margin-top:8px">
          🖨️ Imprimir esta página
        </button>
        """,
        height=60
    )

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
            figsald = None
        else:
            figsald = px.bar(gsec, x="SECRETARIA", y="SALDO_LIVRE",
                             text=[format_brl(v) for v in gsec["SALDO_LIVRE"]], color="SECRETARIA")
            figsald.update_traces(hovertemplate="<b>%{x}</b><br>Saldo: %{y:,.2f}")
            figsald.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80), font=dict(size=PLOTLY_FONT_SIZE))
            st.plotly_chart(figsald, use_container_width=True)

        st.divider()
        st.subheader("📋 Contas (filtradas)")
        sal_display = sal_f.copy()
        sal_display["SALDO BANCARIO"] = sal_display["SALDO BANCARIO"].apply(format_brl)
        st.dataframe(sal_display, use_container_width=True)

        st.subheader("📥 Exportar / Imprimir")
        bsal = io.BytesIO(); sal_f.to_excel(bsal, index=False); bsal.seek(0)
        st.download_button("📊 Excel (saldos filtrados)", data=bsal,
                           file_name="saldos_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        pdf_sal = sal_display.rename(columns={"SALDO BANCARIO":"SALDO (BRL)"})
        pdf2 = gerar_pdf_listagem(pdf_sal, "Saldos - Contas Filtradas")
        st.download_button("📄 PDF (saldos filtrados - tabela)", data=pdf2,
                           file_name="saldos_filtrados.pdf", mime="application/pdf")
        sal_metrics = {
            "Saldo total": format_brl(sal_f["SALDO BANCARIO"].sum()),
            "Contas": str(len(sal_f)),
            "Secretarias": str(sal_f['SECRETARIA'].nunique())
        }
        pdf_sald_dash = gerar_pdf_dashboard(
            "Dashboard - Saldos",
            sal_metrics,
            [("Saldos por Secretaria", figsald)]
        )
        st.download_button("📄 Baixar PDF do Dashboard (imprimir)", data=pdf_sald_dash,
                           file_name="dashboard_saldos.pdf", mime="application/pdf")
        components.html(
            """
            <button onclick="window.print()" style="padding:8px 12px;margin-top:8px">
              🖨️ Imprimir esta página
            </button>
            """,
            height=60
        )
