# app.py
# ============================================================
# Relatório de Débitos • Saldos (2025)
# Dashboards de Débitos e Saldos com:
# - Upload CSV/XLS/XLSX
# - Mapeador de colunas (quando cabeçalhos diferem)
# - Validações (datas/valores/CNPJ), consolidação e outliers
# - Exportações: Excel (com formatação BRL + Resumo), PDF (tabelado)
# - Exportar gráficos: PNG + PDF do dashboard (via kaleido + fpdf2)
# Requisitos (requirements.txt): streamlit, pandas, plotly, fpdf2, openpyxl, kaleido
# ============================================================

import io
import tempfile
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from fpdf import FPDF

# ========= Plotly (para PNG via kaleido) =========
import plotly.io as pio  # noqa: F401  (import necessário para .to_image funcionar)

# ================================
# Config geral
# ================================
st.set_page_config(layout="wide", page_title="Débitos • Saldos 2025")
st.title("📊 Débitos • 🏦 Saldos — 2025")
st.caption("Dashboards por abas. Exports (Excel/PDF). Mapeamento de colunas, validações, duplicados, outliers e exportação de gráficos.")

# ================================
# Utilidades / Helpers
# ================================
BRL_EXCEL_FMT = u'[$R$-416] #,##0.00'

def format_brl(v):
    """R$ 1.234,56 sem depender de locale."""
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

@st.cache_data(show_spinner=False, ttl=300)
def load_table(uploaded_file) -> pd.DataFrame:
    """Lê CSV/XLS/XLSX e normaliza cabeçalhos (CAIXA ALTA, trim)."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=None, engine="python")
    else:
        df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()
    return df

def cast_types_debitos(df: pd.DataFrame) -> pd.DataFrame:
    """DATA robusta (dayfirst) + VALOR aceita '1.234,56' + validações básicas."""
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
    df["VALOR"] = v1.clip(lower=0)

    # Texto
    df["FORNECEDOR"] = df["FORNECEDOR"].astype(str).str.strip()
    df["SECRETARIA"] = df["SECRETARIA"].astype(str).str.strip()

    # CNPJ (se existir)
    if "CNPJ" in df.columns:
        df["CNPJ"] = df["CNPJ"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(14)

    # Limpeza
    df = df.dropna(subset=["DATA", "VALOR", "FORNECEDOR", "SECRETARIA"]).copy()
    df["VALOR"] = df["VALOR"].round(2)

    # Tipos leves
    df["FORNECEDOR"] = df["FORNECEDOR"].astype("category")
    df["SECRETARIA"] = df["SECRETARIA"].astype("category")
    return df

def validar_debitos_cols(cols):
    req = ["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]
    miss = [c for c in req if c not in cols]
    return len(miss)==0, miss, req

def validar_saldos_cols(cols):
    req = ["CONTA","NOME DA CONTA","SECRETARIA","BANCO","TIPO DE RECURSO","SALDO BANCARIO"]
    miss = [c for c in req if c not in cols]
    return len(miss)==0, miss, req

def preparar_saldos(df_raw, apenas_livre=True):
    df = df_raw.copy()
    df.columns = df.columns.str.strip().str.upper()
    if apenas_livre and "TIPO DE RECURSO" in df.columns:
        df = df[df["TIPO DE RECURSO"].str.upper()=="LIVRE"]
    df["SALDO BANCARIO"] = pd.to_numeric(df["SALDO BANCARIO"], errors="coerce").fillna(0.0)
    df["SECRETARIA"] = df["SECRETARIA"].astype(str).str.strip().astype("category")
    df["BANCO"] = df["BANCO"].astype(str).str.strip().astype("category")
    if "TIPO DE RECURSO" in df.columns:
        df["TIPO DE RECURSO"] = df["TIPO DE RECURSO"].astype(str).str.strip().astype("category")
    return df

def saldo_por_secretaria(df_saldos):
    return (df_saldos.groupby("SECRETARIA", as_index=False)["SALDO BANCARIO"]
            .sum().rename(columns={"SALDO BANCARIO":"SALDO_LIVRE"}))

# ===== Mapeador de Colunas =====
def coluna_mapper_ui(cols_atual, req_cols, key_prefix):
    st.info("Mapeie suas colunas para o modelo esperado.")
    mapeamento = {}
    for alvo in req_cols:
        mapeamento[alvo] = st.selectbox(
            f"Coluna no arquivo para **{alvo}**",
            options=["(não existe)"] + list(cols_atual),
            index=(["(não existe)"]+list(cols_atual)).index(alvo) if alvo in cols_atual else 0,
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
    df_m = pd.DataFrame(cols_novas)
    return df_m

# ===== PDF seguro (em colunas, com rodapé) =====
class PDFListagem(FPDF):
    def footer(self):
        self.set_y(-12)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

def _pdf_to_bytesio(pdf_obj):
    out = pdf_obj.output(dest="S")
    return out if isinstance(out, (bytes, bytearray)) else out.encode("latin-1", "ignore")

def _chunk_long_words(text, maxlen=30):
    s = "" if pd.isna(text) else str(text)
    parts = []
    for w in s.split():
        if len(w) > maxlen:
            parts.extend([w[i:i+maxlen] for i in range(0, len(w), maxlen)])
        else:
            parts.append(w)
    return " ".join(parts)

def gerar_pdf_tabelado(df: pd.DataFrame, titulo="Relatório", quebra_por="SECRETARIA"):
    pdf = PDFListagem()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, txt=titulo, ln=True, align="C")
    pdf.ln(2)

    if df.empty:
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(0, 7, "Nenhum registro.")
        return _pdf_to_bytesio(pdf)

    cols = list(df.columns)
    epw = pdf.w - 2 * pdf.l_margin
    widths = [epw / len(cols)] * len(cols)

    grupos = [(None, df)]
    if quebra_por in df.columns:
        grupos = list(df.groupby(quebra_por, sort=True))

    total_cols = [c for c in cols if any(k in c.upper() for k in ["VALOR","SALDO"])]

    for gnome, gdf in grupos:
        pdf.set_font("Arial", 'B', 11)
        if gnome is not None:
            pdf.cell(0, 8, f"{quebra_por}: {gnome}", ln=True)
        pdf.set_font("Arial", 'B', 10)
        for c, w in zip(cols, widths):
            pdf.multi_cell(w, 7, _chunk_long_words(c, 20), border=0, new_x="RIGHT", new_y="TOP")
        pdf.multi_cell(0, 2, "", border=0, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Arial", size=10)

        for _, row in gdf.iterrows():
            for c, w in zip(cols, widths):
                txt = row[c]
                if isinstance(txt, (int, float)) and any(k in c.upper() for k in ["VALOR","SALDO"]):
                    txt = format_brl(txt)
                txt = _chunk_long_words(txt, 30)
                pdf.multi_cell(w, 6, txt, border=0, new_x="RIGHT", new_y="TOP")
            pdf.multi_cell(0, 2, "", border=0, new_x="LMARGIN", new_y="NEXT")

        # totais por grupo
        if total_cols:
            pdf.set_font("Arial", 'B', 10)
            tot_line = " | ".join([f"{c}: {format_brl(gdf[c].sum())}" for c in total_cols])
            pdf.multi_cell(0, 8, f"Totais do grupo → {tot_line}", border=0)
            pdf.ln(2)

    return _pdf_to_bytesio(pdf)

# ===== Exportar imagens dos gráficos (PNG) =====
def fig_to_png_bytes(fig, scale=2):
    """Gera PNG (bytes) de um gráfico Plotly (precisa de 'kaleido')."""
    return fig.to_image(format="png", scale=scale)

# ===== Templates =====
def gerar_template_debitos() -> io.BytesIO:
    cols = ["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]
    df = pd.DataFrame(columns=cols)
    df.loc[0] = ["01/01/2025","Fornecedor Exemplo LTDA","12345678000199", "1.234,56","SAÚDE"]
    df.loc[1] = ["05/01/2025","ACME Serviços","11222333000188", "987,10","EDUCAÇÃO"]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Debitos")
        ws = xw.sheets["Debitos"]
        for row in range(2, 1002):
            ws[f"D{row}"].number_format = BRL_EXCEL_FMT
    buf.seek(0)
    return buf

def gerar_template_saldos() -> io.BytesIO:
    cols = ["CONTA","NOME DA CONTA","SECRETARIA","BANCO","TIPO DE RECURSO","SALDO BANCARIO"]
    df = pd.DataFrame(columns=cols)
    df.loc[0] = ["123-4","Conta Saúde","SAÚDE","Banco X","LIVRE", 150000.00]
    df.loc[1] = ["987-0","Conta Educação","EDUCAÇÃO","Banco Y","VINCULADO", 50000.00]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Saldos")
        ws = xw.sheets["Saldos"]
        for row in range(2, 1002):
            ws[f"F{row}"].number_format = BRL_EXCEL_FMT
    buf.seek(0)
    return buf

# ================================
# ABAS
# ================================
tab_deb, tab_sald = st.tabs(["📈 Dashboard Débitos", "🏦 Dashboard Saldos"])

# -------------------- Débitos --------------------
with tab_deb:
    st.subheader("📥 Entrada de Dados — Débitos")
    c1, c2 = st.columns([2,1])
    with c1:
        up_deb = st.file_uploader("Envie Débitos (CSV/XLS/XLSX)", type=["csv","xls","xlsx"], key="deb_upload")
    with c2:
        st.markdown("**Modelos**")
        st.download_button("📄 Baixar Template Débitos", data=gerar_template_debitos(),
                           file_name="template_debitos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if not up_deb:
        st.info("Envie a planilha de Débitos para ver o dashboard.")
    else:
        df_raw = load_table(up_deb)
        ok, miss, req = validar_debitos_cols(df_raw.columns)
        if not ok:
            st.warning("Faltam colunas obrigatórias. Use o mapeador abaixo.")
            mapa = coluna_mapper_ui(df_raw.columns, req, key_prefix="deb")
            df_m = aplicar_mapeamento(df_raw, mapa)
        else:
            df_m = df_raw[req].copy()

        # Opções de pré-processamento
        st.markdown("### ⚙️ Opções")
        colA, colB, colC = st.columns(3)
        with colA:
            consolidar = st.checkbox("Consolidar duplicados (DATA, FORNECEDOR, SECRETARIA)", value=False)
        with colB:
            marcar_outliers = st.checkbox("Marcar outliers (> p95 por secretaria)", value=True)
        with colC:
            limpar_filtros_click = st.button("🧹 Limpar filtros")

        # Converte tipos
        df = cast_types_debitos(df_m)

        # Consolidar duplicados
        if consolidar:
            df = (df.groupby(["DATA","FORNECEDOR","CNPJ","SECRETARIA"], as_index=False)["VALOR"]
                    .sum().sort_values("DATA"))

        # Outliers
        if marcar_outliers and not df.empty:
            p95 = df.groupby("SECRETARIA")["VALOR"].transform(lambda s: s.quantile(0.95))
            df["ALERTA_OUTLIER"] = (df["VALOR"] > p95).map({True:"ALTO", False:""})
        else:
            df["ALERTA_OUTLIER"] = ""

        # -------- Filtros (persistentes) --------
        if limpar_filtros_click:
            for k in ["deb_secs","deb_forns","deb_dini","deb_dfim"]:
                st.session_state.pop(k, None)

        st.sidebar.header("🔎 Filtros (Débitos)")
        secs_opt = sorted(df["SECRETARIA"].astype(str).unique().tolist())
        forns_opt = sorted(df["FORNECEDOR"].astype(str).unique().tolist())

        din_default = pd.to_datetime(df["DATA"].min()).date()
        dfi_default = pd.to_datetime(df["DATA"].max()).date()

        secs = st.sidebar.multiselect("Secretaria", secs_opt, default=st.session_state.get("deb_secs", []), key="deb_secs")
        forns = st.sidebar.multiselect("Fornecedor", forns_opt, default=st.session_state.get("deb_forns", []), key="deb_forns")
        c1, c2 = st.sidebar.columns(2)
        with c1:
            din = st.date_input("Data inicial", st.session_state.get("deb_dini", din_default), key="deb_dini")
        with c2:
            dfim = st.date_input("Data final", st.session_state.get("deb_dfim", dfi_default), key="deb_dfim")

        if din > dfim:
            st.sidebar.error("Data inicial > Data final."); st.stop()

        df_f = df[(df["DATA"]>=pd.to_datetime(din)) & (df["DATA"]<=pd.to_datetime(dfim))].copy()
        if secs: df_f = df_f[df_f["SECRETARIA"].astype(str).isin(secs)]
        if forns: df_f = df_f[df_f["FORNECEDOR"].astype(str).isin(forns)]

        # KPIs
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Valor total filtrado", format_brl(df_f["VALOR"].sum() if not df_f.empty else 0))
        k2.metric("Registros", f"{len(df_f)}")
        k3.metric("Fornecedores", f"{df_f['FORNECEDOR'].nunique()}")
        k4.metric("Secretarias", f"{df_f['SECRETARIA'].nunique()}")

        st.divider()
        g1c,g2c = st.columns(2)
        with g1c:
            st.subheader("Débitos por Secretaria")
            if df_f.empty:
                st.info("Sem dados.")
            else:
                g1 = df_f.groupby("SECRETARIA", as_index=False)["VALOR"].sum().sort_values("VALOR")
                fig1 = px.bar(g1, x="VALOR", y="SECRETARIA", orientation="h",
                              text=[format_brl(v) for v in g1["VALOR"]], color="SECRETARIA")
                fig1.update_traces(hovertemplate="<b>%{y}</b><br>Valor: %{x:,.2f}")
                fig1.update_layout(showlegend=False, margin=dict(l=10,r=10,t=30,b=10))
                st.plotly_chart(fig1, use_container_width=True)
        with g2c:
            st.subheader("Top 10 Fornecedores")
            if df_f.empty:
                st.info("Sem dados.")
            else:
                g2 = (df_f.groupby("FORNECEDOR", as_index=False)["VALOR"]
                      .sum().sort_values("VALOR", ascending=False).head(10))
                fig2 = px.bar(g2, x="FORNECEDOR", y="VALOR",
                              text=[format_brl(v) for v in g2["VALOR"]], color="FORNECEDOR")
                fig2.update_traces(hovertemplate="<b>%{x}</b><br>Valor: %{y:,.2f}")
                fig2.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80))
                st.plotly_chart(fig2, use_container_width=True)

        # ====== Exportar imagens dos gráficos (Débitos) ======
        st.subheader("🖼️ Exportar gráficos (Débitos)")
        png1 = fig_to_png_bytes(fig1) if 'fig1' in locals() and df_f.shape[0] > 0 else None
        png2 = fig_to_png_bytes(fig2) if 'fig2' in locals() and df_f.shape[0] > 0 else None

        col_img1, col_img2, col_img3 = st.columns(3)
        with col_img1:
            if png1:
                st.download_button("⬇️ PNG — Débitos por Secretaria",
                                   data=png1, file_name="debitos_por_secretaria.png", mime="image/png")
        with col_img2:
            if png2:
                st.download_button("⬇️ PNG — Top 10 Fornecedores",
                                   data=png2, file_name="top10_fornecedores.png", mime="image/png")
        with col_img3:
            if png1 or png2:
                pdf = FPDF(orientation="L", unit="mm", format="A4")
                if png1:
                    pdf.add_page()
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp1:
                        tmp1.write(png1); tmp1.flush()
                        pdf.image(tmp1.name, x=10, y=10, w=277)
                if png2:
                    pdf.add_page()
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp2:
                        tmp2.write(png2); tmp2.flush()
                        pdf.image(tmp2.name, x=10, y=10, w=277)
                out = _pdf_to_bytesio(pdf)
                st.download_button("📄 PDF — Dashboard Débitos",
                                   data=out,
                                   file_name="dashboard_debitos_graficos.pdf",
                                   mime="application/pdf")

        st.divider()
        st.subheader("📋 Dados Filtrados")
        df_disp = df_f.copy()
        df_disp["VALOR"] = df_disp["VALOR"].apply(format_brl)
        st.dataframe(df_disp, use_container_width=True)
        st.markdown(f"**Total exibido:** {format_brl(df_f['VALOR'].sum() if not df_f.empty else 0)}")

        st.subheader("📥 Exportar (Débitos)")
        # Excel com aba Resumo e formatação BRL
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
            df_f.to_excel(xw, index=False, sheet_name="Debitos")
            ws = xw.sheets["Debitos"]
            # Coluna VALOR (4ª) → BRL
            for row in range(2, len(df_f)+2):
                ws[f"D{row}"].number_format = BRL_EXCEL_FMT
            # Resumo
            resumo = pd.DataFrame({
                "Métrica":["Total filtrado","Registros","Fornecedores","Secretarias"],
                "Valor":[df_f["VALOR"].sum(), len(df_f), df_f["FORNECEDOR"].nunique(), df_f["SECRETARIA"].nunique()]
            })
            resumo.to_excel(xw, index=False, sheet_name="Resumo")
            ws2 = xw.sheets["Resumo"]
            ws2["B2"].number_format = BRL_EXCEL_FMT
        xbuf.seek(0)
        st.download_button("📊 Excel (dados filtrados + Resumo)", data=xbuf,
                           file_name="debitos_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF com quebra por secretaria
        pdf_df = df_f.copy()
        pdf_df["VALOR"] = pdf_df["VALOR"].round(2)
        pdf_bytes = gerar_pdf_tabelado(pdf_df[["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]],
                                       "Débitos — Dados Filtrados", quebra_por="SECRETARIA")
        st.download_button("📄 PDF (quebrado por Secretaria)", data=pdf_bytes,
                           file_name="debitos_filtrados.pdf", mime="application/pdf")

# -------------------- Saldos --------------------
with tab_sald:
    st.subheader("📥 Entrada de Dados — Saldos")
    c1, c2 = st.columns([2,1])
    with c1:
        up_sald = st.file_uploader("Envie Saldos (CSV/XLS/XLSX)", type=["csv","xls","xlsx"], key="sald_upload")
    with c2:
        st.markdown("**Modelos**")
        st.download_button("📄 Baixar Template Saldos", data=gerar_template_saldos(),
                           file_name="template_saldos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    apenas_livre_ck = st.checkbox("Considerar apenas Recurso LIVRE", value=True)

    if not up_sald:
        st.info("Envie a planilha de Saldos para ver o dashboard.")
    else:
        sal_raw = load_table(up_sald)
        ok_s, miss_s, req_s = validar_saldos_cols(sal_raw.columns)
        if not ok_s:
            st.warning("Faltam colunas obrigatórias. Use o mapeador abaixo.")
            mapa = coluna_mapper_ui(sal_raw.columns, req_s, key_prefix="sal")
            sal_map = aplicar_mapeamento(sal_raw, mapa)
        else:
            sal_map = sal_raw[req_s].copy()

        sal = preparar_saldos(sal_map, apenas_livre=apenas_livre_ck)

        # Limpar filtros
        if st.button("🧹 Limpar filtros (Saldos)"):
            for k in ["sal_secs","sal_bancos","sal_tipos"]:
                st.session_state.pop(k, None)

        st.sidebar.header("🔎 Filtros (Saldos)")
        secs_opt = sorted(sal["SECRETARIA"].astype(str).unique().tolist())
        bancos_opt = sorted(sal["BANCO"].astype(str).unique().tolist())
        tipos_opt = sorted(sal["TIPO DE RECURSO"].astype(str).unique().tolist()) if "TIPO DE RECURSO" in sal.columns else []

        secs = st.sidebar.multiselect("Secretaria (saldos)", secs_opt, default=st.session_state.get("sal_secs", []), key="sal_secs")
        bancos = st.sidebar.multiselect("Banco", bancos_opt, default=st.session_state.get("sal_bancos", []), key="sal_bancos")
        tipos = st.sidebar.multiselect("Tipo de Recurso", tipos_opt, default=st.session_state.get("sal_tipos", []), key="sal_tipos")

        sal_f = sal.copy()
        if secs: sal_f = sal_f[sal_f["SECRETARIA"].astype(str).isin(secs)]
        if bancos: sal_f = sal_f[sal_f["BANCO"].astype(str).isin(bancos)]
        if tipos and "TIPO DE RECURSO" in sal_f.columns:
            sal_f = sal_f[sal_f["TIPO DE RECURSO"].astype(str).isin(tipos)]

        # KPIs
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

        # ====== Exportar imagem/PDF do gráfico (Saldos) ======
        st.subheader("🖼️ Exportar gráficos (Saldos)")
        png_saldos = fig_to_png_bytes(fig) if 'fig' in locals() and not gsec.empty else None

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            if png_saldos:
                st.download_button("⬇️ PNG — Saldos por Secretaria",
                                   data=png_saldos, file_name="saldos_por_secretaria.png", mime="image/png")
        with col_s2:
            if png_saldos:
                pdf_s = FPDF(orientation="L", unit="mm", format="A4")
                pdf_s.add_page()
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    tmp.write(png_saldos); tmp.flush()
                    pdf_s.image(tmp.name, x=10, y=10, w=277)
                out = _pdf_to_bytesio(pdf_s)
                st.download_button("📄 PDF — Dashboard Saldos",
                                   data=out,
                                   file_name="dashboard_saldos_grafico.pdf",
                                   mime="application/pdf")

        st.divider()
        st.subheader("📋 Contas (filtradas)")
        sal_disp = sal_f.copy()
        sal_disp["SALDO BANCARIO"] = sal_disp["SALDO BANCARIO"].apply(format_brl)
        st.dataframe(sal_disp, use_container_width=True)
        st.markdown(f"**Total exibido:** {format_brl(sal_f['SALDO BANCARIO'].sum())}")

        st.subheader("📥 Exportar (Saldos)")
        # Excel (dados + resumo)
        bsal = io.BytesIO()
        with pd.ExcelWriter(bsal, engine="openpyxl") as xw:
            sal_f.to_excel(xw, index=False, sheet_name="Saldos")
            ws = xw.sheets["Saldos"]
            # Coluna SALDO (6ª) → BRL
            for row in range(2, len(sal_f)+2):
                ws[f"F{row}"].number_format = BRL_EXCEL_FMT
            resumo = pd.DataFrame({
                "Métrica":["Saldo total","Contas","Secretarias"],
                "Valor":[sal_f["SALDO BANCARIO"].sum(), len(sal_f), sal_f["SECRETARIA"].nunique()]
            })
            resumo.to_excel(xw, index=False, sheet_name="Resumo")
            ws2 = xw.sheets["Resumo"]
            ws2["B2"].number_format = BRL_EXCEL_FMT
        bsal.seek(0)
        st.download_button("📊 Excel (saldos filtrados + Resumo)", data=bsal,
                           file_name="saldos_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF (tabelado, quebrado por secretaria)
        pdf_sal = gerar_pdf_tabelado(
            sal_f[["CONTA","NOME DA CONTA","SECRETARIA","BANCO","TIPO DE RECURSO","SALDO BANCARIO"]],
            "Saldos — Contas Filtradas", quebra_por="SECRETARIA"
        )
        st.download_button("📄 PDF (quebrado por Secretaria)", data=pdf_sal,
                           file_name="saldos_filtrados.pdf", mime="application/pdf")
