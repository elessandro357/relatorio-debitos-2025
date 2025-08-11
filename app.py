# app.py — Débitos e Saldos (versão estável)
# --------------------------------------------------------------
# • Abas: Débitos e Saldos
# • Upload de planilhas (CSV/XLS/XLSX)
# • Mapeador de colunas quando cabeçalhos diferem
# • Filtros na sidebar + KPIs
# • Gráficos interativos (Plotly) — sem exportar imagem
# • Exportação: Excel (com Resumo) e PDF tabelado
# --------------------------------------------------------------

import io
import pandas as pd
import streamlit as st
import plotly.express as px
from fpdf import FPDF

# ----------------------- Config -----------------------
st.set_page_config(layout="wide", page_title="Débitos e Saldos — Painel")
st.title("📊 Débitos • 🏦 Saldos — Painel")
st.caption("Upload de planilhas, filtros, KPIs, gráficos e exportações (Excel/PDF).")

# --------------------- Utilidades ---------------------
BRL_EXCEL_FMT = u'[$R$-416] #,##0.00'

def format_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

@st.cache_data(show_spinner=False, ttl=300)
def load_table(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        # rápido e robusto para CSV
        try:
            df = pd.read_csv(uploaded_file)
        except Exception:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=";")
    else:
        df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()
    return df

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

def cast_types_debitos(df: pd.DataFrame) -> pd.DataFrame:
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

# --------------------- PDF Tabelado ---------------------
class PDFListagem(FPDF):
    def footer(self):
        self.set_y(-12)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

def _pdf_to_bytes(pdf_obj):
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
        return _pdf_to_bytes(pdf)

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

        if total_cols:
            pdf.set_font("Arial", 'B', 10)
            tot_line = " | ".join([f"{c}: {format_brl(gdf[c].sum())}" for c in total_cols])
            pdf.multi_cell(0, 8, f"Totais do grupo → {tot_line}", border=0)
            pdf.ln(2)

    return _pdf_to_bytes(pdf)

# ====================== ABAS ======================
tab_deb, tab_sald = st.tabs(["📈 Débitos", "🏦 Saldos"])

# -------------------- Débitos --------------------
with tab_deb:
    st.subheader("📥 Envie a planilha de Débitos")
    up_deb = st.file_uploader("Arquivos: CSV/XLS/XLSX", type=["csv","xls","xlsx"], key="deb")

    if not up_deb:
        st.info("Envie a planilha de Débitos para ver o painel.")
    else:
        raw = load_table(up_deb)
        ok, miss, req = validar_debitos_cols(raw.columns)
        if not ok:
            st.warning("Faltam colunas obrigatórias. Use o mapeador abaixo.")
            mapa = coluna_mapper_ui(raw.columns, req, key_prefix="deb_map")
            df_m = aplicar_mapeamento(raw, mapa)
        else:
            df_m = raw[req].copy()

        df = cast_types_debitos(df_m)

        # -------- Filtros --------
        st.sidebar.header("🔎 Filtros — Débitos")
        dmin = pd.to_datetime(df["DATA"]).min().date()
        dmax = pd.to_datetime(df["DATA"]).max().date()
        c1, c2 = st.sidebar.columns(2)
        with c1:
            di = st.date_input("Data inicial", dmin, key="deb_di")
        with c2:
            dfim = st.date_input("Data final", dmax, key="deb_df")
        if di > dfim:
            st.sidebar.error("Data inicial > Data final."); st.stop()

        secs = sorted(df["SECRETARIA"].unique().tolist())
        forns = sorted(df["FORNECEDOR"].unique().tolist())
        f_secs = st.sidebar.multiselect("Secretaria", secs)
        f_forns = st.sidebar.multiselect("Fornecedor", forns)

        df_f = df[(df["DATA"] >= pd.to_datetime(di)) & (df["DATA"] <= pd.to_datetime(dfim))].copy()
        if f_secs:
            df_f = df_f[df_f["SECRETARIA"].isin(f_secs)]
        if f_forns:
            df_f = df_f[df_f["FORNECEDOR"].isin(f_forns)]

        # -------- KPIs --------
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Total (filtrado)", format_brl(df_f["VALOR"].sum() if not df_f.empty else 0))
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

        st.divider()
        st.subheader("📋 Dados Filtrados — Débitos")
        df_show = df_f.copy()
        df_show["VALOR"] = df_show["VALOR"].apply(format_brl)
        st.dataframe(df_show, use_container_width=True)
        st.markdown(f"**Total exibido:** {format_brl(df_f['VALOR'].sum() if not df_f.empty else 0)}")

        st.subheader("📥 Exportar — Débitos")
        # Excel + Resumo
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
            df_f.to_excel(xw, index=False, sheet_name="Debitos")
            ws = xw.sheets["Debitos"]
            # Coluna D (VALOR) com formato BRL
            for row in range(2, len(df_f)+2):
                ws[f"D{row}"].number_format = BRL_EXCEL_FMT
            resumo = pd.DataFrame({
                "Métrica":["Total filtrado","Registros","Fornecedores","Secretarias"],
                "Valor":[df_f["VALOR"].sum(), len(df_f), df_f["FORNECEDOR"].nunique(), df_f["SECRETARIA"].nunique()]
            })
            resumo.to_excel(xw, index=False, sheet_name="Resumo")
            xw.sheets["Resumo"]["B2"].number_format = BRL_EXCEL_FMT
        xbuf.seek(0)
        st.download_button("⬇️ Excel (dados + Resumo)", data=xbuf,
                           file_name="debitos_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF (quebrado por Secretaria)
        pdf_df = df_f.copy()
        pdf_bytes = gerar_pdf_tabelado(pdf_df[["DATA","FORNECEDOR","CNPJ","VALOR","SECRETARIA"]],
                                       "Débitos — Dados Filtrados", quebra_por="SECRETARIA")
        st.download_button("⬇️ PDF (quebrado por Secretaria)", data=pdf_bytes,
                           file_name="debitos_filtrados.pdf", mime="application/pdf")

# -------------------- Saldos --------------------
with tab_sald:
    st.subheader("📥 Envie a planilha de Saldos")
    up_sald = st.file_uploader("Arquivos: CSV/XLS/XLSX", type=["csv","xls","xlsx"], key="sald")

    apenas_livre = st.checkbox("Considerar apenas Recurso LIVRE", value=True)

    if not up_sald:
        st.info("Envie a planilha de Saldos para ver o painel.")
    else:
        raw = load_table(up_sald)
        ok, miss, req = validar_saldos_cols(raw.columns)
        if not ok:
            st.warning("Faltam colunas obrigatórias. Use o mapeador abaixo.")
            mapa = coluna_mapper_ui(raw.columns, req, key_prefix="sald_map")
            df_m = aplicar_mapeamento(raw, mapa)
        else:
            df_m = raw[req].copy()

        sal = preparar_saldos(df_m)
        if apenas_livre and "TIPO DE RECURSO" in sal.columns:
            sal = sal[sal["TIPO DE RECURSO"].str.upper()=="LIVRE"].copy()

        # -------- Filtros --------
        st.sidebar.header("🔎 Filtros — Saldos")
        secs = sorted(sal["SECRETARIA"].astype(str).unique().tolist())
        bancos = sorted(sal["BANCO"].astype(str).unique().tolist())
        tipos = sorted(sal["TIPO DE RECURSO"].astype(str).unique().tolist()) if "TIPO DE RECURSO" in sal.columns else []

        f_secs = st.sidebar.multiselect("Secretaria (saldos)", secs)
        f_bancos = st.sidebar.multiselect("Banco", bancos)
        f_tipos = st.sidebar.multiselect("Tipo de Recurso", tipos)

        sal_f = sal.copy()
        if f_secs: sal_f = sal_f[sal_f["SECRETARIA"].astype(str).isin(f_secs)]
        if f_bancos: sal_f = sal_f[sal_f["BANCO"].astype(str).isin(f_bancos)]
        if f_tipos and "TIPO DE RECURSO" in sal_f.columns:
            sal_f = sal_f[sal_f["TIPO DE RECURSO"].astype(str).isin(f_tipos)]

        # -------- KPIs --------
        k1,k2,k3 = st.columns(3)
        k1.metric("Saldo total (filtrado)", format_brl(sal_f["SALDO BANCARIO"].sum()))
        k2.metric("Contas", f"{len(sal_f)}")
        k3.metric("Secretarias", f"{sal_f['SECRETARIA'].nunique()}")

        st.divider()
        st.subheader("Saldos por Secretaria")
        gsec = (sal_f.groupby("SECRETARIA", as_index=False)["SALDO BANCARIO"]
                .sum().sort_values("SALDO BANCARIO", ascending=False))
        if gsec.empty:
            st.info("Sem dados.")
        else:
            fig = px.bar(gsec, x="SECRETARIA", y="SALDO BANCARIO",
                         text=[format_brl(v) for v in gsec["SALDO BANCARIO"]],
                         color="SECRETARIA")
            fig.update_traces(hovertemplate="<b>%{x}</b><br>Saldo: %{y:,.2f}")
            fig.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80))
            st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.subheader("📋 Contas — Dados Filtrados")
        sal_show = sal_f.copy()
        sal_show["SALDO BANCARIO"] = sal_show["SALDO BANCARIO"].apply(format_brl)
        st.dataframe(sal_show, use_container_width=True)
        st.markdown(f"**Total exibido:** {format_brl(sal_f['SALDO BANCARIO'].sum())}")

        st.subheader("📥 Exportar — Saldos")
        # Excel + Resumo
        bsal = io.BytesIO()
        with pd.ExcelWriter(bsal, engine="openpyxl") as xw:
            sal_f.to_excel(xw, index=False, sheet_name="Saldos")
            ws = xw.sheets["Saldos"]
            # Coluna F (SALDO) com formato BRL
            for row in range(2, len(sal_f)+2):
                ws[f"F{row}"].number_format = BRL_EXCEL_FMT
            resumo = pd.DataFrame({
                "Métrica":["Saldo total","Contas","Secretarias"],
                "Valor":[sal_f["SALDO BANCARIO"].sum(), len(sal_f), sal_f["SECRETARIA"].nunique()]
            })
            resumo.to_excel(xw, index=False, sheet_name="Resumo")
            xw.sheets["Resumo"]["B2"].number_format = BRL_EXCEL_FMT
        bsal.seek(0)
        st.download_button("⬇️ Excel (dados + Resumo)", data=bsal,
                           file_name="saldos_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF (quebrado por Secretaria)
        pdf_sal = gerar_pdf_tabelado(
            sal_f[["CONTA","NOME DA CONTA","SECRETARIA","BANCO","TIPO DE RECURSO","SALDO BANCARIO"]],
            "Saldos — Contas Filtradas", quebra_por="SECRETARIA"
        )
        st.download_button("⬇️ PDF (quebrado por Secretaria)", data=pdf_sal,
                           file_name="saldos_filtrados.pdf", mime="application/pdf")
