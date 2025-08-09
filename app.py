import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

# ================================
# Config geral
# ================================
st.set_page_config(layout="wide", page_title="Débitos • Saldos • Plano 2025")
st.title("📊 Débitos • 🏦 Saldos • 💸 Plano de Pagamento (2025)")
st.caption("Dashboards por abas + plano de pagamento proporcional por secretaria. Exporta Excel/PDF.")

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
def load_excel(f) -> pd.DataFrame:
    df = pd.read_excel(f)
    df.columns = df.columns.str.strip().str.upper()
    return df

def cast_types_debitos(df: pd.DataFrame) -> pd.DataFrame:
    """DATA robusta (dayfirst) + VALOR aceita '1.234,56'."""
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
    df["FORNECEDOR"] = df["FORNECEDOR"].astype(str).str.strip()
    df["SECRETARIA"] = df["SECRETARIA"].astype(str).str.strip()

    # Limpeza
    df = df.dropna(subset=["DATA", "VALOR", "FORNECEDOR", "SECRETARIA"]).copy()
    df["VALOR"] = df["VALOR"].round(2)
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
    if apenas_livre and "TIPO DE RECURSO" in df.columns:
        df = df[df["TIPO DE RECURSO"].str.upper()=="LIVRE"]
    df["SALDO BANCARIO"] = pd.to_numeric(df["SALDO BANCARIO"], errors="coerce").fillna(0.0)
    return df

def saldo_por_secretaria(df_saldos):
    return (df_saldos.groupby("SECRETARIA", as_index=False)["SALDO BANCARIO"]
            .sum().rename(columns={"SALDO BANCARIO":"SALDO_LIVRE"}))

def debito_por_secretaria(df_debitos):
    return (df_debitos.groupby("SECRETARIA", as_index=False)["VALOR"]
            .sum().rename(columns={"VALOR":"TOTAL_DEBITO"}))

def proportional_allocation(total, serie_debitos):
    """Rateio proporcional com teto por débito + redistribuição de sobras."""
    if total <= 0 or serie_debitos.sum() == 0:
        return pd.Series(0.0, index=serie_debitos.index)
    base = total * (serie_debitos / serie_debitos.sum())
    pago = base.clip(upper=serie_debitos)
    sobra = total - pago.sum()
    for _ in range(8):
        if sobra <= 1e-4: break
        restante = serie_debitos - pago
        eleg = restante[restante > 0]
        if eleg.empty: break
        add = sobra * (eleg / eleg.sum())
        novo = pago.add(add, fill_value=0)
        pago = pd.concat([novo, serie_debitos], axis=1).min(axis=1)
        sobra = total - pago.sum()
    return pago.round(2)

def plano_por_secretaria(df_debitos, df_saldos_livres):
    """Resumo secretaria + rateio por fornecedor dentro de cada secretaria."""
    deb_sec = debito_por_secretaria(df_debitos)
    sal_sec = saldo_por_secretaria(df_saldos_livres)
    quadro = deb_sec.merge(sal_sec, on="SECRETARIA", how="outer").fillna(0.0)
    quadro["PAGAMENTO_PREVISTO"] = quadro[["TOTAL_DEBITO","SALDO_LIVRE"]].min(axis=1).round(2)
    quadro["RESTANTE"] = (quadro["TOTAL_DEBITO"] - quadro["PAGAMENTO_PREVISTO"]).clip(lower=0).round(2)

    det = (df_debitos.groupby(["SECRETARIA","FORNECEDOR","CNPJ"], as_index=False)["VALOR"]
           .sum().rename(columns={"VALOR":"DEBITO_FORNECEDOR"}))

    planos = []
    for sec, grupo in det.groupby("SECRETARIA"):
        saldo_sec = float(quadro.loc[quadro["SECRETARIA"]==sec, "SALDO_LIVRE"].sum())
        debitos_series = grupo.set_index(["FORNECEDOR","CNPJ"])["DEBITO_FORNECEDOR"]
        pagar = proportional_allocation(saldo_sec, debitos_series)
        tmp = pagar.reset_index().rename(columns={0:"PAGAR_AGORA"})
        tmp["SECRETARIA"] = sec
        tmp = tmp.merge(grupo, on=["SECRETARIA","FORNECEDOR","CNPJ"], how="left")
        tmp["RESTANTE"] = (tmp["DEBITO_FORNECEDOR"] - tmp["PAGAR_AGORA"]).round(2)
        planos.append(tmp)

    plano = pd.concat(planos, ignore_index=True) if planos else pd.DataFrame(
        columns=["FORNECEDOR","CNPJ","PAGAR_AGORA","SECRETARIA","DEBITO_FORNECEDOR","RESTANTE"]
    )
    return quadro, plano

# ===== PDF seguro (em colunas) =====
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
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, txt=titulo, ln=True, align="C")
    pdf.ln(2)

    if df.empty:
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(0, 7, "Nenhum registro.")
        return _pdf_to_bytesio(pdf)

    pdf.set_font("Arial", size=10)
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

    pdf.set_font("Arial", 'B', 10)
    for c, w in zip(cols, widths):
        pdf.multi_cell(w, 7, _chunk_long_words(c, 20), border=0, new_x="RIGHT", new_y="TOP")
    pdf.multi_cell(0, 2, "", border=0, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Arial", size=10)

    for _, row in df.iterrows():
        for c, w in zip(cols, widths):
            txt = row[c]
            if isinstance(txt, (int, float)) and c.upper().startswith("VALOR"):
                txt = format_brl(txt)
            txt = _chunk_long_words(txt, 30)
            pdf.multi_cell(w, 6, txt, border=0, new_x="RIGHT", new_y="TOP")
        pdf.multi_cell(0, 2, "", border=0, new_x="LMARGIN", new_y="NEXT")

    return _pdf_to_bytesio(pdf)

# ================================
# ABAS
# ================================
tab_dash, tab_saldos, tab_plano = st.tabs(["📈 Dashboard Débitos", "🏦 Dashboard Saldos", "💸 Plano de Pagamento"])

# --------- Aba Débitos ---------
with tab_dash:
    up_deb = st.file_uploader("📁 Envie a planilha de **Débitos** (DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA)", type=["xlsx"], key="deb_dashboard")
    if not up_deb:
        st.info("Envie a planilha de Débitos para ver o dashboard.")
    else:
        df_raw = load_excel(up_deb)
        ok, miss = validar_debitos_cols(df_raw)
        if not ok:
            st.error(f"Faltam colunas em Débitos: {', '.join(miss)}"); st.stop()
        df = cast_types_debitos(df_raw)

        st.sidebar.header("🔎 Filtros (Débitos)")
        secs = st.sidebar.multiselect("Secretaria", sorted(df["SECRETARIA"].unique()))
        forn = st.sidebar.multiselect("Fornecedor", sorted(df["FORNECEDOR"].unique()))
        dmin = pd.to_datetime(df["DATA"].min()).date()
        dmax = pd.to_datetime(df["DATA"].max()).date()
        c1, c2 = st.sidebar.columns(2)
        with c1: din = st.date_input("Data inicial", dmin, key="d1")
        with c2: dfi = st.date_input("Data final", dmax, key="d2")
        if din > dfi:
            st.sidebar.error("Data inicial > Data final."); st.stop()

        df_f = df[(df["DATA"]>=pd.to_datetime(din)) & (df["DATA"]<=pd.to_datetime(dfi))].copy()
        if secs: df_f = df_f[df_f["SECRETARIA"].isin(secs)]
        if forn: df_f = df_f[df_f["FORNECEDOR"].isin(forn)]

        k1,k2,k3 = st.columns(3)
        k1.metric("Valor total filtrado", format_brl(df_f["VALOR"].sum() if not df_f.empty else 0))
        k2.metric("Registros", f"{len(df_f)}")
        k3.metric("Fornecedores", f"{df_f['FORNECEDOR'].nunique()}")

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
                g2 = df_f.groupby("FORNECEDOR", as_index=False)["VALOR"].sum().sort_values("VALOR", ascending=False).head(10)
                fig2 = px.bar(g2, x="FORNECEDOR", y="VALOR",
                              text=[format_brl(v) for v in g2["VALOR"]], color="FORNECEDOR")
                fig2.update_traces(hovertemplate="<b>%{x}</b><br>Valor: %{y:,.2f}")
                fig2.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80))
                st.plotly_chart(fig2, use_container_width=True)

        st.divider()
        st.subheader("📋 Dados Filtrados")
        df_disp = df_f.copy()
        df_disp["VALOR"] = df_disp["VALOR"].apply(format_brl)
        st.dataframe(df_disp, use_container_width=True)

        st.subheader("📥 Exportar (Débitos)")
        xbuf = io.BytesIO(); df_f.to_excel(xbuf, index=False); xbuf.seek(0)
        st.download_button("📊 Excel (dados filtrados)", data=xbuf,
                           file_name="debitos_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        pdf_df = df_disp.rename(columns={"VALOR":"VALOR (BRL)"})
        pdf = gerar_pdf_listagem(pdf_df, "Débitos - Dados Filtrados")
        st.download_button("📄 PDF (dados filtrados)", data=pdf,
                           file_name="debitos_filtrados.pdf", mime="application/pdf")

# --------- Aba Saldos ---------
with tab_saldos:
    up_saldos = st.file_uploader(
        "🏦 Envie a planilha de **Saldos** (CONTA, NOME DA CONTA, SECRETARIA, BANCO, TIPO DE RECURSO, SALDO BANCARIO)",
        type=["xlsx"], key="saldos_tab")
    apenas_livre = st.checkbox("Considerar apenas Recurso LIVRE", value=True)

    if not up_saldos:
        st.info("Envie a planilha de Saldos para ver o dashboard.")
    else:
        sal_raw = load_excel(up_saldos)
        ok_s, miss_s = validar_saldos_cols(sal_raw)
        if not ok_s:
            st.error(f"Saldos inválidos. Faltam: {', '.join(miss_s)}"); st.stop()
        sal = preparar_saldos(sal_raw, apenas_livre=apenas_livre)

        st.sidebar.header("🔎 Filtros (Saldos)")
        secs_sal = st.sidebar.multiselect("Secretaria (saldos)", sorted(sal["SECRETARIA"].dropna().unique()))
        bancos = st.sidebar.multiselect("Banco", sorted(sal["BANCO"].dropna().unique()))
        tipos = st.sidebar.multiselect("Tipo de Recurso", sorted(sal["TIPO DE RECURSO"].dropna().unique()))

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
        fig = px.bar(gsec, x="SECRETARIA", y="SALDO_LIVRE",
                     text=[format_brl(v) for v in gsec["SALDO_LIVRE"]], color="SECRETARIA")
        fig.update_traces(hovertemplate="<b>%{x}</b><br>Saldo: %{y:,.2f}")
        fig.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10,r=10,t=30,b=80))
        st.plotly_chart(fig, use_container_width=True)

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
        pdf2 = gerar_pdf_listagem(pdf_sal, "Saldos - Contas Filtradas")
        st.download_button("📄 PDF (saldos filtrados)", data=pdf2,
                           file_name="saldos_filtrados.pdf", mime="application/pdf")

# --------- Aba Plano ---------
with tab_plano:
    st.subheader("💸 Plano de Pagamento por Secretaria (Recurso LIVRE)")
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        up_deb2 = st.file_uploader("📁 Débitos (DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA)", type=["xlsx"], key="deb_plano2")
    with c_up2:
        up_sal2 = st.file_uploader("🏦 Saldos (CONTA, NOME DA CONTA, SECRETARIA, BANCO, TIPO DE RECURSO, SALDO BANCARIO)", type=["xlsx"], key="saldo_plano2")

    apenas_livre_plano = st.checkbox("Considerar apenas Recurso LIVRE (saldos)", value=True)

    if (up_deb2 is None) or (up_sal2 is None):
        st.info("Envie as duas planilhas para calcular o plano.")
    else:
        deb_raw = load_excel(up_deb2)
        okd, missd = validar_debitos_cols(deb_raw)
        if not okd:
            st.error(f"Débitos inválidos. Faltam: {', '.join(missd)}"); st.stop()
        deb = cast_types_debitos(deb_raw)

        sal_raw = load_excel(up_sal2)
        oks, misss = validar_saldos_cols(sal_raw)
        if not oks:
            st.error(f"Saldos inválidos. Faltam: {', '.join(misss)}"); st.stop()
        sal = preparar_saldos(sal_raw, apenas_livre=apenas_livre_plano)

        quadro_sec, plano_for = plano_por_secretaria(deb, sal)

        # KPIs
        k1,k2,k3 = st.columns(3)
        k1.metric("Saldo livre considerado", format_brl(quadro_sec["SALDO_LIVRE"].sum()))
        k2.metric("Pagamento previsto (total)", format_brl(quadro_sec["PAGAMENTO_PREVISTO"].sum()))
        k3.metric("Restante após pagamento", format_brl(quadro_sec["RESTANTE"].sum()))

        st.divider()
        st.subheader("📋 Resumo por Secretaria")
        qdisp = quadro_sec.copy()
        for c in ["TOTAL_DEBITO","SALDO_LIVRE","PAGAMENTO_PREVISTO","RESTANTE"]:
            qdisp[c] = qdisp[c].apply(format_brl)
        st.dataframe(qdisp, use_container_width=True)

        st.subheader("📋 Detalhe por Fornecedor (rateio dentro da secretaria)")
        pdisp = plano_for.copy()
        for c in ["DEBITO_FORNECEDOR","PAGAR_AGORA","RESTANTE"]:
            pdisp[c] = pdisp[c].apply(format_brl)
        st.dataframe(pdisp[["SECRETARIA","FORNECEDOR","CNPJ","DEBITO_FORNECEDOR","PAGAR_AGORA","RESTANTE"]],
                     use_container_width=True)

        st.subheader("📥 Exportar (Plano)")
        b1 = io.BytesIO(); quadro_sec.to_excel(b1, index=False); b1.seek(0)
        st.download_button("📊 Excel - Resumo por Secretaria", data=b1,
                           file_name="plano_resumo_secretaria.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        b2 = io.BytesIO(); plano_for.to_excel(b2, index=False); b2.seek(0)
        st.download_button("📊 Excel - Detalhe por Fornecedor", data=b2,
                           file_name="plano_detalhe_fornecedor.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDFs
        pdf_q = qdisp.rename(columns={
            "TOTAL_DEBITO":"TOTAL DEBITO (BRL)", "SALDO_LIVRE":"SALDO LIVRE (BRL)",
            "PAGAMENTO_PREVISTO":"PAGAMENTO PREVISTO (BRL)", "RESTANTE":"RESTANTE (BRL)"
        })
        st.download_button("📄 PDF - Resumo por Secretaria",
                           data=gerar_pdf_listagem(pdf_q, "Plano - Resumo por Secretaria"),
                           file_name="plano_resumo_secretaria.pdf", mime="application/pdf")

        pdf_p = pdisp.rename(columns={
            "DEBITO_FORNECEDOR":"DEBITO (BRL)",
            "PAGAR_AGORA":"PAGAR AGORA (BRL)",
            "RESTANTE":"RESTANTE (BRL)"
        })[["SECRETARIA","FORNECEDOR","CNPJ","DEBITO (BRL)","PAGAR AGORA (BRL)","RESTANTE (BRL)"]]
        st.download_button("📄 PDF - Detalhe por Fornecedor",
                           data=gerar_pdf_listagem(pdf_p, "Plano - Detalhe por Fornecedor"),
                           file_name="plano_detalhe_fornecedor.pdf", mime="application/pdf")
