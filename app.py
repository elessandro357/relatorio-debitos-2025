import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

st.set_page_config(layout="wide", page_title="Débitos & Plano de Pagamento 2025")
st.title("📊 Débitos por Secretaria + 💸 Plano de Pagamento (Recurso Livre)")
st.caption("Use as abas abaixo. Exporte Excel/PDF. Rateio: quem devo mais, recebe mais (sem pagar acima do devido).")

# ========= Utilidades =========
def format_brl(v):
    try:
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return v

@st.cache_data(show_spinner=False)
def load_excel(f) -> pd.DataFrame:
    df = pd.read_excel(f)
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
        errors="coerce",
    )
    v1.loc[precisa_brl] = v2
    df["VALOR"] = v1

    df["FORNECEDOR"] = df["FORNECEDOR"].astype(str).str.strip()
    df["SECRETARIA"] = df["SECRETARIA"].astype(str).str.strip()
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

def gerar_pdf_listagem(df, titulo):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt=titulo, ln=True, align="C")
    pdf.set_font("Arial", size=10)
    pdf.ln(8)
    if df.empty:
        pdf.multi_cell(0, 8, "Nenhum registro.")
    else:
        pdf.set_font("Arial", 'B', 10)
        pdf.multi_cell(0, 7, " | ".join(df.columns))
        pdf.set_font("Arial", size=10)
        pdf.ln(2)
        for _, r in df.iterrows():
            pdf.multi_cell(0, 7, " | ".join(str(r[c]) for c in df.columns))
    return io.BytesIO(pdf.output(dest="S").encode("latin-1"))

def proportional_allocation(total, debitos_series: pd.Series) -> pd.Series:
    if total <= 0 or debitos_series.sum() == 0:
        return pd.Series(0.0, index=debitos_series.index)
    base = total * (debitos_series / debitos_series.sum())
    pago = base.clip(upper=debitos_series)
    sobra = total - pago.sum()
    for _ in range(10):
        if sobra <= 0.0001: break
        restantes = debitos_series - pago
        elegiveis = restantes[restantes > 0]
        if elegiveis.empty: break
        add = sobra * (elegiveis / elegiveis.sum())
        novo = pago.add(add, fill_value=0)
        pago = pd.concat([novo, debitos_series], axis=1).min(axis=1)
        sobra = total - pago.sum()
    return pago.round(2)

# ========= Abas =========
tab_dash, tab_plano = st.tabs(["📈 Dashboard", "💸 Plano de Pagamento"])

# ======= Aba Dashboard =======
with tab_dash:
    up_deb = st.file_uploader("📁 Envie a planilha de **Débitos** (DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA)", type=["xlsx"], key="deb_dashboard")
    if not up_deb:
        st.info("Envie a planilha de Débitos para ver o dashboard.")
    else:
        df_raw = load_excel(up_deb)
        ok, miss = validar_debitos_cols(df_raw)
        if not ok:
            st.error(f"Faltam colunas em Débitos: {', '.join(miss)}")
            st.stop()
        df = cast_types_debitos(df_raw)

        # Filtros
        st.sidebar.header("🔎 Filtros (Dashboard)")
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

        # KPIs
        k1,k2,k3 = st.columns(3)
        k1.metric("Valor total filtrado", format_brl(df_f["VALOR"].sum() if not df_f.empty else 0))
        k2.metric("Registros", f"{len(df_f)}")
        k3.metric("Fornecedores", f"{df_f['FORNECEDOR'].nunique()}")

        st.divider()
        g1c,g2c = st.columns(2)
        with g1c:
            st.subheader("Débitos por Secretaria")
            if df_f.empty: st.info("Sem dados."); 
            else:
                g1 = df_f.groupby("SECRETARIA", as_index=False)["VALOR"].sum().sort_values("VALOR")
                fig1 = px.bar(g1, x="VALOR", y="SECRETARIA", orientation="h",
                              text=[format_brl(v) for v in g1["VALOR"]], color="SECRETARIA")
                fig1.update_traces(hovertemplate="<b>%{y}</b><br>Valor: %{x:,.2f}")
                fig1.update_layout(showlegend=False, margin=dict(l=10,r=10,t=30,b=10))
                st.plotly_chart(fig1, use_container_width=True)
        with g2c:
            st.subheader("Top 10 Fornecedores")
            if df_f.empty: st.info("Sem dados."); 
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

        st.subheader("📥 Exportar (Dashboard)")
        xbuf = io.BytesIO(); df_f.to_excel(xbuf, index=False); xbuf.seek(0)
        st.download_button("📊 Baixar Excel (dados filtrados)", data=xbuf,
                           file_name="dashboard_dados_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        pdf = gerar_pdf_listagem(df_disp.rename(columns={"VALOR":"VALOR (BRL)"}), "Débitos - Dados Filtrados (Dashboard)")
        st.download_button("📄 Baixar PDF (dados filtrados)", data=pdf,
                           file_name="dashboard_dados_filtrados.pdf", mime="application/pdf")

# ======= Aba Plano de Pagamento =======
with tab_plano:
    st.subheader("💸 Plano de Pagamento com Recurso Livre")
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        up_deb2 = st.file_uploader("📁 Débitos (mesmo modelo do Dashboard)", type=["xlsx"], key="deb_plano")
    with c_up2:
        up_sal = st.file_uploader("🏦 Saldos Bancários (CONTA, NOME DA CONTA, SECRETARIA, BANCO, TIPO DE RECURSO, SALDO BANCARIO)", type=["xlsx"], key="saldo_plano")

    if (up_deb2 is None) or (up_sal is None):
        st.info("Envie as duas planilhas para calcular o plano.")
    else:
        deb_raw = load_excel(up_deb2)
        okd, missd = validar_debitos_cols(deb_raw)
        if not okd:
            st.error(f"Débitos inválidos. Faltam: {', '.join(missd)}"); st.stop()
        deb = cast_types_debitos(deb_raw)

        sal_raw = load_excel(up_sal)
        oks, misss = validar_saldos_cols(sal_raw)
        if not oks:
            st.error(f"Saldos inválidos. Faltam: {', '.join(misss)}"); st.stop()

        st.markdown("**Configurações**")
        c1,c2 = st.columns(2)
        with c1:
            so_livre = st.checkbox("Considerar apenas Tipo de Recurso = LIVRE", value=True)
        with c2:
            secs_sal = st.multiselect("Filtrar saldos por Secretaria (opcional)",
                                      sorted(sal_raw["SECRETARIA"].dropna().unique().tolist()))

        sal = sal_raw.copy()
        if so_livre and "TIPO DE RECURSO" in sal.columns:
            sal = sal[sal["TIPO DE RECURSO"].str.upper()=="LIVRE"]
        if secs_sal:
            sal = sal[sal["SECRETARIA"].isin(secs_sal)]

        total_livre = pd.to_numeric(sal["SALDO BANCARIO"], errors="coerce").fillna(0).sum().round(2)
        due = deb.groupby(["FORNECEDOR","CNPJ"], as_index=False)["VALOR"].sum().rename(columns={"VALOR":"DEBITO"})

        st.write(f"**Recurso disponível (considerando filtros):** {format_brl(total_livre)}")
        st.write(f"**Total de débitos (todos fornecedores):** {format_brl(due['DEBITO'].sum())}")

        if total_livre <= 0:
            st.warning("Não há saldo livre para rateio."); st.stop()

        pagos = proportional_allocation(total_livre, due.set_index(["FORNECEDOR","CNPJ"])["DEBITO"])
        pagos = pagos.reset_index().rename(columns={0:"PAGAR_AGORA"})
        plano = due.merge(pagos, on=["FORNECEDOR","CNPJ"], how="left")
        plano["RESTANTE"] = (plano["DEBITO"] - plano["PAGAR_AGORA"]).round(2)

        plano_disp = plano.copy()
        for col in ["DEBITO","PAGAR_AGORA","RESTANTE"]:
            plano_disp[col] = plano_disp[col].apply(format_brl)

        st.subheader("📋 Plano de Pagamento (Rateio Proporcional)")
        st.dataframe(plano_disp, use_container_width=True)

        k1,k2,k3 = st.columns(3)
        k1.metric("Total a pagar agora", format_brl(plano["PAGAR_AGORA"].sum()))
        k2.metric("Fornecedores contemplados", f"{(plano['PAGAR_AGORA']>0).sum()}")
        k3.metric("Débito que permanece", format_brl(plano['RESTANTE'].clip(lower=0).sum()))

        st.subheader("📥 Exportar (Plano)")
        xbuf2 = io.BytesIO(); plano.to_excel(xbuf2, index=False); xbuf2.seek(0)
        st.download_button("📊 Baixar Excel do Plano", data=xbuf2,
                           file_name="plano_pagamento_rateio.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        pdf2 = gerar_pdf_listagem(
            plano_disp[["FORNECEDOR","CNPJ","DEBITO","PAGAR_AGORA","RESTANTE"]],
            "Plano de Pagamento - Rateio Proporcional (Recurso Livre)"
        )
        st.download_button("📄 Baixar PDF do Plano", data=pdf2,
                           file_name="plano_pagamento_rateio.pdf", mime="application/pdf")
