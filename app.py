import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

# =========================================
# Configuração do app
# =========================================
st.set_page_config(layout="wide", page_title="Débitos & Plano de Pagamento 2025")
st.title("📊 Débitos por Secretaria + 💸 Plano de Pagamento com Saldos Livres (2025)")
st.caption("Envie as planilhas e use os filtros. Exporte Excel/PDF. Rateio proporcional: quem devo mais recebe mais.")

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
def load_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip().str.upper()
    return df

def cast_types_debitos(df: pd.DataFrame) -> pd.DataFrame:
    """Conversões robustas para a planilha de débitos (DATA/VALOR PT-BR)."""
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

    # Limpeza mínima
    df["FORNECEDOR"] = df["FORNECEDOR"].astype(str).str.strip()
    df["SECRETARIA"] = df["SECRETARIA"].astype(str).str.strip()

    df = df.dropna(subset=["DATA", "VALOR", "FORNECEDOR", "SECRETARIA"]).copy()
    df["VALOR"] = df["VALOR"].round(2)
    return df

def gerar_pdf_listagem(dataframe: pd.DataFrame, titulo="Relatório"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt=titulo, ln=True, align="C")
    pdf.set_font("Arial", size=10)
    pdf.ln(8)

    if dataframe.empty:
        pdf.multi_cell(0, 8, "Nenhum registro.")
    else:
        # Cabeçalho
        pdf.set_font("Arial", 'B', 10)
        cols = list(dataframe.columns)
        pdf.multi_cell(0, 7, " | ".join(cols))
        pdf.set_font("Arial", size=10)
        pdf.ln(2)

        for _, row in dataframe.iterrows():
            linha = " | ".join(str(row[c]) for c in cols)
            pdf.multi_cell(0, 7, linha)

    pdf_bytes = pdf.output(dest="S").encode("latin-1")
    return io.BytesIO(pdf_bytes)

def validate_debitos_cols(df: pd.DataFrame):
    required = ["DATA", "FORNECEDOR", "CNPJ", "VALOR", "SECRETARIA"]
    missing = [c for c in required if c not in df.columns]
    return (len(missing) == 0, missing)

def validate_saldos_cols(df: pd.DataFrame):
    required = ["CONTA", "NOME DA CONTA", "SECRETARIA", "BANCO", "TIPO DE RECURSO", "SALDO BANCARIO"]
    missing = [c for c in required if c not in df.columns]
    return (len(missing) == 0, missing)

def proportional_allocation(total_recurso: float, debitos: pd.Series) -> pd.Series:
    """
    Rateio proporcional limitado ao valor devido.
    1) Aloca proporcionalmente ao total.
    2) Trunca para não ultrapassar dívida.
    3) Redistribui sobras até esgotar ou todo mundo bater no teto.
    """
    if total_recurso <= 0 or debitos.sum() == 0:
        return pd.Series(0.0, index=debitos.index)

    # Passo 1: rateio proporcional
    base = total_recurso * (debitos / debitos.sum())
    # Passo 2: teto por dívida
    pago = base.clip(upper=debitos)

    # Passo 3: redistribuição de sobras (iterativa simples)
    sobra = total_recurso - pago.sum()
    # para evitar loop infinito, limite iterações
    for _ in range(10):
        if sobra <= 0.0001:
            break
        restantes = debitos - pago
        elegiveis = restantes[restantes > 0]
        if elegiveis.empty:
            break
        add = sobra * (elegiveis / elegiveis.sum())
        novo_pago = pago.add(add, fill_value=0)
        pago = pd.concat([novo_pago, debitos], axis=1).min(axis=1)  # ainda respeita teto
        sobra = total_recurso - pago.sum()

    return pago.round(2)

# =========================================
# Uploads
# =========================================
tab_dash, tab_plano = st.tabs(["📈 Dashboard", "💸 Plano de Pagamento"])

with tab_dash:
    uploaded_debitos = st.file_uploader(
        "📁 Envie a planilha de **Débitos** (colunas: DATA, FORNECEDOR, CNPJ, VALOR, SECRETARIA)",
        type=["xlsx"],
        key="deb",
    )

    if not uploaded_debitos:
        st.info("Envie a planilha de Débitos para visualizar o dashboard.")
    else:
        df_raw = load_excel(uploaded_debitos)
        ok, miss = validate_debitos_cols(df_raw)
        if not ok:
            st.error(f"Débitos inválidos. Faltam as colunas: {', '.join(miss)}")
            st.stop()

        df = cast_types_debitos(df_raw)

        # Diagnóstico
        with st.expander("🔎 Diagnóstico: linhas ignoradas na importação (Débitos)"):
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
                st.info("Nenhuma linha ignorada por problemas de dados.")

        # Filtros
        st.sidebar.header("🔎 Filtros (Dashboard)")
        secretarias_sel = st.sidebar.multiselect("Secretaria", sorted(df["SECRETARIA"].unique().tolist()))
        fornecedores_sel = st.sidebar.multiselect("Fornecedor", sorted(df["FORNECEDOR"].unique().tolist()))
        data_min_default = pd.to_datetime(df["DATA"].min()).date()
        data_max_default = pd.to_datetime(df["DATA"].max()).date()
        col_d1, col_d2 = st.sidebar.columns(2)
        with col_d1:
            data_ini = st.date_input("Data inicial", data_min_default, key="di1")
        with col_d2:
            data_fim = st.date_input("Data final", data_max_default, key="df1")

        if data_ini > data_fim:
            st.sidebar.error("A data inicial não pode ser maior que a data final.")
            st.stop()

        df_filt = df[
            (df["DATA"] >= pd.to_datetime(data_ini)) &
            (df["DATA"] <= pd.to_datetime(data_fim))
        ].copy()

        if secretarias_sel:
            df_filt = df_filt[df_filt["SECRETARIA"].isin(secretarias_sel)]
        if fornecedores_sel:
            df_filt = df_filt[df_filt["FORNECEDOR"].isin(fornecedores_sel)]

        # KPIs
        col_k1, col_k2, col_k3 = st.columns(3)
        total_valor = df_filt["VALOR"].sum() if not df_filt.empty else 0.0
        qtd_linhas = len(df_filt)
        qtd_fornec = df_filt["FORNECEDOR"].nunique()
        col_k1.metric("Valor total filtrado", format_brl(total_valor))
        col_k2.metric("Registros", f"{qtd_linhas}")
        col_k3.metric("Fornecedores", f"{qtd_fornec}")

        st.divider()

        # Gráficos
        col_g1, col_g2 = st.columns(2)

        with col_g1:
            st.subheader("Débitos por Secretaria")
            if df_filt.empty:
                st.info("Sem dados para os filtros selecionados.")
            else:
                g1 = df_filt.groupby("SECRETARIA", as_index=False)["VALOR"].sum().sort_values("VALOR")
                fig1 = px.bar(g1, x="VALOR", y="SECRETARIA", orientation="h",
                              text=[format_brl(v) for v in g1["VALOR"]], color="SECRETARIA")
                fig1.update_traces(hovertemplate="<b>%{y}</b><br>Valor: %{x:,.2f}")
                fig1.update_layout(showlegend=False, margin=dict(l=10, r=10, t=30, b=10))
                st.plotly_chart(fig1, use_container_width=True)

        with col_g2:
            st.subheader("Top 10 Fornecedores")
            if df_filt.empty:
                st.info("Sem dados para os filtros selecionados.")
            else:
                g2 = (df_filt.groupby("FORNECEDOR", as_index=False)["VALOR"]
                      .sum().sort_values(by="VALOR", ascending=False).head(10))
                fig2 = px.bar(g2, x="FORNECEDOR", y="VALOR",
                              text=[format_brl(v) for v in g2["VALOR"]], color="FORNECEDOR")
                fig2.update_traces(hovertemplate="<b>%{x}</b><br>Valor: %{y:,.2f}")
                fig2.update_layout(showlegend=False, xaxis_tickangle=45, margin=dict(l=10, r=10, t=30, b=80))
                st.plotly_chart(fig2, use_container_width=True)

        st.divider()
        st.subheader("📋 Dados Filtrados")
        df_display = df_filt.copy()
        df_display["VALOR"] = df_display["VALOR"].apply(format_brl)
        st.dataframe(df_display, use_container_width=True)

        st.subheader("📥 Exportar (Dashboard)")
        # Excel
        excel_buffer = io.BytesIO()
        df_filt.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        st.download_button("📊 Baixar Excel (dados filtrados)", data=excel_buffer,
                           file_name="dashboard_dados_filtrados.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF
        pdf_df = df_display.rename(columns={
            "DATA":"DATA",
            "FORNECEDOR":"FORNECEDOR",
            "CNPJ":"CNPJ",
            "VALOR":"VALOR",
            "SECRETARIA":"SECRETARIA"
        })
        pdf_bytes = gerar_pdf_listagem(pdf_df, titulo="Débitos - Dados Filtrados (Dashboard)")
        st.download_button("📄 Baixar PDF (dados filtrados)", data=pdf_bytes,
                           file_name="dashboard_dados_filtrados.pdf", mime="application/pdf")

with tab_plano:
    st.subheader("💸 Plano de Pagamento com Recurso Livre")
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        up_debitos2 = st.file_uploader("📁 Envie a planilha de **Débitos**", type=["xlsx"], key="deb2")
    with col_up2:
        up_saldos = st.file_uploader("🏦 Envie a planilha de **Saldos Bancários**", type=["xlsx"], key="sal")

    if (up_debitos2 is None) or (up_saldos is None):
        st.info("Envie as duas planilhas para calcular o plano de pagamento.")
        st.stop()

    # Carregar e validar
    deb_raw = load_excel(up_debitos2)
    okd, missd = validate_debitos_cols(deb_raw)
    if not okd:
        st.error(f"Débitos inválidos. Faltam as colunas: {', '.join(missd)}")
        st.stop()
    deb = cast_types_debitos(deb_raw)

    sal_raw = load_excel(up_saldos)
    oks, misss = validate_saldos_cols(sal_raw)
    if not oks:
        st.error(f"Saldos inválidos. Faltam as colunas: {', '.join(misss)}")
        st.stop()

    # Filtro de recurso livre e (opcional) por secretaria
    st.markdown("**Configurações do rateio**")
    colf1, colf2 = st.columns(2)
    with colf1:
        considerar_so_livre = st.checkbox("Considerar apenas 'Tipo de Recurso = Livre'", value=True)
    with colf2:
        filtrar_secretaria = st.multiselect("Filtrar saldos por Secretaria (opcional)",
                                            sorted(sal_raw["SECRETARIA"].dropna().unique().tolist()))

    sal = sal_raw.copy()
    if considerar_so_livre and "TIPO DE RECURSO" in sal.columns:
        sal = sal[sal["TIPO DE RECURSO"].str.upper().eq("LIVRE")]

    if filtrar_secretaria:
        sal = sal[sal["SECRETARIA"].isin(filtrar_secretaria)]

    total_livre = pd.to_numeric(sal["SALDO BANCARIO"], errors="coerce").fillna(0).sum().round(2)

    # Total devido por fornecedor (agregado geral)
    due_by_vendor = deb.groupby(["FORNECEDOR", "CNPJ"], as_index=False)["VALOR"].sum().rename(columns={"VALOR":"DEBITO"})
    soma_debitos = due_by_vendor["DEBITO"].sum().round(2)

    st.write(f"**Recurso disponível para pagamento:** {format_brl(total_livre)}")
    st.write(f"**Total de débitos (todos fornecedores):** {format_brl(soma_debitos)}")

    if total_livre <= 0:
        st.warning("Não há saldo livre para ratear.")
        st.stop()

    # Rateio proporcional
    pagos = proportional_allocation(total_livre, due_by_vendor.set_index(["FORNECEDOR","CNPJ"])["DEBITO"])
    pagos = pagos.reset_index().rename(columns={0:"PAGAR_AGORA"})
    plano = due_by_vendor.merge(pagos, on=["FORNECEDOR","CNPJ"], how="left")
    plano["RESTANTE"] = (plano["DEBITO"] - plano["PAGAR_AGORA"]).round(2)

    # Exibição formatada
    plano_display = plano.copy()
    plano_display["DEBITO"] = plano_display["DEBITO"].apply(format_brl)
    plano_display["PAGAR_AGORA"] = plano_display["PAGAR_AGORA"].apply(format_brl)
    plano_display["RESTANTE"] = plano_display["RESTANTE"].apply(format_brl)

    st.subheader("📋 Plano de Pagamento (Rateio Proporcional)")
    st.dataframe(plano_display, use_container_width=True)

    # KPIs do plano
    colp1, colp2, colp3 = st.columns(3)
    colp1.metric("Total a pagar agora", format_brl(plano["PAGAR_AGORA"].sum()))
    colp2.metric("Fornecedores contemplados", f"{(plano['PAGAR_AGORA']>0).sum()}")
    colp3.metric("Débito que permanece", format_brl(plano['RESTANTE'].clip(lower=0).sum()))

    # Exportações
    st.subheader("📥 Exportar (Plano de Pagamento)")
    # Excel (numérico)
    excel_plano = io.BytesIO()
    plano.to_excel(excel_plano, index=False)
    excel_plano.seek(0)
    st.download_button("📊 Baixar Excel do Plano", data=excel_plano,
                       file_name="plano_pagamento_rateio.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # PDF (formatado)
    pdf_df = plano.copy()
    pdf_df["DEBITO"] = pdf_df["DEBITO"].apply(format_brl)
    pdf_df["PAGAR_AGORA"] = pdf_df["PAGAR_AGORA"].apply(format_brl)
    pdf_df["RESTANTE"] = pdf_df["RESTANTE"].apply(format_brl)
    pdf_bytes_plano = gerar_pdf_listagem(pdf_df[["FORNECEDOR","CNPJ","DEBITO","PAGAR_AGORA","RESTANTE"]],
                                         titulo="Plano de Pagamento - Rateio Proporcional (Recurso Livre)")
    st.download_button("📄 Baixar PDF do Plano", data=pdf_bytes_plano,
                       file_name="plano_pagamento_rateio.pdf", mime="application/pdf")
