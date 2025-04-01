import streamlit as st
import pandas as pd
import pyodbc
import plotly.express as px
import base64

# ====================
# CONFIG INICIAL
# ====================
st.set_page_config(page_title="Pão Quente", layout="wide")

# ====================
# CSS FIXO PARA TOPO
# ====================
st.markdown("""
    <style>
    .header-fixed {
        position: sticky;
        top: 0;
        background-color: #2e2e2e;
        z-index: 999;
        padding: 10px 2rem 5px 2rem;
        border-bottom: 2px solid #862E3A;
        display: flex;
        align-items: center;
    }
    .header-fixed img {
        width: 80px;
        margin-right: 20px;
    }
    .header-title h1 {
        color: #862E3A;
        margin: 0;
        font-size: 32px;
    }
    .header-title h4 {
        color: #A4B494;
        margin: 0;
        font-size: 18px;
    }
    </style>
""", unsafe_allow_html=True)

# ====================
# TOPO FIXO COM LOGO E TÍTULO
# ====================
def get_base64_image(img_path):
    with open(img_path, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode()

logo_base64 = get_base64_image("logo.png")

st.markdown(f"""
<div class="header-fixed">
  <img src="data:image/png;base64,{logo_base64}" alt="Logo">
  <div class="header-title">
    <h1>Dashboard de Vendas</h1>
    <h4>Padaria Pão Quente</h4>
  </div>
</div>
""", unsafe_allow_html=True)

# ====================
# CONEXÃO COM BANCO
# ====================
@st.cache_data(ttl=600)
def carregar_dados():
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=sx-global.database.windows.net;'
        'DATABASE=sx_comercial;'
        'UID=paulo.ferraz;'
        'PWD=Gs!^42j$G0f0^EI#ZjRv'
    )
    df_vendas = pd.read_sql("SELECT * FROM PQ_VENDAS", conn)
    df_metas = pd.read_sql("SELECT * FROM PQ_METAS", conn)
    conn.close()
    return df_vendas, df_metas

# ====================
# CARGA E PREPARO
# ====================
with st.spinner("🔄 Carregando dados..."):
    df, metas = carregar_dados()

df.columns = df.columns.str.strip().str.upper()
df["DATA"] = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["DATA"])
df["ANO_MES"] = df["DATA"].dt.to_period("M").astype(str)
df["DIA"] = df["DATA"].dt.day

metas.columns = metas.columns.str.strip().str.upper()

# ====================
# FILTROS
# ====================
st.sidebar.header("🔎 Filtros")
todas_uns = sorted(df["UN"].dropna().unique())
un_selecionadas = st.sidebar.multiselect("Unidades:", todas_uns, default=todas_uns)
todos_meses = sorted(df["ANO_MES"].unique())
meses_selecionados = st.sidebar.multiselect("Ano/Mês:", todos_meses, default=todos_meses)

df_filt = df[(df["UN"].isin(un_selecionadas)) & (df["ANO_MES"].isin(meses_selecionados))]
metas_filt = metas[(metas["LOJA"].isin(un_selecionadas)) & (metas["ANO-MES"].isin(meses_selecionados))]



# ====================
# EDIÇÃO DE METAS
# ====================
st.markdown("---")
with st.container(border=True):
    if st.button("✏️ Editar Metas de Faturamento"):
        st.session_state['editar_metas'] = True

if st.session_state.get('editar_metas', False):
    st.markdown("### 🗓️ Edição de Metas por Unidade - Ano 2025")

    meses_2025 = pd.date_range(start="2025-01-01", end="2025-12-01", freq='MS').strftime('%Y-%m')
    unidades = ["P1", "P3", "P4", "P5"]

    # Montar base com dados existentes
    df_input = pd.DataFrame(index=meses_2025, columns=unidades)
    df_input.index.name = "ANO-MES"

    metas_2025 = metas[(metas["ANO-MES"].isin(meses_2025)) & (metas["LOJA"].isin(unidades))]
    for _, row in metas_2025.iterrows():
        df_input.at[row["ANO-MES"], row["LOJA"]] = row["VALOR_META"]

    edited_df = st.data_editor(df_input, num_rows="dynamic", use_container_width=True, key="meta_editor")

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        if st.button("💾 Salvar Metas"):
            try:
                conn = pyodbc.connect(
                    'DRIVER={ODBC Driver 17 for SQL Server};'
                    'SERVER=sx-global.database.windows.net;'
                    'DATABASE=sx_comercial;'
                    'UID=paulo.ferraz;'
                    'PWD=Gs!^42j$G0f0^EI#ZjRv'
                )
                cursor = conn.cursor()
                for ano_mes in edited_df.index:
                    for loja in unidades:
                        valor = edited_df.loc[ano_mes, loja]
                        if valor != "" and not pd.isna(valor):
                            cursor.execute("""
                                IF EXISTS (SELECT 1 FROM PQ_METAS WHERE [ANO-MES] = ? AND LOJA = ?)
                                    UPDATE PQ_METAS SET VALOR_META = ? WHERE [ANO-MES] = ? AND LOJA = ?
                                ELSE
                                    INSERT INTO PQ_METAS ([ANO-MES], LOJA, VALOR_META) VALUES (?, ?, ?)
                            """, ano_mes, loja, valor, ano_mes, loja, ano_mes, loja, valor)
                conn.commit()
                conn.close()
                st.success("✅ Metas salvas com sucesso!")
                st.session_state['editar_metas'] = False
            except Exception as e:
                st.error(f"Erro ao salvar as metas: {e}")
    with col_btn2:
        if st.button("❌ Cancelar"):
            st.session_state['editar_metas'] = False

# ====================
# DASHBOARD PRINCIPAL
# ====================

col1, col2, col3 = st.columns([0.7, 2.3, 2])

with col1:
    fat_total = df_filt["TOTAL"].sum()
    qtd_vendas = df_filt["COD_VENDA"].nunique()
    ticket = fat_total / qtd_vendas if qtd_vendas > 0 else 0
    meta_total = metas_filt["VALOR_META"].sum()

    with st.container(border=True):
        st.metric("💰 Faturamento Total", f"R$ {fat_total:,.2f}".replace(",", "."))

    with st.container(border=True):
        st.metric("🎯 Meta de Faturamento", f"R$ {meta_total:,.2f}".replace(",", "."))

    with st.container(border=True):
        st.metric("📊 Qtde de Vendas", qtd_vendas)

    with st.container(border=True):
        st.metric("💳 Ticket Médio", f"R$ {ticket:,.2f}".replace(",", "."))

with col2:
    with st.container(border=True):
        df_mes = df_filt.groupby("ANO_MES")["TOTAL"].sum().reset_index()
        df_meta_mes = metas_filt.groupby("ANO-MES")["VALOR_META"].sum().reset_index()
        df_meta_mes.rename(columns={"ANO-MES": "ANO_MES"}, inplace=True)
        df_merged = pd.merge(df_mes, df_meta_mes, on="ANO_MES", how="outer").fillna(0)
        df_merged["PCT"] = df_merged["TOTAL"] / df_merged["VALOR_META"]

        fig1 = px.bar(df_merged, x="ANO_MES", y=["VALOR_META", "TOTAL"],
                      title="Faturamento vs Meta por Mês", barmode="group",
                      text_auto=True, color_discrete_sequence=["#A4B494", "#FE9C37"])
        fig1.add_scatter(x=df_merged["ANO_MES"], y=df_merged["PCT"],
                         mode="lines+markers+text", name="% Realizado",
                         text=df_merged["PCT"].map(lambda x: f"{x:.0%}"),
                         textposition="top center",
                         line=dict(color="#862E3A", dash="dot"), yaxis="y2")

        fig1.update_layout(
            yaxis=dict(title="R$", tickprefix="R$ ", tickformat=",.0f"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
            yaxis2=dict(overlaying="y", side="right", tickformat=".0%", title="%", range=[0, 1.5]),
            xaxis=dict(type='category', tickangle=-45)
        )
        st.plotly_chart(fig1, use_container_width=True)

with col3:
    with st.container(border=True):
        df_un = df_filt.groupby("UN")["TOTAL"].sum().reset_index().sort_values("TOTAL")

        col3a, col3b = st.columns(2)

        with col3a:
            fig2 = px.bar(df_un, x="TOTAL", y="UN", orientation='h',
                          text_auto=True, title="Faturamento por UN", color_discrete_sequence=["#A4B494"])
            fig2.update_layout(xaxis_tickprefix="R$ ", xaxis_tickformat=",.2f")
            st.plotly_chart(fig2, use_container_width=True)

        with col3b:
            fig3 = px.pie(df_un, names="UN", values="TOTAL", hole=0.5,
                          title="Distribuição % por UN",
                          color_discrete_sequence=px.colors.sequential.RdBu)
            fig3.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

col4, col5 = st.columns(2)  # ← Essa linha deve vir ANTES de usar col4 e col5

with col4:
    with st.container(border=True):
        # Faturamento acumulado por dia
        df_dia = df_filt.groupby("DATA")["TOTAL"].sum().reset_index()
        df_dia["FAT_ACUM"] = df_dia["TOTAL"].cumsum()

        if not metas_filt.empty:
            meta_mes_total = metas_filt["VALOR_META"].sum()
            dias_do_mes = df_dia["DATA"].dt.days_in_month.iloc[0]  # assume mesmo mês
            meta_dia = meta_mes_total / dias_do_mes

            df_dia["META_DIA"] = meta_dia
            df_dia["META_ACUM"] = df_dia["META_DIA"].cumsum()
            df_dia["PCT"] = df_dia["FAT_ACUM"] / df_dia["META_ACUM"]

        # Gráfico principal
        fig4 = px.line(df_dia, x="DATA", y="FAT_ACUM", title="Faturamento Acumulado no Mês",
                       markers=True, color_discrete_sequence=["#862E3A"])

        # Meta acumulada
        if "META_ACUM" in df_dia.columns:
            fig4.add_scatter(x=df_dia["DATA"], y=df_dia["META_ACUM"],
                             mode="lines+markers", name="Meta Acumulada",
                             line=dict(color="#A4B494", dash="dot"))

            # Cor dinâmica da linha % realizado
            cor_pct = "#A4B494" if df_dia["PCT"].iloc[-1] >= 1 else "#862E3A"

            fig4.add_scatter(x=df_dia["DATA"], y=df_dia["PCT"],
                             mode="lines+markers", name="% Realizado",
                             line=dict(color=cor_pct, dash="dot"),
                             yaxis="y2")

        # Layout final
        fig4.update_layout(
            yaxis=dict(title="R$", tickprefix="R$ ", tickformat=",.0f"),
            yaxis2=dict(overlaying="y", side="right", tickformat=".0%", title="%", range=[0, 1.5]),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig4, use_container_width=True)



with col5:
    with st.container(border=True):
        df_ticket = df_filt.groupby("DATA").agg({"TOTAL": "sum", "COD_VENDA": "nunique"}).reset_index()
        df_ticket["TICKET"] = df_ticket["TOTAL"] / df_ticket["COD_VENDA"]
        df_ticket["MM_TICKET"] = df_ticket["TICKET"].rolling(window=7).mean()

        fig5 = px.line(df_ticket, x="DATA", y="TICKET", title="Evolução do Ticket Médio",
                       markers=True, color_discrete_sequence=["#37392E"])
        fig5.add_scatter(x=df_ticket["DATA"], y=df_ticket["MM_TICKET"], mode="lines", name="Média Móvel (7 dias)",
                         line=dict(color="#FE9C37", dash="dot"))
        fig5.update_layout(yaxis_tickprefix="R$ ", yaxis_tickformat=",.2f")
        st.plotly_chart(fig5, use_container_width=True)


with st.expander("📊 Ver dados detalhados"):
    st.dataframe(df_filt, use_container_width=True)
