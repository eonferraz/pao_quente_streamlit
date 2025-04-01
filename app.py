import streamlit as st
import pandas as pd
import pyodbc
import plotly.express as px
from PIL import Image

# ====================
# CONFIG INICIAL
# ====================
st.set_page_config(page_title="Pão Quente", layout="wide")

# Logo + Título
logo = Image.open("logo.png")
col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.image(logo, width=90)
with col_title:
    st.markdown("""
        <h1 style='color: #862E3A; margin-bottom: 0;'>Dashboard de Vendas</h1>
        <h4 style='color: #37392E; margin-top: 0;'>Padaria Pão Quente</h4>
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
    df = pd.read_sql("SELECT * FROM PQ_VENDAS", conn)
    conn.close()
    return df

# ====================
# CARGA E PREPARO
# ====================
with st.spinner("Carregando dados..."):
    df = carregar_dados()

df.columns = df.columns.str.strip().str.upper()
df["DATA"] = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["DATA"])
df["ANO_MES"] = df["DATA"].dt.to_period("M").astype(str)
df["DIA"] = df["DATA"].dt.day

# ====================
# FILTROS
# ====================
st.sidebar.header("🔍 Filtros")
todas_uns = sorted(df["UN"].dropna().unique())
un_selecionadas = st.sidebar.multiselect("Unidades:", todas_uns, default=todas_uns)
todos_meses = sorted(df["ANO_MES"].unique())
meses_selecionados = st.sidebar.multiselect("Ano/Mês:", todos_meses, default=todos_meses)

df_filt = df[(df["UN"].isin(un_selecionadas)) & (df["ANO_MES"].isin(meses_selecionados))]

# ====================
# 1ª FAIXA (CARDS + 3 GRÁFICOS)
# ====================
col1, col2, col3 = st.columns([1, 2, 2])

# Cards
with col1:
    fat_total = df_filt["TOTAL"].sum()
    qtd_vendas = df_filt["COD_VENDA"].nunique()
    ticket = fat_total / qtd_vendas if qtd_vendas > 0 else 0

    st.metric("💰 Faturamento Total", f"R$ {fat_total:,.2f}".replace(",", "."))
    st.metric("📊 Qtde de Vendas", qtd_vendas)
    st.metric("💳 Ticket Médio", f"R$ {ticket:,.2f}".replace(",", "."))

# Gráfico 1: Faturamento por Ano-Mês
with col2:
    df_mes = df_filt.groupby("ANO_MES")["TOTAL"].sum().reset_index()
    fig1 = px.bar(df_mes, x="ANO_MES", y="TOTAL", title="Faturamento por Mês",
                  color_discrete_sequence=["#FE9C37"])
    st.plotly_chart(fig1, use_container_width=True)

# Gráfico 2: Barras Horizontais + Rosca
with col3:
    df_un = df_filt.groupby("UN")["TOTAL"].sum().reset_index().sort_values("TOTAL")

    fig2 = px.bar(df_un, x="TOTAL", y="UN", orientation='h',
                  title="Faturamento por UN", color_discrete_sequence=["#A4B494"])
    st.plotly_chart(fig2, use_container_width=True)

    fig3 = px.pie(df_un, names="UN", values="TOTAL", hole=0.5,
                  title="Distribuição % por UN",
                  color_discrete_sequence=px.colors.sequential.RdBu)
    fig3.update_traces(textposition="inside", textinfo="percent+label")
    st.plotly_chart(fig3, use_container_width=True)

# ====================
# 2ª FAIXA (ACUMULADO E TICKET)
# ====================
st.markdown("---")

col4, col5 = st.columns(2)

with col4:
    df_dia = df_filt.groupby("DATA")["TOTAL"].sum().cumsum().reset_index(name="FAT_ACUM")
    fig4 = px.line(df_dia, x="DATA", y="FAT_ACUM", title="Faturamento Acumulado no Mês",
                   markers=True, color_discrete_sequence=["#862E3A"])
    st.plotly_chart(fig4, use_container_width=True)

with col5:
    df_ticket = df_filt.groupby("DATA").agg({"TOTAL": "sum", "COD_VENDA": "nunique"}).reset_index()
    df_ticket["TICKET"] = df_ticket["TOTAL"] / df_ticket["COD_VENDA"]
    fig5 = px.line(df_ticket, x="DATA", y="TICKET", title="Evolução do Ticket Médio",
                   markers=True, color_discrete_sequence=["#37392E"])
    st.plotly_chart(fig5, use_container_width=True)

# ====================
# EXPANSÍVEL - TABELA DETALHADA
# ====================
with st.expander("📊 Ver dados detalhados"):
    st.dataframe(df_filt, use_container_width=True)
