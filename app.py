import streamlit as st
import pandas as pd
import pyodbc
import plotly.express as px

st.set_page_config(page_title="Dashboard SX Comercial", layout="wide")

# Função de conexão
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

# Carregar dados
with st.spinner("🔄 Carregando dados..."):
    df = carregar_dados()

# Padronizar colunas
df.columns = df.columns.str.strip().str.upper()

# Ajuste da data
df["DATA"] = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
df["ANO_MES"] = df["DATA"].dt.to_period("M").astype(str)

# Filtro de UNs com opção "Selecionar todas"
todas_uns = sorted(df["UN"].dropna().unique())
un_selecionadas = st.multiselect("🏢 Selecionar UN(s):", options=todas_uns, default=todas_uns)

df_filtrado = df[df["UN"].isin(un_selecionadas)]

# Cálculos
faturamento = df_filtrado.groupby("UN")["TOTAL"].sum().reset_index(name="FATURAMENTO")
clientes = df_filtrado.groupby("UN")["COD_VENDA"].nunique().reset_index(name="CLIENTES")
ticket = faturamento.merge(clientes, on="UN")
ticket["TICKET_MEDIO"] = ticket["FATURAMENTO"] / ticket["CLIENTES"]

# Layout em colunas
col1, col2, col3 = st.columns(3)

# Gráfico 1 - Faturamento
with col1:
    st.subheader("💰 Faturamento por UN")
    fig1 = px.bar(
        faturamento,
        x="UN",
        y="FATURAMENTO",
        text_auto=".2s",
        color_discrete_sequence=["#3A2E86"]
    )
    st.plotly_chart(fig1, use_container_width=True)

# Gráfico 2 - Clientes
with col2:
    st.subheader("👥 Clientes por UN")
    fig2 = px.bar(
        clientes,
        x="UN",
        y="CLIENTES",
        text_auto=True,
        color_discrete_sequence=["#379CFE"]
    )
    st.plotly_chart(fig2, use_container_width=True)

# Gráfico 3 - Ticket Médio
with col3:
    st.subheader("💳 Ticket Médio por UN")
    fig3 = px.bar(
        ticket,
        x="UN",
        y="TICKET_MEDIO",
        text_auto=".2f",
        color_discrete_sequence=["#94B4A4"]
    )
    st.plotly_chart(fig3, use_container_width=True)

# Mostrar tabela detalhada abaixo, se quiser
with st.expander("🔍 Ver dados detalhados"):
    st.dataframe(df_filtrado, use_container_width=True)
