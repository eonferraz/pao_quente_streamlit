import streamlit as st
import pandas as pd
import pyodbc
import plotly.express as px
from PIL import Image

# Logo
logo = Image.open("logo.png")  # Troque pelo nome do seu arquivo

# Layout superior com título e logo
col1, col2 = st.columns([4, 1])
with col1:
    st.markdown("<h1 style='color: #3A2E86;'>Dashboard de Vendas - SX Comercial</h1>", unsafe_allow_html=True)
with col2:
    st.image(logo, width=120)


st.set_page_config(page_title="Pão Quente", layout="wide")

# Conexão com SQL Server
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

# Converter datas e criar ANO_MES
df["DATA"] = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
df["ANO_MES"] = df["DATA"].dt.to_period("M").astype(str)

# Sidebar com filtros
st.sidebar.header("🔎 Filtros")

# Filtro de UNs
todas_uns = sorted(df["UN"].dropna().unique())
un_selecionadas = st.sidebar.multiselect("Selecionar UN(s):", options=todas_uns, default=todas_uns)

# Filtro de ANO_MES
todos_meses = sorted(df["ANO_MES"].dropna().unique())
meses_selecionados = st.sidebar.multiselect("Selecionar Ano/Mês:", options=todos_meses, default=todos_meses)

# Aplicar filtros
df_filtrado = df[(df["UN"].isin(un_selecionadas)) & (df["ANO_MES"].isin(meses_selecionados))]

# =======================
# 🔷 GRÁFICO EMPILHADO
# =======================
st.subheader("📊 Faturamento Mensal por UN (Empilhado)")
df_empilhado = df_filtrado.groupby(["ANO_MES", "UN"])["TOTAL"].sum().reset_index()

fig_empilhado = px.bar(
    df_empilhado,
    x="ANO_MES",
    y="TOTAL",
    color="UN",
    text_auto=".2s",
    title="Faturamento Empilhado por Mês e UN",
    color_discrete_sequence=px.colors.qualitative.Dark24
)
st.plotly_chart(fig_empilhado, use_container_width=True)

# =======================
# 🔶 GRÁFICO DE ROSCA
# =======================
st.subheader("🍩 Participação no Faturamento por UN")
faturamento_total = df_filtrado.groupby("UN")["TOTAL"].sum().reset_index()

fig_rosca = px.pie(
    faturamento_total,
    names="UN",
    values="TOTAL",
    hole=0.5,
    title="Distribuição de Faturamento",
    color_discrete_sequence=px.colors.sequential.Teal
)
fig_rosca.update_traces(textposition="inside", textinfo="percent+label")
st.plotly_chart(fig_rosca, use_container_width=True)

# =======================
# 🔍 TABELA DETALHADA
# =======================
with st.expander("📋 Ver dados detalhados"):
    st.dataframe(df_filtrado, use_container_width=True)
