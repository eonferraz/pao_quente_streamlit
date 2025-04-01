import streamlit as st
import pandas as pd
import pyodbc

st.set_page_config(page_title="Dashboard de Vendas - SX Comercial", layout="wide")

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

st.title("📊 Dashboard de Faturamento - SX Comercial")

with st.spinner("🔄 Carregando dados..."):
    df = carregar_dados()

# Padronizar nomes de colunas
df.columns = df.columns.str.strip().str.upper()

# Filtro por Unidade (Loja)
unidades = df["LOJA"].dropna().unique()
unidades.sort()
un_selecionada = st.selectbox("🔎 Selecione a UN (Loja):", unidades)

df_filtrado = df[df["LOJA"] == un_selecionada]

# Faturamento por mês
faturamento_mes = df_filtrado.groupby("ANO_MES")["TOTAL"].sum().reset_index()
faturamento_mes = faturamento_mes.sort_values(by="ANO_MES")

# Gráfico
st.subheader(f"💰 Faturamento Mensal - UN: {un_selecionada}")
st.bar_chart(faturamento_mes.set_index("ANO_MES"))
