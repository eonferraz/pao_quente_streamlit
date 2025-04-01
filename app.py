import streamlit as st
import pandas as pd
import pyodbc

st.set_page_config(page_title="Dashboard de Vendas", layout="wide")

# Conexão com SQL Server
@st.cache_data(ttl=600)
def carregar_dados():
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        f'SERVER=sx-global.database.windows.net;'
        f'DATABASE=sx_comercial;'
        f'UID=paulo.ferraz;'
        f'PWD=Gs!^42j$G0f0^EI#ZjRv'
    )
    df = pd.read_sql("SELECT * FROM PQ_VENDAS", conn)
    conn.close()
    return df

st.title("📊 Dashboard de Vendas - SX Comercial")

with st.spinner("Carregando dados..."):
    df = carregar_dados()

# Filtros
meses = df['ANO_MES'].dropna().unique()
meses.sort()
mes_selecionado = st.selectbox("Filtrar por Ano/Mês:", meses)

df_filtrado = df[df["ANO_MES"] == mes_selecionado]

# Exibir dados
st.subheader(f"📄 Dados de Vendas - {mes_selecionado}")
st.dataframe(df_filtrado, use_container_width=True)

# Faturamento por loja
st.subheader("💰 Faturamento por Loja")
faturamento = df_filtrado.groupby("Loja")["TOTAL"].sum().reset_index()
faturamento = faturamento.sort_values(by="TOTAL", ascending=False)

st.bar_chart(faturamento.set_index("Loja"))

