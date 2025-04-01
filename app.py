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
        background-color: #FFFFFF;
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
        color: #37392E;
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
st.sidebar.header("🔎 Filtros")
todas_uns = sorted(df["UN"].dropna().unique())
un_selecionadas = st.sidebar.multiselect("Unidades:", todas_uns, default=todas_uns)
todos_meses = sorted(df["ANO_MES"].unique())
meses_selecionados = st.sidebar.multiselect("Ano/Mês:", todos_meses, default=todos_meses)

df_filt = df[(df["UN"].isin(un_selecionadas)) & (df["ANO_MES"].isin(meses_selecionados))]

# ====================
# DASHBOARD
# ====================
col1, col2, col3 = st.columns([1, 2, 2])

with col1:
    fat_total = df_filt["TOTAL"].sum()
    qtd_vendas = df_filt["COD_VENDA"].nunique()
    ticket = fat_total / qtd_vendas if qtd_vendas > 0 else 0

    card_style = "height: 35px; display: flex; flex-direction: column; justify-content: center;"

    with st.container(border=True):
        st.markdown(f"<div style='{card_style}'>", unsafe_allow_html=True)
        st.metric("Faturamento", f"R$ {fat_total:,.2f}".replace(",", "."))
        st.markdown("</div>", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"<div style='{card_style}'>", unsafe_allow_html=True)
        st.metric("Qtde de Vendas", qtd_vendas)
        st.markdown("</div>", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"<div style='{card_style}'>", unsafe_allow_html=True)
        st.metric("Ticket Médio", f"R$ {ticket:,.2f}".replace(",", "."))
        st.markdown("</div>", unsafe_allow_html=True)

with col2:
    with st.container(border=True):
        df_mes = df_filt.groupby("ANO_MES")["TOTAL"].sum().reset_index()
        fig1 = px.bar(df_mes, x="ANO_MES", y="TOTAL", title="Faturamento por Mês",
                      text_auto=True, color_discrete_sequence=["#FE9C37"])
        fig1.update_layout(
            yaxis=dict(showticklabels=False, showgrid=False, zeroline=False, visible=False),
            xaxis=dict(type='category'),
            xaxis_tickangle=-45,
            yaxis_tickprefix="R$ ",
            yaxis_tickformat=",.2f"
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

col4, col5 = st.columns(2)

with col4:
    with st.container(border=True):
        df_dia = df_filt.groupby("DATA")["TOTAL"].sum().cumsum().reset_index(name="FAT_ACUM")
        fig4 = px.line(df_dia, x="DATA", y="FAT_ACUM", title="Faturamento Acumulado no Mês",
                       markers=True, color_discrete_sequence=["#862E3A"])
        fig4.update_layout(yaxis_tickprefix="R$ ", yaxis_tickformat=",.2f")
        st.plotly_chart(fig4, use_container_width=True)

with col5:
    with st.container(border=True):
        df_ticket = df_filt.groupby("DATA").agg({"TOTAL": "sum", "COD_VENDA": "nunique"}).reset_index()
        df_ticket["TICKET"] = df_ticket["TOTAL"] / df_ticket["COD_VENDA"]
        fig5 = px.line(df_ticket, x="DATA", y="TICKET", title="Evolução do Ticket Médio",
                       markers=True, color_discrete_sequence=["#37392E"])
        fig5.update_layout(yaxis_tickprefix="R$ ", yaxis_tickformat=",.2f")
        st.plotly_chart(fig5, use_container_width=True)

with st.expander("📊 Ver dados detalhados"):
    st.dataframe(df_filt, use_container_width=True)
