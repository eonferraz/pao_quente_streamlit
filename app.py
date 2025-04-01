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

# (restante do código permanece inalterado)
