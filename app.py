import streamlit as st
import pandas as pd
import pyodbc
import plotly.express as px
import base64
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from datetime import datetime
import numpy as np
from plotly import graph_objects as go
import networkx as nx
import scipy

# ====================
# CONFIG INICIAL
# ====================
st.set_page_config(page_title="Pão Quente", layout="wide")


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

# CSS atualizado para cabeçalho com alinhamento correto
st.markdown("""
    <style>
        .fixed-header {
            position: sticky;
            top: 0;
            background-color: white;
            z-index: 999;
            padding: 10px 20px 5px 20px;
            border-bottom: 1px solid #ccc;
            box-shadow: 0px 2px 6px rgba(0,0,0,0.05);
        }

        .header-flex {
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .logo {
            height: 50px;
        }

        .title {
            font-size: 26px;
            font-weight: bold;
            color: #862E3A;
            margin: 0 auto;
        }

        .filters {
            display: flex;
            gap: 15px;
        }

        .block-container {
            padding-top: 0rem;
        }
    </style>
""", unsafe_allow_html=True)

# ====================
# HEADER ÚNICO COM LOGO, TÍTULO E FILTROS
# ====================
with st.container():
    st.markdown("<div class='fixed-header'>", unsafe_allow_html=True)
    st.markdown("<div class='header-flex'>", unsafe_allow_html=True)

    # Logo
    st.image("logo.png", width=90)

    # Título
    st.markdown("""
        <div style="display: flex; align-items: center; justify-content: center;">
            <img src="logo.png" style="height: 50px; margin-right: 12px;">
            <span style="font-size: 26px; font-weight: bold; color: #862E3A;">
                Dashboard Comercial - Pão Quente
            </span>
        </div>
    """, unsafe_allow_html=True)

    # Filtros (divididos em 2 colunas lado a lado)
    col1, col2 = st.columns(2)

    with col1:
        todas_uns = sorted(df["UN"].dropna().unique())
        un_selecionadas = st.multiselect("Unidades:", todas_uns, default=todas_uns, key="filtros_un")

    with col2:
        todos_meses = sorted(df["ANO_MES"].unique())
        meses_selecionados = st.multiselect("Ano/Mês:", todos_meses, default=todos_meses, key="filtros_mes")

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ====================
# APLICAÇÃO DOS FILTROS
# ====================
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

col1, col2 = st.columns([1, 4.2])


# CARDS
#=====================================================================================================================================================================
with col1:
    fat_total = df_filt["TOTAL"].sum()
    qtd_vendas = df_filt["COD_VENDA"].nunique()
    ticket = fat_total / qtd_vendas if qtd_vendas > 0 else 0
    meta_total = metas_filt["VALOR_META"].sum()
    progresso = (fat_total / meta_total) * 100 if meta_total > 0 else 0

    def metric_card(titulo, valor):
        st.markdown(
            f"""
            <div style="border: 1px solid #DDD; border-radius: 10px; padding: 10px; margin-bottom: 10px; text-align: center;">
                <div style="font-size: 13px; color: gray;">{titulo}</div>
                <div style="font-size: 20px; font-weight: bold;">{valor}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # Formatação dos números com ponto para milhar e vírgula para decimal
    def format_brl(value):
        return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def format_percent(value):
        return f"{value:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")

    # Cards
    metric_card("💰 Faturamento Total", format_brl(fat_total))
    metric_card("🎯 Meta de Faturamento", format_brl(meta_total))
    metric_card("📈 Progresso da Meta", format_percent(progresso))
    metric_card("📊 Qtde de Vendas", f"{qtd_vendas:,}".replace(",", "."))
    metric_card("💳 Ticket Médio", format_brl(ticket))
#=====================================================================================================================================================================

with col2:
    with st.container(border=True):
        df_mes = df_filt.groupby("ANO_MES")["TOTAL"].sum().reset_index()
        df_meta_mes = metas_filt.groupby("ANO-MES")["VALOR_META"].sum().reset_index()
        df_meta_mes.rename(columns={"ANO-MES": "ANO_MES"}, inplace=True)
        df_merged = pd.merge(df_mes, df_meta_mes, on="ANO_MES", how="outer").fillna(0)
        df_merged["PCT"] = df_merged["TOTAL"] / df_merged["VALOR_META"]

        fig1 = px.bar(
            df_merged, x="ANO_MES", y=["VALOR_META", "TOTAL"],
            title="Faturamento vs Meta por Mês", barmode="group",
            color_discrete_sequence=["#A4B494", "#FE9C37"]
        )

        fig1.update_traces(
            texttemplate="R$ %{y:,.0f}",
            textposition="inside",
            textangle=-90,
            textfont_size=16,
            insidetextanchor="start"
        )    

        # Linha % realizado (sem texto, faremos manualmente)
        fig1.add_scatter(
            x=df_merged["ANO_MES"],
            y=df_merged["PCT"],
            mode="lines+markers",
            name="% Realizado",
            line=dict(color="#862E3A", dash="dot"),
            yaxis="y2",
            marker=dict(size=8)
        )

        # Adicionando % como annotations (fundo vermelho, texto branco)
        for i, row in df_merged.iterrows():
            fig1.add_annotation(
                x=row["ANO_MES"],
                y=row["PCT"],
                text=f"{row['PCT']:.0%}",
                showarrow=False,
                font=dict(color="white", size=12),
                align="center",
                bgcolor="#862E3A",
                borderpad=4,
                yanchor="bottom",  # ancora na base do texto
                yshift=12           # empurra o texto para cima da linha
            )

        # Layout geral
        fig1.update_layout(
            yaxis=dict(
                title="R$",
                tickprefix="R$ ",
                tickformat=",.0f",
                showticklabels=False,
                showline=False,
                zeroline=False,
                showgrid=False
            ),
            yaxis2=dict(
                overlaying="y",
                side="right",
                tickformat=".0%",
                title="%",
                range=[0, 1.5],
                showticklabels=False,
                showline=False,
                zeroline=False,
                showgrid=False
            ),
            xaxis=dict(
                type='category',
                tickangle=-45,
                showgrid=False
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            ),
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        st.plotly_chart(fig1, use_container_width=True)

st.markdown("---")

col3a, col3b = st.columns(2)  # Desempacota as colunas corretamente

with col3a:
    with st.container(border=True):
        df_un = df_filt.groupby("UN")["TOTAL"].sum().reset_index().sort_values("TOTAL")

        fig2 = px.bar(
            df_un, x="TOTAL", y="UN", orientation='h',
            text_auto=True, title="Faturamento por UN",
            color_discrete_sequence=["#A4B494"]
        )
        fig2.update_layout(xaxis_tickprefix="R$ ", xaxis_tickformat=",.2f")
        st.plotly_chart(fig2, use_container_width=True)

with col3b:
    with st.container(border=True):
        fig3 = px.pie(
            df_un, names="UN", values="TOTAL", hole=0.5,
            title="Distribuição % por UN",
            color_discrete_sequence=px.colors.sequential.RdBu
        )
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

        fig5.add_scatter(x=df_ticket["DATA"], y=df_ticket["MM_TICKET"],
                         mode="lines", name="Média Móvel (7 dias)",
                         line=dict(color="#FE9C37", dash="dot"))

        fig5.update_layout(
            yaxis_tickprefix="R$ ",
            yaxis_tickformat=",.2f",
            yaxis_range=[0, df_ticket["TICKET"].max() * 1.1],
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
        )

        st.plotly_chart(fig5, use_container_width=True)

st.markdown("---")




# ANÁLISE DE PRODUTOS
with st.container(border=True):
    st.markdown("<h4 style='color:#862E3A;'>🏆 Top 10 Produtos e Produtos Associados</h4>", unsafe_allow_html=True)

    col1, col2 = st.columns([1.2, 1.8])

    # ================= COLUNA 1 - BARRAS =================
    with col1:
        df_top = df_filt.groupby("DESCRICAO_PRODUTO")["TOTAL"].sum().reset_index()
        df_top = df_top.sort_values("TOTAL", ascending=False).head(10)
        top_produtos = df_top["DESCRICAO_PRODUTO"].tolist()

        produto_selecionado = st.selectbox("🧠 Selecione um produto:", top_produtos)

        fig_top10 = px.bar(df_top.sort_values("TOTAL"),
                           x="TOTAL", y="DESCRICAO_PRODUTO",
                           orientation='h',
                           text_auto=True,
                           title="Top 10 Produtos",
                           color="TOTAL", color_continuous_scale="OrRd")

        fig_top10.update_layout(yaxis=dict(categoryorder="total ascending"),
                                xaxis_tickprefix="R$ ", xaxis_tickformat=",.2f",
                                margin=dict(t=40, l=10, r=10, b=10),
                                title_font=dict(size=16),
                                height=400)

        st.plotly_chart(fig_top10, use_container_width=True)

    # ================= COLUNA 2 - GRAFO =================
    with col2:
        df_assoc = df_filt[["COD_VENDA", "DESCRICAO_PRODUTO"]].drop_duplicates()

        vendas_com_produto = df_assoc[df_assoc["DESCRICAO_PRODUTO"] == produto_selecionado]["COD_VENDA"].unique()
        df_relacionados = df_assoc[df_assoc["COD_VENDA"].isin(vendas_com_produto)]

        total_vendas_produto = len(vendas_com_produto)

        relacionados = df_relacionados[df_relacionados["DESCRICAO_PRODUTO"] != produto_selecionado]
        freq_relacionados = relacionados["DESCRICAO_PRODUTO"].value_counts().head(5).reset_index()
        freq_relacionados.columns = ["PRODUTO", "FREQ"]
        freq_relacionados["PCT"] = freq_relacionados["FREQ"] / total_vendas_produto

        import networkx as nx
        import plotly.graph_objects as go

        G = nx.Graph()
        G.add_node(produto_selecionado, size=100)

        for _, row in freq_relacionados.iterrows():
            G.add_node(row["PRODUTO"], size=row["PCT"] * 100)
            G.add_edge(produto_selecionado, row["PRODUTO"], weight=row["PCT"])

        pos = nx.spring_layout(G, seed=42, k=0.8)

        edge_x, edge_y = [], []
        for edge in G.edges():
            x0, y0 = pos[edge[0]]
            x1, y1 = pos[edge[1]]
            edge_x += [x0, x1, None]
            edge_y += [y0, y1, None]

        edge_trace = go.Scatter(
            x=edge_x, y=edge_y,
            line=dict(width=1.5, color="#A4B494"),
            hoverinfo="none",
            mode="lines"
        )

        node_x, node_y, node_text, node_size = [], [], [], []
        for node in G.nodes():
            x, y = pos[node]
            node_x.append(x)
            node_y.append(y)

            if node == produto_selecionado:
                legenda = ""
                tamanho = 50
                texto = f"<b>{node}</b>"
            else:
                pct = G[produto_selecionado][node]['weight']
                legenda = f"aparece em {pct:.1%} das vendas de {produto_selecionado}"
                tamanho = 20 + pct * 100
                texto = f"<b>{node}</b><br><span style='font-size:18px; color:#333;'>{legenda}</span>"
            
            node_text.append(texto)


            
            #if node == produto_selecionado:
            #    legenda = "Produto Selecionado"
            #    tamanho = 50
            #else:
            #    pct = G[produto_selecionado][node]['weight']
            #    legenda = f"{pct:.1%} das vendas com {produto_selecionado}"
            #    tamanho = 20 + pct * 100

            #node_text.append(f"<b>{node}</b><br><span style='font-size:13px; color:#333;'>{legenda}</span>")
            node_size.append(tamanho)

        node_trace = go.Scatter(
            x=node_x, y=node_y,
            mode="markers+text",
            hoverinfo="skip",
            text=node_text,
            textposition="bottom center",
            marker=dict(
                showscale=False,
                color=node_size,
                size=node_size,
                colorscale="OrRd",
                line_width=2
            )
        )

        fig_grafo = go.Figure(data=[edge_trace, node_trace],
                              layout=go.Layout(
                                  title=dict(text=f"Produtos Relacionados a: {produto_selecionado}", font=dict(size=16)),
                                  showlegend=False,
                                  margin=dict(t=40, l=0, r=0, b=0),
                                  hovermode="closest",
                                  xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                                  yaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                                  height=400
                              ))

        st.plotly_chart(fig_grafo, use_container_width=True)


#Evolução de venda por dia da semana
with st.container(border=True):
    st.markdown("<h4 style='color:#862E3A;'>📊 Evolução de Faturamento por Dia da Semana (Drilldown Mensal com Cores)</h4>", unsafe_allow_html=True)

    df_filt["MES_ANO"] = df_filt["DATA"].dt.to_period("M").astype(str)
    meses_disp = sorted(df_filt["MES_ANO"].unique())
    meses_selecionados = st.multiselect("Selecionar Mês(es):", meses_disp, default=[meses_disp[-1]])

    df_mes = df_filt[df_filt["MES_ANO"].isin(meses_selecionados)].copy()
    df_mes["SEMANA"] = df_mes["DATA"].dt.isocalendar().week
    df_mes["ANO"] = df_mes["DATA"].dt.year
    dias_traduzidos = {
        "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
        "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
    }
    df_mes["DIA_SEMANA"] = df_mes["DATA"].dt.day_name().map(dias_traduzidos)

    df_mes["INICIO_SEMANA"] = df_mes["DATA"] - pd.to_timedelta(df_mes["DATA"].dt.weekday, unit="d")
    df_mes["FIM_SEMANA"] = df_mes["INICIO_SEMANA"] + pd.Timedelta(days=6)
    df_mes["PERIODO"] = df_mes["INICIO_SEMANA"].dt.strftime('%d/%m') + " à " + df_mes["FIM_SEMANA"].dt.strftime('%d/%m')

    df_grouped = df_mes.groupby(["SEMANA", "PERIODO", "DIA_SEMANA"])["TOTAL"].sum().reset_index()
    df_pivot = df_grouped.pivot(index="DIA_SEMANA", columns="PERIODO", values="TOTAL").fillna(0)
    ordem = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]
    df_pivot = df_pivot.reindex(ordem)
    df_pivot = df_pivot[sorted(df_pivot.columns, key=lambda x: datetime.strptime(x.split(" à ")[0], "%d/%m"))]

    df_formatada = pd.DataFrame(index=df_pivot.index)
    colunas = df_pivot.columns.tolist()
    variacoes_pct = pd.DataFrame(index=df_pivot.index)

    for i, col in enumerate(colunas):
        col_fmt = []
        var_list = []
        for idx in df_pivot.index:
            valor = df_pivot.loc[idx, col]
            texto = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            variacao = None

            if i > 0:
                valor_ant = df_pivot.loc[idx, colunas[i - 1]]
                if valor_ant > 0:
                    variacao = (valor - valor_ant) / valor_ant
                    cor = "green" if variacao > 0 else "red"
                    texto += f"<br><span style='color:{cor}; font-size: 12px'>{variacao:+.2%}</span>"
            col_fmt.append(texto)
            var_list.append(variacao)
        df_formatada[col] = col_fmt
        variacoes_pct[col] = var_list

    # === Tabela HTML
    tabela_html = "<table style='border-collapse: collapse; width: 100%; text-align: center;'>"
    tabela_html += "<thead><tr><th style='padding: 6px; border: 1px solid #555;'>DIA_SEMANA</th>"

    for col in colunas:
        tabela_html += f"<th style='padding: 6px; border: 1px solid #555;'>{col}</th>"
    tabela_html += "</tr></thead><tbody>"

    for idx in df_formatada.index:
        tabela_html += f"<tr><td style='padding: 6px; border: 1px solid #555; font-weight: bold'>{idx}</td>"
        for col in colunas:
            celula = df_formatada.loc[idx, col]
            pct = variacoes_pct.loc[idx, col]

            # Cor de fundo com proteção
            if pct is None or pd.isna(pct):
                fundo = "#f0f0f0"
            else:
                try:
                    if pct >= 0:
                        intensidade = int(255 - min(pct, 1) * 155)
                        fundo = f"rgb({intensidade}, 255, {intensidade})"
                    else:
                        intensidade = int(255 - min(abs(pct), 1) * 155)
                        fundo = f"rgb(255, {intensidade}, {intensidade})"
                except:
                    fundo = "#f0f0f0"

            tabela_html += f"<td style='padding: 6px; border: 1px solid #555; background-color: {fundo}; color: #111;'>{celula}</td>"
        tabela_html += "</tr>"
    tabela_html += "</tbody></table>"

    st.markdown(tabela_html, unsafe_allow_html=True)

    # === Exportação para Excel
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativo"

    # Cabeçalhos
    ws.append(["DIA_SEMANA"] + colunas)

    # Dados com variações simples
    for idx in df_pivot.index:
        linha = [idx]
        for col in colunas:
            val = df_pivot.loc[idx, col]
            linha.append(round(val, 2))
        ws.append(linha)

    # Estilização no Excel
    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            pct_row = cell.row - 2
            pct_col = cell.column - 2
            try:
                pct = variacoes_pct.iloc[pct_row, pct_col]
                if pct is not None:
                    if pct >= 0:
                        fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
                    else:
                        fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                    cell.fill = fill
            except:
                continue
            cell.alignment = Alignment(horizontal="center")

    wb.save(output)
    st.download_button(
        label="📥 Baixar Excel",
        data=output.getvalue(),
        file_name="comparativo_dia_da_semana.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# === Evolução da Quantidade de Vendas por Dia da Semana (Drilldown Semanal)
st.markdown("---")
with st.container(border=True):
    st.markdown("<h4 style='color:#862E3A;'>📊 Evolução da Qtde de Vendas por Dia da Semana (Drilldown Mensal com Cores)</h4>", unsafe_allow_html=True)

    df_filt["MES_ANO"] = df_filt["DATA"].dt.to_period("M").astype(str)
    meses_disp = sorted(df_filt["MES_ANO"].unique())
    meses_selecionados_vendas = st.multiselect("Selecionar Mês(es):", meses_disp, default=[meses_disp[-1]], key="meses_vendas")

    df_mes_vendas = df_filt[df_filt["MES_ANO"].isin(meses_selecionados_vendas)].copy()
    df_mes_vendas["SEMANA"] = df_mes_vendas["DATA"].dt.isocalendar().week
    df_mes_vendas["ANO"] = df_mes_vendas["DATA"].dt.year
    dias_traduzidos = {
        "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
        "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
    }
    df_mes_vendas["DIA_SEMANA"] = df_mes_vendas["DATA"].dt.day_name().map(dias_traduzidos)
    df_mes_vendas["INICIO_SEMANA"] = df_mes_vendas["DATA"] - pd.to_timedelta(df_mes_vendas["DATA"].dt.weekday, unit="d")
    df_mes_vendas["FIM_SEMANA"] = df_mes_vendas["INICIO_SEMANA"] + pd.Timedelta(days=6)
    df_mes_vendas["PERIODO"] = df_mes_vendas["INICIO_SEMANA"].dt.strftime('%d/%m') + " à " + df_mes_vendas["FIM_SEMANA"].dt.strftime('%d/%m')

    df_grouped_vendas = df_mes_vendas.groupby(["SEMANA", "PERIODO", "DIA_SEMANA"])["COD_VENDA"].nunique().reset_index(name="QTD_VENDAS")
    df_pivot_vendas = df_grouped_vendas.pivot(index="DIA_SEMANA", columns="PERIODO", values="QTD_VENDAS").fillna(0)

    ordem = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]
    df_pivot_vendas = df_pivot_vendas.reindex(ordem)
    df_pivot_vendas = df_pivot_vendas[sorted(df_pivot_vendas.columns, key=lambda x: datetime.strptime(x.split(" à ")[0], "%d/%m"))]

    df_formatada_vendas = pd.DataFrame(index=df_pivot_vendas.index)
    colunas_vendas = df_pivot_vendas.columns.tolist()
    variacoes_pct_vendas = pd.DataFrame(index=df_pivot_vendas.index)

    for i, col in enumerate(colunas_vendas):
        col_fmt = []
        var_list = []
        for idx in df_pivot_vendas.index:
            valor = df_pivot_vendas.loc[idx, col]
            texto = f"{int(valor):,}".replace(",", ".")
            variacao = None
            if i > 0:
                valor_ant = df_pivot_vendas.loc[idx, colunas_vendas[i - 1]]
                if valor_ant > 0:
                    variacao = (valor - valor_ant) / valor_ant
                    cor = "green" if variacao > 0 else "red"
                    texto += f"<br><span style='color:{cor}; font-size: 14px'>{variacao:+.2%}</span>"
            col_fmt.append(texto)
            var_list.append(variacao)
        df_formatada_vendas[col] = col_fmt
        variacoes_pct_vendas[col] = var_list

    # === Tabela HTML
    tabela_vendas = "<table style='border-collapse: collapse; width: 100%; text-align: center;'>"
    tabela_vendas += "<thead><tr><th style='padding: 6px; border: 1px solid #555;'>DIA_SEMANA</th>"

    for col in colunas_vendas:
        tabela_vendas += f"<th style='padding: 6px; border: 1px solid #555;'>{col}</th>"
    tabela_vendas += "</tr></thead><tbody>"

    for idx in df_formatada_vendas.index:
        tabela_vendas += f"<tr><td style='padding: 6px; border: 1px solid #555; font-weight: bold'>{idx}</td>"
        for col in colunas_vendas:
            celula = df_formatada_vendas.loc[idx, col]
            pct = variacoes_pct_vendas.loc[idx, col]

            # Cor de fundo com proteção
            if pct is None or pd.isna(pct):
                fundo = "#f0f0f0"
            else:
                try:
                    if pct >= 0:
                        intensidade = int(255 - min(pct, 1) * 155)
                        fundo = f"rgb({intensidade}, 255, {intensidade})"
                    else:
                        intensidade = int(255 - min(abs(pct), 1) * 155)
                        fundo = f"rgb(255, {intensidade}, {intensidade})"
                except:
                    fundo = "#f0f0f0"

            tabela_vendas += f"<td style='padding: 6px; border: 1px solid #555; background-color: {fundo}; color: #111;'>{celula}</td>"
        tabela_vendas += "</tr>"
    tabela_vendas += "</tbody></table>"

    st.markdown(tabela_vendas, unsafe_allow_html=True)

# === Evolução do Ticket Médio por Dia da Semana (Drilldown Semanal)
st.markdown("---")
with st.container(border=True):
    st.markdown("<h4 style='color:#862E3A;'>📊 Evolução do Ticket Médio por Dia da Semana (Drilldown Mensal com Cores)</h4>", unsafe_allow_html=True)

    df_filt["MES_ANO"] = df_filt["DATA"].dt.to_period("M").astype(str)
    meses_disp = sorted(df_filt["MES_ANO"].unique())
    meses_selecionados_ticket = st.multiselect("Selecionar Mês(es):", meses_disp, default=[meses_disp[-1]], key="meses_ticket")

    df_mes_ticket = df_filt[df_filt["MES_ANO"].isin(meses_selecionados_ticket)].copy()
    df_mes_ticket["SEMANA"] = df_mes_ticket["DATA"].dt.isocalendar().week
    df_mes_ticket["ANO"] = df_mes_ticket["DATA"].dt.year

    dias_traduzidos = {
        "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
        "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
    }
    df_mes_ticket["DIA_SEMANA"] = df_mes_ticket["DATA"].dt.day_name().map(dias_traduzidos)

    df_mes_ticket["INICIO_SEMANA"] = df_mes_ticket["DATA"] - pd.to_timedelta(df_mes_ticket["DATA"].dt.weekday, unit="d")
    df_mes_ticket["FIM_SEMANA"] = df_mes_ticket["INICIO_SEMANA"] + pd.Timedelta(days=6)
    df_mes_ticket["PERIODO"] = df_mes_ticket["INICIO_SEMANA"].dt.strftime('%d/%m') + " à " + df_mes_ticket["FIM_SEMANA"].dt.strftime('%d/%m')

    df_grouped_ticket = df_mes_ticket.groupby(["SEMANA", "PERIODO", "DIA_SEMANA"]).agg({"TOTAL": "sum", "COD_VENDA": "nunique"}).reset_index()
    df_grouped_ticket["TICKET_MEDIO"] = df_grouped_ticket["TOTAL"] / df_grouped_ticket["COD_VENDA"]

    df_pivot_ticket = df_grouped_ticket.pivot(index="DIA_SEMANA", columns="PERIODO", values="TICKET_MEDIO").fillna(0)

    ordem = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]
    df_pivot_ticket = df_pivot_ticket.reindex(ordem)
    df_pivot_ticket = df_pivot_ticket[sorted(df_pivot_ticket.columns, key=lambda x: datetime.strptime(x.split(" à ")[0], "%d/%m"))]

    df_formatada_ticket = pd.DataFrame(index=df_pivot_ticket.index)
    colunas_ticket = df_pivot_ticket.columns.tolist()
    variacoes_pct_ticket = pd.DataFrame(index=df_pivot_ticket.index)

    for i, col in enumerate(colunas_ticket):
        col_fmt = []
        var_list = []
        for idx in df_pivot_ticket.index:
            valor = df_pivot_ticket.loc[idx, col]
            texto = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            variacao = None
            if i > 0:
                valor_ant = df_pivot_ticket.loc[idx, colunas_ticket[i - 1]]
                if valor_ant > 0:
                    variacao = (valor - valor_ant) / valor_ant
                    cor = "green" if variacao > 0 else "red"
                    texto += f"<br><span style='color:{cor}; font-size: 14px'>{variacao:+.2%}</span>"
            col_fmt.append(texto)
            var_list.append(variacao)
        df_formatada_ticket[col] = col_fmt
        variacoes_pct_ticket[col] = var_list

    # === Tabela HTML
    tabela_ticket = "<table style='border-collapse: collapse; width: 100%; text-align: center;'>"
    tabela_ticket += "<thead><tr><th style='padding: 6px; border: 1px solid #555;'>DIA_SEMANA</th>"

    for col in colunas_ticket:
        tabela_ticket += f"<th style='padding: 6px; border: 1px solid #555;'>{col}</th>"
    tabela_ticket += "</tr></thead><tbody>"

    for idx in df_formatada_ticket.index:
        tabela_ticket += f"<tr><td style='padding: 6px; border: 1px solid #555; font-weight: bold'>{idx}</td>"
        for col in colunas_ticket:
            celula = df_formatada_ticket.loc[idx, col]
            pct = variacoes_pct_ticket.loc[idx, col]

            # Cor de fundo com proteção
            try:
                if pd.isna(pct) or not isinstance(pct, (int, float)):
                    raise ValueError("Valor inválido para pct")
                if pct >= 0:
                    intensidade = int(255 - min(pct, 1) * 155)
                    fundo = f"rgb({intensidade}, 255, {intensidade})"
                else:
                    intensidade = int(255 - min(abs(pct), 1) * 155)
                    fundo = f"rgb(255, {intensidade}, {intensidade})"
            except:
                fundo = "#f0f0f0"

            tabela_ticket += f"<td style='padding: 6px; border: 1px solid #555; background-color: {fundo}; color: #111;'>{celula}</td>"
        tabela_ticket += "</tr>"
    tabela_ticket += "</tbody></table>"

st.markdown(tabela_ticket, unsafe_allow_html=True)
