import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Configuração da página para ocupar toda a tela e ter o título correto
st.set_page_config(page_title="Conecta+ Strategy Cockpit", layout="wide", initial_sidebar_state="collapsed")

# Customização de CSS para o tema escuro (inspirado na imagem)
st.markdown("""
    <style>
    .reportview-container { background: #0b0f19; color: white; }
    .stMetric { background-color: #1a2235; padding: 15px; border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

st.title("Conecta+ Strategy Cockpit")
st.subheader("Efetividade, custo e performance por estratégia")

# --- SIMULAÇÃO DE CARREGAMENTO DE DADOS ---
# Na prática, você usaria: pd.read_csv("url_do_seu_sheets_export") 
# ou a biblioteca gspread para ler a aba 'dashboard_fact' e 'strategy'
@st.cache_data
def load_data():
    # URL de exportação direta do Google Sheets (exemplo para a primeira aba)
    url = "https://docs.google.com/spreadsheets/d/16qSTNR6z920Rp0LMdwpBp1pcKvfXZp-jSjUmfxIN95A/export?format=csv"
    try:
        df = pd.read_csv(url)
        return df
    except:
        # Fallback de dados baseados na estrutura que você enviou
        return pd.DataFrame({
            "Estratégia": ["Custo Eficiente", "Máximo Contato", "Cobrança Progressiva"],
            "Números Únicos": [48320, 41872, 35238],
            "% Contatos Produtivos": [42.1, 36.8, 35.9],
            "Custo Total": [11280.50, 13940.20, 3230.05],
            "Custo por Contato": [0.52, 0.78, 0.46]
        })

df = load_data()

# --- FILTROS ---
col_periodo, col_estrategia, _ = st.columns([2, 2, 6])
with col_periodo:
    st.selectbox("Período", ["01 Jan 2026 - 31 Jan 2026"])
with col_estrategia:
    st.selectbox("Estratégia", ["Todas", "Custo Eficiente", "Máximo Contato", "Cobrança Progressiva"])

# --- KPIs PRINCIPAIS ---
st.write("---")
kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)

kpi1.metric(label="Números únicos", value="125.430", delta="+12,4%")
kpi2.metric(label="Contatos produtivos", value="48.219", delta="+18,7%")
kpi3.metric(label="Contatos improdutivos", value="32.105", delta="+6,1%")
kpi4.metric(label="Sem contato", value="45.106", delta="-8,3%")
kpi5.metric(label="Custo total", value="R$ 28.450,75", delta="+4,9%")
kpi6.metric(label="Custo por contato efetivo", value="R$ 0,59", delta="-11,3%")

st.write("---")

# --- SEÇÃO INFERIOR: VISÃO POR ESTRATÉGIA E DRILL DOWN ---
col_esq, col_dir = st.columns([1, 1])

with col_esq:
    st.markdown("### Visão por Estratégia")
    st.dataframe(df, use_container_width=True, hide_index=True)
    
    st.markdown("### Funil da Estratégia Selecionada")
    # Gráfico de Funil usando Plotly
    fig_funnel = go.Figure(go.Funnel(
        y=["Números únicos", "Contatos produtivos", "Contatos improdutivos", "Sem contato"],
        x=[48320, 20342, 12450, 15528],
        textinfo="value+percent initial",
        marker={"color": ["#1f77b4", "#00bfa5", "#e91e63", "#607d8b"]}
    ))
    fig_funnel.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'))
    st.plotly_chart(fig_funnel, use_container_width=True)

with col_dir:
    st.markdown("### Configuração da Estratégia Selecionada: Custo Eficiente")
    st.info("**Sequência:** WhatsApp Texto ➔ Ligação Tradicional (+15 min) ➔ Branded Call (+2 h) ➔ WhatsApp Call (+24 h)")
    
    # Gráficos de Donut (Drill Down)
    col_donut1, col_donut2 = st.columns(2)
    
    with col_donut1:
        st.markdown("**Drill down — Contatos Produtivos**")
        labels = ['Chamada tradicional', 'Branded Calls', 'WhatsApp Call', 'Agendamento WhatsApp']
        values = [34.2, 28.6, 22.1, 13.5]
        fig_donut1 = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.6)])
        fig_donut1.update_layout(showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), margin=dict(t=0, b=0, l=0, r=0))
        st.plotly_chart(fig_donut1, use_container_width=True)

    with col_donut2:
        st.markdown("**Drill down — Contatos Improdutivos**")
        labels = ['00-30 seg', '30 seg - 1 min', '1 - 2 min']
        values = [52.3, 32.8, 14.9]
        fig_donut2 = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.6)])
        fig_donut2.update_layout(showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), margin=dict(t=0, b=0, l=0, r=0))
        st.plotly_chart(fig_donut2, use_container_width=True)