import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import io

# Configuração da página e tema escuro integral
st.set_page_config(page_title="Conecta+ Strategy Cockpit", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    .stApp { background-color: #0b0f19; }
    .stMetric { background-color: #1a2235; padding: 15px; border-radius: 10px; }
    div[data-testid="stMetricValue"], div[data-testid="stMetricDelta"] { color: white; }
    h1, h2, h3, h4, p, span, div, label { color: #e2e8f0; }
    .stDataFrame { background-color: #1a2235; border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

st.title("Conecta+ Strategy Cockpit")
st.markdown("Efetividade, custo e performance por estratégia")

# --- CONEXÃO COM SEUS DADOS REAIS ---
@st.cache_data(ttl=600) # Atualiza a cada 10 min
def load_data():
    url = "https://docs.google.com/spreadsheets/d/16qSTNR6z920Rp0LMdwpBp1pcKvfXZp-jSjUmfxIN95A/export?format=xlsx"
    
    # Faz o download fingindo ser um navegador comum para evitar bloqueios (HTTPError)
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    response = requests.get(url, headers=headers)
    response.raise_for_status() # Garante que o download funcionou
    
    # Carrega o arquivo excel em memória
    excel_data = io.BytesIO(response.content)
    
    df_fact = pd.read_excel(excel_data, sheet_name='dashboard_fact')
    df_strat = pd.read_excel(excel_data, sheet_name='strategy')
    df_steps = pd.read_excel(excel_data, sheet_name='strategy_steps')
    
    # Tratamento da coluna de custos (transformando "R$ 0.34" em número)
    if 'attempt_cost' in df_fact.columns:
        df_fact['custo_num'] = df_fact['attempt_cost'].astype(str).str.replace('R$', '', regex=False).str.replace(' ', '', regex=False).str.replace(',', '.', regex=False)
        df_fact['custo_num'] = pd.to_numeric(df_fact['custo_num'], errors='coerce').fillna(0)
    else:
        df_fact['custo_num'] = 0.0

    return df_fact, df_strat, df_steps

df_fact, df_strat, df_steps = load_data()

# --- FILTROS ---
estrategias_disp = ["Todas"] + list(df_fact['strategy_name'].dropna().unique())
col_periodo, col_estrategia, _ = st.columns([2, 2, 6])
with col_periodo:
    st.selectbox("Período", ["01 Set 2026 - 30 Set 2026"])
with col_estrategia:
    estr_selecionada = st.selectbox("Estratégia", estrategias_disp)

if estr_selecionada != "Todas":
    df_filtered = df_fact[df_fact['strategy_name'] == estr_selecionada]
else:
    df_filtered = df_fact

# --- CÁLCULO DE KPIs REAIS ---
numeros_unicos = df_filtered['contact_id'].nunique()
contatos_prod = df_filtered[df_filtered['productive_flag'] == 1]['contact_id'].nunique()
contatos_improd = df_filtered[df_filtered['unproductive_flag'] == 1]['contact_id'].nunique()

sem_contato = numeros_unicos - contatos_prod - contatos_improd
if sem_contato < 0: sem_contato = 0

custo_total = df_filtered['custo_num'].sum()
custo_por_efetivo = custo_total / contatos_prod if contatos_prod > 0 else 0

st.write("---")
kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
kpi1.metric(label="Números únicos", value=f"{numeros_unicos:,.0f}".replace(',','.'))
kpi2.metric(label="Contatos produtivos", value=f"{contatos_prod:,.0f}".replace(',','.'))
kpi3.metric(label="Contatos improdutivos", value=f"{contatos_improd:,.0f}".replace(',','.'))
kpi4.metric(label="Sem contato", value=f"{sem_contato:,.0f}".replace(',','.'))
kpi5.metric(label="Custo total", value=f"R$ {custo_total:,.2f}".replace('.',','))
kpi6.metric(label="Custo por contato efetivo", value=f"R$ {custo_por_efetivo:,.2f}".replace('.',','))
st.write("---")

# --- VISÃO POR ESTRATÉGIA ---
col_esq, col_dir = st.columns([1, 1])

with col_esq:
    st.markdown("### Visão por Estratégia")
    
    df_grp = df_fact.groupby('strategy_name').agg(
        unicos=('contact_id', 'nunique'),
        custo_tot=('custo_num', 'sum')
    ).reset_index()
    
    prod_grp = df_fact[df_fact['productive_flag']==1].groupby('strategy_name')['contact_id'].nunique().reset_index()
    prod_grp.rename(columns={'contact_id': 'produtivos'}, inplace=True)
    
    df_grp = df_grp.merge(prod_grp, on='strategy_name', how='left').fillna(0)
    df_grp['% contato produtivo'] = (df_grp['produtivos'] / df_grp['unicos']) * 100
    df_grp['Custo por efetivo'] = df_grp['custo_tot'] / df_grp['produtivos']
    
    df_grp_show = df_grp[['strategy_name', 'unicos', '% contato produtivo', 'custo_tot', 'Custo por efetivo']].copy()
    df_grp_show.columns = ['Estratégia', 'Números Únicos', '% Produtivos', 'Custo Total (R$)', 'Custo / Efetivo (R$)']
    
    st.dataframe(df_grp_show.style.format({
        '% Produtivos': '{:.1f}%',
        'Custo Total (R$)': 'R$ {:.2f}',
        'Custo / Efetivo (R$)': 'R$ {:.2f}'
    }), use_container_width=True, hide_index=True)

    st.markdown(f"### Funil da Estratégia: {estr_selecionada}")
    fig_funnel = go.Figure(go.Funnel(
        y=["Números únicos", "Contatos produtivos", "Contatos improdutivos", "Sem contato"],
        x=[numeros_unicos, contatos_prod, contatos_improd, sem_contato],
        textinfo="value+percent initial",
        marker={"color": ["#1f77b4", "#00bfa5", "#e91e63", "#607d8b"]}
    ))
    fig_funnel.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), margin=dict(t=20, b=20))
    st.plotly_chart(fig_funnel, use_container_width=True)

with col_dir:
    st.markdown(f"### Configuração da Estratégia Selecionada")
    if estr_selecionada != "Todas":
        strat_id = df_strat[df_strat['strategy_name'] == estr_selecionada]['strategy_id'].iloc[0]
        steps = df_steps[df_steps['strategy_id'] == strat_id].sort_values('step_order')
        step_str = " ➔ ".join([f"{row['channel']} (+{row['wait_minutes']}m)" for idx, row in steps.iterrows()])
        st.info(f"**Sequência de abordagem:** {step_str}")
    else:
        st.info("Selecione uma estratégia específica no filtro do topo para ver a configuração.")

    col_donut1, col_donut2 = st.columns(2)
    
    with col_donut1:
        st.markdown("**Drill down — Produtivos**")
        df_prod = df_filtered[df_filtered['productive_flag'] == 1]
        if not df_prod.empty:
            df_chan = df_prod['channel'].value_counts().reset_index()
            df_chan.columns = ['Canal', 'Contatos']
            fig_donut1 = px.pie(df_chan, values='Contatos', names='Canal', hole=.6, color_discrete_sequence=['#1f77b4', '#9467bd', '#00bfa5'])
            fig_donut1.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.1), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), margin=dict(t=0, b=0, l=0, r=0))
            st.plotly_chart(fig_donut1, use_container_width=True)
        else:
            st.write("Sem contatos produtivos")

    with col_donut2:
        st.markdown("**Drill down — Improdutivos**")
        df_improd = df_filtered[df_filtered['unproductive_flag'] == 1]
        if not df_improd.empty:
            df_dur = df_improd['duration_band'].value_counts().reset_index()
            df_dur.columns = ['Duração', 'Contatos']
            fig_donut2 = px.pie(df_dur, values='Contatos', names='Duração', hole=.6, color_discrete_sequence=['#2ca02c', '#d62728', '#ff7f0e'])
            fig_donut2.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.1), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), margin=dict(t=0, b=0, l=0, r=0))
            st.plotly_chart(fig_donut2, use_container_width=True)
        else:
            st.write("Sem contatos improdutivos")

st.write("---")

# --- SEÇÃO INFERIOR: CUSTOS DA ESTRATÉGIA NO PERÍODO ---
st.markdown("### Custos da Estratégia no Período")
col_c1, col_c2, col_c3 = st.columns(3)

custo_improd_tot = df_filtered[df_filtered['unproductive_flag'] == 1]['custo_num'].sum()
custo_por_improd = custo_improd_tot / contatos_improd if contatos_improd > 0 else 0

col_c1.metric(label="Custo por contato efetivo", value=f"R$ {custo_por_efetivo:,.2f}".replace('.',','))
col_c2.metric(label="Custo por contato improdutivo", value=f"R$ {custo_por_improd:,.2f}".replace('.',','))
col_c3.metric(label="Custo total da estratégia", value=f"R$ {custo_total:,.2f}".replace('.',','))
