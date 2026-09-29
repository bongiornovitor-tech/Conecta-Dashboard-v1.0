import streamlit as st
import pandas as pd
import requests
import io
import plotly.graph_objects as go
import base64
import os

st.set_page_config(page_title="Conecta+ Strategy Cockpit", layout="wide", initial_sidebar_state="collapsed")

# 1. INJEÇÃO DE CSS GLOBAL AVANÇADO
st.markdown("""
<style>
/* Fundo e tipografia geral */
.stApp { background-color: #050b14; color: #e2e8f0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }

/* Esconde elementos padrão do Streamlit */
header {visibility: hidden;}
footer {visibility: hidden;}
.css-18e3th9 {padding-top: 0rem;}

/* Cabecalho e Logo */
.header-container { display: flex; align-items: center; padding: 15px 0 25px 0; margin-bottom: 20px; border-bottom: 1px solid rgba(255,255,255,0.05); }
.header-text h1 { color: white; margin: 0; font-size: 34px; font-weight: 700; display: inline-block; letter-spacing: -0.5px;}
.header-text h1 span { color: #8b5cf6; } 
.header-text p { color: #9ca3af; margin: 4px 0 0 0; font-size: 15px; font-weight: 400;}

/* WIDGETS NEON */
.kpi-wrapper { display: flex; gap: 15px; justify-content: space-between; margin-bottom: 30px; }
.kpi-card { 
    flex: 1; background: linear-gradient(145deg, #0b1121, #060913); border-radius: 12px; padding: 18px;
    border: 1px solid rgba(255,255,255,0.04); box-shadow: 0 4px 20px rgba(0,0,0,0.4);
    display: flex; flex-direction: column; justify-content: space-between; min-width: 160px;
}
.kpi-top { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 15px; }
.kpi-icon { 
    width: 42px; height: 42px; border-radius: 10px; display: flex; align-items: center; justify-content: center; 
    font-size: 20px; color: white; flex-shrink: 0;
}
.kpi-text-group { display: flex; flex-direction: column; }
.kpi-title { font-size: 12px; color: #9ca3af; font-weight: 500; margin-bottom: 2px; }
.kpi-value { font-size: 24px; color: white; font-weight: 700; margin: 0; line-height: 1.1; letter-spacing: -0.5px;}
.kpi-bottom { display: flex; justify-content: space-between; align-items: flex-end; }
.kpi-delta-group { display: flex; flex-direction: column; }
.kpi-delta { font-size: 12px; font-weight: 700; display: flex; align-items: center; gap: 4px;}
.kpi-delta-desc { font-size: 10px; color: #6b7280; margin-top: 2px;}
.kpi-sparkline { margin-bottom: -5px; }

/* VISÃO POR ESTRATÉGIA MODERNA (Cards Limpos sem Linhas) */
.strat-card {
    background: linear-gradient(145deg, #0b1121, #060913);
    border: 1px solid rgba(255,255,255,0.04);
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 12px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.2);
}
.strat-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px; }
.strat-name { font-size: 16px; font-weight: bold; color: white; }
.strat-desc { font-size: 12px; color: #9ca3af; margin-top: 2px; }
.strat-metrics-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; text-align: left; }
.strat-metric-label { font-size: 10px; text-transform: uppercase; color: #64748b; font-weight: 600; margin-bottom: 2px; }
.strat-metric-val { font-size: 14px; font-weight: 600; color: white; }
.p-bar-bg { width: 100%; background-color: #1e293b; border-radius: 10px; height: 6px; margin-top: 6px; overflow: hidden; }
.p-bar-fill { height: 100%; border-radius: 10px; }

/* Fluxo da Estratégia */
.flow-wrapper { display: flex; align-items: center; justify-content: center; gap: 20px; background: #0b1120; padding: 25px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.03); margin-top: 15px;}
.flow-step { display: flex; flex-direction: column; align-items: center; width: 90px; }
.flow-icon { width: 50px; height: 50px; border-radius: 12px; display: flex; align-items: center; justify-content: center; margin-bottom: 10px; font-size: 24px; color: white; box-shadow: 0 5px 15px rgba(0,0,0,0.3);}
.flow-label { font-size: 12px; color: #cbd5e1; font-weight: 500; text-align: center; line-height: 1.2;}
.flow-arrow { color: #475569; font-size: 18px; margin-top: -20px;}

/* Funil Customizado 3D */
.funnel-container { display: flex; flex-direction: column; align-items: center; gap: 5px; margin-top: 30px; width: 100%;}
.funnel-layer { position: relative; display: flex; justify-content: center; align-items: center; text-align: center; color: white; font-weight: bold; font-size: 14px; text-shadow: 1px 1px 2px rgba(0,0,0,0.8); }
.funnel-layer span { position: relative; z-index: 2; line-height: 1.2;}
.f1 { width: 100%; height: 60px; background: linear-gradient(90deg, #1e3a8a, #3b82f6); clip-path: polygon(0 0, 100% 0, 85% 100%, 15% 100%); }
.f2 { width: 70%; height: 60px; background: linear-gradient(90deg, #0f766e, #14b8a6); clip-path: polygon(0 0, 100% 0, 80% 100%, 20% 100%); }
.f3 { width: 42%; height: 50px; background: linear-gradient(90deg, #be123c, #f43f5e); clip-path: polygon(0 0, 100% 0, 75% 100%, 25% 100%); }
.f4 { width: 21%; height: 40px; background: linear-gradient(90deg, #334155, #64748b); clip-path: polygon(0 0, 100% 0, 100% 100%, 0 100%); border-radius: 0 0 8px 8px;}

/* Cabeçalhos de Seção */
.section-title { font-size: 18px; color: white; font-weight: 600; margin: 20px 0 15px 0; border-left: 4px solid #3b82f6; padding-left: 10px; }
</style>
""", unsafe_allow_html=True)

# 2. LOGO LIDO DIRETAMENTE DO ARQUIVO LOCAL DO GITHUB
def carregar_logo():
    caminho_arquivo = "Conecta+ Logo.png"
    if os.path.exists(caminho_arquivo):
        with open(caminho_arquivo, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            return f"data:image/png;base64,{encoded_string}"
    else:
        return "https://raw.githubusercontent.com/bongiornovitor-tech/Conecta-Dashboard-v1.0/main/Conecta%2B%20Logo.png"

st.markdown(f"""
<div class="header-container">
<img src="{carregar_logo()}" width="160" style="margin-right: 25px;">
<div class="header-text">
<h1>Conecta+ <span>Strategy Cockpit</span></h1>
<p>Efetividade, custo e performance por estratégia</p>
</div>
</div>
""", unsafe_allow_html=True)

# --- CONEXÃO COM DADOS REAIS ---
@st.cache_data(ttl=600)
def load_data():
    url = "https://docs.google.com/spreadsheets/d/16qSTNR6z920Rp0LMdwpBp1pcKvfXZp-jSjUmfxIN95A/export?format=xlsx"
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers)
    excel_data = io.BytesIO(response.content)
    df_fact = pd.read_excel(excel_data, sheet_name='dashboard_fact')
    df_strat = pd.read_excel(excel_data, sheet_name='strategy')
    df_steps = pd.read_excel(excel_data, sheet_name='strategy_steps')
    
    if 'attempt_cost' in df_fact.columns:
        df_fact['custo_num'] = df_fact['attempt_cost'].astype(str).str.replace('R$', '', regex=False).str.replace(' ', '', regex=False).str.replace(',', '.', regex=False)
        df_fact['custo_num'] = pd.to_numeric(df_fact['custo_num'], errors='coerce').fillna(0)
    return df_fact, df_strat, df_steps

df_fact, df_strat, df_steps = load_data()

# --- FILTROS ---
estrategias_disp = ["Todas"] + list(df_fact['strategy_name'].dropna().unique())
col_p, col_e, _ = st.columns([2, 2, 8])
with col_p:
    st.selectbox("Período", ["01 Set 2026 - 30 Set 2026"], label_visibility="collapsed")
with col_e:
    estr_selecionada = st.selectbox("Estratégia", estrategias_disp, label_visibility="collapsed")

df_filtered = df_fact[df_fact['strategy_name'] == estr_selecionada] if estr_selecionada != "Todas" else df_fact

# --- CÁLCULO DOS KPIs ---
n_unicos = df_filtered['contact_id'].nunique()
n_prod = df_filtered[df_filtered['productive_flag'] == 1]['contact_id'].nunique()
n_improd = df_filtered[df_filtered['unproductive_flag'] == 1]['contact_id'].nunique()
n_sem = max(0, n_unicos - n_prod - n_improd)
c_total = df_filtered['custo_num'].sum()
c_efetivo = c_total / n_prod if n_prod > 0 else 0

# --- GERADOR DE SPARKLINE (Tendência) EM SVG ---
def get_sparkline_svg(color, points):
    return f"""<svg viewBox="0 0 100 30" width="70" height="25" preserveAspectRatio="none">
<defs>
<linearGradient id="grad-{color.replace('#','')}" x1="0" y1="0" x2="0" y2="1">
<stop offset="0%" stop-color="{color}" stop-opacity="0.5"/>
<stop offset="100%" stop-color="{color}" stop-opacity="0.0"/>
</linearGradient>
</defs>
<polyline points="{points}" fill="none" stroke="{color}" stroke-width="2.5" stroke-linecap="round"/>
<polygon points="0,30 {points} 100,30" fill="url(#grad-{color.replace('#','')})"/>
</svg>"""

# 3. CONSTRUÇÃO DINÂMICA DOS CARDS NEON
cards_data = [
    {"titulo": "Números únicos", "valor": f"{n_unicos:,.0f}", "cor": "#3b82f6", "icone": "👥", "delta_val": "+12,4%", "delta_cor": "#10b981", "delta_seta": "▲", "desc": "vs. período anterior", "pontos": "0,20 20,22 40,15 60,20 80,10 100,5"},
    {"titulo": "Contatos produtivos", "valor": f"{n_prod:,.0f}", "cor": "#14b8a6", "icone": "📞", "delta_val": "+18,7%", "delta_cor": "#10b981", "delta_seta": "▲", "desc": "38,4% da base", "pontos": "0,25 20,20 40,22 60,10 80,12 100,2"},
    {"titulo": "Contatos improdutivos", "valor": f"{n_improd:,.0f}", "cor": "#f43f5e", "icone": "📵", "delta_val": "+6,1%", "delta_cor": "#10b981", "delta_seta": "▲", "desc": "25,6% da base", "pontos": "0,25 20,24 40,20 60,22 80,15 100,5"},
    {"titulo": "Sem contato", "valor": f"{n_sem:,.0f}", "cor": "#6b7280", "icone": "📴", "delta_val": "-8,3%", "delta_cor": "#f43f5e", "delta_seta": "▼", "desc": "36,0% da base", "pontos": "0,5 20,10 40,8 60,15 80,12 100,25"},
    {"titulo": "Custo total", "valor": f"R$ {c_total:,.2f}", "cor": "#8b5cf6", "icone": "🪙", "delta_val": "+4,9%", "delta_cor": "#10b981", "delta_seta": "▲", "desc": "vs. período anterior", "pontos": "0,25 20,26 40,20 60,15 80,18 100,5"}
]

html_cards = '<div class="kpi-wrapper">\n'
for c in cards_data:
    valor_fmt = c["valor"].replace(',', 'X').replace('.', ',').replace('X', '.')
    svg = get_sparkline_svg(c["cor"], c["pontos"])
    html_cards += f'''<div class="kpi-card" style="border-top: 2px solid {c['cor']}40;">
<div class="kpi-top">
<div class="kpi-icon" style="background-color: {c['cor']}; box-shadow: 0 0 15px {c['cor']}60;">{c['icone']}</div>
<div class="kpi-text-group">
<div class="kpi-title">{c['titulo']}</div>
<div class="kpi-value">{valor_fmt}</div>
</div>
</div>
<div class="kpi-bottom">
<div class="kpi-delta-group">
<div class="kpi-delta" style="color: {c['delta_cor']};"><span>{c['delta_seta']}</span> {c['delta_val']}</div>
<div class="kpi-delta-desc">{c['desc']}</div>
</div>
<div class="kpi-sparkline">{svg}</div>
</div>
</div>\n'''
html_cards += '</div>'
st.markdown(html_cards, unsafe_allow_html=True)


col_esq, col_dir = st.columns([1, 1])

with col_esq:
    st.markdown('<div class="section-title">Visão por Estratégia</div>', unsafe_allow_html=True)
    
    # 4. VISÃO POR ESTRATÉGIA MODERNA (Com Descrições e Barras Proporcionais)
    df_grp = df_fact.groupby('strategy_name').agg(unicos=('contact_id', 'nunique'), custo_tot=('custo_num', 'sum')).reset_index()
    prod_grp = df_fact[df_fact['productive_flag']==1].groupby('strategy_name')['contact_id'].nunique().reset_index()
    prod_grp.rename(columns={'contact_id': 'produtivos'}, inplace=True)
    df_grp = df_grp.merge(prod_grp, on='strategy_name', how='left').fillna(0)
    df_grp = df_grp.merge(df_strat[['strategy_name', 'objective']], on='strategy_name', how='left')
    
    max_custo_ef = 3.0 # Limite máximo proporcional para a barra de custo unitário
    
    for _, row in df_grp.iterrows():
        pct_prod = (row['produtivos'] / row['unicos']) * 100 if row['unicos'] > 0 else 0
        custo_ef = row['custo_tot'] / row['produtivos'] if row['produtivos'] > 0 else 0
        pct_custo_ef = min(100, (custo_ef / max_custo_ef) * 100) # Proporção exata baseada no teto
        
        unicos_fmt = f"{row['unicos']:,.0f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        custo_tot_fmt = f"R$ {row['custo_tot']:,.2f}".replace('.', ',')
        custo_ef_fmt = f"R$ {custo_ef:,.2f}".replace('.', ',')
        
        st.markdown(f'''
        <div class="strat-card">
            <div class="strat-header">
                <div>
                    <div class="strat-name">{row['strategy_name']}</div>
                    <div class="strat-desc">{row['objective']}</div>
                </div>
            </div>
            <div class="strat-metrics-grid">
                <div>
                    <div class="strat-metric-label">Números únicos</div>
                    <div class="strat-metric-val">{unicos_fmt}</div>
                </div>
                <div>
                    <div class="strat-metric-label">% Produtivo</div>
                    <div class="strat-metric-val">{pct_prod:.1f}%</div>
                    <div class="p-bar-bg"><div class="p-bar-fill" style="width:{pct_prod}%; background:#14b8a6;"></div></div>
                </div>
                <div>
                    <div class="strat-metric-label">Custo total</div>
                    <div class="strat-metric-val">{custo_tot_fmt}</div>
                </div>
                <div>
                    <div class="strat-metric-label">Custo / Efetivo</div>
                    <div class="strat-metric-val">{custo_ef_fmt}</div>
                    <div class="p-bar-bg"><div class="p-bar-fill" style="width:{pct_custo_ef}%; background:#8b5cf6;"></div></div>
                </div>
            </div>
        </div>
        ''', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Funil da Estratégia</div>', unsafe_allow_html=True)
    
    # 5. FUNIL 3D
    pct_prod = (n_prod/n_unicos*100) if n_unicos>0 else 0
    pct_improd = (n_improd/n_unicos*100) if n_unicos>0 else 0
    pct_sem = (n_sem/n_unicos*100) if n_unicos>0 else 0
    
    st.markdown(f"""<div class="funnel-container">
<div class="funnel-layer f1"><span>{n_unicos:,.0f}<br><span style="font-size:11px; font-weight:normal;">Números únicos</span></span></div>
<div class="funnel-layer f2"><span>{n_prod:,.0f} ({pct_prod:.1f}%)<br><span style="font-size:11px; font-weight:normal;">Contatos produtivos</span></span></div>
<div class="funnel-layer f3"><span>{n_improd:,.0f} ({pct_improd:.1f}%)<br><span style="font-size:11px; font-weight:normal;">Contatos improdutivos</span></span></div>
<div class="funnel-layer f4"><span>{n_sem:,.0f} ({pct_sem:.1f}%)<br><span style="font-size:11px; font-weight:normal;">Sem contato</span></span></div>
</div>""".replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)

with col_dir:
    st.markdown(f'<div class="section-title">Configuração da Estratégia: {estr_selecionada}</div>', unsafe_allow_html=True)
    
    # 6. DIAGRAMA DE FLUXO
    if estr_selecionada != "Todas":
        strat_id = df_strat[df_strat['strategy_name'] == estr_selecionada]['strategy_id'].iloc[0]
        steps = df_steps[df_steps['strategy_id'] == strat_id].sort_values('step_order')
        
        flow_html = '<div class="flow-wrapper">\n'
        for i, row in steps.iterrows():
            canal = row['channel']
            if 'whatsapp' in canal.lower():
                bg, icon, label = "background: #10b981;", "💬", "WhatsApp"
            elif 'branded' in canal.lower():
                bg, icon, label = "background: #8b5cf6;", "📊", "Branded Call"
            else:
                bg, icon, label = "background: #3b82f6;", "📞", "Ligação Trad"
                
            flow_html += f'''<div class="flow-step">
<div class="flow-icon" style="{bg}">{icon}</div>
<div class="flow-label">{label}</div>
</div>\n'''
            if i < len(steps) - 1: flow_html += '<div class="flow-arrow">➔</div>\n'
        flow_html += '</div>'
        st.markdown(flow_html, unsafe_allow_html=True)
    else:
        st.info("Selecione uma estratégia para ver o fluxo.")

    # 7. DRILL DOWNS
    st.markdown("<br>", unsafe_allow_html=True)
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        st.markdown("<div style='text-align:center; color:#cbd5e1; font-size:14px; margin-bottom:10px;'>Distribuição Produtivos</div>", unsafe_allow_html=True)
        df_prod = df_filtered[df_filtered['productive_flag'] == 1]
        if not df_prod.empty:
            df_chan = df_prod['channel'].value_counts().reset_index()
            df_chan.columns = ['Canal', 'Contatos']
            fig_donut1 = go.Figure(data=[go.Pie(labels=df_chan['Canal'], values=df_chan['Contatos'], hole=.75, textinfo='percent', textposition='outside')])
            fig_donut1.update_traces(marker=dict(colors=['#14b8a6', '#8b5cf6', '#3b82f6', '#f43f5e']))
            fig_donut1.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.2), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), height=280, margin=dict(t=0,b=0,l=0,r=0))
            st.plotly_chart(fig_donut1, use_container_width=True)

    with col_d2:
        st.markdown("<div style='text-align:center; color:#cbd5e1; font-size:14px; margin-bottom:10px;'>Distribuição Improdutivos</div>", unsafe_allow_html=True)
        df_improd = df_filtered[df_filtered['unproductive_flag'] == 1]
        if not df_improd.empty:
            df_dur = df_improd['duration_band'].value_counts().reset_index()
            df_dur.columns = ['Duração', 'Contatos']
            fig_donut2 = go.Figure(data=[go.Pie(labels=df_dur['Duração'], values=df_dur['Contatos'], hole=.75, textinfo='percent', textposition='outside')])
            fig_donut2.update_traces(marker=dict(colors=['#8b5cf6', '#3b82f6', '#14b8a6']))
            fig_donut2.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.2), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), height=280, margin=dict(t=0,b=0,l=0,r=0))
            st.plotly_chart(fig_donut2, use_container_width=True)
