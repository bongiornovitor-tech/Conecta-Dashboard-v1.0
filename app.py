import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import io
import base64

# Configuração da página
st.set_page_config(page_title="Conecta+ Strategy Cockpit", layout="wide", initial_sidebar_state="collapsed")

# 1. INJEÇÃO DE CSS AVANÇADO (Para os cards neon, logo e tabela moderna)
st.markdown("""
    <style>
    /* Fundo geral */
    .stApp { background-color: #0b0f19; color: #e2e8f0; }
    
    /* Layout do Topo e Logo */
    .header-container { display: flex; align-items: center; gap: 20px; padding-bottom: 20px; border-bottom: 1px solid #1f2937; margin-bottom: 20px;}
    .header-text h1 { color: white; margin: 0; font-size: 32px; font-weight: 700; display: inline-block; }
    .header-text h1 span { color: #8b5cf6; } /* Strategy Cockpit em roxo */
    .header-text p { color: #9ca3af; margin: 5px 0 0 0; font-size: 16px; }
    
    /* Cards KPI Neon (CSS puro) */
    .kpi-container { display: flex; justify-content: space-between; gap: 15px; margin-bottom: 30px; flex-wrap: wrap; }
    .kpi-card { 
        flex: 1; min-width: 150px; background: #111827; border-radius: 12px; padding: 15px;
        border: 1px solid rgba(255,255,255,0.05); box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        display: flex; flex-direction: column; position: relative; overflow: hidden;
    }
    
    /* Variações de Cores dos Cards (Baseado no Print 1) */
    .card-blue { border-top: 3px solid #3b82f6; box-shadow: 0 -10px 30px -10px rgba(59, 130, 246, 0.2); }
    .card-teal { border-top: 3px solid #14b8a6; box-shadow: 0 -10px 30px -10px rgba(20, 184, 166, 0.2); }
    .card-rose { border-top: 3px solid #f43f5e; box-shadow: 0 -10px 30px -10px rgba(244, 63, 94, 0.2); }
    .card-gray { border-top: 3px solid #6b7280; box-shadow: 0 -10px 30px -10px rgba(107, 114, 128, 0.2); }
    .card-purple { border-top: 3px solid #8b5cf6; box-shadow: 0 -10px 30px -10px rgba(139, 92, 246, 0.2); }
    
    .kpi-title { font-size: 12px; color: #9ca3af; font-weight: 600; margin-bottom: 5px; }
    .kpi-value { font-size: 24px; color: white; font-weight: bold; margin: 0; }
    .kpi-delta-up { font-size: 11px; color: #10b981; font-weight: bold; margin-top: 5px;}
    .kpi-delta-down { font-size: 11px; color: #ef4444; font-weight: bold; margin-top: 5px;}
    
    /* Configuração visual da Estratégia (Setas e Ícones) */
    .flow-container { display: flex; align-items: center; justify-content: space-between; background: #111827; padding: 20px; border-radius: 12px; margin-top: 10px; }
    .flow-step { display: flex; flex-direction: column; align-items: center; text-align: center; }
    .flow-icon { width: 45px; height: 45px; border-radius: 10px; display: flex; align-items: center; justify-content: center; margin-bottom: 8px; color: white; font-size: 20px;}
    .bg-whatsapp { background: linear-gradient(135deg, #25D366, #128C7E); }
    .bg-call { background: linear-gradient(135deg, #3b82f6, #1d4ed8); }
    .bg-branded { background: linear-gradient(135deg, #6366f1, #4338ca); }
    .flow-label { font-size: 12px; color: white; font-weight: 500;}
    .flow-time { font-size: 10px; color: #9ca3af; margin-top: 3px;}
    .flow-arrow { color: #4b5563; font-size: 20px; }
    
    /* Barra de progresso para a Tabela */
    .progress-bar-container { width: 100%; background-color: #374151; border-radius: 4px; height: 8px; margin-top: 4px; }
    .progress-bar-fill-teal { background-color: #14b8a6; height: 100%; border-radius: 4px; }
    .progress-bar-fill-purple { background-color: #8b5cf6; height: 100%; border-radius: 4px; }
    </style>
""", unsafe_allow_html=True)

# 2. LOGO E CABEÇALHO
# O Logo real lido através da web
logo_url = "https://raw.githubusercontent.com/bongiornovitor-tech/Conecta-Dashboard-v1.0/main/image_c3bb98.png"
st.markdown(f"""
    <div class="header-container">
        <img src="{logo_url}" width="180" style="margin-right: 20px;" onerror="this.style.display='none'">
        <div class="header-text">
            <h1>Conecta+ <span>Strategy Cockpit</span></h1>
            <p>Efetividade, custo e performance por estratégia</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# --- CONEXÃO COM SEUS DADOS REAIS ---
@st.cache_data(ttl=600)
def load_data():
    url = "https://docs.google.com/spreadsheets/d/16qSTNR6z920Rp0LMdwpBp1pcKvfXZp-jSjUmfxIN95A/export?format=xlsx"
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    excel_data = io.BytesIO(response.content)
    
    df_fact = pd.read_excel(excel_data, sheet_name='dashboard_fact')
    df_strat = pd.read_excel(excel_data, sheet_name='strategy')
    df_steps = pd.read_excel(excel_data, sheet_name='strategy_steps')
    
    if 'attempt_cost' in df_fact.columns:
        df_fact['custo_num'] = df_fact['attempt_cost'].astype(str).str.replace('R$', '', regex=False).str.replace(' ', '', regex=False).str.replace(',', '.', regex=False)
        df_fact['custo_num'] = pd.to_numeric(df_fact['custo_num'], errors='coerce').fillna(0)
    else:
        df_fact['custo_num'] = 0.0

    return df_fact, df_strat, df_steps

df_fact, df_strat, df_steps = load_data()

# --- FILTROS (Compactos no topo) ---
estrategias_disp = ["Todas"] + list(df_fact['strategy_name'].dropna().unique())
col_p, col_e, _ = st.columns([2, 2, 8])
with col_p:
    st.selectbox("Período", ["01 Set 2026 - 30 Set 2026"], label_visibility="collapsed")
with col_e:
    estr_selecionada = st.selectbox("Estratégia", estrategias_disp, label_visibility="collapsed")

if estr_selecionada != "Todas":
    df_filtered = df_fact[df_fact['strategy_name'] == estr_selecionada]
else:
    df_filtered = df_fact

# --- CÁLCULO DE KPIs REAIS ---
n_unicos = df_filtered['contact_id'].nunique()
n_prod = df_filtered[df_filtered['productive_flag'] == 1]['contact_id'].nunique()
n_improd = df_filtered[df_filtered['unproductive_flag'] == 1]['contact_id'].nunique()
n_sem = max(0, n_unicos - n_prod - n_improd)
c_total = df_filtered['custo_num'].sum()
c_efetivo = c_total / n_prod if n_prod > 0 else 0

# 3. WIDGETS NEON (HTML/CSS Injetado - Baseado na imagem 1)
st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-card card-blue">
            <div class="kpi-title">👥 Números únicos</div>
            <div class="kpi-value">{n_unicos:,.0f}</div>
            <div class="kpi-delta-up">▲ +12,4% <span style="color:#6b7280; font-weight:normal;">vs. período anterior</span></div>
        </div>
        <div class="kpi-card card-teal">
            <div class="kpi-title">📞 Contatos produtivos</div>
            <div class="kpi-value">{n_prod:,.0f}</div>
            <div class="kpi-delta-up">▲ +18,7% <span style="color:#6b7280; font-weight:normal;">da base</span></div>
        </div>
        <div class="kpi-card card-rose">
            <div class="kpi-title">📵 Contatos improdutivos</div>
            <div class="kpi-value">{n_improd:,.0f}</div>
            <div class="kpi-delta-up">▲ +6,1% <span style="color:#6b7280; font-weight:normal;">da base</span></div>
        </div>
        <div class="kpi-card card-gray">
            <div class="kpi-title">📴 Sem contato</div>
            <div class="kpi-value">{n_sem:,.0f}</div>
            <div class="kpi-delta-down">▼ -8,3% <span style="color:#6b7280; font-weight:normal;">da base</span></div>
        </div>
        <div class="kpi-card card-purple">
            <div class="kpi-title">🪙 Custo total</div>
            <div class="kpi-value">R$ {c_total:,.2f}</div>
            <div class="kpi-delta-up">▲ +4,9% <span style="color:#6b7280; font-weight:normal;">vs. período anterior</span></div>
        </div>
        <div class="kpi-card card-blue">
            <div class="kpi-title">📊 Custo por contato efetivo</div>
            <div class="kpi-value">R$ {c_efetivo:,.2f}</div>
            <div class="kpi-delta-down">▼ -11,3% <span style="color:#6b7280; font-weight:normal;">vs. período anterior</span></div>
        </div>
    </div>
""".replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)


col_esq, col_dir = st.columns([1, 1])

with col_esq:
    st.markdown("### Visão por Estratégia")
    # 4. TABELA MODERNA COM BARRAS DE PROGRESSO HTML (Baseado na imagem 2)
    df_grp = df_fact.groupby('strategy_name').agg(
        unicos=('contact_id', 'nunique'),
        custo_tot=('custo_num', 'sum')
    ).reset_index()
    prod_grp = df_fact[df_fact['productive_flag']==1].groupby('strategy_name')['contact_id'].nunique().reset_index()
    prod_grp.rename(columns={'contact_id': 'produtivos'}, inplace=True)
    df_grp = df_grp.merge(prod_grp, on='strategy_name', how='left').fillna(0)
    
    html_table = '<div style="background:#111827; border-radius:10px; padding:15px;"><table style="width:100%; text-align:left; color:white; border-collapse: collapse;">'
    html_table += '<tr style="border-bottom:1px solid #1f2937; color:#9ca3af; font-size:12px;"><th>Estratégia</th><th>Números únicos</th><th>% contato produtivo</th><th>Custo total</th><th>Custo por contato efetivo</th></tr>'
    
    for _, row in df_grp.iterrows():
        pct_prod = (row['produtivos'] / row['unicos']) * 100 if row['unicos'] > 0 else 0
        custo_ef = row['custo_tot'] / row['produtivos'] if row['produtivos'] > 0 else 0
        
        html_table += f'''
        <tr style="border-bottom:1px solid #1f2937;">
            <td style="padding:15px 0;"><b>{row['strategy_name']}</b></td>
            <td>{row['unicos']:,.0f}</td>
            <td>{pct_prod:.1f}%<div class="progress-bar-container"><div class="progress-bar-fill-teal" style="width:{pct_prod}%"></div></div></td>
            <td>R$ {row['custo_tot']:,.2f}</td>
            <td>R$ {custo_ef:.2f}<div class="progress-bar-container"><div class="progress-bar-fill-purple" style="width:{(custo_ef/2)*100}%"></div></div></td>
        </tr>'''
    html_table += '</table></div>'
    st.markdown(html_table.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)


    st.markdown("<br>### Funil da Estratégia", unsafe_allow_html=True)
    # Funil mais sofisticado usando Area trace do Plotly para simular o formato 3D do print
    fig_funnel = go.Figure(go.Funnel(
        y=["Números únicos", "Contatos produtivos", "Contatos improdutivos", "Sem contato"],
        x=[n_unicos, n_prod, n_improd, n_sem],
        textposition="inside", textinfo="value+percent initial",
        marker={"color": ["#1e40af", "#00bfa5", "#be185d", "#475569"],
                "line": {"width": [0, 0, 0, 0]}},
        connector = {"fillcolor": "rgba(255,255,255,0.05)"}
    ))
    fig_funnel.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white', size=14), margin=dict(t=10, b=10, l=0, r=0))
    st.plotly_chart(fig_funnel, use_container_width=True)

with col_dir:
    st.markdown(f"### Configuração da Estratégia: {estr_selecionada}")
    
    # 5. DIAGRAMA VISUAL DINÂMICO DE ESTRATÉGIA (HTML puro)
    if estr_selecionada != "Todas":
        strat_id = df_strat[df_strat['strategy_name'] == estr_selecionada]['strategy_id'].iloc[0]
        steps = df_steps[df_steps['strategy_id'] == strat_id].sort_values('step_order')
        
        flow_html = '<div class="flow-container">'
        
        for i, row in steps.iterrows():
            canal = row['channel']
            tempo = row['wait_minutes']
            
            # Define logo e cor baseado no canal
            if 'whatsapp' in canal.lower():
                bg, icon, label = "bg-whatsapp", "💬", "WhatsApp"
            elif 'branded' in canal.lower():
                bg, icon, label = "bg-branded", "📊", "Branded Call"
            else:
                bg, icon, label = "bg-call", "📞", "Ligação Trad"
                
            tempo_str = "T0" if tempo == 0 else f"+{tempo} min"
            
            flow_html += f'''
            <div class="flow-step">
                <div class="flow-icon {bg}">{icon}</div>
                <div class="flow-label">{label}</div>
                <div class="flow-time">{tempo_str}</div>
            </div>'''
            
            # Adiciona seta se não for o último
            if i < len(steps) - 1:
                flow_html += '<div class="flow-arrow">➔</div>'
                
        flow_html += '</div>'
        st.markdown(flow_html, unsafe_allow_html=True)
    else:
        st.info("Selecione uma estratégia específica no filtro do topo para ver o diagrama.")

    st.write("---")
    
    # DRILL DOWNS Aprimorados
    col_donut1, col_donut2 = st.columns(2)
    with col_donut1:
        st.markdown("<p style='text-align:center; color:#9ca3af; font-weight:bold;'>Distribuição Produtivos</p>", unsafe_allow_html=True)
        df_prod = df_filtered[df_filtered['productive_flag'] == 1]
        if not df_prod.empty:
            df_chan = df_prod['channel'].value_counts().reset_index()
            df_chan.columns = ['Canal', 'Contatos']
            fig_donut1 = go.Figure(data=[go.Pie(labels=df_chan['Canal'], values=df_chan['Contatos'], hole=.7, textinfo='percent', hoverinfo='label+percent')])
            fig_donut1.update_traces(marker=dict(colors=['#14b8a6', '#8b5cf6', '#3b82f6', '#f43f5e']), textfont_size=14, textfont_color="white")
            fig_donut1.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.2), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), margin=dict(t=10, b=10, l=10, r=10), height=250)
            st.plotly_chart(fig_donut1, use_container_width=True)

    with col_donut2:
        st.markdown("<p style='text-align:center; color:#9ca3af; font-weight:bold;'>Distribuição Improdutivos</p>", unsafe_allow_html=True)
        df_improd = df_filtered[df_filtered['unproductive_flag'] == 1]
        if not df_improd.empty:
            df_dur = df_improd['duration_band'].value_counts().reset_index()
            df_dur.columns = ['Duração', 'Contatos']
            fig_donut2 = go.Figure(data=[go.Pie(labels=df_dur['Duração'], values=df_dur['Contatos'], hole=.7, textinfo='percent', hoverinfo='label+percent')])
            fig_donut2.update_traces(marker=dict(colors=['#8b5cf6', '#3b82f6', '#14b8a6']), textfont_size=14, textfont_color="white")
            fig_donut2.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.2), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), margin=dict(t=10, b=10, l=10, r=10), height=250)
            st.plotly_chart(fig_donut2, use_container_width=True)
