"""Conecta+ Strategy Cockpit — layout responsivo inspirado no mockup.
Execute: streamlit run app.py. Dependências: streamlit, pandas, requests, openpyxl.
A planilha e as abas do app original são preservadas.
CONFIGURAÇÃO DA GRAVAÇÃO (uma vez no Streamlit Cloud):
1. Adicione google-auth ao requirements.txt (junto de streamlit, pandas,
   requests e openpyxl).
2. No Google Cloud, habilite Google Sheets API, crie uma conta de serviço
   e sua chave JSON. Compartilhe APENAS a planilha com o client_email dessa
   chave, como Editor.
3. Em Streamlit > Settings > Secrets, adicione os campos da chave sob
   [gcp_service_account], incluindo type, project_id, private_key_id,
   private_key, client_email, client_id, token_uri. Para private_key use
   uma string TOML multilinha com as quebras de linha da chave JSON.
4. Opcionalmente defina spreadsheet_id nos Secrets. O padrão é a planilha
   original. Não coloque a chave no GitHub.
Configurações > Estratégias cria/edita strategy e strategy_steps numa operação
atômica. Novas estratégias copiam uma linha de cost_parameters do modelo
selecionado. As sequências antigas são adequadas automaticamente na primeira
conexão autenticada: texto deixa de ser etapa independente e fica vinculado ao
WhatsApp Call no campo goal. Nomes, IDs, ANI e status existentes são mantidos.
Configurações > Custos altera cinco células da estratégia em cost_parameters.
Custos históricos e tentativas em dashboard_fact NÃO são reescritos.
Não exige novas dependências além das já utilizadas, incluindo google-auth.
IA: configure [ai] nos Secrets, com provider="gemini", model="gemini-3.8-flash"
e api_key. A chamada REST usa requests; não precisa instalar SDK adicional.
O botão IA analisa agregações sem IDs pessoais. Só chama a API ao clicar Analisar.
Não existe troca automática de modelo/plano; o projeto da chave deve estar no
Free Tier no Google AI Studio se o usuário não quiser cobrança.
"""
import html
import io
import math
import os
import re
import uuid
import json
import hashlib
import time
import unicodedata
from urllib.parse import quote

import pandas as pd
import altair as alt
import requests
import streamlit as st

st.set_page_config(page_title="Conecta+ Strategy Cockpit", page_icon="☎", layout="wide", initial_sidebar_state="expanded")
SHEET_URL = os.getenv("CONECTA_SHEET_URL", "https://docs.google.com/spreadsheets/d/16qSTNR6z920Rp0LMdwpBp1pcKvfXZp-jSjUmfxIN95A/export?format=xlsx")
COLORS = ["#168bff", "#983bff", "#00dcc0", "#8aa8ff", "#f33b91"]
COST_LABELS = {
    "cost_branded_call": "Chamada com identificação da marca (R$/chamada)",
    "cost_whatsapp_template": "Mensagem de consentimento WhatsApp (R$/mensagem)",
    "cost_meta_minute": "Chamada pelo WhatsApp — Meta (R$/minuto)",
    "cost_productive_minute": "Ligação produtiva (R$/minuto)",
    "cost_unproductive_minute": "Ligação improdutiva (R$/minuto)",
}
CHANNELS = {"traditional_call": "Chamada tradicional", "branded_call": "Branded Call", "whatsapp_call": "WhatsApp Call", "whatsapp_text": "WhatsApp texto"}

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
:root {color-scheme:dark; --muted:#a5b8df; --line:#164579; --ink:#f2f5ff;}
.stApp {background:radial-gradient(ellipse at 55% 0%,#061b39 0%,#001126 48%,#000d20 100%); color:var(--ink);}
.stApp, .stApp button, .stApp input {font-family:'Plus Jakarta Sans',sans-serif;}
[data-testid="stHeader"] {background:transparent; height:0;}
[data-testid="stToolbar"], #MainMenu, footer {display:none;}
[data-testid="stMainBlockContainer"], .block-container {max-width:1800px; padding:1.1rem 1.15rem 2rem;}
[data-testid="stVerticalBlock"] {gap:.65rem;}
[data-testid="stMarkdownContainer"] p {margin:0;}
[data-testid="stWidgetLabel"] p {color:#b8cbef;font-size:12px;}
[data-baseweb="select"] > div, [data-testid="stDateInput"] input {background:#021b36!important; color:#e9efff!important; border-color:#2264a7!important;border-radius:7px!important;}
[data-testid="stDateInput"] > div {background:#021b36; border-color:#2264a7;}
[data-testid="stDateInput"] svg, [data-baseweb="select"] svg {fill:#a9caff;}
.cockpit * {box-sizing:border-box;}
.cockpit {font-family:'Plus Jakarta Sans',sans-serif; color:var(--ink); font-size:12px; line-height:1.4;}
.cockpit h1,.cockpit h2,.cockpit h3,.cockpit p {margin:0; padding:0; font-family:inherit; color:inherit;}
.hero {display:grid;grid-template-columns:235px 1fr auto;gap:24px;align-items:center; min-height:88px; border-bottom:1px solid #123a62; padding:0 12px 15px; margin-bottom:2px;position:relative;overflow:hidden;}
.hero:after {content:'';position:absolute;width:360px;height:85px;right:25%;top:8px;border-top:2px solid #5234ee;border-radius:50%;transform:rotate(-12deg);box-shadow:0 -12px 45px #2548fb24;pointer-events:none;}
.brand {font-size:23px;font-weight:800;letter-spacing:-.7px;}.brand span{color:#4461ff}.brand small{display:block;font-size:12px;font-weight:400;color:var(--muted);letter-spacing:0;margin-top:5px;max-width:170px;}
.hero h1{font-size:clamp(24px,2.65vw,44px);font-weight:800;letter-spacing:-1.5px;line-height:1.15;position:relative;z-index:1;}
.hero h1 span{background:linear-gradient(100deg,#078dff,#a236ff);background-clip:text;-webkit-background-clip:text;color:transparent;}
.hero p{font-size:clamp(13px,1.2vw,19px);color:#c3d0f5;margin-top:6px;}
.use-cases{display:flex;gap:22px;}.use-case{display:flex;align-items:center;gap:9px;font-size:11px;}.use-case small{display:block;color:var(--muted);font-size:10px;}
.icon {width:44px;height:44px;display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;border-radius:10px;background:linear-gradient(145deg,var(--accent),color-mix(in srgb,var(--accent) 45%,#001a44));border:1px solid color-mix(in srgb,var(--accent) 70%,white);box-shadow:0 0 22px color-mix(in srgb,var(--accent) 25%,transparent),inset 0 0 12px #ffffff12;color:white;}
.icon svg{width:24px;height:24px;}.use-case .icon{width:32px;height:32px;background:#031732;box-shadow:0 0 12px #2650ff22;}.use-case .icon svg{width:19px;height:19px;}
.kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin:10px 0;}
.kpi,.panel{background:linear-gradient(130deg,#031e3c 0%,#00152b 65%,#041a36 100%);border:1px solid var(--line);border-radius:11px;box-shadow:inset 0 0 18px #0867e80a,0 5px 16px #00000016;min-width:0;overflow:hidden;}
.kpi{padding:12px;position:relative;min-height:122px;}.kpi-head{display:flex;gap:12px;align-items:center;}.kpi-label{font-size:11px;font-weight:600;color:#dce5ff;}.kpi-value{font-size:clamp(20px,1.55vw,28px);font-weight:800;white-space:nowrap;margin-top:4px;letter-spacing:-.5px;}
.kpi-foot{display:flex;justify-content:space-between;align-items:end;gap:4px;margin-top:13px;}.delta{font-size:12px;font-weight:700;}.sub{font-size:10px;color:var(--muted);margin-top:3px;}.spark{width:46%;height:35px;overflow:visible;}
.dashboard{display:grid;grid-template-columns:minmax(0,1.12fr) minmax(0,1fr);gap:10px;}.panel-title{padding:10px 13px 8px;font-size:16px;font-weight:800;border-bottom:1px solid #174475;display:flex;justify-content:space-between;align-items:center;gap:8px;}.panel-body{padding:10px 13px;}.badge{font-size:10px;font-weight:500;white-space:nowrap;color:#00eac6;background:#00373e;border:1px solid #00606a;padding:4px 8px;border-radius:6px;}.tag{font-size:11px;color:#d5e1ff;font-weight:500;background:#061c3a;padding:3px 9px;border:1px solid #164579;border-radius:5px;}
.strategy-table{width:100%;border-collapse:collapse;table-layout:fixed;}.strategy-table th{font-size:10px;font-weight:500;color:#c8d5f4;text-align:left;padding:8px 8px;}.strategy-table td{padding:10px 8px;border-top:1px solid #164579;vertical-align:middle;font-size:12px;}.strategy-table th:first-child{width:34%;}.strategy-table tr.active{background:linear-gradient(90deg,#082571,#061c42);box-shadow:inset 0 0 16px #225cff55;}.strategy-name{font-weight:700;font-size:13px;}.strategy-name:before{content:'○';font-size:19px;color:#acc1ff;margin-right:7px;}.active .strategy-name:before{content:'◉';color:#2993ff;}.strategy-desc{font-size:10px;color:var(--muted);padding-left:23px;margin-top:2px;}.bar{height:7px;background:#103365;border-radius:3px;margin-top:7px;overflow:hidden;}.bar>i{display:block;height:100%;border-radius:3px;background:linear-gradient(90deg,var(--accent),color-mix(in srgb,var(--accent) 65%,white));}
.config-head{display:flex;align-items:center;justify-content:space-between;gap:15px;margin-bottom:12px;}.config-head h3{font-size:22px;font-weight:800;}.meta{display:flex;gap:6px;}.meta>div{background:#031b38;border:1px solid #153e68;border-radius:6px;padding:5px 9px;}.meta small{color:var(--muted);font-size:10px;display:block;}.sequence-label{color:#d4e2ff;font-size:12px;margin-bottom:7px;}.flow{display:flex;gap:8px;align-items:center;margin-bottom:13px;}.step{flex:1;min-width:0;}.step-card{border:1px solid #164579;background:#041e3a;border-radius:7px;display:flex;align-items:center;gap:8px;padding:8px;font-size:10px;min-height:52px;}.step-card .icon{width:31px;height:31px;border-radius:8px;}.step-card svg{width:18px;height:18px;}.arrow{color:#9cbcef;font-size:20px;}.cost-settings{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));border:1px solid #164579;border-radius:6px;overflow:hidden;}.cost-settings>div{text-align:center;padding:6px 3px;background:#041e3b;border-right:1px solid #164579;}.cost-settings small{display:block;font-size:9px;color:var(--muted);min-height:25px;}.cost-settings b{font-size:12px;}
.bottom{display:grid;grid-template-columns:minmax(0,.73fr) minmax(0,1fr) minmax(0,.75fr);gap:10px;margin-top:10px;align-items:stretch;}.funnel-panel{grid-row:span 2;}.funnel{padding:17px 8px 4px;}.funnel-row{display:flex;align-items:center;gap:9px;height:63px;}.funnel-shape{width:77%;display:flex;justify-content:center;}.funnel-layer{height:59px;display:flex;flex-direction:column;align-items:center;justify-content:center;clip-path:polygon(0 0,100% 0,91% 91%,87% 100%,13% 100%,9% 91%);background:linear-gradient(100deg,color-mix(in srgb,var(--accent) 65%,#002060),var(--accent),color-mix(in srgb,var(--accent) 70%,#001342));border-top:3px solid #ffffff44;filter:drop-shadow(0 0 7px var(--accent));font-size:11px;text-align:center;}.funnel-layer b{font-size:20px;line-height:1.2;}.funnel-pct{flex:1;color:#d9e5ff;font-size:12px;position:relative;}.funnel-pct:before{content:'';display:block;width:100%;border-top:1px dashed #789ed4;margin-bottom:3px;}.funnel-note{background:#06203e;border:1px solid #174579;border-radius:7px;padding:9px 11px;margin:12px 3px 3px;color:#b9ccec;font-size:10px;}
.drill-body{display:grid;grid-template-columns:minmax(0,1.3fr) minmax(0,1fr);gap:10px;padding:10px 12px;}.donut-wrap{display:flex;align-items:center;gap:10px;min-width:0;}.donut{width:125px;min-width:95px;max-width:44%;aspect-ratio:1;position:relative;border-radius:50%;background:var(--segments);box-shadow:inset 0 0 18px #ffffff20;}.donut:after{content:'';position:absolute;inset:21%;border-radius:50%;background:#00172e;box-shadow:0 0 10px #0008;}.donut-center{position:absolute;inset:23%;z-index:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;font-size:9px;line-height:1.2;}.donut-center b{font-size:16px;}.legend{flex:1;min-width:0;}.legend-row{display:flex;gap:5px;align-items:center;font-size:9px;margin:7px 0;}.dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;box-shadow:inset 0 0 3px #fff7;}.legend-name{flex:1;overflow-wrap:anywhere;}.legend-row b{font-size:9px;white-space:nowrap;}.mini-title{font-size:10px;color:#c8d8f7;margin-bottom:8px;}.duration-table{width:100%;border-collapse:collapse;font-size:9px;background:#041e3a;}.duration-table td,.duration-table th{padding:6px 5px;border:1px solid #153c65;text-align:left;font-weight:400;}.duration-table th{color:var(--muted);}.duration-table td:last-child{text-align:right;white-space:nowrap;}
.cost-panel{grid-column:2/4;}.cost-body{display:grid;grid-template-columns:1.25fr 1fr;gap:14px;padding:9px;}.cost-cards{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:6px;}.mini-kpi{padding:15px 12px;min-height:110px;container-type:inline-size;display:flex;flex-direction:column;justify-content:flex-start;align-items:flex-start;background:#05223f;border:1px solid #163e65;border-radius:7px;}.mini-kpi small{font-size:11px;color:#c5d6f7;display:block;line-height:1.4;min-height:3.6em;width:100%;}.mini-kpi b{font-size:min(32px,var(--value-size,16cqw));font-weight:800;letter-spacing:-.5px;display:block;margin-top:8px;line-height:1.2;white-space:nowrap;overflow-wrap:normal;word-break:normal;}.stacked{display:flex;height:23px;border-radius:6px;overflow:hidden;}.stacked span{display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;min-width:0;}.cost-legend .legend-row b{font-size:13px;}.cost-legend{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:3px 12px;margin-top:6px;}.empty{color:var(--muted);padding:25px 10px;text-align:center;font-size:12px;}.caption{color:#7f99bc;font-size:10px;margin-top:10px;}
@media(min-width:1600px){.kpi{min-height:130px;}.panel-title{font-size:18px;}.strategy-table td{padding:13px 8px;}.drill-body{padding:13px;}.donut{width:150px;}.funnel-row{height:69px;}.funnel-layer{height:65px;}}
