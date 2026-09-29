"""Conecta+ Strategy Cockpit — layout responsivo inspirado no mockup.
Execute: streamlit run app.py. Dependências: streamlit, pandas, requests, openpyxl.
A planilha e as três abas do app original são preservadas.
"""
import html
import io
import math
import os
import re

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Conecta+ Strategy Cockpit", page_icon="☎", layout="wide", initial_sidebar_state="collapsed")
SHEET_URL = os.getenv("CONECTA_SHEET_URL", "https://docs.google.com/spreadsheets/d/16qSTNR6z920Rp0LMdwpBp1pcKvfXZp-jSjUmfxIN95A/export?format=xlsx")
COLORS = ["#168bff", "#983bff", "#00dcc0", "#8aa8ff", "#f33b91"]
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
.config-head{display:flex;align-items:center;justify-content:space-between;gap:15px;margin-bottom:12px;}.config-head h3{font-size:22px;font-weight:800;}.meta{display:flex;gap:6px;}.meta>div{background:#031b38;border:1px solid #153e68;border-radius:6px;padding:5px 9px;}.meta small{color:var(--muted);font-size:10px;display:block;}.sequence-label{color:#d4e2ff;font-size:12px;margin-bottom:7px;}.flow{display:flex;gap:8px;align-items:center;margin-bottom:13px;}.step{flex:1;min-width:0;}.step-card{border:1px solid #164579;background:#041e3a;border-radius:7px;display:flex;align-items:center;gap:8px;padding:8px;font-size:10px;min-height:52px;}.step-card .icon{width:31px;height:31px;border-radius:8px;}.step-card svg{width:18px;height:18px;}.step-time{text-align:center;font-size:10px;color:var(--muted);margin-top:5px;}.arrow{color:#9cbcef;font-size:20px;}.cost-settings{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));border:1px solid #164579;border-radius:6px;overflow:hidden;}.cost-settings>div{text-align:center;padding:6px 3px;background:#041e3b;border-right:1px solid #164579;}.cost-settings small{display:block;font-size:9px;color:var(--muted);min-height:25px;}.cost-settings b{font-size:12px;}
.bottom{display:grid;grid-template-columns:minmax(0,.73fr) minmax(0,1fr) minmax(0,.75fr);gap:10px;margin-top:10px;align-items:stretch;}.funnel-panel{grid-row:span 2;}.funnel{padding:17px 8px 4px;}.funnel-row{display:flex;align-items:center;gap:9px;height:63px;}.funnel-shape{width:77%;display:flex;justify-content:center;}.funnel-layer{height:59px;display:flex;flex-direction:column;align-items:center;justify-content:center;clip-path:polygon(0 0,100% 0,91% 91%,87% 100%,13% 100%,9% 91%);background:linear-gradient(100deg,color-mix(in srgb,var(--accent) 65%,#002060),var(--accent),color-mix(in srgb,var(--accent) 70%,#001342));border-top:3px solid #ffffff44;filter:drop-shadow(0 0 7px var(--accent));font-size:11px;text-align:center;}.funnel-layer b{font-size:20px;line-height:1.2;}.funnel-pct{flex:1;color:#d9e5ff;font-size:12px;position:relative;}.funnel-pct:before{content:'';display:block;width:100%;border-top:1px dashed #789ed4;margin-bottom:3px;}.funnel-note{background:#06203e;border:1px solid #174579;border-radius:7px;padding:9px 11px;margin:12px 3px 3px;color:#b9ccec;font-size:10px;}
.drill-body{display:grid;grid-template-columns:minmax(0,1.3fr) minmax(0,1fr);gap:10px;padding:10px 12px;}.donut-wrap{display:flex;align-items:center;gap:10px;min-width:0;}.donut{width:125px;min-width:95px;max-width:44%;aspect-ratio:1;position:relative;border-radius:50%;background:var(--segments);box-shadow:inset 0 0 18px #ffffff20;}.donut:after{content:'';position:absolute;inset:21%;border-radius:50%;background:#00172e;box-shadow:0 0 10px #0008;}.donut-center{position:absolute;inset:23%;z-index:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;font-size:9px;line-height:1.2;}.donut-center b{font-size:16px;}.legend{flex:1;min-width:0;}.legend-row{display:flex;gap:5px;align-items:center;font-size:9px;margin:7px 0;}.dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;box-shadow:inset 0 0 3px #fff7;}.legend-name{flex:1;overflow-wrap:anywhere;}.legend-row b{font-size:9px;white-space:nowrap;}.mini-title{font-size:10px;color:#c8d8f7;margin-bottom:8px;}.duration-table{width:100%;border-collapse:collapse;font-size:9px;background:#041e3a;}.duration-table td,.duration-table th{padding:6px 5px;border:1px solid #153c65;text-align:left;font-weight:400;}.duration-table th{color:var(--muted);}.duration-table td:last-child{text-align:right;white-space:nowrap;}
.cost-panel{grid-column:2/4;}.cost-body{display:grid;grid-template-columns:1.25fr 1fr;gap:14px;padding:9px;}.cost-cards{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:6px;}.mini-kpi{padding:9px;background:#05223f;border:1px solid #163e65;border-radius:7px;}.mini-kpi small{font-size:9px;color:#c5d6f7;display:block;}.mini-kpi b{font-size:17px;display:block;margin-top:6px;white-space:nowrap;}.stacked{display:flex;height:23px;border-radius:6px;overflow:hidden;}.stacked span{display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;min-width:0;}.cost-legend{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:3px 12px;margin-top:6px;}.empty{color:var(--muted);padding:25px 10px;text-align:center;font-size:12px;}.caption{color:#7f99bc;font-size:10px;margin-top:10px;}
@media(min-width:1600px){.kpi{min-height:130px;}.panel-title{font-size:18px;}.strategy-table td{padding:13px 8px;}.drill-body{padding:13px;}.donut{width:150px;}.funnel-row{height:69px;}.funnel-layer{height:65px;}}
@media(max-width:1250px){.hero{grid-template-columns:180px 1fr;}.use-cases{display:none;}.kpis{grid-template-columns:repeat(3,minmax(0,1fr));}.kpi-value{font-size:25px;}.bottom{grid-template-columns:minmax(0,.85fr) minmax(0,1.3fr);}.funnel-panel{grid-row:span 2;}.cost-panel{grid-column:1/3;}.drill-body{grid-template-columns:1.2fr 1fr;}.config-head{flex-wrap:wrap;}.meta{width:100%;}.meta>div{flex:1;}.flow{gap:5px;}.step-card{flex-direction:column;text-align:center;padding:7px 3px;}.cost-body{grid-template-columns:1fr 1fr;}}
@media(max-width:850px){.dashboard{grid-template-columns:1fr;}.hero{gap:14px;grid-template-columns:150px 1fr;}.hero h1{font-size:28px;}.brand{font-size:19px;}.bottom{grid-template-columns:1fr 1fr;}.funnel-panel{grid-row:auto;grid-column:1/3;}.funnel{max-width:500px;margin:auto;}.drill-body{grid-template-columns:1fr;}.cost-body{grid-template-columns:1fr;}.donut{width:125px;}.step-card{flex-direction:row;text-align:left;}.panel-title{font-size:14px;}}
@media(max-width:560px){[data-testid="stMainBlockContainer"],.block-container{padding:.7rem .6rem 1.3rem;}.hero{grid-template-columns:1fr;gap:12px;padding:2px 4px 16px;}.brand small{display:none;}.hero h1{font-size:28px;letter-spacing:-1px;}.hero p{font-size:12px;}.kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:7px;}.kpi{padding:10px;min-height:120px;}.kpi-head{gap:7px;}.kpi .icon{width:32px;height:32px;border-radius:8px;}.kpi .icon svg{width:20px;height:20px;}.kpi-label{font-size:9px;}.kpi-value{font-size:clamp(15px,4.6vw,19px);white-space:normal;overflow-wrap:anywhere;}.kpi-head>div{min-width:0;}.kpi-foot{margin-top:12px;}.delta{font-size:10px;}.sub{font-size:9px;}.bottom{grid-template-columns:1fr;}.funnel-panel,.cost-panel{grid-column:auto;}.drill-body{grid-template-columns:1fr 1fr;}.donut-wrap{flex-direction:column;}.donut{width:130px;max-width:100%;}.legend{width:100%;}.config-head h3{font-size:20px;}.flow{flex-wrap:wrap;}.step{flex:1 1 40%;}.arrow{display:none;}.step-card{font-size:11px;}.cost-settings{grid-template-columns:repeat(2,minmax(0,1fr));gap:1px;border:0;}.cost-settings>div{border:1px solid #164579;border-radius:5px;}.cost-settings small{min-height:0;margin-bottom:3px;}.strategy-table th{font-size:8px;padding:7px 4px;}.strategy-table td{font-size:10px;padding:10px 4px;}.strategy-table th:first-child{width:31%;}.strategy-name{font-size:10px;}.strategy-name:before{display:none;}.strategy-desc{padding-left:0;font-size:8px;}.cost-cards{gap:4px;}.mini-kpi{padding:7px 5px;}.mini-kpi b{font-size:14px;}.panel-title{font-size:13px;padding:9px;}.tag{font-size:9px;}.badge{font-size:8px;}.panel-body{padding:10px;}}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def esc(value):
    return html.escape(str(value)) if pd.notna(value) else "—"


def br(value, decimals=0):
    return f"{value:,.{decimals}f}".replace(",", "_").replace(".", ",").replace("_", ".")


def money(value):
    return "R$ " + br(value, 2) if value is not None else "—"


def icon(kind, color="#168bff"):
    paths = {
        "users": '<circle cx="9" cy="7" r="3"/><path d="M3 20v-4a6 6 0 0 1 12 0v4M16 4a3 3 0 0 1 0 6M19 20v-4a5 5 0 0 0-3-4"/>',
        "phone": '<path d="M5 3h4l2 5-3 2a16 16 0 0 0 6 6l2-3 5 2v4c0 2-2 3-4 2A22 22 0 0 1 3 7c-1-2 0-4 2-4Z"/>',
        "off": '<path d="m3 3 18 18M5 3h4l2 5-2 1M8 12a16 16 0 0 0 6 4l2-3 5 2v4c0 2-2 3-4 2A22 22 0 0 1 3 7"/>',
        "bars": '<path d="M4 21V13h3v8ZM11 21V7h3v14ZM18 21V3h3v18Z"/>',
        "coin": '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 4 16 4 16 0V5M4 10c0 4 16 4 16 0M4 15c0 4 16 4 16 0"/>',
        "chat": '<path d="M21 11a9 9 0 0 1-13 8l-5 2 2-5A9 9 0 1 1 21 11Z"/><path d="M8 8h8M8 12h5"/>',
        "bag": '<rect x="4" y="7" width="16" height="14" rx="2"/><path d="M8 9V6a4 4 0 0 1 8 0v3"/>',
    }
    return f'<span class="icon" style="--accent:{color}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{paths.get(kind, paths["phone"])}</svg></span>'


def numeric(value):
    if pd.isna(value):
        return 0.0
    if isinstance(value, (float, int)):
        return float(value)
    text = re.sub(r"[^\d.,-]", "", str(value))
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        result = float(text)
        return result if math.isfinite(result) else 0.0
    except ValueError:
        return 0.0


@st.cache_data(ttl=600, show_spinner=False)
def load_data():
    response = requests.get(SHEET_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    response.raise_for_status()
    sheets = pd.read_excel(io.BytesIO(response.content), sheet_name=["dashboard_fact", "strategy", "strategy_steps"])
    fact = sheets["dashboard_fact"].copy()
    required = {"contact_id", "strategy_name", "channel", "productive_flag", "unproductive_flag", "attempt_cost"}
    missing = required.difference(fact.columns)
    if missing:
        raise ValueError("Colunas ausentes em dashboard_fact: " + ", ".join(sorted(missing)))
    fact = fact[fact["contact_id"].notna()].copy()
    fact["custo_num"] = fact["attempt_cost"].map(numeric)
    for column in ["productive_flag", "unproductive_flag"]:
        fact[column] = pd.to_numeric(fact[column], errors="coerce").fillna(0).astype(int)
    # Aceita os nomes usuais da base e não inventa uma data fixa.
    candidates = ["date", "attempt_date", "attempt_datetime", "attempt_timestamp", "attempt_at", "data", "contact_date", "created_at", "timestamp", "event_date", "data_tentativa"]
    date_column = next((c for c in candidates if c in fact), None)
    if date_column:
        fact["_date"] = pd.to_datetime(fact[date_column], errors="coerce", dayfirst=False, utc=True).dt.tz_convert("America/Sao_Paulo").dt.tz_localize(None).dt.normalize() if "timestamp" in date_column or "datetime" in date_column else pd.to_datetime(fact[date_column], errors="coerce").dt.normalize()
    else:
        fact["_date"] = pd.NaT
    return fact, sheets["strategy"], sheets["strategy_steps"]


def outcome_frames(df):
    """Uma pessoa pertence a um único resultado: produtivo > improdutivo > sem contato.
    Atribuição ao canal da primeira tentativa que alcançou esse resultado.
    """
    sort_column = next((c for c in ["attempt_timestamp", "attempt_datetime", "_date"] if c in df), None)
    ordered = df.sort_values(sort_column, kind="stable") if sort_column else df
    productive = ordered[ordered["productive_flag"].eq(1)].drop_duplicates("contact_id")
    unproductive = ordered[ordered["unproductive_flag"].eq(1) & ~ordered["contact_id"].isin(productive["contact_id"])].drop_duplicates("contact_id")
    no_contact = ordered[~ordered["contact_id"].isin(pd.concat([productive["contact_id"], unproductive["contact_id"]]))].drop_duplicates("contact_id")
    return productive, unproductive, no_contact


def metrics(df):
    productive, unproductive, no_contact = outcome_frames(df)
    total = df["custo_num"].sum()
    return [df["contact_id"].nunique(), len(productive), len(unproductive), len(no_contact), total, total / len(productive) if len(productive) else None]


def channel_name(value):
    return CHANNELS.get(str(value), str(value).replace("_", " ").title())


def spark(values, color, ident):
    vals = [float(v or 0) for v in values]
    if len(vals) < 2:
        return ""
    low, high = min(vals), max(vals)
    points = " ".join(f"{i*100/(len(vals)-1):.1f},{30-(v-low)/(high-low or 1)*23:.1f}" for i, v in enumerate(vals))
    return f'<svg class="spark" viewBox="0 0 100 36" aria-label="Tendência diária"><defs><linearGradient id="{ident}" x1="0" y1="0" x2="0" y2="1"><stop stop-color="{color}" stop-opacity=".4"/><stop offset="1" stop-color="{color}" stop-opacity="0"/></linearGradient></defs><polygon points="0,36 {points} 100,36" fill="url(#{ident})"/><polyline points="{points}" fill="none" stroke="{color}" stroke-width="1.8"/></svg>'


def panel(title, body, extra="", cls=""):
    return f'<section class="panel {cls}"><div class="panel-title"><span>{title}</span>{extra}</div>{body}</section>'


def donut(series, label, palette=COLORS):
    series = series[series > 0]
    if series.empty:
        return '<div class="empty">Sem contatos neste período.</div>'
    total = series.sum()
    start, segments, legend = 0, [], []
    for i, (name, value) in enumerate(series.items()):
        pct = value / total * 100
        color = palette[i % len(palette)]
        segments.append(f"{color} {start:.3f}% {start+pct:.3f}%")
        start += pct
        legend.append(f'<div class="legend-row"><i class="dot" style="background:{color}"></i><span class="legend-name">{esc(name)}</span><b>{br(pct,1)}%</b></div>')
    return f'<div class="donut-wrap"><div class="donut" role="img" aria-label="{esc(label)}: {br(total)}" style="--segments:conic-gradient({",".join(segments)})"><div class="donut-center"><b>{br(total)}</b><span>{label}</span></div></div><div class="legend">{"".join(legend)}</div></div>'


def duration_table(df):
    col = next((c for c in ["duration_seconds", "duration_sec", "call_duration_seconds", "duration", "duration_s", "talk_time_seconds"] if c in df), None)
    if col is None or df.empty:
        return '<div class="mini-title">Duração média por canal</div><div class="empty">Duração em segundos não disponível na base.</div>'
    copy = df.assign(_duration=pd.to_numeric(df[col], errors="coerce"))
    rows = []
    for channel, seconds in copy.groupby("channel")["_duration"].mean().items():
        if pd.notna(seconds):
            sec = int(round(seconds))
            duration = f"{sec//60} min {sec%60:02d} seg" if sec >= 60 else f"{sec} seg"
            rows.append(f'<tr><td>{esc(channel_name(channel))}</td><td>{duration}</td></tr>')
    return '<div class="mini-title">Duração média por canal</div><table class="duration-table"><thead><tr><th>Canal</th><th>Duração média</th></tr></thead><tbody>' + "".join(rows) + '</tbody></table>'


hero = '<div class="cockpit"><div class="hero"><div class="brand">Nuveto <span>| Conecta+</span><small>Inteligência que conecta<br>os seus resultados.</small></div><div><h1>Conecta+ <span>Strategy Cockpit</span></h1><p>Efetividade, custo e performance por estratégia</p></div><div class="use-cases">' + "".join(f'<div class="use-case">{icon(kind,"#4759ff")}<div><b>{name}</b><small>{desc}</small></div></div>' for kind, name, desc in [("chat", "Marketing", "Mais oportunidades"), ("bars", "Vendas", "Mais conversões"), ("bag", "Cobrança", "Mais resultados")]) + '</div></div></div>'
st.markdown(hero, unsafe_allow_html=True)
try:
    with st.spinner("Carregando indicadores…"):
        df_fact, df_strat, df_steps = load_data()
except Exception as exc:
    st.error("Não foi possível carregar a planilha. Verifique o compartilhamento, as abas e a conexão.")
    with st.expander("Detalhes do carregamento"):
        st.code(str(exc))
    st.stop()

if df_fact.empty:
    st.info("A aba dashboard_fact ainda não contém tentativas.")
    st.stop()

names = list(df_fact["strategy_name"].dropna().unique())
names.sort(key=lambda n: (n != "Custo Eficiente", n != "Máximo Contato", str(n)))
valid_dates = df_fact["_date"].dropna()
f1, f2, f3 = st.columns([1.1, 1.1, 2.8])
with f1:
    if valid_dates.empty:
        st.selectbox("Período", ["Todo o período disponível"], disabled=True)
        date_start = date_end = None
    else:
        selection = st.date_input("Período", value=(valid_dates.min().date(), valid_dates.max().date()), min_value=valid_dates.min().date(), max_value=valid_dates.max().date(), format="DD/MM/YYYY")
        if len(selection) != 2:
            st.info("Escolha a data final do período.")
            st.stop()
        date_start, date_end = map(pd.Timestamp, selection)
with f2:
    selected = st.selectbox("Estratégia", ["Todas"] + names)
with f3:
    # Escolha separada para os painéis inferiores quando os KPIs estão consolidados.
    detail_name = st.selectbox("Detalhar estratégia", names, index=names.index(selected) if selected in names else 0, disabled=selected != "Todas")
    if selected != "Todas":
        detail_name = selected

period_df = df_fact if date_start is None else df_fact[df_fact["_date"].between(date_start, date_end)]
filtered = period_df if selected == "Todas" else period_df[period_df["strategy_name"].eq(selected)]
detail = period_df[period_df["strategy_name"].eq(detail_name)]
current = metrics(filtered)
previous = None
if date_start is not None:
    days = (date_end - date_start).days + 1
    previous_df = df_fact[df_fact["_date"].between(date_start-pd.Timedelta(days=days), date_start-pd.Timedelta(days=1))]
    if selected != "Todas":
        previous_df = previous_df[previous_df["strategy_name"].eq(selected)]
    if not previous_df.empty:
        previous = metrics(previous_df)

# Tendências e variações reais. Não exibimos os deltas fictícios do mockup.
daily = []
if not filtered["_date"].dropna().empty:
    for day in pd.date_range(filtered["_date"].min(), filtered["_date"].max()):
        daily.append(metrics(filtered[filtered["_date"].eq(day)]))
kpi_html = []
labels = ["Números únicos", "Contatos produtivos", "Contatos improdutivos", "Sem contato", "Custo total", "Custo por contato efetivo"]
for i, (label, kind, color) in enumerate(zip(labels, ["users", "phone", "off", "off", "coin", "bars"], ["#168bff", "#00cfb2", "#ee3585", "#6389c5", "#853aff", "#168bff"])):
    value = money(current[i]) if i >= 4 else br(current[i])
    desc = f"{br(current[i]/current[0]*100 if current[0] else 0,1)}% da base" if i in [1,2,3] else "no período selecionado"
    delta, delta_color = "", "#a5b8df"
    if previous and previous[i] and current[i] is not None:
        change = (current[i] / previous[i]-1)*100
        good = change >= 0 if i in [0,1] else change <= 0
        delta_color = "#00dcc0" if good else "#f33b91"
        delta = f'{"▲" if change >= 0 else "▼"} {"+" if change >= 0 else ""}{br(change,1)}%'
        desc = "vs. período anterior"
    trend = spark([row[i] for row in daily], color, f"spark-{i}")
    kpi_html.append(f'<article class="kpi"><div class="kpi-head">{icon(kind,color)}<div><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div></div><div class="kpi-foot"><div><div class="delta" style="color:{delta_color}">{delta or "&nbsp;"}</div><div class="sub">{desc}</div></div>{trend}</div></article>')

summary = [(name, metrics(period_df[period_df["strategy_name"].eq(name)])) for name in names]
max_unique = max([m[0] for _, m in summary] + [1])
max_cost = max([m[4] for _, m in summary] + [1])
max_unit = max([m[5] or 0 for _, m in summary] + [1])
rows = []
for name, m in summary:
    meta = df_strat[df_strat["strategy_name"].eq(name)]
    objective = meta.iloc[0].get("objective", "") if not meta.empty else ""
    vals = [(br(m[0]),m[0]/max_unique*100,"#00acff"),(br(m[1]/m[0]*100 if m[0] else 0,1)+"%",m[1]/m[0]*100 if m[0] else 0,"#00d6b4"),(money(m[4]),m[4]/max_cost*100,"#477dff"),(money(m[5]),(m[5] or 0)/max_unit*100,"#b659ff")]
    cells = "".join(f'<td>{value}<div class="bar" style="--accent:{color}"><i style="width:{pct:.2f}%"></i></div></td>' for value,pct,color in vals)
    rows.append(f'<tr class="{"active" if name==detail_name else ""}"><td><div class="strategy-name">{esc(name)}</div><div class="strategy-desc">{esc(objective)}</div></td>{cells}</tr>')
strategy_panel = panel("Visão por Estratégia", '<div class="panel-body" style="padding:0 6px 6px"><table class="strategy-table"><thead><tr><th>Estratégia</th><th>Números únicos</th><th>% contato produtivo</th><th>Custo total</th><th>Custo por contato efetivo</th></tr></thead><tbody>'+"".join(rows)+'</tbody></table></div>')
meta = df_strat[df_strat["strategy_name"].eq(detail_name)]
record = meta.iloc[0] if not meta.empty else pd.Series(dtype=object)
steps = df_steps[df_steps["strategy_id"].eq(record.get("strategy_id"))].sort_values("step_order") if "strategy_id" in df_steps and "step_order" in df_steps else pd.DataFrame()
flow = []
for pos, (_, row) in enumerate(steps.iterrows()):
    channel = str(row.get("channel", ""))
    color = "#00cdb2" if "whatsapp" in channel.lower() else "#168bff"
    kind = "chat" if "text" in channel.lower() else "bars" if "branded" in channel.lower() else "phone"
    delay = next((row[c] for c in ["delay_minutes", "wait_minutes", "retry_delay_minutes", "delay_min"] if c in row and pd.notna(row[c])), None)
    timing = f"+{br(float(delay))} min" if delay is not None else f"Etapa {pos+1}"
    flow.append(f'<div class="step"><div class="step-card">{icon(kind,color)}<span>{esc(channel_name(channel))}</span></div><div class="step-time">{timing}</div></div>')
flow_html = '<div class="flow">'+ '<span class="arrow">→</span>'.join(flow) + '</div>' if flow else '<div class="empty">Sequência não cadastrada.</div>'
# Tarifas só aparecem quando fornecidas na aba strategy, nunca como valores fictícios.
cost_fields = [("Custo Branded Call",["branded_call_cost", "cost_branded_call"]),("Custo template WhatsApp",["whatsapp_template_cost", "template_cost", "cost_whatsapp_template"]),("Custo minuto Meta",["meta_minute_cost", "cost_meta_minute"]),("Custo minuto produtivo",["productive_minute_cost", "cost_productive_minute"]),("Custo minuto improdutivo",["unproductive_minute_cost", "cost_unproductive_minute"])]
settings = []
for label, candidates in cost_fields:
    tariff = next((record[c] for c in candidates if c in record and pd.notna(record[c])), None)
    if tariff is None:
        for field in candidates:
            if field in detail and not detail[field].dropna().empty:
                rates = detail[field].dropna().map(numeric).unique()
                if len(rates) == 1:
                    tariff = rates[0]
                break
    settings.append(f'<div><small>{label}</small><b>{money(numeric(tariff)) if tariff is not None else "—"}</b></div>')
ani = next((record[c] for c in ["ani", "caller_id", "bina"] if c in record and pd.notna(record[c])), "Não informado")
config_body = f'<div class="panel-body"><div class="config-head"><h3>{esc(detail_name)}</h3><div class="meta"><div><small>ANI</small>{esc(ani)}</div><div><small>Objetivo central</small>{esc(record.get("objective","—"))}</div></div></div><div class="sequence-label">Sequência de abordagem</div>{flow_html}<div class="sequence-label">Custos configurados da estratégia</div><div class="cost-settings">{"".join(settings)}</div></div>'
config_panel = panel("Configuração da Estratégia Selecionada", config_body, '<span class="badge">● Estratégia selecionada</span>')

m = metrics(detail)
prod, improd, no_contact = outcome_frames(detail)
funnel_rows = []
for value,label,width,color in [(m[0],"Números únicos",100,"#087aff"),(m[1],"Contatos produtivos",81,"#00bffc"),(m[2],"Contatos improdutivos",63,"#e438ad"),(m[3],"Sem contato",47,"#8b9cae")]:
    pct = value/m[0]*100 if m[0] else 0
    funnel_rows.append(f'<div class="funnel-row"><div class="funnel-shape"><div class="funnel-layer" style="width:{width}%;--accent:{color}"><span>{label}</span><b>{br(value)}</b></div></div><div class="funnel-pct">{br(pct,1)}%</div></div>')
funnel_body = '<div class="panel-body"><div class="funnel">'+"".join(funnel_rows)+'</div><div class="funnel-note">Resultados únicos por pessoa: produtivo tem prioridade sobre improdutivo. Sem contato representa os demais números da base filtrada.</div></div>'
funnel_panel = panel("Funil da Estratégia Selecionada", funnel_body, f'<span class="tag">{esc(detail_name)}</span>', "funnel-panel")
prod_series = prod.groupby("channel")["contact_id"].nunique().sort_values(ascending=False)
prod_series.index = prod_series.index.map(channel_name)
prod_body = '<div class="drill-body"><div><div class="mini-title">Distribuição por canal</div>'+donut(prod_series,"contatos<br>produtivos")+'</div><div>'+duration_table(prod)+'</div></div>'
prod_panel = panel("Drill down — Contatos Produtivos",prod_body,f'<span class="tag">Total: {br(m[1])}</span>')
improd_series = improd.groupby("duration_band")["contact_id"].nunique() if "duration_band" in improd else pd.Series(dtype=float)
improd_body = '<div class="drill-body"><div><div class="mini-title">Distribuição por faixa de duração</div>'+donut(improd_series,"contatos<br>improdutivos",["#168bff","#983bff","#a6a4ff","#00dcc0"])+'</div><div>'+duration_table(improd)+'</div></div>'
improd_panel = panel("Drill down — Contatos Improdutivos",improd_body,f'<span class="tag">Total: {br(m[2])}</span>')
cost_series = detail.groupby("channel")["custo_num"].sum().sort_values(ascending=False)
segments, cost_legend = [], []
for i,(channel,value) in enumerate(cost_series.items()):
    pct = value/m[4]*100 if m[4] else 0
    color = COLORS[i % len(COLORS)]
    segments.append(f'<span title="{esc(channel_name(channel))}: {money(value)}" style="width:{pct:.3f}%;background:{color}">{br(pct,1)+"%" if pct>=10 else ""}</span>')
    cost_legend.append(f'<div class="legend-row"><i class="dot" style="background:{color}"></i><span class="legend-name">{esc(channel_name(channel))}<br><b>{money(value)}</b></span></div>')
mini_cards = "".join(f'<div class="mini-kpi"><small>{label}</small><b>{money(value)}</b></div>' for label,value in [("Custo por contato efetivo",m[5]),("Custo por contato improdutivo",m[4]/m[2] if m[2] else None),("Custo total da estratégia",m[4])])
cost_body = '<div class="cost-body"><div class="cost-cards">'+mini_cards+'</div><div><div class="mini-title">Composição do custo por canal</div><div class="stacked">'+"".join(segments)+'</div><div class="cost-legend">'+"".join(cost_legend)+'</div></div></div>'
cost_panel = panel("Custos da Estratégia no Período",cost_body,f'<span class="tag">{esc(detail_name)}</span>',"cost-panel")
output = '<div class="cockpit"><div class="kpis">'+"".join(kpi_html)+'</div><div class="dashboard">'+strategy_panel+config_panel+'</div><div class="bottom">'+funnel_panel+prod_panel+improd_panel+cost_panel+'</div><div class="caption">Fonte: dashboard_fact · Custos incluem todas as tentativas do período. Distribuições por contato único. Variações exibidas somente com histórico comparável. Tarifas ausentes aparecem como —.</div></div>'
# Uma única árvore HTML mantém o grid coeso e evita tags abertas entre blocos Streamlit.
st.markdown(output, unsafe_allow_html=True)
