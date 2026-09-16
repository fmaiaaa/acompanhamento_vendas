# -*- coding: utf-8 -*-
"""
Ferramenta para cálculo da representatividade diária dos indicadores.
Extrai 36 meses de histórico do Salesforce, roda regressões OLS isoladas 
(dia da semana, dia do mês, mês e tendência linear por ano) e converte o volume aditivo 
esperado na participação percentual de cada dia dentro do mês alvo.
Busca metas no Google Sheets, converte reversamente para todos os indicadores
e plota comparativos Projetado x Realizado com Intervalos de Confiança.
"""
import os
import io
import time
import base64
import html
import calendar
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from datetime import date, timedelta

# -----------------------------------------------------------------------------
# Constantes Base e Configurações de Design
# -----------------------------------------------------------------------------
FUNIL_ETAPAS = ("agendamentos", "visitas", "pastas", "pastas_aprovadas", "vendas")
FUNIL_LABELS = {
    "agendamentos": "Agendamentos",
    "visitas": "Visitas",
    "pastas": "Pastas",
    "pastas_aprovadas": "Pastas Aprovadas",
    "vendas": "Vendas",
}

DIAS_SEMANA_PT = {
    0: "segunda", 1: "terça", 2: "quarta", 3: "quinta",
    4: "sexta", 5: "sábado", 6: "domingo"
}
MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"

COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_TEXTO_PRETO = "#000000"
COR_BORDA = "#eef2f6"
COR_INPUT_BG = "#f0f2f6"
RGB_AZUL_CSS = "4, 66, 143"
RGB_VERMELHO_CSS = "203, 9, 53"

# -----------------------------------------------------------------------------
# Funções de Estilização (Design Direcional)
# -----------------------------------------------------------------------------
def _resolver_png_raiz(nome: str) -> Path | None:
    for base in (_DIR_APP, _DIR_APP.parent):
        p = base / nome
        if p.is_file(): return p
    return None

def _resolver_imagem_fundo_local(nome: str) -> Path | None:
    for base in (_DIR_APP, _DIR_APP.parent):
        for ext in (".jpg", ".jpeg", ".JPG", ".JPEG", ".png", ".PNG"):
            stem = Path(nome).stem
            p = base / f"{stem}{ext}"
            if p.is_file(): return p
        p = base / nome
        if p.is_file(): return p
    return None

def _css_url_fundo_cadastro() -> str:
    p = _resolver_imagem_fundo_local(FUNDO_CADASTRO_ARQUIVO)
    if p and p.is_file():
        try:
            raw = p.read_bytes()
            suf = p.suffix.lower()
            mime = "image/jpeg" if suf in (".jpg", ".jpeg") else "image/png"
            b64 = base64.b64encode(raw).decode("ascii")
            return f"data:{mime};base64,{b64}"
        except OSError: pass
    return "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?auto=format&fit=crop&w=1920&q=80"

def _logo_arquivo_local() -> str | None:
    p_topo = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
    if p_topo: return str(p_topo)
    for name in ("logo_direcional.png", "logo_direcional.jpg", "logo.png"):
        p = _DIR_APP / "assets" / name
        if p.is_file(): return str(p)
    return None

def _exibir_logo_topo() -> None:
    path = _logo_arquivo_local()
    try:
        if path:
            ext = Path(path).suffix.lower().lstrip(".")
            mime = "image/png" if ext == "png" else "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            st.markdown(f'<div class="ficha-logo-wrap"><img src="data:{mime};base64,{b64}" alt="Direcional" /></div>', unsafe_allow_html=True)
            return
        u = "https://logodownload.org/wp-content/uploads/2021/04/direcional-engenharia-logo.png"
        st.markdown(f'<div class="ficha-logo-wrap"><img src="{html.escape(u)}" alt="Direcional" /></div>', unsafe_allow_html=True)
    except Exception: pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack"><div class="ficha-hero">'
        f'<p class="ficha-title">Representatividade Sazonal & Metas</p></div>'
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true"><div class="ficha-hero-bar"></div></div></div>',
        unsafe_allow_html=True,
    )

def aplicar_estilo() -> None:
    bg_url = _css_url_fundo_cadastro()
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap');
        @keyframes fichaFadeIn {{ from {{ opacity: 0; transform: translateY(18px); }} to {{ opacity: 1; transform: translateY(0); }} }}
        @keyframes fichaShimmer {{ 0% {{ background-position: 0% 50%; }} 100% {{ background-position: 200% 50%; }} }}
        html, body, :root, [data-testid="stApp"] {{ color-scheme: light !important; }}
        .stDataFrame, [data-testid="stTable"] {{ background-color: #ffffff !important; color-scheme: light !important; }}
        [data-testid="stTable"] th, [data-testid="stTable"] td {{ background-color: #ffffff !important; color: #000000 !important; }}
        html, body {{ font-family: 'Inter', sans-serif; color: {COR_TEXTO_PRETO}; background: transparent !important; }}
        .stApp, [data-testid="stApp"] {{
            background: linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%),
            url("{bg_url}") center / cover no-repeat !important;
            background-attachment: scroll !important;
        }}
        [data-testid="stHeader"] {{ background: transparent !important; border: none !important; }}
        [data-testid="stSidebar"] {{ display: none !important; }}
        [data-testid="stSidebarCollapsedControl"] {{ display: none !important; }}
        [data-testid="stMain"] {{ padding: clamp(12px, 3.5vh, 40px) clamp(14px, 5vw, 56px) !important; }}
        .block-container {{
            max-width: 1700px !important; margin: clamp(4px, 1vh, 14px) auto !important;
            padding: 1.45rem 2.25rem 1.55rem 2.25rem !important;
            background: rgba(255, 255, 255, 0.78) !important;
            backdrop-filter: blur(18px) saturate(1.15); -webkit-backdrop-filter: blur(18px) saturate(1.15);
            border-radius: 24px !important; border: 1px solid rgba(255, 255, 255, 0.45) !important;
            box-shadow: 0 4px 6px -1px rgba({RGB_AZUL_CSS}, 0.06), 0 24px 48px -12px rgba({RGB_AZUL_CSS}, 0.18), inset 0 1px 0 rgba(255, 255, 255, 0.55) !important;
            animation: fichaFadeIn 0.7s cubic-bezier(0.22, 1, 0.36, 1) both;
        }}
        h1, h2, h3, h4, .stHeading {{ font-family: 'Montserrat', sans-serif !important; color: {COR_AZUL_ESC} !important; font-weight: 800 !important; text-align: center !important; }}
        h5, h6 {{ font-family: 'Montserrat', sans-serif !important; color: {COR_TEXTO_PRETO} !important; font-weight: 700 !important; text-align: center !important; }}
        p, label, li, span {{ color: {COR_TEXTO_PRETO} !important; }}
        div[data-testid="stButton"] button[kind="primary"] *, div[data-testid="stDownloadButton"] button * {{ color: #ffffff !important; }}
        .ficha-logo-wrap {{ text-align: center; padding: 0.1rem 0 0.45rem 0; }}
        .ficha-logo-wrap img {{ max-height: 72px; width: auto; max-width: min(280px, 85vw); }}
        .ficha-hero {{ text-align: center; padding: 0.5rem 0 0 0; margin: 0 auto; max-width: 640px; }}
        .ficha-hero .ficha-title {{ font-family: 'Montserrat', sans-serif; font-size: clamp(1.35rem, 3.5vw, 1.75rem); font-weight: 900; color: {COR_AZUL_ESC}; margin: 0; }}
        .ficha-hero-bar-wrap {{ width: 100%; margin: clamp(0.85rem, 2.4vw, 1.2rem) 0; }}
        .ficha-hero-bar {{ height: 4px; border-radius: 999px; background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC}); background-size: 200% 100%; animation: fichaShimmer 4s ease-in-out infinite alternate; }}
        div[data-baseweb="input"], div[data-baseweb="select"] {{ border-radius: 10px !important; border: 1px solid #e2e8f0 !important; background-color: {COR_INPUT_BG} !important; }}
        </style>
        """, unsafe_allow_html=True
    )

# -----------------------------------------------------------------------------
# Google Sheets (Metas)
# -----------------------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def buscar_meta_vendas_gsheets(ano_alvo: int, mes_alvo: int) -> float:
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        if "gcp_service_account" in st.secrets:
            sec = dict(st.secrets["gcp_service_account"])
        elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            sec = dict(st.secrets["connections"]["gsheets"])
        else:
            return 0.0
            
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds = Credentials.from_service_account_info(sec, scopes=scopes)
        gc = gspread.authorize(creds)
        ws = gc.open_by_key("1cseWbys3GXd7Q70irMK_8Iw3xz5iQy0wmArgkSdd5BA").worksheet("Meta Mensal")
        df = pd.DataFrame(ws.get_all_records())
        mask = (df["Ano"].astype(int) == int(ano_alvo)) & (df["Mês"].astype(int) == int(mes_alvo))
        if mask.any():
            return float(df.loc[mask, "Meta Vendas"].iloc[0])
    except Exception as e:
        pass
    return 0.0

# -----------------------------------------------------------------------------
# Conexão e Extração Salesforce
# -----------------------------------------------------------------------------
def conectar_sf():
    try:
        sec = st.secrets["salesforce"]
        from simple_salesforce import Salesforce
        kwargs = {"username": sec.get("USER", ""), "password": sec.get("PASSWORD", ""), "domain": sec.get("DOMAIN", "login")}
        if sec.get("TOKEN", ""): kwargs["security_token"] = sec.get("TOKEN", "")
        return Salesforce(**kwargs)
    except Exception:
        return None

@st.cache_data(ttl=3600, show_spinner=False)
def extrair_dados_sf_cached(ano_alvo: int, mes_alvo: int, _progress_callback=None):
    if _progress_callback: _progress_callback(10, "Conectando ao Salesforce...")
    sf = conectar_sf()
    if not sf: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    data_alvo_inicio = date(ano_alvo, mes_alvo, 1)
    dias_no_mes = calendar.monthrange(ano_alvo, mes_alvo)[1]
    data_alvo_fim = date(ano_alvo, mes_alvo, dias_no_mes)
    
    hoje = date.today()
    fim_extracao = data_alvo_fim if data_alvo_fim <= hoje else hoje
    ini_treino = date(ano_alvo - 3, mes_alvo, 1)
    
    desde_str = f"{ini_treino.isoformat()}T00:00:00Z"
    desde_date_str = ini_treino.isoformat()
    ate_str = f"{fim_extracao.isoformat()}T23:59:59Z"
    ate_date_str = fim_extracao.isoformat()

    soql_vendas = (
        "SELECT Id, ContratoGeradoEm__c FROM Opportunity "
        "WHERE DirecionalVendas__c = true AND Empreendimento__r.Regional__c = 'RJ' "
        "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' AND Imobiliaria__r.Name LIKE 'DIR%' "
        f"AND ContratoGeradoEm__c >= {desde_date_str} AND ContratoGeradoEm__c <= {ate_date_str}"
    )
    if _progress_callback: _progress_callback(30, "Extraindo Vendas (Oportunidades)...")
    df_ven = pd.DataFrame(sf.query_all(soql_vendas).get("records", []))
        
    soql_pastas = (
        "SELECT Name, dataPrimeiroEnvioAnalise__c, dataAprovacaoSAFI__c FROM Avaliacao_credito__c "
        "WHERE Empreendimento__r.Regional__c = 'RJ' AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND CreatedDate >= {desde_str} AND CreatedDate <= {ate_str}"
    )
    if _progress_callback: _progress_callback(55, "Extraindo Pastas (Avaliações de Crédito)...")
    df_pas = pd.DataFrame(sf.query_all(soql_pastas).get("records", []))

    soql_ag = (
        "SELECT Codigo_do_agendamento__c, CreatedDate, Data_da_Visita__c FROM Event "
        "WHERE Unidade_de_negocio__c = 'Direcional' AND Regional__c = 'RJ' "
        "AND Empreendimento_de_interesse__c != null "
        f"AND CreatedDate >= {desde_str} AND CreatedDate <= {ate_str}"
    )
    if _progress_callback: _progress_callback(80, "Extraindo Agendamentos e Visitas (Eventos)...")
    df_ag = pd.DataFrame(sf.query_all(soql_ag).get("records", []))

    return df_ag, df_pas, df_ven

# -----------------------------------------------------------------------------
# Processamento e Calendário
# -----------------------------------------------------------------------------
def formatar_data(serie):
    return pd.to_datetime(serie, errors="coerce").dt.date

def montar_calendario(df_ag, df_pas, df_ven, inicio, fim):
    if not df_ven.empty:
        df_ven["dt_contrato"] = formatar_data(df_ven.get("ContratoGeradoEm__c"))
        vendas_count = df_ven.dropna(subset=["dt_contrato"]).drop_duplicates("Id")["dt_contrato"].value_counts()
    else: vendas_count = pd.Series(dtype=float)

    if not df_pas.empty:
        df_pas["dt_envio"] = formatar_data(df_pas.get("dataPrimeiroEnvioAnalise__c"))
        df_pas["dt_safi"] = formatar_data(df_pas.get("dataAprovacaoSAFI__c"))
        pas_count = df_pas.dropna(subset=["dt_envio"]).drop_duplicates("Name")["dt_envio"].value_counts()
        aprov_count = df_pas.dropna(subset=["dt_safi"]).drop_duplicates("Name")["dt_safi"].value_counts()
    else:
        pas_count = pd.Series(dtype=float)
        aprov_count = pd.Series(dtype=float)

    if not df_ag.empty:
        df_ag["dt_criacao"] = formatar_data(df_ag.get("CreatedDate"))
        df_ag["dt_visita"] = formatar_data(df_ag.get("Data_da_Visita__c"))
        ag_count = df_ag.dropna(subset=["dt_criacao"]).drop_duplicates("Codigo_do_agendamento__c")["dt_criacao"].value_counts()
        vis_count = df_ag.dropna(subset=["dt_visita"]).drop_duplicates("Codigo_do_agendamento__c")["dt_visita"].value_counts()
    else:
        ag_count = pd.Series(dtype=float)
        vis_count = pd.Series(dtype=float)

    idx = pd.date_range(inicio, fim, freq="D")
    cal = pd.DataFrame({"data": [d.date() for d in idx]})
    cal["agendamentos"] = cal["data"].map(ag_count).fillna(0.0)
    cal["visitas"] = cal["data"].map(vis_count).fillna(0.0)
    cal["pastas"] = cal["data"].map(pas_count).fillna(0.0)
    cal["pastas_aprovadas"] = cal["data"].map(aprov_count).fillna(0.0)
    cal["vendas"] = cal["data"].map(vendas_count).fillna(0.0)
    
    cal["dia_mes"] = cal["data"].map(lambda d: d.day)
    cal["dia_semana"] = cal["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    cal["mes"] = cal["data"].map(lambda d: MESES_PT[d.month])
    return cal

# -----------------------------------------------------------------------------
# Regressão OLS de Efeitos Relativos (Com Tendência Linear e IC)
# -----------------------------------------------------------------------------
def matriz_explicativas_relativa(df: pd.DataFrame, t_start: int = 0) -> np.ndarray:
    n = len(df)
    X = np.zeros((n, 30 + 6 + 11 + 1 + 1), dtype=float)
    X[:, -1] = 1.0 # Intercepto
    # Tendência linear (evolução anual)
    X[:, -2] = (np.arange(n, dtype=float) + t_start) / 365.25 

    dias_semana_idx = {nome: i for i, nome in DIAS_SEMANA_PT.items()}
    meses_idx = {nome: i for i, nome in MESES_PT.items()}

    for i, row in enumerate(df.itertuples(index=False)):
        dia = int(row.dia_mes)
        if 2 <= dia <= 31: X[i, dia - 2] = 1.0
        ds = dias_semana_idx.get(str(row.dia_semana), None)
        if ds is not None and ds >= 1: X[i, 30 + (ds - 1)] = 1.0
        ms = meses_idx.get(str(row.mes), None)
        if ms is not None and ms >= 2: X[i, 30 + 6 + (ms - 2)] = 1.0
    return X

def estimar_efeitos_sazonais(treino: pd.DataFrame):
    if treino.empty or float(treino["qtd"].sum()) <= 0:
        return None
    X = matriz_explicativas_relativa(treino, t_start=0)
    y = treino["qtd"].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    
    y_hat = X @ coef
    res = y - y_hat
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    ss_res = np.sum(res ** 2)
    r2 = 1.0 - (ss_res / ss_tot) if ss_tot > 0 else 0.0
    mae = float(np.mean(np.abs(res)))
    
    n, k = X.shape
    df_err = n - k
    if df_err > 0:
        s2 = ss_res / df_err
        try:
            cov_matrix = s2 * np.linalg.pinv(X.T @ X)
            se = np.sqrt(np.diagonal(cov_matrix))
            t_stats = coef / se
            import math
            p_values = [1.0 - math.erf(abs(t) / math.sqrt(2.0)) for t in t_stats]
        except np.linalg.LinAlgError:
            cov_matrix = np.zeros((k, k))
            p_values = [np.nan] * k
    else:
        s2 = 0
        cov_matrix = np.zeros((k, k))
        p_values = [np.nan] * k
        
    nomes_features = []
    for d in range(2, 32): nomes_features.append(f"Dia do Mês {d}")
    for i in range(1, 7): nomes_features.append(f"Dia da Semana: {DIAS_SEMANA_PT[i].capitalize()}")
    for m in range(2, 13): nomes_features.append(f"Mês: {MESES_PT[m].capitalize()}")
    nomes_features.append("Tendência Linear (Efeito por Ano)")
    nomes_features.append("Intercepto (Dia 1, Segunda, Janeiro)")
    
    df_stats = pd.DataFrame({"Variável": nomes_features, "Beta": coef, "Valor-p": p_values})
    def sign_level(p):
        if pd.isna(p): return ""
        if p < 0.01: return "***"
        if p < 0.05: return "**"
        if p < 0.1: return "*"
        return ""
    df_stats["Significância"] = df_stats["Valor-p"].apply(sign_level)
    
    return {
        "coef": coef, "cov_matrix": cov_matrix, "s2": s2,
        "r2": r2, "mae": mae, "df_stats": df_stats
    }

# -----------------------------------------------------------------------------
# Plotly Helpers
# -----------------------------------------------------------------------------
def _plot_comparativo_representatividade(etapa: str, df: pd.DataFrame, is_current_month: bool):
    label = FUNIL_LABELS.get(etapa, etapa)
    col_proj = f"{label} Projetado (%)"
    col_real = f"{label} Realizado (%)"
    col_ic_up = f"{label} IC Upper (%)"
    col_ic_lo = f"{label} IC Lower (%)"
    
    if col_real not in df.columns or df[col_real].sum() <= 0:
        return 
        
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=df["Dia do Mês"], y=df[col_ic_up],
        mode="lines", line=dict(width=0), showlegend=False,
        hoverinfo="skip"
    ))
    fig.add_trace(go.Scatter(
        x=df["Dia do Mês"], y=df[col_ic_lo],
        mode="lines", line=dict(width=0), fill='tonexty',
        fillcolor='rgba(203, 9, 53, 0.15)', name="Intervalo de Confiança",
        hoverinfo="skip"
    ))
    
    df_real = df.dropna(subset=[col_real]) if is_current_month else df
    
    fig.add_trace(go.Scatter(
        x=df_real["Dia do Mês"], y=df_real[col_real],
        mode="lines+markers", name="Realizado",
        line=dict(color=COR_AZUL_ESC, width=3),
        marker=dict(size=7, color=COR_AZUL_ESC),
        hovertemplate="%{x}º dia: %{y:.2f}%<extra></extra>"
    ))
    
    fig.add_trace(go.Scatter(
        x=df["Dia do Mês"], y=df[col_proj],
        mode="lines+markers", name="Projetado",
        line=dict(color=COR_VERMELHO, width=3, dash="dash"),
        marker=dict(size=7, color=COR_VERMELHO),
        hovertemplate="%{x}º dia: %{y:.2f}%<extra></extra>"
    ))
    
    fig.update_layout(
        title=dict(text=f"{label} — Projetado x Realizado (%)", font=dict(family="Montserrat", color=COR_TEXTO_PRETO, size=16)),
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5),
        hovermode="x unified",
        height=380,
    )
    fig.update_xaxes(title_text="Dia do Mês", dtick=1, showgrid=False)
    fig.update_yaxes(title_text="Representatividade (%)", showgrid=True, gridcolor="rgba(226,232,240,0.5)")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def _plot_metas_acumuladas(etapa: str, df: pd.DataFrame, is_current_month: bool):
    label = FUNIL_LABELS.get(etapa, etapa)
    col_meta = f"Meta {label} (Acumulada)"
    col_real = f"{label} Realizado Absoluto (Acumulado)"
    
    if col_real not in df.columns or col_meta not in df.columns: return
        
    fig = go.Figure()
    df_real = df.dropna(subset=[col_real]) if is_current_month else df
    
    fig.add_trace(go.Scatter(
        x=df_real["Dia do Mês"], y=df_real[col_real],
        mode="lines+markers", name="Realizado (Acum.)",
        line=dict(color=COR_AZUL_ESC, width=3),
        marker=dict(size=7, color=COR_AZUL_ESC),
        hovertemplate="%{x}º dia: %{y:.0f} unidades<extra></extra>"
    ))
    
    fig.add_trace(go.Scatter(
        x=df["Dia do Mês"], y=df[col_meta],
        mode="lines+markers", name="Meta Projetada (Acum.)",
        line=dict(color=COR_VERMELHO, width=3, dash="dash"),
        marker=dict(size=7, color=COR_VERMELHO),
        hovertemplate="%{x}º dia: %{y:.0f} unidades<extra></extra>"
    ))
    
    fig.update_layout(
        title=dict(text=f"{label} — Meta x Realizado (Volume Acumulado)", font=dict(family="Montserrat", color=COR_TEXTO_PRETO, size=16)),
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5),
        hovermode="x unified", height=380,
    )
    fig.update_xaxes(title_text="Dia do Mês", dtick=1, showgrid=False)
    fig.update_yaxes(title_text="Volume Acumulado", showgrid=True, gridcolor="rgba(226,232,240,0.5)")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def to_excel(df: pd.DataFrame, df_conv: pd.DataFrame, df_metas: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Representatividade')
        df_metas.to_excel(writer, index=False, sheet_name='Metas Diárias')
        df_conv.to_excel(writer, index=False, sheet_name='Conversões')
        
        for sheet in writer.sheets:
            ws = writer.sheets[sheet]
            d = df if sheet == 'Representatividade' else (df_metas if sheet == 'Metas Diárias' else df_conv)
            for i, col in enumerate(d.columns):
                max_len = max(d[col].astype(str).map(len).max(), len(col)) + 2
                ws.set_column(i, i, min(max_len, 40))
    return output.getvalue()

def main():
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(page_title="Representatividade Sazonal", page_icon=str(fav) if fav else None, layout="wide")
    
    aplicar_estilo()
    _cabecalho_pagina()
    
    hoje = date.today()
    col1, col2 = st.columns(2)
    with col1:
        ano_alvo = st.number_input("Ano da Projeção", min_value=2020, max_value=2040, value=hoje.year)
    with col2:
        mes_alvo = st.selectbox("Mês da Projeção", options=list(range(1, 13)), format_func=lambda x: MESES_PT[x].capitalize(), index=hoje.month - 1)

    is_past_month = (ano_alvo < hoje.year) or (ano_alvo == hoje.year and mes_alvo <= hoje.month)
    is_current_month = (ano_alvo == hoje.year and mes_alvo == hoje.month)

    prog_placeholder = st.empty()
    t_start = time.time()
    
    def update_progress(pct, msg):
        elapsed = time.time() - t_start
        grad = f"linear-gradient(90deg, {COR_AZUL_ESC} 0%, {COR_VERMELHO} 100%)"
        
        if pct < 100:
            icone = f'''
            <svg style="animation: prog-spin 1s linear infinite; width: 1.25rem; height: 1.25rem; margin-right: 10px; color: {COR_AZUL_ESC};" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" stroke-dasharray="31.4 31.4" stroke-linecap="round" opacity="0.25"></circle>
                <path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"></path>
            </svg>
            '''
        else:
            icone = '''
            <svg style="width: 1.25rem; height: 1.25rem; margin-right: 10px; color: #10b981;" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M5 13l4 4L19 7" />
            </svg>
            '''

        html_str = f'''
        <style>
            @keyframes prog-spin {{ 100% {{ transform: rotate(360deg); }} }}
            @keyframes prog-shimmer {{ 0% {{ background-position: -200% 0; }} 100% {{ background-position: 200% 0; }} }}
        </style>
        <div style="margin: 1.5rem 0; padding: 1.25rem; background: rgba(255,255,255,0.95); border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 12px rgba(0,0,0,0.06);">
            <div style="display: flex; justify-content: space-between; align-items: center; font-size: 0.95rem; color: #1e293b; margin-bottom: 0.8rem; font-family: 'Inter', sans-serif; font-weight: 600;">
                <div style="display: flex; align-items: center;">
                    {icone}
                    <span>{msg}</span>
                </div>
                <span style="font-family: monospace; color: #475569; font-size: 0.85rem; background: #f1f5f9; padding: 3px 10px; border-radius: 6px; font-weight: 700;">
                    {pct}% &nbsp;|&nbsp; {elapsed:.1f}s
                </span>
            </div>
            <div style="width: 100%; background-color: #cbd5e1; border-radius: 999px; overflow: hidden; height: 12px; position: relative;">
                <div style="width: {pct}%; background: {grad}; height: 100%; transition: width 0.3s ease; border-radius: 999px; position: relative; overflow: hidden;">
                    <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: linear-gradient(90deg, transparent, rgba(255,255,255,0.35), transparent); background-size: 200% 100%; animation: prog-shimmer 1.5s infinite linear;"></div>
                </div>
            </div>
        </div>
        '''
        prog_placeholder.markdown(html_str, unsafe_allow_html=True)
        
    update_progress(5, "Iniciando inteligência sazonal...")
    
    # Executa cached fetching do Salesforce e Google Sheets
    df_ag_raw, df_pas_raw, df_ven_raw = extrair_dados_sf_cached(ano_alvo, mes_alvo, _progress_callback=update_progress)
    
    update_progress(85, "Buscando Metas no Google Sheets...")
    meta_vendas = buscar_meta_vendas_gsheets(ano_alvo, mes_alvo)
    
    if df_ven_raw.empty:
        prog_placeholder.empty()
        st.warning("Sem dados suficientes no Salesforce para projetar.")
        return
        
    update_progress(95, "Gerando projeções e painéis de dados...")
    data_alvo_inicio = date(ano_alvo, mes_alvo, 1)
    dias_no_mes = calendar.monthrange(ano_alvo, mes_alvo)[1]
    data_alvo_fim = date(ano_alvo, mes_alvo, dias_no_mes)
    
    fim_treino = data_alvo_inicio - timedelta(days=1)
    ini_treino = date(ano_alvo - 3, mes_alvo, 1)
    
    cal_total = montar_calendario(df_ag_raw, df_pas_raw, df_ven_raw, ini_treino, data_alvo_fim if is_past_month else fim_treino)
    cal_treino = cal_total[cal_total['data'] <= fim_treino].copy()
    cal_alvo = cal_total[(cal_total['data'] >= data_alvo_inicio) & (cal_total['data'] <= (hoje - timedelta(days=1) if is_current_month else data_alvo_fim))].copy() if is_past_month else pd.DataFrame()
    
    datas_alvo = [date(ano_alvo, mes_alvo, d) for d in range(1, dias_no_mes + 1)]
    
    df_resultado = pd.DataFrame({"Data": datas_alvo})
    df_resultado["Dia do Mês"] = df_resultado["Data"].map(lambda d: d.day)
    df_resultado["Dia da Semana"] = df_resultado["Data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()].capitalize())
    
    df_alvo_dummy = df_resultado.copy()
    df_alvo_dummy.rename(columns={"Dia do Mês": "dia_mes", "Dia da Semana": "dia_semana"}, inplace=True)
    df_alvo_dummy["dia_semana"] = df_alvo_dummy["dia_semana"].str.lower()
    df_alvo_dummy["mes"] = df_alvo_dummy["Data"].map(lambda d: MESES_PT[d.month])
    
    df_stats_all = pd.DataFrame()
    
    for etapa in FUNIL_ETAPAS:
        df_temp = cal_treino[["data", "dia_mes", "dia_semana", "mes", etapa]].copy()
        df_temp.rename(columns={etapa: "qtd"}, inplace=True)
        
        efeitos = estimar_efeitos_sazonais(df_temp)
        label_proj = f"{FUNIL_LABELS[etapa]} Projetado (%)"
        label_ic_up = f"{FUNIL_LABELS[etapa]} IC Upper (%)"
        label_ic_lo = f"{FUNIL_LABELS[etapa]} IC Lower (%)"
        label_real = f"{FUNIL_LABELS[etapa]} Realizado (%)"
        label_real_abs = f"{FUNIL_LABELS[etapa]} Realizado Absoluto"
        
        if not efeitos:
            for l in [label_proj, label_ic_up, label_ic_lo, label_real, label_real_abs]: df_resultado[l] = 0.0
            continue
            
        ef_stats = efeitos["df_stats"].copy()
        ef_stats.insert(0, "Indicador", FUNIL_LABELS[etapa])
        df_stats_all = pd.concat([df_stats_all, ef_stats])
        
        X_futuro = matriz_explicativas_relativa(df_alvo_dummy, t_start=len(cal_treino))
        pred_raw = X_futuro @ efeitos["coef"]
        var_pred = np.sum((X_futuro @ efeitos["cov_matrix"]) * X_futuro, axis=1) + efeitos["s2"]
        se_pred = np.sqrt(var_pred)
        
        esp_abs = np.maximum(pred_raw, 0.0)
        soma_esp = np.sum(esp_abs)
        
        df_resultado[label_proj] = (esp_abs / soma_esp * 100.0) if soma_esp > 0 else 0.0
        df_resultado[label_ic_up] = (np.maximum(pred_raw + 1.96 * se_pred, 0.0) / soma_esp * 100.0) if soma_esp > 0 else 0.0
        df_resultado[label_ic_lo] = (np.maximum(pred_raw - 1.96 * se_pred, 0.0) / soma_esp * 100.0) if soma_esp > 0 else 0.0
        
        df_resultado[label_real_abs] = np.nan
        df_resultado[label_real] = np.nan
        
        if is_past_month and not cal_alvo.empty:
            mapa_real = dict(zip(cal_alvo["data"], cal_alvo[etapa]))
            reais_diarios = [mapa_real.get(d) for d in datas_alvo if d in mapa_real]
            soma_real = sum([r for r in reais_diarios if r is not None])
            
            for i, d in enumerate(datas_alvo):
                if d in mapa_real:
                    v = mapa_real[d]
                    df_resultado.loc[i, label_real_abs] = v
                    df_resultado.loc[i, label_real] = (v / soma_real * 100.0) if soma_real > 0 else 0.0

    df_mensal = cal_treino.copy()
    df_mensal['ano_mes'] = pd.to_datetime(df_mensal['data']).dt.to_period('M')
    df_mensal_agg = df_mensal.groupby(['ano_mes', 'mes'], as_index=False)[list(FUNIL_ETAPAS)].sum()
    
    res_conv = []
    historico_plots = {}
    stats_mensais = {}
    
    for etapa in FUNIL_ETAPAS[:-1]:
        hist_y = []
        hist_x = []
        for _, row in df_mensal_agg.iterrows():
            v_etapa = float(row[etapa])
            v_vendas = float(row['vendas'])
            c = (v_vendas / v_etapa * 100.0) if v_etapa > 0 else 0.0
            hist_y.append(c)
            hist_x.append(str(row['ano_mes']))
        historico_plots[etapa] = {"x": hist_x, "y": hist_y}
        
        # Matriz de conversão com Evolução Anual
        X = np.zeros((len(df_mensal_agg), 11 + 1 + 1), dtype=float)
        X[:, -1] = 1.0 # Intercepto
        X[:, -2] = np.arange(len(df_mensal_agg), dtype=float) / 12.0 # Tendência Linear Anual
        
        meses_idx = {nome: i for i, nome in MESES_PT.items()}
        for i, m_str in enumerate(df_mensal_agg['mes']):
            ms = meses_idx.get(m_str, None)
            if ms is not None and ms >= 2: X[i, ms - 2] = 1.0
                
        y = np.array([y_val / 100.0 for y_val in hist_y])
        coef, *_ = np.linalg.lstsq(X, y, rcond=None)
        
        y_hat = X @ coef
        res_err = y - y_hat
        df_err = len(y) - X.shape[1]
        s2 = np.sum(res_err**2) / df_err if df_err > 0 else 0
        try:
            cov_matrix = s2 * np.linalg.pinv(X.T @ X)
        except:
            cov_matrix = np.zeros((X.shape[1], X.shape[1]))
        
        intercepto = float(coef[-1])
        tendencia = float(coef[-2])
        m_alvo_idx = meses_idx.get(MESES_PT[mes_alvo], 1)
        e_mes = float(coef[m_alvo_idx - 2]) if m_alvo_idx >= 2 else 0.0
        
        t_futuro = float(len(df_mensal_agg)) / 12.0
        esperado = max(intercepto + e_mes + tendencia * t_futuro, 0.0)
        
        real = None
        if is_past_month and not cal_alvo.empty:
            s_vendas = cal_alvo['vendas'].sum()
            s_etapa = cal_alvo[etapa].sum()
            real = (s_vendas / s_etapa) if s_etapa > 0 else 0.0
            
        dic_res = {
            "Indicador": f"{FUNIL_LABELS[etapa]} → Vendas",
            "Projetado (%)": esperado * 100.0,
        }
        if is_past_month:
            dic_res["Realizado (%)"] = real * 100.0 if real is not None else 0.0
        
        dic_res["Média Histórica (%)"] = float(np.mean(hist_y)) if hist_y else 0.0
        dic_res["Mediana Histórica (%)"] = float(np.median(hist_y)) if hist_y else 0.0
        res_conv.append(dic_res)
        
    df_res_conv = pd.DataFrame(res_conv)

    metas_etapa = {"vendas": meta_vendas}
    for etapa in FUNIL_ETAPAS[:-1]:
        conv_val = df_res_conv.loc[df_res_conv["Indicador"] == f"{FUNIL_LABELS[etapa]} → Vendas", "Projetado (%)"].iloc[0]
        metas_etapa[etapa] = (meta_vendas / (conv_val / 100.0)) if conv_val > 0 else 0.0

    df_metas_resumo = []
    for etapa in FUNIL_ETAPAS:
        label_proj_pct = f"{FUNIL_LABELS[etapa]} Projetado (%)"
        label_meta_dia = f"Meta {FUNIL_LABELS[etapa]} (Dia)"
        label_meta_acum = f"Meta {FUNIL_LABELS[etapa]} (Acumulada)"
        label_real_abs = f"{FUNIL_LABELS[etapa]} Realizado Absoluto"
        label_real_acum = f"{FUNIL_LABELS[etapa]} Realizado Absoluto (Acumulado)"
        
        # Meta Diária Absoluta
        df_resultado[label_meta_dia] = df_resultado[label_proj_pct] / 100.0 * metas_etapa[etapa]
        df_resultado[label_meta_acum] = df_resultado[label_meta_dia].cumsum()
        
        # Realizado Acumulado
        df_resultado[label_real_acum] = df_resultado[label_real_abs].cumsum()
        
        real_tot = df_resultado[label_real_abs].sum(skipna=True)
        meta_tot = metas_etapa[etapa]
        pct_ating = (real_tot / meta_tot * 100.0) if meta_tot > 0 else 0.0
        df_metas_resumo.append({
            "Indicador": FUNIL_LABELS[etapa],
            "Meta Mensal Estimada": meta_tot,
            "Realizado Acumulado": real_tot,
            "Atingimento (%)": pct_ating
        })

    df_metas_res = pd.DataFrame(df_metas_resumo)

    update_progress(100, "Concluído com sucesso!")
    time.sleep(0.5)
    prog_placeholder.empty()
    
    st.success(f"Análise concluída em {time.time() - t_start:.1f}s! Base de treino: {ini_treino.strftime('%m/%Y')} a {fim_treino.strftime('%m/%Y')}.")
    
    tab_metas, tab_diaria, tab_mensal, tab_stats = st.tabs(["Metas x Realizado", "Representatividade Diária", "Conversão Mensal", "Estatísticas OLS"])
    
    with tab_metas:
        st.subheader(f"Acompanhamento de Metas de {MESES_PT[mes_alvo].capitalize()}/{ano_alvo}")
        st.markdown(
            "<p style='color:#475569;font-size:0.9rem;'>As metas de topo de funil foram retrocalculadas usando a <b>Meta de Vendas do Google Sheets</b> "
            "e dividindo-a pela taxa de conversão projetada para o mês.</p>", unsafe_allow_html=True
        )
        
        st.dataframe(df_metas_res.style.format({"Meta Mensal Estimada": "{:,.0f}", "Realizado Acumulado": "{:,.0f}", "Atingimento (%)": "{:.1f}%"}), use_container_width=True, hide_index=True)
        
        if is_past_month:
            st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)
            for etapa in FUNIL_ETAPAS:
                _plot_metas_acumuladas(etapa, df_resultado, is_current_month)

    with tab_diaria:
        formatacao = {c: "{:.2f}%" for c in df_resultado.columns if "(%)" in c}
        st.dataframe(df_resultado[[c for c in df_resultado.columns if "Meta" not in c and "Absoluto" not in c]].style.format(formatacao), use_container_width=True, hide_index=True)
        
        if is_past_month and not cal_alvo.empty:
            st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)
            for etapa in FUNIL_ETAPAS:
                _plot_comparativo_representatividade(etapa, df_resultado, is_current_month)
                
    with tab_mensal:
        st.subheader("Projeção de Conversão em Vendas")
        fmt_conv = {c: "{:.2f}%" for c in df_res_conv.columns if "(%)" in c}
        st.dataframe(df_res_conv.style.format(fmt_conv), use_container_width=True, hide_index=True)
        
        st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)
        col_g1, col_g2 = st.columns(2)
        graficos_cols = [col_g1, col_g2, col_g1, col_g2]
        
        for i, etapa in enumerate(FUNIL_ETAPAS[:-1]):
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=historico_plots[etapa]["x"], y=historico_plots[etapa]["y"],
                mode="lines+markers", line=dict(color=COR_AZUL_ESC, width=2),
                hovertemplate="%{x}<br>Conversão: %{y:.2f}%<extra></extra>"
            ))
            
            projetado_val = df_res_conv.loc[df_res_conv["Indicador"] == f"{FUNIL_LABELS[etapa]} → Vendas", "Projetado (%)"].iloc[0]
            fig.add_hline(y=projetado_val, line_dash="dash", line_color=COR_VERMELHO, annotation_text=f"Projetado: {projetado_val:.1f}%")
            
            if is_past_month and "Realizado (%)" in df_res_conv.columns:
                realizado_val = df_res_conv.loc[df_res_conv["Indicador"] == f"{FUNIL_LABELS[etapa]} → Vendas", "Realizado (%)"].iloc[0]
                if realizado_val is not None:
                    fig.add_hline(y=realizado_val, line_dash="dot", line_color="#0f766e", annotation_text=f"Real: {realizado_val:.1f}%")

            fig.update_layout(title=dict(text=f"{FUNIL_LABELS[etapa]} → Vendas"), margin=dict(l=20, r=20, t=40, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=280, showlegend=False)
            fig.update_yaxes(title_text="Conversão (%)", showgrid=True, gridcolor="rgba(226,232,240,0.5)")
            with graficos_cols[i]: st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    
    with tab_stats:
        st.subheader("Painel de Correlação OLS (Sazonalidade Diária)")
        st.dataframe(df_stats_all.style.format({"Beta": "{:.4f}", "Valor-p": "{:.4f}"}), use_container_width=True, hide_index=True)

    # Download Global (Botão único remanescente)
    st.markdown("<br/>", unsafe_allow_html=True)
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        st.download_button(
            label=f"⬇️ Baixar Relatório em Excel Completo",
            data=to_excel(df_resultado, df_res_conv, df_metas_res),
            file_name=f"Representatividade_Metas_{MESES_PT[mes_alvo]}_{ano_alvo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True
        )

if __name__ == "__main__":
    main()
