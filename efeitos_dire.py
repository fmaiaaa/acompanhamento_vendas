# -*- coding: utf-8 -*-
"""
Ferramenta para cálculo da representatividade diária dos indicadores.
Extrai 36 meses de histórico do Salesforce, roda regressões OLS isoladas 
(dia da semana, dia do mês e mês) e converte o volume aditivo esperado
na participação percentual de cada dia dentro do mês alvo.
Inclui comparativo Realizado x Projetado para meses fechados, com 
estilização Gaps Style (Direcional).
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
        # Fallback URL se não achar a logo na pasta local
        u = "https://logodownload.org/wp-content/uploads/2021/04/direcional-engenharia-logo.png"
        st.markdown(f'<div class="ficha-logo-wrap"><img src="{html.escape(u)}" alt="Direcional" /></div>', unsafe_allow_html=True)
    except Exception: pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack"><div class="ficha-hero">'
        f'<p class="ficha-title">Representatividade Sazonal</p></div>'
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
# Conexão e Extração Salesforce
# -----------------------------------------------------------------------------
def conectar_salesforce_app():
    try:
        sec = st.secrets["salesforce"]
        username = sec.get("USER", "")
        password = sec.get("PASSWORD", "")
        token = sec.get("TOKEN", "")
        domain = sec.get("DOMAIN", "login")
        
        from simple_salesforce import Salesforce
        kwargs = {"username": username, "password": password, "domain": domain}
        if token: kwargs["security_token"] = token
        return Salesforce(**kwargs)
    except Exception as e:
        st.error(f"Erro ao conectar no Salesforce: {e}")
        return None

def buscar_dados_salesforce(sf, inicio_treino: date, fim_alvo: date):
    """Extrai dados necessários cobrindo desde o inicio do treino até o fim do mês alvo."""
    desde_str = f"{inicio_treino.isoformat()}T00:00:00Z"
    desde_date_str = inicio_treino.isoformat()
    ate_str = f"{fim_alvo.isoformat()}T23:59:59Z"
    ate_date_str = fim_alvo.isoformat()

    with st.spinner("Extraindo Vendas (Opportunity)..."):
        soql_vendas = (
            "SELECT Id, ContratoGeradoEm__c FROM Opportunity "
            "WHERE DirecionalVendas__c = true "
            "AND Empreendimento__r.Regional__c = 'RJ' "
            "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
            f"AND ContratoGeradoEm__c >= {desde_date_str} AND ContratoGeradoEm__c <= {ate_date_str}"
        )
        df_ven = pd.DataFrame(sf.query_all(soql_vendas).get("records", []))
        
    with st.spinner("Extraindo Pastas (Avaliacao_credito__c)..."):
        soql_pastas = (
            "SELECT Name, dataPrimeiroEnvioAnalise__c, dataAprovacaoSAFI__c "
            "FROM Avaliacao_credito__c "
            "WHERE Empreendimento__r.Regional__c = 'RJ' "
            "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
            f"AND CreatedDate >= {desde_str} AND CreatedDate <= {ate_str}"
        )
        df_pas = pd.DataFrame(sf.query_all(soql_pastas).get("records", []))

    with st.spinner("Extraindo Agendamentos e Visitas (Event)..."):
        soql_ag = (
            "SELECT Codigo_do_agendamento__c, CreatedDate, Data_da_Visita__c "
            "FROM Event "
            "WHERE Unidade_de_negocio__c = 'Direcional' "
            "AND Regional__c = 'RJ' "
            "AND Empreendimento_de_interesse__c != null "
            f"AND CreatedDate >= {desde_str} AND CreatedDate <= {ate_str}"
        )
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
    else:
        vendas_count = pd.Series(dtype=float)

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
# Regressão OLS de Efeitos Relativos
# -----------------------------------------------------------------------------
def matriz_explicativas_relativa(df: pd.DataFrame) -> np.ndarray:
    n = len(df)
    X = np.zeros((n, 30 + 6 + 11 + 1), dtype=float)
    X[:, -1] = 1.0

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
    X = matriz_explicativas_relativa(treino)
    y = treino["qtd"].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    
    intercepto = float(coef[-1])
    efeito_dm = {"1": 0.0}
    for d in range(2, 32): efeito_dm[str(d)] = float(coef[d - 2])
    efeito_ds = {DIAS_SEMANA_PT[0]: 0.0}
    for i in range(1, 7): efeito_ds[DIAS_SEMANA_PT[i]] = float(coef[30 + (i - 1)])
    efeito_mes = {MESES_PT[1]: 0.0}
    for m in range(2, 13): efeito_mes[MESES_PT[m]] = float(coef[30 + 6 + (m - 2)])

    return {
        "intercepto": intercepto,
        "dia_mes": efeito_dm,
        "dia_semana": efeito_ds,
        "mes": efeito_mes,
    }

# -----------------------------------------------------------------------------
# Funções de Plotagem e Interface
# -----------------------------------------------------------------------------
def _plot_comparativo_representatividade(etapa: str, df: pd.DataFrame):
    label = FUNIL_LABELS.get(etapa, etapa)
    col_proj = f"{label} Projetado (%)"
    col_real = f"{label} Realizado (%)"
    
    if col_real not in df.columns or df[col_real].sum() <= 0:
        return # Não plota se não há dado real
        
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=df["Dia do Mês"], y=df[col_real],
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

def to_excel(df: pd.DataFrame, df_conv: pd.DataFrame = None) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Representatividade')
        worksheet = writer.sheets['Representatividade']
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
            
        if df_conv is not None:
            df_conv.to_excel(writer, index=False, sheet_name='Conversões')
            ws_conv = writer.sheets['Conversões']
            for i, col in enumerate(df_conv.columns):
                max_len = max(df_conv[col].astype(str).map(len).max(), len(col)) + 2
                ws_conv.set_column(i, i, max_len)
                
    return output.getvalue()

def main():
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(page_title="Representatividade Sazonal", page_icon=str(fav) if fav else None, layout="wide")
    
    aplicar_estilo()
    _cabecalho_pagina()
    
    st.markdown("""
    <p style="text-align:center; max-width:800px; margin: 0 auto 2rem auto; font-size:1rem; color: #475569;">
    Selecione o mês para projetar. A ferramenta capta os <b>36 meses anteriores</b> fechados, rodando uma 
    regressão linear de calendário para projetar o peso esperado de cada dia. 
    Se o mês alvo for um mês que já possui dados reais, a ferramenta incluirá o comparativo 
    Projetado vs Realizado.
    </p>
    """, unsafe_allow_html=True)
    
    hoje = date.today()
    col1, col2 = st.columns(2)
    with col1:
        ano_alvo = st.number_input("Ano da Projeção", min_value=2020, max_value=2040, value=hoje.year)
    with col2:
        mes_alvo = st.selectbox("Mês da Projeção", options=list(range(1, 13)), format_func=lambda x: MESES_PT[x].capitalize(), index=hoje.month - 1)

    is_past_month = (ano_alvo < hoje.year) or (ano_alvo == hoje.year and mes_alvo <= hoje.month)

    if st.button("Buscar Dados e Calcular", type="primary", use_container_width=True):
        sf = conectar_salesforce_app()
        if not sf: return
        
        # O treino precisa ser dos 36 meses EXATAMENTE ANTERIORES ao mês alvo.
        data_alvo_inicio = date(ano_alvo, mes_alvo, 1)
        dias_no_mes = calendar.monthrange(ano_alvo, mes_alvo)[1]
        data_alvo_fim = date(ano_alvo, mes_alvo, dias_no_mes)
        
        fim_treino = data_alvo_inicio - timedelta(days=1)
        ini_treino = date(ano_alvo - 3, mes_alvo, 1)
        
        # Vamos baixar os dados cobrindo o treino e o mes alvo (se houver dados)
        df_ag, df_pas, df_ven = buscar_dados_salesforce(sf, ini_treino, data_alvo_fim if is_past_month else fim_treino)
        
        with st.spinner("Construindo calendário e processando modelo..."):
            cal_total = montar_calendario(df_ag, df_pas, df_ven, ini_treino, data_alvo_fim if is_past_month else fim_treino)
            
            cal_treino = cal_total[cal_total['data'] <= fim_treino].copy()
            cal_alvo = cal_total[cal_total['data'] >= data_alvo_inicio].copy() if is_past_month else pd.DataFrame()
            
            datas_alvo = [date(ano_alvo, mes_alvo, d) for d in range(1, dias_no_mes + 1)]
            
            df_resultado = pd.DataFrame()
            df_resultado["Data"] = datas_alvo
            df_resultado["Dia do Mês"] = df_resultado["Data"].map(lambda d: d.day)
            df_resultado["Dia da Semana"] = df_resultado["Data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()].capitalize())
            
            for etapa in FUNIL_ETAPAS:
                df_temp = cal_treino[["data", "dia_mes", "dia_semana", "mes", etapa]].copy()
                df_temp.rename(columns={etapa: "qtd"}, inplace=True)
                
                efeitos = estimar_efeitos_sazonais(df_temp)
                label_proj = f"{FUNIL_LABELS[etapa]} Projetado (%)"
                label_real = f"{FUNIL_LABELS[etapa]} Realizado (%)"
                
                if not efeitos:
                    df_resultado[label_proj] = 0.0
                    if is_past_month: df_resultado[label_real] = 0.0
                    continue
                
                intercepto = efeitos["intercepto"]
                esperados_diarios = []
                
                for d in datas_alvo:
                    e_mes = efeitos["mes"].get(MESES_PT[d.month], 0.0)
                    e_dm = efeitos["dia_mes"].get(str(d.day), 0.0)
                    e_ds = efeitos["dia_semana"].get(DIAS_SEMANA_PT[d.weekday()], 0.0)
                    esperados_diarios.append(max(intercepto + e_mes + e_dm + e_ds, 0.0))
                
                soma_esp = sum(esperados_diarios)
                df_resultado[label_proj] = [(v / soma_esp * 100.0) if soma_esp > 0 else 0.0 for v in esperados_diarios]
                
                # Se for mês atual ou passado, pega a distribuição REAL do mês
                if is_past_month and not cal_alvo.empty:
                    mapa_real = dict(zip(cal_alvo["data"], cal_alvo[etapa]))
                    reais_diarios = [mapa_real.get(d, 0.0) for d in datas_alvo]
                    soma_real = sum(reais_diarios)
                    df_resultado[label_real] = [(v / soma_real * 100.0) if soma_real > 0 else 0.0 for v in reais_diarios]

        st.success(f"Cálculo concluído! Base de treino: {ini_treino.strftime('%m/%Y')} a {fim_treino.strftime('%m/%Y')}.")
        
        tab_diaria, tab_mensal = st.tabs(["Representatividade Diária", "Conversão Mensal"])
        
        df_mensal = cal_treino.copy()
        df_mensal['ano_mes'] = pd.to_datetime(df_mensal['data']).dt.to_period('M')
        df_mensal_agg = df_mensal.groupby(['ano_mes', 'mes'], as_index=False)[list(FUNIL_ETAPAS)].sum()
        
        res_conv = []
        historico_plots = {}
        
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
            
            X = np.zeros((len(df_mensal_agg), 11 + 1), dtype=float)
            X[:, -1] = 1.0
            meses_idx = {nome: i for i, nome in MESES_PT.items()}
            for i, m_str in enumerate(df_mensal_agg['mes']):
                ms = meses_idx.get(m_str, None)
                if ms is not None and ms >= 2:
                    X[i, ms - 2] = 1.0
            y = np.array([y_val / 100.0 for y_val in hist_y])
            
            coef, *_ = np.linalg.lstsq(X, y, rcond=None)
            intercepto = float(coef[-1])
            m_alvo_idx = meses_idx.get(MESES_PT[mes_alvo], 1)
            e_mes = float(coef[m_alvo_idx - 2]) if m_alvo_idx >= 2 else 0.0
            esperado = max(intercepto + e_mes, 0.0)
            
            real = None
            if is_past_month and not cal_alvo.empty:
                s_vendas = cal_alvo['vendas'].sum()
                s_etapa = cal_alvo[etapa].sum()
                real = (s_vendas / s_etapa) if s_etapa > 0 else 0.0
                
            dic_res = {
                "Indicador": f"{FUNIL_LABELS[etapa]} → Vendas",
                "Projetado (%)": esperado * 100.0
            }
            if is_past_month:
                dic_res["Realizado (%)"] = real * 100.0 if real is not None else 0.0
            res_conv.append(dic_res)
            
        df_res_conv = pd.DataFrame(res_conv)
        
        with tab_diaria:
            # Mostra formatado
            formatacao = {c: "{:.2f}%" for c in df_resultado.columns if "(%)" in c}
            st.dataframe(df_resultado.style.format(formatacao), use_container_width=True, hide_index=True)
            
            # Gráficos Comparativos
            if is_past_month and not cal_alvo.empty:
                st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)
                st.subheader("Gráficos Comparativos: Projetado vs Realizado")
                for etapa in FUNIL_ETAPAS:
                    _plot_comparativo_representatividade(etapa, df_resultado)
                    
        with tab_mensal:
            st.subheader("Projeção de Conversão em Vendas")
            st.markdown(
                "<p style='color:#475569;font-size:0.9rem;'>Estimativa da taxa de conversão mensal, "
                "baseada em regressão linear simples (dummies de mês) do histórico de 36 meses.</p>", 
                unsafe_allow_html=True
            )
            
            fmt_conv = {c: "{:.2f}%" for c in df_res_conv.columns if "(%)" in c}
            st.dataframe(df_res_conv.style.format(fmt_conv), use_container_width=True, hide_index=True)
            
            st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)
            st.subheader("Histórico de Conversão (Últimos 36 meses)")
            
            col_g1, col_g2 = st.columns(2)
            graficos_cols = [col_g1, col_g2, col_g1, col_g2]
            
            for i, etapa in enumerate(FUNIL_ETAPAS[:-1]):
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=historico_plots[etapa]["x"], 
                    y=historico_plots[etapa]["y"],
                    mode="lines+markers",
                    name="Histórico",
                    line=dict(color=COR_AZUL_ESC, width=2),
                    marker=dict(size=6, color=COR_AZUL_ESC),
                    hovertemplate="%{x}<br>Conversão: %{y:.2f}%<extra></extra>"
                ))
                
                projetado_val = df_res_conv.loc[df_res_conv["Indicador"] == f"{FUNIL_LABELS[etapa]} → Vendas", "Projetado (%)"].iloc[0]
                fig.add_hline(y=projetado_val, line_dash="dash", line_color=COR_VERMELHO, annotation_text=f"Projetado {MESES_PT[mes_alvo].capitalize()}: {projetado_val:.1f}%")
                
                if is_past_month and "Realizado (%)" in df_res_conv.columns:
                    realizado_val = df_res_conv.loc[df_res_conv["Indicador"] == f"{FUNIL_LABELS[etapa]} → Vendas", "Realizado (%)"].iloc[0]
                    fig.add_hline(y=realizado_val, line_dash="dot", line_color="#0f766e", annotation_text=f"Real: {realizado_val:.1f}%", annotation_position="bottom right")

                fig.update_layout(
                    title=dict(text=f"{FUNIL_LABELS[etapa]} → Vendas", font=dict(family="Montserrat", color=COR_TEXTO_PRETO, size=14)),
                    margin=dict(l=20, r=20, t=40, b=20),
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(family="Inter", color=COR_TEXTO_PRETO),
                    height=280,
                    showlegend=False
                )
                fig.update_yaxes(title_text="Conversão (%)", showgrid=True, gridcolor="rgba(226,232,240,0.5)")
                fig.update_xaxes(showgrid=False, tickangle=-45, type='category')
                
                with graficos_cols[i]:
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        
        # Download
        excel_bytes = to_excel(df_resultado, df_res_conv)
        nome_arquivo = f"Representatividade_{MESES_PT[mes_alvo]}_{ano_alvo}.xlsx"
        
        st.markdown("<br/>", unsafe_allow_html=True)
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            st.download_button(
                label=f"⬇️ Baixar Relatório em Excel ({nome_arquivo})",
                data=excel_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

    st.markdown(
        f'<div style="text-align:center;padding:2rem 0 1rem 0;color:{COR_TEXTO_PRETO};font-size:0.82rem;">'
        f"Direcional Engenharia · Representatividade Sazonal</div>",
        unsafe_allow_html=True,
    )

if __name__ == "__main__":
    main()
