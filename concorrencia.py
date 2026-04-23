# -*- coding: utf-8 -*-
"""
Análise de Concorrência — Inteligência Competitiva e Performance.
Foco: Comparação Direta entre Empreendimentos (BD GERAL e BD DETALHADA).
"""
from __future__ import annotations

import base64
import html
import os
import re
import math
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

# -----------------------------------------------------------------------------
# Identificação da planilha e Arquivos Visuais
# -----------------------------------------------------------------------------
SPREADSHEET_ID_CONC = "1nwRSz-ixnHncsT7UxRkjA7jBwe31ZiJdEfcM0Dm5aYE"

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"
URL_LOGO_DIRECIONAL_EMAIL = "https://logodownload.org/wp-content/uploads/2021/04/direcional-engenharia-logo.png"

COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_BORDA = "#eef2f6"
COR_TEXTO_MUTED = "#64748b"
COR_TEXTO_LABEL = "#1e293b"
COR_INPUT_BG = "#f0f2f6"

def _hex_rgb_triplet(hex_color: str) -> str:
    x = (hex_color or "").strip().lstrip("#")
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}" if len(x) == 6 else "0,0,0"

RGB_AZUL_CSS = _hex_rgb_triplet(COR_AZUL_ESC)
RGB_VERMELHO_CSS = _hex_rgb_triplet(COR_VERMELHO)

# -----------------------------------------------------------------------------
# Funções de Design (Padrão Gaps)
# -----------------------------------------------------------------------------
def _resolver_png_raiz(nome: str) -> Path | None:
    for base in (_DIR_APP, _DIR_APP.parent):
        p = base / nome
        if p.is_file(): return p
    return None

def _css_url_fundo_cadastro() -> str:
    p = _resolver_png_raiz(FUNDO_CADASTRO_ARQUIVO)
    if p and p.is_file():
        b64 = base64.b64encode(p.read_bytes()).decode("ascii")
        return f"data:image/jpeg;base64,{b64}"
    return ""

def aplicar_estilo() -> None:
    bg_url = _css_url_fundo_cadastro()
    st.markdown(f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap');
        @keyframes fichaFadeIn {{ from {{ opacity: 0; transform: translateY(18px); }} to {{ opacity: 1; transform: translateY(0); }} }}
        html, body, :root, [data-testid="stApp"] {{ color-scheme: light !important; }}
        html, body {{ font-family: 'Inter', sans-serif; color: {COR_TEXTO_LABEL}; background: transparent !important; }}
        .stApp {{
            background: linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%),
                url("{bg_url}") center / cover no-repeat !important;
            background-attachment: scroll !important;
        }}
        [data-testid="stAppViewContainer"] {{ background: transparent !important; }}
        header[data-testid="stHeader"] {{ background: transparent !important; border: none !important; box-shadow: none !important; }}
        [data-testid="stSidebar"] {{ display: none !important; }}
        .block-container {{
            max-width: 1700px !important; margin-left: auto !important; margin-right: auto !important;
            padding: 1.45rem 2.25rem 1.55rem 2.25rem !important; background: rgba(255, 255, 255, 0.78) !important;
            backdrop-filter: blur(18px) saturate(1.15); border-radius: 24px !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            box-shadow: 0 4px 6px -1px rgba({RGB_AZUL_CSS}, 0.06), 0 24px 48px -12px rgba({RGB_AZUL_CSS}, 0.18) !important;
            animation: fichaFadeIn 0.7s both;
        }}
        h1, h2, h3, h4 {{ font-family: 'Montserrat', sans-serif !important; color: {COR_AZUL_ESC} !important; font-weight: 800 !important; text-align: center !important; }}
        .ficha-logo-wrap {{ text-align: center; padding: 0.1rem 0 0.45rem 0; }}
        .ficha-logo-wrap img {{ max-height: 72px; width: auto; max-width: min(280px, 85vw); object-fit: contain; }}
        .ficha-hero-stack {{ width: 100%; margin-bottom: 0.35rem; }}
        .ficha-hero {{ text-align: center; padding: 0.5rem 0 0 0; max-width: 640px; margin: 0 auto; }}
        .ficha-hero .ficha-title {{ font-family: 'Montserrat', sans-serif; font-size: clamp(1.35rem, 3.5vw, 1.75rem); font-weight: 900; color: {COR_AZUL_ESC}; margin: 0; }}
        .ficha-hero .ficha-sub {{ color: #475569; font-size: 0.95rem; margin: 0.45rem 0 0 0; }}
        .ficha-hero-bar-wrap {{ width: 100%; margin: clamp(0.85rem, 2.4vw, 1.2rem) 0; }}
        .ficha-hero-bar {{ height: 4px; border-radius: 999px; background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC}); background-size: 200% 100%; }}
        .vel-kpi-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 1.25rem; }}
        .vel-kpi {{
            flex: 1 1 18%; background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.82) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9); border-radius: 14px; padding: 14px 16px; text-align: center;
            box-shadow: 0 2px 8px rgba({RGB_AZUL_CSS}, 0.06); transition: transform 0.3s ease;
        }}
        .vel-kpi:hover {{ transform: translateY(-4px); box-shadow: 0 10px 20px -5px rgba({RGB_AZUL_CSS}, 0.15); }}
        .vel-kpi .lbl {{ font-size: 0.72rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: {COR_AZUL_ESC}; opacity: 0.85; }}
        .vel-kpi .val {{ font-family: 'Montserrat', sans-serif; font-size: 1.35rem; font-weight: 800; color: {COR_AZUL_ESC}; margin-top: 6px; }}
        .vel-kpi .val--red {{ color: {COR_VERMELHO} !important; }}
        div[data-baseweb="input"], div[data-baseweb="select"] {{ border-radius: 10px !important; border: 1px solid #e2e8f0 !important; background-color: {COR_INPUT_BG} !important; }}
        </style>
        """, unsafe_allow_html=True)

def _exibir_logo_topo() -> None:
    path = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
    try:
        if path:
            ext = Path(path).suffix.lower().lstrip(".")
            mime = "image/png" if ext == "png" else "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            st.markdown(f'<div class="ficha-logo-wrap"><img src="data:{mime};base64,{b64}" alt="Direcional" /></div>', unsafe_allow_html=True)
    except Exception: pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Inteligencia Competitiva - Estudo Comparativo</p>'
        f'<p class="ficha-sub">Analise Direcional vs Concorrencia via BD GERAL e DETALHADA.</p>'
        f"</div>"
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true"><div class="ficha-hero-bar"></div></div>'
        f"</div>",
        unsafe_allow_html=True,
    )

# -----------------------------------------------------------------------------
# Pipeline de Dados
# -----------------------------------------------------------------------------
def parse_val(v):
    if not v: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace("R$", "").replace(".", "").replace(",", ".").replace(" ", "").strip()
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

@st.cache_data(ttl=300, show_spinner=False)
def load_base_master() -> dict[str, pd.DataFrame]:
    import gspread
    from google.oauth2.service_account import Credentials
    sec = st.secrets["connections"]["gsheets"]
    info = {k: v for k, v in sec.items() if v}
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID_CONC)
    
    dfs = {}
    for ws in sh.worksheets():
        if ws.title in ["BD GERAL", "BD DETALHADA", "DADOS DIRECIONAL"]:
            data = ws.get_all_values()
            if data:
                df = pd.DataFrame(data[1:], columns=data[0])
                df.columns = [str(c).strip() for c in df.columns]
                dfs[ws.title] = df
    return dfs

def process_pipeline():
    sheets = load_base_master()
    df_geral = sheets.get("BD GERAL", pd.DataFrame())
    df_det = sheets.get("BD DETALHADA", pd.DataFrame())
    df_ref_dir = sheets.get("DADOS DIRECIONAL", pd.DataFrame())
    
    if df_geral.empty:
        raise ValueError("A aba BD GERAL nao foi encontrada ou esta vazia.")

    # Normalizar chaves
    for df in [df_geral, df_det, df_ref_dir]:
        if "CHAVE" in df.columns: df["CHAVE"] = df["CHAVE"].astype(str).str.strip().str.upper()
        elif "Chave" in df.columns: df["CHAVE"] = df["Chave"].astype(str).str.strip().str.upper()
        elif "Nome do Empreendimento (Chave)" in df.columns: df["CHAVE"] = df["Nome do Empreendimento (Chave)"].astype(str).str.strip().str.upper()

    # 1. Limpeza BD GERAL (Performance)
    cols_num = ["VENDAS", "ESTOQUE", "ESTOQUE INICIAL"]
    for col in cols_num:
        df_geral[col] = pd.to_numeric(df_geral[col], errors='coerce').fillna(0)
    
    df_geral["VGV_Vendido"] = df_geral["VGV"].apply(parse_val)
    df_geral["PRECO_MEDIO"] = df_geral["PREÇO MÉDIO"].apply(parse_val)
    df_geral["DATA_DT"] = pd.to_datetime(df_geral["DATA"], dayfirst=True, errors='coerce')
    df_geral = df_geral[df_geral["DATA_DT"].notna()].sort_values(["CHAVE", "DATA_DT"])
    
    # Indicadores Temporais
    df_geral["Absorcao"] = df_geral["VENDAS"] / df_geral["ESTOQUE INICIAL"].replace(0, np.nan)
    df_geral["Velocidade"] = df_geral["VENDAS"] / (df_geral["ESTOQUE INICIAL"] - df_geral["VENDAS"]).replace(0, np.nan)
    df_geral["VGV_Rate"] = df_geral["VGV_Vendido"] / (df_geral["VGV_Vendido"] + (df_geral["ESTOQUE"] * df_geral["PRECO_MEDIO"])).replace(0, np.nan)
    
    # 2. Limpeza BD DETALHADA (Precificacao por Tipologia)
    df_det["Preco_m2"] = df_det["PREÇO_M2"].apply(parse_val)
    df_det["DATA_DT"] = pd.to_datetime(df_det["DATA"], dayfirst=True, errors='coerce')
    df_det = df_det[df_det["DATA_DT"].notna()]
    
    # 3. Identificar Direcional
    direcional_keys = [str(x).strip().upper() for x in df_ref_dir["CHAVE"].unique() if x]
    df_geral["Is_Direcional"] = (
        (df_geral["CONSTRUTORA"].astype(str).str.strip().str.upper() == "DIRECIONAL") |
        (df_geral["CHAVE"].isin(direcional_keys))
    )
    
    return df_geral, df_det

# -----------------------------------------------------------------------------
# Main Application
# -----------------------------------------------------------------------------
def main():
    aplicar_estilo()
    _cabecalho_pagina()
    
    try:
        df_master, df_detalhada = process_pipeline()
    except Exception as e:
        st.error(f"Erro no pipeline de dados: {e}")
        return

    # Filtros de Topo (Seleção Manual)
    st.markdown("<div style='text-align: center; font-weight: bold; margin-bottom: 1rem;'>Configuracao da Analise Comparativa</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    
    with c1:
        # Seleção do empreendimento Direcional
        lista_direcional = sorted(df_master[df_master["Is_Direcional"]]["EMPREENDIMENTO"].unique())
        if not lista_direcional:
            st.error("Nenhum empreendimento Direcional identificado.")
            lista_direcional = sorted(df_master["EMPREENDIMENTO"].unique())
        alvo_selecionado = st.selectbox("Selecione o Empreendimento Direcional Alvo", lista_direcional)
        
    with c2:
        # Seleção de Concorrentes (Qualquer outro empreendimento)
        lista_concorrentes = sorted(df_master[df_master["EMPREENDIMENTO"] != alvo_selecionado]["EMPREENDIMENTO"].unique())
        concorrentes_selecionados = st.multiselect("Selecione Empreendimento(s) Concorrente(s)", lista_concorrentes)

    # Consolidação da Seleção
    nomes_selecionados = [alvo_selecionado] + concorrentes_selecionados
    df_f = df_master[df_master["EMPREENDIMENTO"].isin(nomes_selecionados)].copy()

    # Filtro de Mês para KPIs
    df_sorted_dates = df_master[["DATA", "DATA_DT"]].drop_duplicates().sort_values("DATA_DT")
    meses_disp = df_sorted_dates["DATA"].tolist()
    mes_estudo = st.selectbox("Mes de Referencia para Resumo", meses_disp, index=len(meses_disp)-1)

    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # KPIs Consolidados (Mês Selecionado)
    # -------------------------------------------------------------------------
    df_mes = df_f[df_f["DATA"] == mes_estudo]
    df_alvo_mes = df_mes[df_mes["EMPREENDIMENTO"] == alvo_selecionado]
    df_conc_mes = df_mes[df_mes["EMPREENDIMENTO"] != alvo_selecionado]
    
    abs_alvo = df_alvo_mes["Absorcao"].iloc[0] * 100 if not df_alvo_mes.empty else 0
    m2_alvo = df_alvo_mes["PRECO_MEDIO"].iloc[0] if not df_alvo_mes.empty else 0
    m2_conc = df_conc_mes["PRECO_MEDIO"].mean() if not df_conc_mes.empty else 0
    vgv_rate_alvo = df_alvo_mes["VGV_Rate"].iloc[0] * 100 if not df_alvo_mes.empty else 0

    st.markdown(f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Preco Medio Concorrencia</div><div class="val">R$ {m2_conc:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Preco Medio Direcional</div><div class="val val--red">R$ {m2_alvo:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Absorcao Direcional</div><div class="val val--red">{abs_alvo:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">VGV Rate (Monetizacao)</div><div class="val val--red">{vgv_rate_alvo:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">Total Vendas Cluster</div><div class="val">{int(df_mes["VENDAS"].sum())}</div></div>
        </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # GRAFICOS DE COLUNAS (BD GERAL)
    # -------------------------------------------------------------------------
    st.subheader("Indicadores de Performance ao Longo do Tempo (BD GERAL)")
    
    if not df_f.empty:
        df_trend = df_f.sort_values("DATA_DT")
        df_trend["DATA_STR"] = df_trend["DATA_DT"].dt.strftime("%m/%Y")

        g1, g2 = st.columns(2)
        
        with g1:
            st.markdown("**Taxa de Absorcao (Barra)**")
            fig_abs = px.bar(df_trend, x="DATA_STR", y="Absorcao", color="EMPREENDIMENTO", barmode="group",
                             color_discrete_sequence=px.colors.qualitative.Prism)
            fig_abs.update_layout(yaxis_tickformat=".1%", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
            st.plotly_chart(fig_abs, use_container_width=True)
            
            st.markdown("**VGV Rate (Barra)**")
            fig_vgv = px.bar(df_trend, x="DATA_STR", y="VGV_Rate", color="EMPREENDIMENTO", barmode="group",
                             color_discrete_sequence=px.colors.qualitative.Prism)
            fig_vgv.update_layout(yaxis_tickformat=".1%", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
            st.plotly_chart(fig_vgv, use_container_width=True)

        with g2:
            st.markdown("**Velocidade de Vendas (Barra)**")
            fig_vel = px.bar(df_trend, x="DATA_STR", y="Velocidade", color="EMPREENDIMENTO", barmode="group",
                             color_discrete_sequence=px.colors.qualitative.Prism)
            fig_vel.update_layout(yaxis_tickformat=".1%", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
            st.plotly_chart(fig_vel, use_container_width=True)
            
            st.markdown("**Preco Medio (Barra)**")
            fig_prc = px.bar(df_trend, x="DATA_STR", y="PRECO_MEDIO", color="EMPREENDIMENTO", barmode="group",
                             color_discrete_sequence=px.colors.qualitative.Prism)
            fig_prc.update_layout(yaxis_tickprefix="R$ ", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
            st.plotly_chart(fig_prc, use_container_width=True)

    # -------------------------------------------------------------------------
    # TABELAS DE PRECO (BD DETALHADA)
    # -------------------------------------------------------------------------
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:2rem 0;'/>", unsafe_allow_html=True)
    
    # Filtrar BD Detalhada pelos empreendimentos selecionados
    df_det_f = df_detalhada[df_detalhada["EMPREENDIMENTO"].isin(nomes_selecionados)].copy()
    
    if not df_det_f.empty:
        df_det_f["Mes_Ano"] = df_det_f["DATA_DT"].dt.strftime("%m/%Y")
        
        # Tabela 1: Preço m2 MEDIANO por Empreendimento
        st.subheader("Tabela 1: Preco m2 Mediano por Empreendimento (BD DETALHADA)")
        pivot_mediano = pd.pivot_table(
            df_det_f, 
            values="Preco_m2", 
            index="EMPREENDIMENTO", 
            columns="Mes_Ano", 
            aggfunc="median"
        ).reset_index()
        
        # Reordenar colunas cronologicamente
        cols_datas = sorted([c for c in pivot_mediano.columns if c != "EMPREENDIMENTO"], 
                           key=lambda x: datetime.strptime(x, "%m/%Y"))
        pivot_mediano = pivot_mediano[["EMPREENDIMENTO"] + cols_datas]
        st.dataframe(pivot_mediano.style.format({c: "R$ {:.2f}" for c in cols_datas}), 
                     use_container_width=True, hide_index=True)

        # Tabela 2: Preço m2 MÉDIO por Tipologia
        st.subheader("Tabela 2: Preco m2 Medio por Tipologia (BD DETALHADA)")
        pivot_medio = pd.pivot_table(
            df_det_f, 
            values="Preco_m2", 
            index="TIPOLOGIA", 
            columns="Mes_Ano", 
            aggfunc="mean"
        ).reset_index()
        
        cols_datas_2 = sorted([c for c in pivot_medio.columns if c != "TIPOLOGIA"], 
                             key=lambda x: datetime.strptime(x, "%m/%Y"))
        pivot_medio = pivot_medio[["TIPOLOGIA"] + cols_datas_2]
        st.dataframe(pivot_medio.style.format({c: "R$ {:.2f}" for c in cols_datas_2}), 
                     use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum dado de precificacao detalhada encontrado para os empreendimentos selecionados.")

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.82rem;">Direcional Engenharia · Inteligencia de Mercado · Fontes: BD GERAL & BD DETALHADA</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
