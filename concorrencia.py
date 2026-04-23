# -*- coding: utf-8 -*-
"""
Análise de Concorrência — Inteligência Competitiva e Performance.
Foco: Estudo de Cluster Regional por Empreendimento Direcional.
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
# Funções de Design
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
    return "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?auto=format&fit=crop&w=1920&q=80"

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
        .block-container {{
            max-width: 1700px !important; margin-left: auto !important; margin-right: auto !important;
            padding: 1.45rem 2.25rem 1.55rem 2.25rem !important; background: rgba(255, 255, 255, 0.78) !important;
            backdrop-filter: blur(18px) saturate(1.15); border-radius: 24px !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            box-shadow: 0 4px 6px -1px rgba({RGB_AZUL_CSS}, 0.06), 0 24px 48px -12px rgba({RGB_AZUL_CSS}, 0.18) !important;
            animation: fichaFadeIn 0.7s cubic-bezier(0.22, 1, 0.36, 1) both;
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
            mime = "image/png" if ext == "png" else "image/jpeg"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            st.markdown(f'<div class="ficha-logo-wrap"><img src="data:{mime};base64,{b64}" alt="Direcional" /></div>', unsafe_allow_html=True)
    except Exception: pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Estudo de Cluster Competitivo</p>'
        f'<p class="ficha-sub">Performance Real Direcional vs Concorrência Regional.</p>'
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
    s = str(v).replace("R$", "").replace(".", "").replace(",", ".").replace(" ", "").strip()
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

@st.cache_data(ttl=300, show_spinner=False)
def load_all_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    import gspread
    from google.oauth2.service_account import Credentials
    
    sec = st.secrets["connections"]["gsheets"]
    info = {k: v for k, v in sec.items() if v}
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID_CONC)
    
    def get_df(name):
        ws = sh.worksheet(name)
        data = ws.get_all_values()
        df = pd.DataFrame(data[1:], columns=data[0])
        df.columns = [str(c).strip() for c in df.columns]
        return df

    return get_df("BD DETALHADA"), get_df("BD GERAL"), get_df("Abr/2026"), get_df("Automação Estudos Concorrentes - Lucas"), get_df("Controle de Vendas RJ - Periodo Integral"), get_df("DADOS DIRECIONAL")

def process_pipeline():
    df_det, df_ger, df_men, df_est_dir, df_ven_dir, df_dados_ref = load_all_data()
    
    # 1. Tratar Concorrência (Até 10k/m2)
    df_det["Preço_Float"] = df_det["PREÇO"].apply(parse_val)
    df_det["Metragem_Float"] = df_det["METRAGEM"].apply(parse_val)
    df_det["Preço_m2"] = np.where(df_det["Metragem_Float"] > 0, df_det["Preço_Float"] / df_det["Metragem_Float"], 0)
    df_det = df_det[df_det["Preço_m2"] <= 10000].copy()
    
    df_master = df_det.merge(df_men[["CHAVE", "Vendas (Qnt.)", "Estoque (Qnt.)", "VGV (R$)"]], on="CHAVE", how="left")
    
    # 2. Tratar Dados Direcional (Estoque)
    status_estoque = ["Disponível", "Fora de venda", "Fora de venda - Comercial", "Mirror"]
    df_est_dir = df_est_dir[df_est_dir["Status da unidade"].isin(status_estoque)].copy()
    
    def calc_liquido(row):
        v = parse_val(row["Valor Final Com Kit"])
        f = parse_val(row["Folga de Tabela"])
        b = parse_val(row["Bônus Adimplência"])
        return v - f - b
    
    df_est_dir["Valor_Liquido"] = df_est_dir.apply(calc_liquido, axis=1)
    df_est_dir["Area_Float"] = df_est_dir["Área privativa total"].apply(parse_val)
    
    # 3. Tratar Dados Direcional (Vendas)
    df_ven_dir = df_ven_dir[pd.to_numeric(df_ven_dir["Venda Comercial?"], errors='coerce') == 1].copy()
    df_ven_dir["VGV_Real"] = df_ven_dir["Valor Real de Venda"].apply(parse_val)
    df_ven_dir["Data_Venda_DT"] = pd.to_datetime(df_ven_dir["Data da venda"], dayfirst=True, errors='coerce')
    
    return df_master, df_est_dir, df_ven_dir, df_dados_ref

# -----------------------------------------------------------------------------
# Main App
# -----------------------------------------------------------------------------
def main():
    aplicar_estilo()
    _cabecalho_pagina()
    
    try:
        df_conc, df_est_dir, df_ven_dir, df_ref = process_pipeline()
    except Exception as e:
        st.error(f"Erro no processamento das bases: {e}")
        return

    # Filtros de Topo
    st.markdown("<div style='text-align: center; font-weight: bold; margin-bottom: 1rem;'>Configuração do Estudo de Cluster</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    
    with c1:
        lista_alvos = sorted(df_ref["Nome do Empreendimento (Chave)"].unique())
        alvo_selecionado = st.selectbox("Selecione o Empreendimento Direcional para Estudo", lista_alvos)
        
    with c2:
        anos_disp = sorted(df_ven_dir["Data_Venda_DT"].dt.year.dropna().unique().astype(int))
        ano_estudo = st.selectbox("Ano de Referência das Vendas", anos_disp, index=len(anos_disp)-1 if anos_disp else 0)

    # Lógica de Cluster Regional
    info_alvo = df_ref[df_ref["Nome do Empreendimento (Chave)"] == alvo_selecionado].iloc[0]
    regiao_cluster = info_alvo["Região"]
    
    # Filtro Mercado
    df_f_conc = df_conc[df_conc["REGIÃO"] == regiao_cluster].copy()
    
    # Cálculos Reais Direcional (Alvo)
    vendas_alvo = len(df_ven_dir[(df_ven_dir["Empreendimento"] == alvo_selecionado) & (df_ven_dir["Data_Venda_DT"].dt.year == ano_estudo)])
    estoque_alvo = len(df_est_dir[df_est_dir["Nome do Empreendimento"] == alvo_selecionado])
    
    # Métrica de Absorção Real Direcional
    absorcao_real_dir = (vendas_alvo / (vendas_alvo + estoque_alvo)) * 100 if (vendas_alvo + estoque_alvo) > 0 else 0
    
    # Preço m2 Médio Direcional Real
    m2_real_dir = df_est_dir[df_est_dir["Nome do Empreendimento"] == alvo_selecionado]["Valor_Liquido"].sum() / \
                  df_est_dir[df_est_dir["Nome do Empreendimento"] == alvo_selecionado]["Area_Float"].sum() \
                  if df_est_dir[df_est_dir["Nome do Empreendimento"] == alvo_selecionado]["Area_Float"].sum() > 0 else 0

    st.markdown(f"### Estudo de Viabilidade: Cluster {regiao_cluster}")
    
    # KPIs Comparativos
    avg_m2_conc = df_f_conc[df_f_conc["CHAVE"] != alvo_selecionado]["Preço_m2"].mean()
    avg_abs_conc = pd.to_numeric(df_f_conc["Vendas (Qnt.)"], errors='coerce').fillna(0).mean() # Simplificação para escala
    
    st.markdown(f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Preço m² Médio (Cluster)</div><div class="val">R$ {avg_m2_conc:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Preço m² Real ({alvo_selecionado})</div><div class="val val--red">R$ {m2_real_dir:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Vendas Direcional ({ano_estudo})</div><div class="val val--red">{vendas_alvo} unid.</div></div>
            <div class="vel-kpi"><div class="lbl">Estoque Direcional Atual</div><div class="val">{estoque_alvo} unid.</div></div>
            <div class="vel-kpi"><div class="lbl">Absorção Real Direcional</div><div class="val val--red">{absorcao_real_dir:.1f}%</div></div>
        </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # MAPA E GRÁFICOS
    # -------------------------------------------------------------------------
    col_mapa, col_graf = st.columns([1.5, 1])
    
    with col_mapa:
        st.subheader("Mapa de Calor do Cluster Competitivo")
        # Como não temos Lat/Long reais, geramos uma visualização baseada na dispersão de tipologias por endereço único
        df_map = df_f_conc.groupby("ENDEREÇO").agg({"Preço_m2": "mean", "EMPREENDIMENTO": "first"}).reset_index()
        # Nota: O Streamlit requer lat/lon. Usaremos um gráfico de dispersão simulando o cluster se as coordenadas não existirem.
        fig_cluster = px.scatter(df_f_conc, x="Metragem_Float", y="Preço_m2", color="CONCORRENTE", size="Preço_Float",
                                 hover_name="EMPREENDIMENTO", title="Dispersão de Produtos no Cluster")
        fig_cluster.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_cluster, use_container_width=True)

    with col_graf:
        st.subheader("Curva de Demanda Regional")
        fig_dem = px.scatter(df_f_conc, x="Preço_m2", y="Vendas (Qnt.)", color="CONCORRENTE", trendline="ols",
                             labels={"Vendas (Qnt.)": "Vendas no Mês", "Preço_m2": "R$ / m²"})
        fig_dem.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_dem, use_container_width=True)

    # -------------------------------------------------------------------------
    # TABELA COMPARATIVA
    # -------------------------------------------------------------------------
    st.markdown("### 📊 Tabela de Benchmarking Detalhado")
    
    # Unificando a visão
    df_tab = df_f_conc.groupby("EMPREENDIMENTO").agg({
        "CONCORRENTE": "first",
        "TIPOLOGIA": "first",
        "Metragem_Float": "mean",
        "Preço_m2": "mean",
        "Vendas (Qnt.)": "first"
    }).reset_index().sort_values("Preço_m2", ascending=False)
    
    st.dataframe(df_tab.rename(columns={
        "EMPREENDIMENTO": "Produto",
        "Metragem_Float": "m² Médio",
        "Preço_m2": "R$/m²",
        "Vendas (Qnt.)": "Vendas Mês"
    }).style.format({"R$/m²": "R$ {:.2f}", "m² Médio": "{:.1f}"}), use_container_width=True, hide_index=True)

    # Diagnóstico Final
    st.markdown("<br>### 💡 Diagnóstico de Competitividade", unsafe_allow_html=True)
    if m2_real_dir < avg_m2_conc:
        st.success(f"💰 **Vantagem de Custo:** O m² da Direcional está {((avg_m2_conc/m2_real_dir)-1)*100:.1f}% abaixo da média do cluster. Há espaço para valorização.")
    else:
        st.warning(f"⚠️ **Pressão de Preço:** O m² da Direcional está acima da média regional. Focar na diferenciação das áreas comuns e tipologia.")

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.82rem;">Direcional Engenharia · Inteligência de Mercado · Estudo por Cluster Regional (Proxy 15km)</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
