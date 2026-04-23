# -*- coding: utf-8 -*-
"""
Análise de Concorrência — Inteligência Competitiva e Performance.
Pipeline: Data Cleaning -> Feature Engineering -> Strategic Analysis -> UI
Planilha: Bases Concorrentes + Direcional.
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
COR_FUNDO_CARD = "rgba(255, 255, 255, 0.78)"
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
# Funções de Design (Padrão Premium)
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
        @keyframes fichaShimmer {{ 0% {{ background-position: 0% 50%; }} 100% {{ background-position: 200% 50%; }} }}
        html, body, :root, [data-testid="stApp"] {{ color-scheme: light !important; }}
        html, body {{ font-family: 'Inter', sans-serif; color: {COR_TEXTO_LABEL}; background: transparent !important; }}
        .stApp, [data-testid="stApp"] {{
            background: linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%),
                url("{bg_url}") center / cover no-repeat !important;
            background-attachment: scroll !important;
        }}
        [data-testid="stAppViewContainer"] {{ background: transparent !important; }}
        header[data-testid="stHeader"] {{ background: transparent !important; border: none !important; box-shadow: none !important; }}
        [data-testid="stMain"] {{
            padding-left: clamp(14px, 5vw, 56px) !important; padding-right: clamp(14px, 5vw, 56px) !important;
            padding-top: clamp(12px, 3.5vh, 40px) !important; padding-bottom: clamp(14px, 4vh, 44px) !important;
        }}
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
        .ficha-hero {{ text-align: center; padding: 0.5rem 0 0 0; max-width: 640px; margin: 0 auto; animation: fichaFadeIn 0.85s both; }}
        .ficha-hero .ficha-title {{ font-family: 'Montserrat', sans-serif; font-size: clamp(1.35rem, 3.5vw, 1.75rem); font-weight: 900; color: {COR_AZUL_ESC}; margin: 0; }}
        .ficha-hero .ficha-sub {{ color: #475569; font-size: 0.95rem; margin: 0.45rem 0 0 0; }}
        .ficha-hero-bar-wrap {{ width: 100%; margin: clamp(0.85rem, 2.4vw, 1.2rem) 0; }}
        .ficha-hero-bar {{ height: 4px; border-radius: 999px; background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC}); background-size: 200% 100%; animation: fichaShimmer 4s infinite alternate; }}
        .vel-kpi-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 1.25rem; }}
        .vel-kpi {{
            flex: 1 1 20%; background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.9) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9); border-radius: 14px; padding: 14px 16px; text-align: center;
            box-shadow: 0 2px 8px rgba({RGB_AZUL_CSS}, 0.06); transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        .vel-kpi:hover {{ transform: translateY(-4px); box-shadow: 0 10px 20px -5px rgba({RGB_AZUL_CSS}, 0.15); }}
        .vel-kpi .lbl {{ font-size: 0.72rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: {COR_AZUL_ESC}; opacity: 0.85; }}
        .vel-kpi .val {{ font-family: 'Montserrat', sans-serif; font-size: 1.35rem; font-weight: 800; color: {COR_AZUL_ESC}; margin-top: 6px; }}
        div[data-baseweb="input"] {{ border-radius: 10px !important; border: 1px solid #e2e8f0 !important; background-color: {COR_INPUT_BG} !important; }}
        div[data-baseweb="input"]:focus-within {{ border-color: rgba({RGB_AZUL_CSS}, 0.35) !important; box-shadow: 0 0 0 3px rgba({RGB_AZUL_CSS}, 0.08) !important; }}
        </style>
        """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Pipeline de Dados (Pandas)
# -----------------------------------------------------------------------------
def _secrets_connections_gsheets() -> Dict[str, Any]:
    try:
        sec = st.secrets
        if hasattr(sec, "get") and sec.get("connections"):
            g = sec["connections"].get("gsheets")
            if g is not None: return dict(g)
    except Exception: pass
    return {}

def montar_service_account_info(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not raw: return None
    out = {k: v for k, v in raw.items() if v}
    if "private_key" in out: out["private_key"] = out["private_key"].replace("\\n", "\n")
    return out

@st.cache_data(ttl=300, show_spinner=False)
def ler_aba(spreadsheet_id: str, worksheet: str) -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials
    raw = _secrets_connections_gsheets()
    info = montar_service_account_info(raw)
    if not info: raise ValueError("Credenciais [connections.gsheets] ausentes.")
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(worksheet)
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    df.columns = [str(c).strip() for c in df.columns]
    return df

def parse_val(v):
    if not v: return 0.0
    s = str(v).replace("R$", "").replace(".", "").replace(",", ".").replace(" ", "").strip()
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

def process_data() -> pd.DataFrame:
    # 1. Carregamento
    df_det = ler_aba(SPREADSHEET_ID_CONC, "BD DETALHADA")
    df_ger = ler_aba(SPREADSHEET_ID_CONC, "BD GERAL")
    df_men = ler_aba(SPREADSHEET_ID_CONC, "Abr/2026")
    df_dados_dir = ler_aba(SPREADSHEET_ID_CONC, "DADOS DIRECIONAL")
    
    # 2. Data Cleaning
    df_det["Preço_Float"] = df_det["PREÇO"].apply(parse_val)
    df_det["Metragem_Float"] = df_det["METRAGEM"].apply(parse_val)
    df_det["Preço_m2"] = df_det["PREÇO_M2"].apply(parse_val)
    
    # Recalcular Preço_m2 se inconsistente
    df_det["Preço_m2"] = np.where(
        (df_det["Preço_m2"] == 0) & (df_det["Metragem_Float"] > 0),
        df_det["Preço_Float"] / df_det["Metragem_Float"],
        df_det["Preço_m2"]
    )
    
    df_ger["Preço_Min"] = df_ger["Venda a Partir"].apply(parse_val)
    
    # 3. Merges
    df_master = df_det.merge(df_ger[["CHAVE", "Preço_Min", "Previsão"]], on="CHAVE", how="left")
    df_master = df_master.merge(df_men[["CHAVE", "Vendas (Qnt.)", "Estoque (Qnt.)", "VGV (R$)", "PREÇO MÉDIO"]], on="CHAVE", how="left")
    
    # 4. Feature Engineering
    # Taxa de Absorção
    vendas = pd.to_numeric(df_master["Vendas (Qnt.)"], errors='coerce').fillna(0)
    estoque = pd.to_numeric(df_master["Estoque (Qnt.)"], errors='coerce').fillna(0)
    df_master["Absorcao"] = vendas / (vendas + estoque)
    df_master["Absorcao"] = df_master["Absorcao"].replace([np.inf, -np.inf], 0).fillna(0)
    
    # Desconto Implícito (Preço Unidade vs Preço Mínimo Projeto)
    df_master["Desconto"] = (df_master["Preço_Float"] - df_master["Preço_Min"]) / df_master["Preço_Min"]
    df_master["Desconto"] = df_master["Desconto"].fillna(0)
    
    # Posicionamento Relativo (Preço m2 vs Média Bairro)
    medias_bairro = df_master.groupby("BAIRRO")["Preço_m2"].transform("mean")
    df_master["Posicionamento_Rel"] = df_master["Preço_m2"] - medias_bairro
    
    # Densidade Competitiva
    df_master["Densidade_Bairro"] = df_master.groupby("BAIRRO")["CHAVE"].transform("nunique")
    
    # Faixas de Metragem
    df_master["Faixa_Metragem"] = pd.cut(df_master["Metragem_Float"], bins=[0, 35, 45, 55, 70, 1000], labels=["Até 35m²", "35-45m²", "45-55m²", "55-70m²", "70m²+"])
    
    # Identificar se é Direcional (Filtro por Endereço solicitado)
    allowed_addresses = [str(x).strip().upper() for x in df_dados_dir["Endereço"].unique() if x]
    df_master["Is_Direcional"] = df_master["ENDEREÇO"].str.strip().str.upper().isin(allowed_addresses)
    
    return df_master

# -----------------------------------------------------------------------------
# Interface e Páginas
# -----------------------------------------------------------------------------
def main():
    aplicar_estilo()
    _exibir_logo_topo()
    
    st.sidebar.markdown(f"<h3 style='color:{COR_AZUL_ESC};'>Menu Estratégico</h3>", unsafe_allow_html=True)
    page = st.sidebar.radio("Navegação", ["Visão Geral", "Mapa Competitivo", "Produto Ideal", "Pricing & Demanda", "Ranking Concorrentes", "Oportunidades (Insights)"])
    
    try:
        df_master = process_data()
    except Exception as e:
        st.error(f"Erro no Pipeline de Dados: {e}")
        return

    # Filtros Globais
    st.sidebar.markdown("---")
    f_regiao = st.sidebar.multiselect("Região", sorted(df_master["REGIÃO"].dropna().unique()))
    f_bairro = st.sidebar.multiselect("Bairro", sorted(df_master["BAIRRO"].dropna().unique()))
    
    df_f = df_master.copy()
    if f_regiao: df_f = df_f[df_f["REGIÃO"].isin(f_regiao)]
    if f_bairro: df_f = df_f[df_f["BAIRRO"].isin(f_bairro)]

    if page == "Visão Geral":
        st.markdown("## Visão Geral do Mercado")
        avg_m2 = df_f["Preço_m2"].mean()
        avg_abs = df_f["Absorcao"].mean() * 100
        num_conc = df_f["CONCORRENTE"].nunique()
        
        st.markdown(f"""
            <div class="vel-kpi-row">
                <div class="vel-kpi"><div class="lbl">Preço Médio / m²</div><div class="val">R$ {avg_m2:,.2f}</div></div>
                <div class="vel-kpi"><div class="lbl">Absorção Média</div><div class="val">{avg_abs:.1f}%</div></div>
                <div class="vel-kpi"><div class="lbl">Nº Concorrentes</div><div class="val">{num_conc}</div></div>
                <div class="vel-kpi"><div class="lbl">Oferta Ativa (Unid.)</div><div class="val">{int(df_f['CHAVE'].count())}</div></div>
            </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Market Share por Concorrente (VGV)")
            df_gv = df_f.groupby("CONCORRENTE")["Preço_Float"].sum().reset_index()
            fig = px.pie(df_gv, values='Preço_Float', names='CONCORRENTE', hole=.4, color_discrete_sequence=px.colors.sequential.RdBu)
            fig.update_layout(paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            st.subheader("Performance por Região")
            df_reg = df_f.groupby("REGIÃO")["Absorcao"].mean().reset_index()
            fig = px.bar(df_reg, x='REGIÃO', y='Absorcao', color_discrete_sequence=[COR_AZUL_ESC])
            fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    elif page == "Mapa Competitivo":
        st.markdown("## Mapa Competitivo e Saturação")
        col1, col2 = st.columns([2, 1])
        with col1:
            st.subheader("Distribuição de Preço/m² por Bairro")
            fig = px.box(df_f, x="BAIRRO", y="Preço_m2", color="CONCORRENTE", points="all")
            fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            st.subheader("Densidade de Projetos")
            df_dens = df_f.groupby("BAIRRO")["CHAVE"].nunique().reset_index().sort_values("CHAVE", ascending=False)
            st.dataframe(df_dens.rename(columns={"CHAVE": "Projetos Ativos"}), use_container_width=True, hide_index=True)

    elif page == "Produto Ideal":
        st.markdown("## Inteligência de Produto")
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("O que mais vende? (Tipologia)")
            df_tipo = df_f.groupby("TIPOLOGIA")["Absorcao"].mean().reset_index().sort_values("Absorcao", ascending=False)
            fig = px.bar(df_tipo, x="Absorcao", y="TIPOLOGIA", orientation='h', color_discrete_sequence=[COR_VERMELHO])
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            st.subheader("Metragem vs Velocidade")
            fig = px.scatter(df_f, x="Metragem_Float", y="Absorcao", color="CONCORRENTE", size="Preço_m2", hover_name="EMPREENDIMENTO")
            st.plotly_chart(fig, use_container_width=True)

    elif page == "Pricing & Demanda":
        st.markdown("## Curva de Demanda e Elasticidade")
        st.subheader("Relação Preço/m² vs Absorção")
        fig = px.scatter(df_f, x="Preço_m2", y="Absorcao", color="CONCORRENTE", trendline="ols", 
                         labels={"Preço_m2": "R$ / m²", "Absorcao": "Taxa de Absorção"},
                         hover_name="EMPREENDIMENTO", text="BAIRRO")
        st.plotly_chart(fig, use_container_width=True)
        st.info("💡 Pontos acima da linha de tendência representam produtos com alta eficiência de preço. Pontos abaixo indicam produtos 'Overpriced' para a demanda local.")

    elif page == "Ranking Concorrentes":
        st.markdown("## Benchmarking de Concorrentes")
        df_rank = df_f.groupby("CONCORRENTE").agg(
            Preco_m2=("Preço_m2", "mean"),
            Absorcao_Media=("Absorcao", "mean"),
            Volume_Vendas=("Vendas (Qnt.)", "sum"),
            N_Projetos=("CHAVE", "nunique")
        ).reset_index()
        
        # Eficiência = Absorção / (Preço_m2 / Média_Mercado) -> penaliza preço alto se a absorção for baixa
        avg_mkt_m2 = df_f["Preço_m2"].mean()
        df_rank["Eficiência"] = df_rank["Absorcao_Media"] / (df_rank["Preco_m2"] / avg_mkt_m2)
        
        st.dataframe(df_rank.sort_values("Eficiência", ascending=False).style.format({
            "Preco_m2": "R$ {:.2f}", "Absorcao_Media": "{:.1%}", "Eficiência": "{:.2f}"
        }), use_container_width=True, hide_index=True)

    elif page == "Oportunidades (Insights)":
        st.markdown("## 🎯 Matriz de Oportunidades")
        
        if not df_f.empty:
            df_insight = df_f.groupby("BAIRRO").agg(
                Metragem_Media=("Metragem_Float", "mean"),
                Preco_m2=("Preço_m2", "mean"),
                Abs_Media=("Absorcao", "mean"),
                Qtd_Projetos=("CHAVE", "nunique")
            ).reset_index()
            
            avg_mkt_abs = df_master["Absorcao"].mean()
            med_mkt_projects = df_insight["Qtd_Projetos"].median()
            
            def strategic_action(row):
                if row['Abs_Media'] > avg_mkt_abs and row['Qtd_Projetos'] <= med_mkt_projects:
                    return "🔥 OPORTUNIDADE: Alta demanda e baixa oferta. Lançar produto similar."
                if row['Abs_Media'] < avg_mkt_abs and row['Qtd_Projetos'] > med_mkt_projects:
                    return "⚠️ RISCO: Mercado saturado ou preço descolado. Evitar lançamentos."
                if row['Preco_m2'] < df_insight['Preco_m2'].mean() * 0.8:
                    return "💎 GAP DE PADRÃO: Preços muito baixos. Oportunidade para upgrade."
                return "MERCADO ESTÁVEL: Monitorar performance."

            df_insight["Ação Recomendada"] = df_insight.apply(strategic_action, axis=1)
            
            st.dataframe(df_insight.sort_values("Abs_Media", ascending=False), use_container_width=True, hide_index=True)

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.82rem;">Direcional Engenharia · Inteligência Competitiva · Developed by Lucas Maia</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
