# -*- coding: utf-8 -*-
"""
Análise de Concorrência — Inteligência Competitiva e Performance.
Foco: Comparação Direta via BD GERAL e BD DETALHADA.
Recursos: Evolução de Preço m2 vs Estoque, Gaps de 2 meses e Design Gaps.
Fórmulas: Explicitação matemática abaixo dos gráficos.
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
from plotly.subplots import make_subplots

# -----------------------------------------------------------------------------
# Identificação da folha de cálculo e Ficheiros Visuais
# -----------------------------------------------------------------------------
SPREADSHEET_ID_CONC = "1nwRSz-ixnHncsT7UxRkjA7jBwe31ZiJdEfcM0Dm5aYE"

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"
URL_LOGO_DIRECIONAL_EMAIL = "https://logodownload.org/wp-content/uploads/2021/04/direcional-engenharia-logo.png"

# Paleta alinhada à Análise de Gaps
COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_FUNDO_CARD = "rgba(255, 255, 255, 0.78)"
COR_BORDA = "#eef2f6"
COR_TEXTO_MUTED = "#64748b"
COR_TEXTO_LABEL = "#1e293b"
COR_INPUT_BG = "#f0f2f6"

def _hex_rgb_triplet(hex_color: str) -> str:
    """Converte #RRGGBB em 'r, g, b' para uso em rgba(...)."""
    x = (hex_color or "").strip().lstrip("#")
    if len(x) != 6:
        return "0, 0, 0"
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}"

RGB_AZUL_CSS = _hex_rgb_triplet(COR_AZUL_ESC)
RGB_VERMELHO_CSS = _hex_rgb_triplet(COR_VERMELHO)

# -----------------------------------------------------------------------------
# Funções de Design (Padrão Gaps)
# -----------------------------------------------------------------------------
def _resolver_png_raiz(nome: str) -> Path | None:
    for base in (_DIR_APP, _DIR_APP.parent):
        p = base / nome
        if p.is_file():
            return p
    return None

def _resolver_imagem_fundo_local(nome: str) -> Path | None:
    for base in (_DIR_APP, _DIR_APP.parent):
        for ext in (".jpg", ".jpeg", ".png", ".PNG"):
            stem = Path(nome).stem
            p = base / f"{stem}{ext}"
            if p.is_file():
                return p
        p = base / nome
        if p.is_file():
            return p
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
        except OSError:
            pass
    return "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?auto=format&fit=crop&w=1920&q=80"

def _logo_arquivo_local() -> str | None:
    p_topo = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
    if p_topo:
        return str(p_topo)
    return None

def _exibir_logo_topo() -> None:
    path = _logo_arquivo_local()
    try:
        if path:
            ext = Path(path).suffix.lower().lstrip(".")
            mime = "image/jpeg" if ext == "png" else "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            st.markdown(f'<div class="ficha-logo-wrap"><img src="data:{mime};base64,{b64}" alt="Direcional" /></div>', unsafe_allow_html=True)
    except Exception:
        pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Inteligência Competitiva - Estudo Direcional</p>'
        f'<p class="ficha-sub">Performance via <strong>BD GERAL</strong> e Precificação via <strong>BD DETALHADA</strong>.</p>'
        f"</div>"
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true"><div class="ficha-hero-bar"></div></div>'
        f"</div>",
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
        html, body {{
            font-family: 'Inter', sans-serif; color: {COR_TEXTO_LABEL};
            background: transparent !important; background-color: transparent !important;
        }}
        .stApp, [data-testid="stApp"] {{
            background: linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%),
                url("{bg_url}") center / cover no-repeat !important;
            background-attachment: scroll !important; background-color: transparent !important;
        }}
        [data-testid="stAppViewContainer"] {{ background: transparent !important; background-color: transparent !important; }}
        header[data-testid="stHeader"] {{ background: transparent !important; border: none !important; box-shadow: none !important; }}
        [data-testid="stSidebar"] {{ display: none !important; }}
        [data-testid="stMain"] {{
            padding-left: clamp(14px, 5vw, 56px) !important; padding-right: clamp(14px, 5vw, 56px) !important;
            padding-top: clamp(12px, 3.5vh, 40px) !important; padding-bottom: clamp(14px, 4vh, 44px) !important;
        }}
        .block-container {{
            max-width: 1700px !important; margin-left: auto !important; margin-right: auto !important;
            margin-top: clamp(4px, 1vh, 14px) !important; margin-bottom: clamp(4px, 1vh, 14px) !important;
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
        .vel-kpi .val--red {{ color: {COR_VERMELHO} !important; }}
        div[data-baseweb="input"], div[data-baseweb="select"] {{ border-radius: 10px !important; border: 1px solid #e2e8f0 !important; background-color: {COR_INPUT_BG} !important; }}
        /* Centralizar fórmulas LaTeX */
        .stLatex {{ text-align: center !important; display: flex; justify-content: center; margin-bottom: 1.5rem; }}
        </style>
        """,
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
        if ws.title in ["BD GERAL", "BD DETALHADA"]:
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
    
    if df_geral.empty: raise ValueError("BD GERAL não encontrada.")

    # Normalização
    for df in [df_geral, df_det]:
        if "CHAVE" in df.columns: df["CHAVE"] = df["CHAVE"].astype(str).str.strip().str.upper()
        if "CONSTRUTORA" in df.columns: df["CONSTRUTORA"] = df["CONSTRUTORA"].astype(str).str.strip().str.upper()
        if "EMPREENDIMENTO" in df.columns: df["EMPREENDIMENTO"] = df["EMPREENDIMENTO"].astype(str).str.strip().str.upper()
        if "CONCORRE COM" in df.columns: df["CONCORRE COM"] = df["CONCORRE COM"].astype(str).str.strip().str.upper()

    cols_num = ["VENDAS", "ESTOQUE", "ESTOQUE INICIAL"]
    for col in cols_num: df_geral[col] = pd.to_numeric(df_geral[col], errors='coerce').fillna(0)
    
    df_geral["PRECO_MEDIO"] = df_geral["PREÇO MÉDIO"].apply(parse_val)
    df_geral["DATA_DT"] = pd.to_datetime(df_geral["DATA"], dayfirst=True, errors='coerce')
    df_geral = df_geral[df_geral["DATA_DT"].notna()].sort_values(["CHAVE", "DATA_DT"])
    
    # Cálculos dos Indicadores (Fórmulas solicitadas)
    df_geral["Absorcao"] = df_geral["VENDAS"] / df_geral["ESTOQUE INICIAL"].replace(0, np.nan)
    df_geral["Velocidade"] = df_geral["VENDAS"] / (df_geral["ESTOQUE INICIAL"] - df_geral["VENDAS"]).replace(0, np.nan)
    df_geral["Escoamento"] = df_geral["ESTOQUE"] / df_geral["ESTOQUE INICIAL"].replace(0, np.nan)
    df_geral["Delta_Preco"] = df_geral.groupby("CHAVE")["PRECO_MEDIO"].pct_change()
    
    df_det["Preco_m2"] = df_det["PREÇO_M2"].apply(parse_val)
    df_det["DATA_DT"] = pd.to_datetime(df_det["DATA"], dayfirst=True, errors='coerce')
    return df_geral, df_det

# -----------------------------------------------------------------------------
# Main Application
# -----------------------------------------------------------------------------
def main():
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(page_title="Análise de Concorrência | Direcional", page_icon=str(fav) if fav else None, layout="wide")
    aplicar_estilo()
    _cabecalho_pagina()
    
    try:
        df_master, df_detalhada = process_pipeline()
    except Exception as e:
        st.error(f"Erro no pipeline: {e}")
        return

    st.markdown("<div style='text-align: center; font-weight: bold; margin-bottom: 1rem;'>Configuração da Análise Direta</div>", unsafe_allow_html=True)
    df_direcional = df_master[df_master["CONSTRUTORA"] == "DIRECIONAL"]
    lista_alvos = sorted(df_direcional["EMPREENDIMENTO"].unique())
    if not lista_alvos:
        st.error("Nenhum projeto DIRECIONAL localizado.")
        return
    alvo_selecionado = st.selectbox("Escolha o Produto Direcional", lista_alvos)

    # Lógica de Concorrência Automática via Coluna
    df_cluster = df_master[
        (df_master["EMPREENDIMENTO"] == alvo_selecionado) | 
        (df_master["CONCORRE COM"].str.contains(alvo_selecionado, na=False, regex=False))
    ].copy()

    df_sorted_dates = df_master[["DATA", "DATA_DT"]].drop_duplicates().sort_values("DATA_DT")
    meses_disp = df_sorted_dates["DATA"].tolist()
    mes_estudo = st.selectbox("Mês de Referência", meses_disp, index=len(meses_disp)-1)

    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)

    df_mes = df_cluster[df_cluster["DATA"] == mes_estudo]
    df_alvo_mes = df_mes[df_mes["EMPREENDIMENTO"] == alvo_selecionado]
    abs_alvo = df_alvo_mes["Absorcao"].iloc[0] * 100 if not df_alvo_mes.empty else 0
    m2_alvo = df_alvo_mes["PRECO_MEDIO"].iloc[0] if not df_alvo_mes.empty else 0
    m2_conc = df_mes[df_mes["EMPREENDIMENTO"] != alvo_selecionado]["PRECO_MEDIO"].mean() if not df_mes[df_mes["EMPREENDIMENTO"] != alvo_selecionado].empty else 0

    st.markdown(f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Preço Médio Alvo</div><div class="val val--red">R$ {m2_alvo:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Preço Médio Concorrentes</div><div class="val">R$ {m2_conc:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Taxa de Absorção (Alvo)</div><div class="val val--red">{abs_alvo:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">Total Concorrentes</div><div class="val">{df_mes[df_mes["EMPREENDIMENTO"] != alvo_selecionado]["EMPREENDIMENTO"].nunique()}</div></div>
        </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # INDICADORES DE PERFORMANCE (BD GERAL)
    # -------------------------------------------------------------------------
    st.subheader("Performance de Vendas e Preço (BD GERAL)")
    if not df_cluster.empty:
        df_trend = df_cluster.sort_values("DATA_DT")
        df_trend["DATA_STR"] = df_trend["DATA_DT"].dt.strftime("%m/%Y")
        
        g1, g2 = st.columns(2)
        with g1:
            # Grafico 1: Absorção
            fig_abs = px.bar(df_trend, x="DATA_STR", y="Absorcao", color="EMPREENDIMENTO", 
                             title="Taxa de Absorção", barmode="group", color_discrete_sequence=px.colors.qualitative.Prism)
            fig_abs.update_layout(title={'x':0.5, 'xanchor': 'center'}, yaxis_tickformat=".1%", 
                                  paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
            st.plotly_chart(fig_abs, use_container_width=True)
            st.latex(r"Absorção_t = \frac{Vendas_t}{Estoque_{t-1} + Vendas_t}")
            
            # Grafico 2: Escoamento
            fig_esc = px.bar(df_trend, x="DATA_STR", y="Escoamento", color="EMPREENDIMENTO", 
                             title="Taxa de Escoamento", barmode="group", color_discrete_sequence=px.colors.qualitative.Prism)
            fig_esc.update_layout(title={'x':0.5, 'xanchor': 'center'}, yaxis_tickformat=".1%", 
                                  paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
            st.plotly_chart(fig_esc, use_container_width=True)
            st.latex(r"Escoamento_t = \frac{Estoque_t}{Estoque_{Inicial}}")
            
        with g2:
            # Grafico 3: Velocidade
            fig_vel = px.bar(df_trend, x="DATA_STR", y="Velocidade", color="EMPREENDIMENTO", 
                             title="Velocidade de Vendas", barmode="group", color_discrete_sequence=px.colors.qualitative.Prism)
            fig_vel.update_layout(title={'x':0.5, 'xanchor': 'center'}, yaxis_tickformat=".1%", 
                                  paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
            st.plotly_chart(fig_vel, use_container_width=True)
            st.latex(r"Velocidade_t = \frac{Vendas_t}{Estoque_{t-1}}")
            
            # Grafico 4: Variação Preço
            fig_delta = px.bar(df_trend, x="DATA_STR", y="Delta_Preco", color="EMPREENDIMENTO", 
                               title="Variação de Preço (MoM)", barmode="group", color_discrete_sequence=px.colors.qualitative.Prism)
            fig_delta.update_layout(title={'x':0.5, 'xanchor': 'center'}, yaxis_tickformat=".1%", 
                                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
            st.plotly_chart(fig_delta, use_container_width=True)
            st.latex(r"\Delta Preço_{m2} = \frac{Preço_t - Preço_{t-1}}{Preço_{t-1}}")

    # -------------------------------------------------------------------------
    # NOVO GRÁFICO: EVOLUÇÃO PREÇO M2 VS ESTOQUE
    # -------------------------------------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    
    df_det_f = df_detalhada[df_detalhada["EMPREENDIMENTO"].isin(df_cluster["EMPREENDIMENTO"].unique())].copy()
    
    if not df_det_f.empty and not df_cluster.empty:
        # Agrupar Preço m2 por data (Mediana do Cluster)
        df_price_trend = df_det_f.groupby("DATA_DT").agg(Preco_m2_Mediana=("Preco_m2", "median")).reset_index()
        # Agrupar Estoque por data (Soma do Cluster)
        df_stock_trend = df_cluster.groupby("DATA_DT").agg(Estoque_Total=("ESTOQUE", "sum")).reset_index()
        
        df_merged_trend = pd.merge(df_price_trend, df_stock_trend, on="DATA_DT", how="outer").sort_values("DATA_DT")
        df_merged_trend["DATA_STR"] = df_merged_trend["DATA_DT"].dt.strftime("%m/%Y")
        
        fig_dual = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig_dual.add_trace(
            go.Scatter(x=df_merged_trend["DATA_STR"], y=df_merged_trend["Preco_m2_Mediana"], 
                       name="Preço m² Mediano (R$)", line=dict(color=COR_VERMELHO, width=4)),
            secondary_y=False,
        )
        
        fig_dual.add_trace(
            go.Bar(x=df_merged_trend["DATA_STR"], y=df_merged_trend["Estoque_Total"], 
                   name="Estoque Total (Unid.)", marker_color=COR_AZUL_ESC, opacity=0.4),
            secondary_y=True,
        )
        
        fig_dual.update_layout(
            title={'text': "Evolução do Mercado: Preço m² (Mediana) e Estoque (Soma)", 'x':0.5, 'xanchor': 'center'},
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=20, r=20, t=50, b=20),
            font=dict(family="Inter", color=COR_TEXTO_LABEL)
        )
        
        fig_dual.update_yaxes(title_text="Preço m² (R$)", secondary_y=False, gridcolor="rgba(226, 232, 240, 0.4)")
        fig_dual.update_yaxes(title_text="Estoque (Unid.)", secondary_y=True)
        
        st.plotly_chart(fig_dual, use_container_width=True, config={"displayModeBar": False})

    # -------------------------------------------------------------------------
    # TABELAS DE PREÇO (BD DETALHADA) - 2 ÚLTIMOS MESES
    # -------------------------------------------------------------------------
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:2rem 0;'/>", unsafe_allow_html=True)
    
    if not df_det_f.empty:
        df_det_f["Mes_Ano"] = df_det_f["DATA_DT"].dt.strftime("%m/%Y")
        all_months = sorted(df_det_f["Mes_Ano"].unique(), key=lambda x: datetime.strptime(x, "%m/%Y"))
        last_two = all_months[-2:] if len(all_months) >= 2 else all_months
        
        st.subheader("Mediana do Preço m² (Últimos 2 meses)")
        p_med = pd.pivot_table(df_det_f, values="Preco_m2", index="EMPREENDIMENTO", columns="Mes_Ano", aggfunc="median")[last_two].reset_index()
        if len(last_two) == 2: 
            p_med["Gap (%)"] = ((p_med[last_two[1]] / p_med[last_two[0]]) - 1) * 100
        st.dataframe(p_med.style.format({**{m: "R$ {:.2f}" for m in last_two}, "Gap (%)": "{:.1f}%"}), use_container_width=True, hide_index=True)

        st.subheader("Média do Preço m² por Tipologia (Últimos 2 meses)")
        p_tip = pd.pivot_table(df_det_f, values="Preco_m2", index="TIPOLOGIA", columns="Mes_Ano", aggfunc="mean")[last_two].reset_index()
        if len(last_two) == 2: 
            p_tip["Gap (%)"] = ((p_tip[last_two[1]] / p_tip[last_two[0]]) - 1) * 100
        st.dataframe(p_tip.style.format({**{m: "R$ {:.2f}" for m in last_two}, "Gap (%)": "{:.1f}%"}), use_container_width=True, hide_index=True)

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.82rem;">Direcional Engenharia · Inteligência de Mercado</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
