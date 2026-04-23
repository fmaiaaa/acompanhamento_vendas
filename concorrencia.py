# -*- coding: utf-8 -*-
"""
Análise de Concorrência — Inteligência Competitiva e Performance.
Planilha: Bases Concorrentes + Direcional.
Dependências: streamlit, pandas, plotly, gspread, google-auth
"""
from __future__ import annotations

import base64
import html
import os
import re
import math
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# -----------------------------------------------------------------------------
# Identificação da planilha e Arquivos Visuais
# -----------------------------------------------------------------------------
SPREADSHEET_ID_CONC = "1nwRSz-ixnHncsT7UxRkjA7jBwe31ZiJdEfcM0Dm5aYE"

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"
URL_LOGO_DIRECIONAL_EMAIL = "https://logodownload.org/wp-content/uploads/2021/04/direcional-engenharia-logo.png"

# Paleta alinhada à Ficha Credenciamento / Gaps
COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_VERMELHO_ESCURO = "#9e0828"
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
# Funções de Design (Padrão Premium)
# -----------------------------------------------------------------------------
def _resolver_png_raiz(nome: str) -> Path | None:
    """Procura o PNG na pasta do app e na pasta pai."""
    for base in (_DIR_APP, _DIR_APP.parent):
        p = base / nome
        if p.is_file():
            return p
    return None

def _resolver_imagem_fundo_local(nome: str) -> Path | None:
    """Imagem JPG/PNG na pasta do app ou na pasta pai."""
    for base in (_DIR_APP, _DIR_APP.parent):
        for ext in (".jpg", ".jpeg", ".JPG", ".JPEG", ".png", ".PNG"):
            stem = Path(nome).stem
            p = base / f"{stem}{ext}"
            if p.is_file():
                return p
        p = base / nome
        if p.is_file():
            return p
    return None

def _css_url_fundo_cadastro() -> str:
    """String para `url(...)` no CSS: data-URL ou URL https (fallback)."""
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
    for name in ("logo_direcional.png", "logo_direcional.jpg", "logo_direcional.jpeg", "logo.png"):
        p = _DIR_APP / "assets" / name
        if p.is_file():
            return str(p)
    return None

def _logo_url_secrets() -> str | None:
    try:
        if hasattr(st, "secrets") and isinstance(st.secrets.get("branding"), dict):
            return (st.secrets["branding"].get("LOGO_URL") or "").strip() or None
    except Exception:
        pass
    return None

def _logo_url_drive_por_id_arquivo() -> str | None:
    fid = (os.environ.get("DIRECIONAL_LOGO_FILE_ID") or "").strip()
    if len(fid) < 10: return None
    return f"https://drive.google.com/uc?export=view&id={fid}"

def _exibir_logo_topo() -> None:
    """Logo centralizada no topo."""
    path = _logo_arquivo_local()
    url = _logo_url_secrets() or _logo_url_drive_por_id_arquivo()
    try:
        if path:
            ext = Path(path).suffix.lower().lstrip(".")
            mime = "image/png" if ext == "png" else "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            st.markdown(f'<div class="ficha-logo-wrap"><img src="data:{mime};base64,{b64}" alt="Direcional" /></div>', unsafe_allow_html=True)
            return
        if url:
            st.markdown(f'<div class="ficha-logo-wrap"><img src="{html.escape(url)}" alt="Direcional" /></div>', unsafe_allow_html=True)
    except Exception:
        pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Inteligência Competitiva e Performance</p>'
        f'<p class="ficha-sub">Análise Realizado X Projetado — Mercado e Concorrência.</p>'
        f"</div>"
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true"><div class="ficha-hero-bar"></div></div>'
        f"</div>",
        unsafe_allow_html=True,
    )

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
        [data-testid="stSidebar"] {{ display: none !important; }}
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
# Lógicas de Dados
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

# -----------------------------------------------------------------------------
# App Principal
# -----------------------------------------------------------------------------
def main():
    aplicar_estilo()
    _cabecalho_pagina()

    with st.spinner("Consolidando bases de inteligência competitiva..."):
        # 1. Carregando Bases
        try:
            df_detalhada = ler_aba(SPREADSHEET_ID_CONC, "BD DETALHADA")
            df_geral = ler_aba(SPREADSHEET_ID_CONC, "BD GERAL")
            df_mensal = ler_aba(SPREADSHEET_ID_CONC, "Abr/2026") # Mês base
            df_dados_dir = ler_aba(SPREADSHEET_ID_CONC, "DADOS DIRECIONAL")
            df_vendas_dir = ler_aba(SPREADSHEET_ID_CONC, "Controle de Vendas RJ - Periodo Integral")
        except Exception as e:
            st.error(f"Erro ao carregar abas: {e}")
            return

    # 2. Processamento e Enriquecimento Concorrentes
    df_detalhada["Preço_Float"] = df_detalhada["PREÇO"].apply(parse_val)
    df_detalhada["Preço_m2"] = df_detalhada["PREÇO_M2"].apply(parse_val)
    
    # Merge Master - CORREÇÃO: Usando 'EMPREENDIMENTO' em vez de 'EMPREENDIMENTOVendas' para evitar KeyError
    df_master = df_detalhada.merge(
        df_geral[["Empreendimento", "Venda a Partir", "Previsão"]], 
        left_on="EMPREENDIMENTO", 
        right_on="Empreendimento", 
        how="left", 
        suffixes=('', '_geral')
    )
    df_master = df_master.merge(df_mensal[["CHAVE", "Vendas (Qnt.)", "Estoque (Qnt.)", "VGV (R$)"]], on="CHAVE", how="left")
    
    # KPI: Taxa de Absorção
    df_master["Vendas_Num"] = pd.to_numeric(df_master["Vendas (Qnt.)"], errors="coerce").fillna(0)
    df_master["Estoque_Num"] = pd.to_numeric(df_master["Estoque (Qnt.)"], errors="coerce").fillna(0)
    # Absorção = Vendas / (Vendas + Estoque)
    df_master["Absorcao"] = df_master["Vendas_Num"] / (df_master["Vendas_Num"] + df_master["Estoque_Num"])
    df_master["Absorcao"] = df_master["Absorcao"].fillna(0)

    # -------------------------------------------------------------------------
    # Filtros Estratégicos
    # -------------------------------------------------------------------------
    st.markdown("<div style='margin-bottom:1rem; text-align: center;'><strong>Filtros Estratégicos de Mercado</strong></div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: f_regiao = st.multiselect("Região", sorted(df_master["REGIÃO"].dropna().unique()))
    with c2: f_bairro = st.multiselect("Bairro", sorted(df_master["BAIRRO"].dropna().unique()))
    with c3: f_conc = st.multiselect("Concorrente", sorted(df_master["CONCORRENTE"].dropna().unique()))
    with c4: f_tipo = st.multiselect("Tipologia", sorted(df_master["TIPOLOGIA"].dropna().unique()))

    df_f = df_master.copy()
    if f_regiao: df_f = df_f[df_f["REGIÃO"].isin(f_regiao)]
    if f_bairro: df_f = df_f[df_f["BAIRRO"].isin(f_bairro)]
    if f_conc: df_f = df_f[df_f["CONCORRENTE"].isin(f_conc)]
    if f_tipo: df_f = df_f[df_f["TIPOLOGIA"].isin(f_tipo)]

    # -------------------------------------------------------------------------
    # KPIs Mercado
    # -------------------------------------------------------------------------
    avg_m2 = df_f["Preço_m2"].mean()
    avg_abs = df_f["Absorcao"].mean() * 100
    total_vendas = df_f["Vendas_Num"].sum()
    
    st.markdown(f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Preço Médio / m²</div><div class="val">R$ {avg_m2:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Absorção Média</div><div class="val">{avg_abs:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">Vendas no Mês (Segmento)</div><div class="val">{int(total_vendas)}</div></div>
            <div class="vel-kpi"><div class="lbl">Nº Projetos Ativos</div><div class="val">{df_f['CHAVE'].nunique()}</div></div>
        </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # Gráficos de Inteligência
    # -------------------------------------------------------------------------
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("Curva de Demanda: Preço vs Absorção")
        fig_demand = go.Figure(data=go.Scatter(
            x=df_f["Preço_m2"], y=df_f["Absorcao"] * 100,
            mode='markers',
            text=df_f["EMPREENDIMENTO"],
            marker=dict(size=14, color=df_f["Preço_Float"], colorscale='Viridis', showscale=True, line=dict(width=1, color='white'))
        ))
        fig_demand.update_layout(xaxis_title="R$ / m²", yaxis_title="Absorção (%)", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_demand, use_container_width=True, config={"displayModeBar": False})

    with col_g2:
        st.subheader("Eficiência Média por Concorrente")
        df_eff = df_f.groupby("CONCORRENTE").agg({"Absorcao": "mean"}).sort_values("Absorcao", ascending=False)
        fig_eff = go.Figure(go.Bar(
            x=df_eff.index, y=df_eff["Absorcao"] * 100,
            marker_color=COR_AZUL_ESC
        ))
        fig_eff.update_layout(yaxis_title="Absorção Média (%)", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_eff, use_container_width=True, config={"displayModeBar": False})

    # -------------------------------------------------------------------------
    # Oportunidades e Gaps de Mercado
    # -------------------------------------------------------------------------
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Gap de Mercado: Análise por Bairro")
    
    if not df_f.empty:
        df_gap = df_f.groupby("BAIRRO").agg(
            Oferta_Ativa=("CHAVE", "nunique"),
            Preco_m2_Medio=("Preço_m2", "mean"),
            Absorcao_Media=("Absorcao", "mean")
        ).reset_index()
        
        med_oferta = df_gap['Oferta_Ativa'].median()
        
        def identificar_status(row):
            if row['Absorcao_Media'] > (avg_abs/100) and row['Oferta_Ativa'] <= med_oferta:
                return "🔥 Oportunidade: Alta Demanda / Baixa Oferta"
            if row['Absorcao_Media'] < (avg_abs/100) and row['Oferta_Ativa'] > med_oferta:
                return "⚠️ Atenção: Mercado Saturado"
            return "Alinhado ao Mercado"

        df_gap["Sugestão Estratégica"] = df_gap.apply(identificar_status, axis=1)
        df_gap["Absorcao_Media"] = (df_gap["Absorcao_Media"] * 100).map("{:.1f}%".format)
        df_gap["Preco_m2_Medio"] = df_gap["Preco_m2_Medio"].map("R$ {:,.2f}".format)
        
        st.dataframe(df_gap.sort_values("Oferta_Ativa", ascending=False), use_container_width=True, hide_index=True)

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.82rem;">Direcional Engenharia · Inteligência de Mercado</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
