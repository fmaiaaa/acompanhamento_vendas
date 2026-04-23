# -*- coding: utf-8 -*-
"""
Análise de Concorrência — Inteligência Competitiva e Performance.
Foco: Estudo de Raio de Atuação por Empreendimento Direcional.
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
# Funções de Design (Padrão Premium Gaps)
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
        .ficha-hero {{ text-align: center; padding: 0.5rem 0 0 0; max-width: 640px; margin: 0 auto; }}
        .ficha-hero .ficha-title {{ font-family: 'Montserrat', sans-serif; font-size: clamp(1.35rem, 3.5vw, 1.75rem); font-weight: 900; color: {COR_AZUL_ESC}; margin: 0; }}
        .ficha-hero .ficha-sub {{ color: #475569; font-size: 0.95rem; margin: 0.45rem 0 0 0; }}
        .ficha-hero-bar-wrap {{ width: 100%; margin: clamp(0.85rem, 2.4vw, 1.2rem) 0; }}
        .ficha-hero-bar {{ height: 4px; border-radius: 999px; background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC}); background-size: 200% 100%; animation: fichaShimmer 4s infinite alternate; }}
        .vel-kpi-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 1.25rem; }}
        .vel-kpi {{
            flex: 1 1 18%; background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.82) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9); border-radius: 14px; padding: 14px 16px; text-align: center;
            box-shadow: 0 2px 8px rgba({RGB_AZUL_CSS}, 0.06); transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        .vel-kpi:hover {{ transform: translateY(-4px); box-shadow: 0 10px 20px -5px rgba({RGB_AZUL_CSS}, 0.15); }}
        .vel-kpi .lbl {{ font-size: 0.72rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: {COR_AZUL_ESC}; opacity: 0.85; }}
        .vel-kpi .val {{ font-family: 'Montserrat', sans-serif; font-size: 1.35rem; font-weight: 800; color: {COR_AZUL_ESC}; margin-top: 6px; }}
        .vel-kpi .val--red {{ color: {COR_VERMELHO} !important; }}
        div[data-baseweb="input"] {{ border-radius: 10px !important; border: 1px solid #e2e8f0 !important; background-color: {COR_INPUT_BG} !important; }}
        div[data-baseweb="input"]:focus-within {{ border-color: rgba({RGB_AZUL_CSS}, 0.35) !important; box-shadow: 0 0 0 3px rgba({RGB_AZUL_CSS}, 0.08) !important; }}
        </style>
        """, unsafe_allow_html=True)

def _exibir_logo_topo() -> None:
    path = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
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
    except Exception: pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Inteligência Competitiva: Direcional X Mercado</p>'
        f'<p class="ficha-sub">Estudo de performance e viabilidade por cluster regional.</p>'
        f"</div>"
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true"><div class="ficha-hero-bar"></div></div>'
        f"</div>",
        unsafe_allow_html=True,
    )

def _logo_url_secrets() -> str | None:
    try:
        if hasattr(st, "secrets") and isinstance(st.secrets.get("branding"), dict):
            return (st.secrets["branding"].get("LOGO_URL") or "").strip() or None
    except Exception: pass
    return None

def _logo_url_drive_por_id_arquivo() -> str | None:
    fid = (os.environ.get("DIRECIONAL_LOGO_FILE_ID") or "").strip()
    if len(fid) < 10: return None
    return f"https://drive.google.com/uc?export=view&id={fid}"

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

def process_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Pipeline focado em unificar dados via CHAVE e filtrar por Região."""
    df_det = ler_aba(SPREADSHEET_ID_CONC, "BD DETALHADA")
    df_ger = ler_aba(SPREADSHEET_ID_CONC, "BD GERAL")
    df_men = ler_aba(SPREADSHEET_ID_CONC, "Abr/2026")
    df_dados_dir = ler_aba(SPREADSHEET_ID_CONC, "DADOS DIRECIONAL")
    
    # 1. Limpeza Concorrentes
    df_det["Preço_Float"] = df_det["PREÇO"].apply(parse_val)
    df_det["Metragem_Float"] = df_det["METRAGEM"].apply(parse_val)
    df_det["Preço_m2"] = df_det["PREÇO_M2"].apply(parse_val)
    
    # Recalcular Preço_m2
    df_det["Preço_m2"] = np.where(
        (df_det["Preço_m2"] == 0) & (df_det["Metragem_Float"] > 0),
        df_det["Preço_Float"] / df_det["Metragem_Float"],
        df_det["Preço_m2"]
    )
    
    # FILTRO: Desconsiderar Preço_m2 > 10.000 (Conforme pedido)
    df_det = df_det[df_det["Preço_m2"] <= 10000]
    
    df_ger["Preço_Min"] = df_ger["Venda a Partir"].apply(parse_val)
    
    # 2. Merge Master
    # Cruzamento de Mensal com Detalhada via CHAVE
    df_master = df_det.merge(df_men[["CHAVE", "Vendas (Qnt.)", "Estoque (Qnt.)", "VGV (R$)", "PREÇO MÉDIO"]], on="CHAVE", how="left")
    
    # Cruzamento com Geral via EMPREENDIMENTO (BD Geral não possui coluna CHAVE)
    df_master = df_master.merge(df_ger[["Empreendimento", "Preço_Min", "Previsão"]], left_on="EMPREENDIMENTO", right_on="Empreendimento", how="left")
    
    # 3. Engenharia de Features
    vendas = pd.to_numeric(df_master["Vendas (Qnt.)"], errors='coerce').fillna(0)
    estoque = pd.to_numeric(df_master["Estoque (Qnt.)"], errors='coerce').fillna(0)
    df_master["Absorcao"] = vendas / (vendas + estoque)
    df_master["Absorcao"] = df_master["Absorcao"].replace([np.inf, -np.inf], 0).fillna(0)
    
    # 4. Identificar Direcional (Usando a lista fornecida de DADOS DIRECIONAL)
    direcional_keys = [str(x).strip().upper() for x in df_dados_dir["Nome do Empreendimento (Chave)"].unique() if x]
    
    df_master["Is_Direcional"] = df_master["CHAVE"].str.strip().str.upper().isin(direcional_keys)
    
    # Se algum item da aba "DADOS DIRECIONAL" não foi encontrado na base detalhada, marcamos erro silencioso
    # mas garantimos que as regiões sejam mapeadas corretamente.
    
    return df_master, df_dados_dir

# -----------------------------------------------------------------------------
# Interface Principal
# -----------------------------------------------------------------------------
def main():
    aplicar_estilo()
    _cabecalho_pagina()
    
    try:
        df_master, df_dados_dir = process_data()
    except Exception as e:
        st.error(f"Erro no Pipeline de Dados: {e}")
        return

    # Filtros Centralizados
    st.markdown("<div style='margin-bottom:1.5rem; text-align: center;'><strong>Filtros do Estudo Comparativo</strong></div>", unsafe_allow_html=True)
    
    f_col1, f_col2 = st.columns(2)
    with f_col1:
        # Seleção do empreendimento Direcional Alvo
        direcional_ativos = sorted(df_master[df_master["Is_Direcional"]]["CHAVE"].unique())
        if not direcional_ativos:
            st.warning("Nenhum empreendimento Direcional encontrado na base detalhada com preço/m² <= 10k.")
            # Fallback para permitir análise de mercado geral se o filtro de 10k removeu o alvo
            direcional_ativos = sorted(df_dados_dir["Nome do Empreendimento (Chave)"].unique())
            
        alvo_direcional = st.selectbox("Escolha o Produto Direcional para análise", direcional_ativos)
    
    with f_col2:
        # Filtro de Concorrente (Opcional para refinar o cluster)
        f_concorrente = st.multiselect("Filtrar Concorrente específico (opcional)", sorted(df_master["CONCORRENTE"].dropna().unique()))

    # Lógica de Raio Competitivo (Baseada na Região do Alvo)
    # 1. Descobrir a região do alvo direcional selecionado
    regiao_alvo = df_master[df_master["CHAVE"] == alvo_direcional]["REGIÃO"].iloc[0] if not df_master[df_master["CHAVE"] == alvo_direcional].empty else "Desconhecida"
    
    if regiao_alvo == "Desconhecida":
        # Tenta buscar na aba DADOS DIRECIONAL se não estiver no cruzamento
        match_dir = df_dados_dir[df_dados_dir["Nome do Empreendimento (Chave)"] == alvo_direcional]
        if not match_dir.empty:
            regioes_possiveis = match_dir["Região"].values
            # Filtra o mercado por essas regiões
            df_f = df_master[df_master["REGIÃO"].isin(regioes_possiveis)].copy()
            regiao_display = regioes_possiveis[0]
        else:
            df_f = df_master.copy()
            regiao_display = "Geral"
    else:
        # Filtra o mercado apenas pela mesma REGIÃO do produto Direcional (Proxy do raio de 15km)
        df_f = df_master[df_master["REGIÃO"] == regiao_alvo].copy()
        regiao_display = regiao_alvo

    if f_concorrente:
        df_f = df_f[df_f["CONCORRENTE"].isin(f_concorrente) | (df_f["CHAVE"] == alvo_direcional)]

    st.markdown(f"<div style='text-align:center; color:{COR_AZUL_ESC};'>Analisando Cluster: <strong>{regiao_display}</strong> (Produtos competindo com {alvo_direcional})</div>", unsafe_allow_html=True)
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # KPIs Comparativos
    # -------------------------------------------------------------------------
    df_conc = df_f[df_f["CHAVE"] != alvo_direcional]
    df_alvo = df_f[df_f["CHAVE"] == alvo_direcional]
    
    avg_m2_conc = df_conc["Preço_m2"].mean() if not df_conc.empty else 0
    avg_m2_alvo = df_alvo["Preço_m2"].mean() if not df_alvo.empty else 0
    
    avg_abs_conc = df_conc["Absorcao"].mean() * 100 if not df_conc.empty else 0
    avg_abs_alvo = df_alvo["Absorcao"].mean() * 100 if not df_alvo.empty else 0

    st.markdown(f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Preço/m² Médio Vizinhos</div><div class="val">R$ {avg_m2_conc:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Preço/m² {alvo_direcional}</div><div class="val val--red">R$ {avg_m2_alvo:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Velocidade Média Vizinhos</div><div class="val">{avg_abs_conc:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">Velocidade {alvo_direcional}</div><div class="val val--red">{avg_abs_alvo:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">Total Concorrentes Regionais</div><div class="val">{df_conc["CHAVE"].nunique()}</div></div>
        </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # Gráficos Estratégicos
    # -------------------------------------------------------------------------
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("Curva de Demanda Regional")
        # Scatter destacado
        fig = px.scatter(df_f, x="Preço_m2", y="Absorcao", color="Is_Direcional", trendline="ols",
                         color_discrete_map={True: COR_VERMELHO, False: COR_AZUL_ESC},
                         hover_name="EMPREENDIMENTO", text="CONCORRENTE",
                         labels={"Preço_m2": "R$ / m²", "Absorcao": "Taxa de Absorção"})
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Nota: A linha de tendência indica a sensibilidade ao preço na região selecionada.")

    with col_g2:
        st.subheader("Benchmark de Metragem vs Velocidade")
        fig = px.scatter(df_f, x="Metragem_Float", y="Absorcao", color="Is_Direcional", 
                         color_discrete_map={True: COR_VERMELHO, False: COR_AZUL_ESC},
                         size="Preço_m2", hover_name="EMPREENDIMENTO", opacity=0.7)
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    # -------------------------------------------------------------------------
    # Tabelas e Insights
    # -------------------------------------------------------------------------
    st.markdown("<br>### 🎯 Matriz de Performance Regional", unsafe_allow_html=True)
    
    # Tabela detalhada comparativa
    df_tab = df_f.groupby("CONCORRENTE").agg(
        Projetos=("CHAVE", "nunique"),
        Preco_m2=("Preço_m2", "mean"),
        Abs_Media=("Absorcao", "mean")
    ).reset_index().sort_values("Abs_Media", ascending=False)
    
    st.dataframe(df_tab.style.format({
        "Preco_m2": "R$ {:.2f}", "Abs_Media": "{:.1%}"
    }), use_container_width=True, hide_index=True)

    # Diagnóstico Inteligente
    st.markdown("<br>### 💡 Diagnóstico de Viabilidade", unsafe_allow_html=True)
    
    if avg_abs_alvo > avg_abs_conc:
        st.success(f"✅ **Liderança Regional:** {alvo_direcional} está com absorção acima da média dos vizinhos. O produto está bem encaixado.")
    else:
        st.error(f"⚠️ **Alerta de Performance:** {alvo_direcional} está abaixo da velocidade média da região. Avaliar elasticidade de preço ou tipologia.")

    if avg_m2_alvo > avg_m2_conc * 1.1:
        st.warning(f"💎 **Posicionamento Premium:** Seu m² está +10% acima da região. A diferenciação de produto precisa justificar este ticket.")
    elif avg_m2_alvo < avg_m2_conc * 0.9:
        st.info(f"💰 **Vantagem de Preço:** Seu m² está competitivo. Oportunidade para aumentar margem se a velocidade permitir.")

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.82rem;">Direcional Engenharia · Inteligência de Mercado · Estudo de Cluster Regional</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
