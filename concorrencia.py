# -*- coding: utf-8 -*-
"""
Análise de Concorrência — Inteligência Competitiva e Performance.
Foco: Estudo Geográfico de Precisão com Raio Dinâmico, Performance Temporal e Monetização.
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
from geopy.geocoders import Nominatim
from geopy.distance import geodesic

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
            max-width: 1700px !important; padding: 1.45rem 2.25rem 1.55rem 2.25rem !important; background: rgba(255, 255, 255, 0.78) !important;
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
        f'<p class="ficha-title">Estudo de Inteligência Competitiva</p>'
        f'<p class="ficha-sub">Análise Realizado vs Projetado — Cluster Temporal e Monetização.</p>'
        f"</div>"
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true"><div class="ficha-hero-bar"></div></div>'
        f"</div>",
        unsafe_allow_html=True,
    )

# -----------------------------------------------------------------------------
# Pipeline de Dados e Geocodificação de Precisão
# -----------------------------------------------------------------------------
def parse_val(v):
    if not v: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace("R$", "").replace(".", "").replace(",", ".").replace(" ", "").strip()
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

@st.cache_data(ttl=86400, show_spinner=False)
def geocode_address_precision(address: str) -> tuple[Optional[float], Optional[float]]:
    if not address or str(address).lower() == "nan": return None, None
    try:
        geolocator = Nominatim(user_agent="direcional_precision_v5")
        clean_addr = str(address).split("(")[0].strip()
        location = geolocator.geocode(clean_addr + ", Rio de Janeiro, RJ, Brasil", timeout=15)
        if location: return location.latitude, location.longitude
    except: pass
    return None, None

@st.cache_data(ttl=300, show_spinner=False)
def load_raw_sheets() -> dict[str, pd.DataFrame]:
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
        data = ws.get_all_values()
        if data:
            df = pd.DataFrame(data[1:], columns=data[0])
            df.columns = [str(c).strip() for c in df.columns]
            dfs[ws.title] = df
    return dfs

def process_pipeline():
    sheets = load_raw_sheets()
    
    df_temp = sheets.get("BD TEMPORAL", pd.DataFrame())
    df_det = sheets.get("BD DETALHADA", pd.DataFrame())
    df_ref_dir = sheets.get("DADOS DIRECIONAL", pd.DataFrame())
    df_procv = sheets.get("PROCV", pd.DataFrame())
    
    # Normalizar chaves para joins seguros
    for df in [df_temp, df_det, df_ref_dir, df_procv]:
        if "CHAVE" in df.columns: df["CHAVE"] = df["CHAVE"].astype(str).str.strip().str.upper()
        elif "Chave" in df.columns: df["CHAVE"] = df["Chave"].astype(str).str.strip().str.upper()
    
    # 1. Processamento BD DETALHADA (Contagem de Estoque Disponível)
    df_det["Preço_Unidade"] = df_det["PREÇO"].apply(parse_val)
    
    # Count de estoque disponível por Chave
    df_estoque_atual = df_det[df_det["DISPONIBILIDADE"].astype(str).str.strip().str.upper() == "DISPONÍVEL"].copy()
    df_estoque_count = df_estoque_atual.groupby("CHAVE").agg(
        Estoque_Count=("DISPONIBILIDADE", "count"),
        VGV_Estoque=("Preço_Unidade", "sum")
    ).reset_index()
    
    # 2. Processamento BD TEMPORAL
    cols_num = ["VENDAS", "ESTOQUE", "ESTOQUE INICIAL"]
    for col in cols_num:
        df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0)
    
    df_temp["VGV_Vendido"] = df_temp["VGV"].apply(parse_val)
    df_temp["PREÇO MÉDIO"] = df_temp["PREÇO MÉDIO"].apply(parse_val)
    df_temp["DATA_DT"] = pd.to_datetime(df_temp["DATA"], dayfirst=True, errors='coerce')
    
    # Limpeza e Ordenação
    df_temp = df_temp[df_temp["DATA_DT"].notna()].sort_values(["CHAVE", "DATA_DT"])
    
    # Cruzamento temporal com Tipologia da Detalhada (pegando a primeira ocorrência por chave)
    df_tipologias = df_det[["CHAVE", "TIPOLOGIA"]].drop_duplicates(subset=["CHAVE"])
    df_temp = df_temp.merge(df_tipologias, on="CHAVE", how="left")
    
    # 3. Engenharia de Features (Cálculos de Indicadores)
    # Taxa de Absorção: vendas_t / (estoque_t-1 + vendas_t) 
    # Usamos ESTOQUE INICIAL como o estoque do início do mês (estoque_t-1 + vendas_t)
    df_temp["Absorcao"] = df_temp["VENDAS"] / df_temp["ESTOQUE INICIAL"].replace(0, np.nan)
    
    # Velocidade de Vendas
    df_temp["Velocidade"] = df_temp["VENDAS"] / (df_temp["ESTOQUE INICIAL"].replace(0, np.nan))
    
    # Taxa de Escoamento
    df_temp["Escoamento"] = df_temp["ESTOQUE"] / df_temp["ESTOQUE INICIAL"].replace(0, np.nan)
    
    # VGV Rate: VGV Vendido / VGV Total do Projeto
    df_temp["VGV_Total_Projeto"] = df_temp["VGV_Vendido"] + (df_temp["ESTOQUE"] * df_temp["PREÇO MÉDIO"])
    df_temp["VGV_Rate"] = df_temp["VGV_Vendido"] / df_temp["VGV_Total_Projeto"].replace(0, np.nan)
    
    # Share de Vendas Mensal
    total_vendas_mes = df_temp.groupby("DATA_DT")["VENDAS"].transform("sum")
    df_temp["Share_Vendas"] = df_temp["VENDAS"] / total_vendas_mes.replace(0, np.nan)
    
    # Identificar Direcional
    direcional_keys = [str(x).strip().upper() for x in df_ref_dir["Nome do Empreendimento (Chave)"].unique() if x]
    df_temp["Is_Direcional"] = df_temp["CHAVE"].isin(direcional_keys)
    
    # Merge Final com estoque actual da Detalhada
    df_master = df_temp.merge(df_estoque_count, on="CHAVE", how="left")
    df_master["Estoque_Count"] = df_master["Estoque_Count"].fillna(0)
    
    return df_master, df_ref_dir

# -----------------------------------------------------------------------------
# Main Application
# -----------------------------------------------------------------------------
def main():
    aplicar_estilo()
    _cabecalho_pagina()
    
    try:
        df_master, df_ref_dir = process_pipeline()
    except Exception as e:
        st.error(f"Erro no processamento das bases: {e}")
        return

    # Filtros de Topo
    st.markdown("<div style='text-align: center; font-weight: bold; margin-bottom: 1rem;'>Configuração do Estudo de Cluster</div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.8, 1, 1.2])
    
    with c1:
        lista_alvos = sorted(df_ref_dir["Nome do Empreendimento (Chave)"].unique())
        alvo_selecionado = st.selectbox("Selecione o Empreendimento Direcional", lista_alvos)
        
    with c2:
        df_sorted_dates = df_master[["DATA", "DATA_DT"]].drop_duplicates().sort_values("DATA_DT")
        meses_disp = df_sorted_dates["DATA"].tolist()
        mes_estudo = st.selectbox("Mês para Visão de Mapa", meses_disp, index=len(meses_disp)-1 if meses_disp else 0)
        
    with c3:
        raio_estudo = st.slider("Raio de Atuação Dinâmico (km)", 1, 50, 15)

    # -------------------------------------------------------------------------
    # Geocodificação e Cluster
    # -------------------------------------------------------------------------
    info_alvo_rows = df_ref_dir[df_ref_dir["Nome do Empreendimento (Chave)"].astype(str).str.strip().str.upper() == str(alvo_selecionado).strip().upper()]
    if info_alvo_rows.empty:
        st.error("Dados de referência do empreendimento não encontrados.")
        return
        
    info_alvo = info_alvo_rows.iloc[0]
    endereco_alvo = info_alvo["Endereço"]
    regiao_alvo = info_alvo["Região"]
    
    with st.spinner("Processando geolocalização e histórico de cluster..."):
        lat_alvo, lon_alvo = geocode_address_precision(endereco_alvo)
        
        if not lat_alvo:
            st.error(f"Localização não encontrada para o endereço: {endereco_alvo}")
            return

        # Filtrar Cluster Regional
        df_regiao = df_master[df_master["REGIÃO"] == regiao_alvo].copy()
        
        # Geocodificar vizinhos únicos no cluster
        addrs_conc = df_regiao["ENDEREÇO"].dropna().unique()
        coords_cache = {addr: geocode_address_precision(addr) for addr in addrs_conc}
        
        df_regiao["lat"] = df_regiao["ENDEREÇO"].map(lambda x: coords_cache.get(x, (None, None))[0])
        df_regiao["lon"] = df_regiao["ENDEREÇO"].map(lambda x: coords_cache.get(x, (None, None))[1])
        
        # Calcular Distância Real
        df_regiao["Distancia_km"] = df_regiao.apply(
            lambda r: geodesic((lat_alvo, lon_alvo), (r["lat"], r["lon"])).km if pd.notna(r["lat"]) else 999, axis=1
        )
        
        df_cluster = df_regiao[df_regiao["Distancia_km"] <= raio_estudo].copy()

    # -------------------------------------------------------------------------
    # KPIs Comparativos
    # -------------------------------------------------------------------------
    df_mes = df_cluster[df_cluster["DATA"] == mes_estudo]
    df_alvo = df_mes[df_mes["CHAVE"] == str(alvo_selecionado).strip().upper()]
    df_vizinhos = df_mes[df_mes["CHAVE"] != str(alvo_selecionado).strip().upper()]
    
    abs_alvo = df_alvo["Absorcao"].iloc[0] * 100 if not df_alvo.empty else 0
    abs_vizinhos = df_vizinhos["Absorcao"].mean() * 100 if not df_vizinhos.empty else 0
    vgv_rate_alvo = df_alvo["VGV_Rate"].iloc[0] * 100 if not df_alvo.empty else 0
    m2_alvo = df_alvo["PREÇO MÉDIO"].iloc[0] if not df_alvo.empty else 0
    m2_vizinhos = df_vizinhos["PREÇO MÉDIO"].mean() if not df_vizinhos.empty else 0

    st.markdown(f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Preço Médio Cluster</div><div class="val">R$ {m2_vizinhos:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Preço Médio Direcional</div><div class="val val--red">R$ {m2_alvo:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Absorção Direcional</div><div class="val val--red">{abs_alvo:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">Monetização (VGV Rate)</div><div class="val val--red">{vgv_rate_alvo:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">Estoque Real (Disponível)</div><div class="val">{int(df_alvo["Estoque_Count"].sum()) if not df_alvo.empty else 0} unid.</div></div>
        </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # MAPA GEOGRÁFICO
    # -------------------------------------------------------------------------
    st.subheader(f"Mapa do Cluster Competitivo (Raio {raio_estudo}km)")
    
    df_map_points = df_mes.groupby("EMPREENDIMENTO").agg({
        "CONSTRUTORA": "first", "lat": "first", "lon": "first", "PREÇO MÉDIO": "mean", "Estoque_Count": "sum", "Is_Direcional": "first"
    }).reset_index()
    
    # Tamanho da bolha pelo estoque, cor pelo preço
    df_map_points["Marker_Size"] = df_map_points["Estoque_Count"].fillna(0) + 15
    
    fig_map = px.scatter_mapbox(df_map_points, lat="lat", lon="lon", 
                                color="PREÇO MÉDIO", 
                                size="Marker_Size",
                                hover_name="EMPREENDIMENTO", 
                                hover_data={"PREÇO MÉDIO": ":.2f", "Estoque_Count": True},
                                color_continuous_scale="Reds", 
                                zoom=12, height=500, center={"lat": lat_alvo, "lon": lon_alvo})
    
    fig_map.update_layout(mapbox_style="carto-positron", margin={"r":0,"t":0,"l":0,"b":0}, paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_map, use_container_width=True)

    # -------------------------------------------------------------------------
    # GRÁFICOS DE INDICADORES TEMPORAIS
    # -------------------------------------------------------------------------
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:2rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Indicadores de Performance ao Longo do Tempo")
    
    df_trend = df_cluster.groupby(["DATA_DT", "Is_Direcional"]).agg({
        "Absorcao": "mean", "Velocidade": "mean", "Escoamento": "mean", "PREÇO MÉDIO": "mean", "VGV_Rate": "mean", "Share_Vendas": "sum"
    }).reset_index().sort_values("DATA_DT")
    df_trend["DATA_STR"] = df_trend["DATA_DT"].dt.strftime("%m/%Y")

    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Taxa de Absorção (Vendas / Estoque Inicial)**")
        fig_abs = px.line(df_trend, x="DATA_STR", y="Absorcao", color="Is_Direcional", color_discrete_map={True: COR_VERMELHO, False: COR_AZUL_ESC}, markers=True)
        fig_abs.update_layout(yaxis_tickformat=".1%", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
        st.plotly_chart(fig_abs, use_container_width=True)
        
        st.markdown("**Taxa de Monetização (VGV Realizado / VGV Total)**")
        fig_vgv_r = px.line(df_trend, x="DATA_STR", y="VGV_Rate", color="Is_Direcional", color_discrete_map={True: COR_VERMELHO, False: COR_AZUL_ESC}, markers=True)
        fig_vgv_r.update_layout(yaxis_tickformat=".1%", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
        st.plotly_chart(fig_vgv_r, use_container_width=True)

    with g2:
        st.markdown("**Absorção por Tipologia no Cluster**")
        df_tipo = df_cluster.groupby(["DATA_DT", "TIPOLOGIA"]).agg({"Absorcao": "mean"}).reset_index().sort_values("DATA_DT")
        df_tipo["DATA_STR"] = df_tipo["DATA_DT"].dt.strftime("%m/%Y")
        fig_tipo = px.line(df_tipo, x="DATA_STR", y="Absorcao", color="TIPOLOGIA", markers=True, color_discrete_sequence=px.colors.qualitative.Safe)
        fig_tipo.update_layout(yaxis_tickformat=".1%", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="", legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig_tipo, use_container_width=True)
        
        st.markdown("**Dinâmica de Preço Médio Praticado**")
        fig_prc = px.line(df_trend, x="DATA_STR", y="PREÇO MÉDIO", color="Is_Direcional", color_discrete_map={True: COR_VERMELHO, False: COR_AZUL_ESC}, markers=True)
        fig_prc.update_layout(yaxis_tickprefix="R$ ", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="")
        st.plotly_chart(fig_prc, use_container_width=True)

    # -------------------------------------------------------------------------
    # TABELA DE BENCHMARKING
    # -------------------------------------------------------------------------
    st.markdown("### Benchmarking de Performance no Cluster")
    
    df_tab = df_mes.groupby("EMPREENDIMENTO").agg({
        "CONSTRUTORA": "first", "TIPOLOGIA": "first", "PREÇO MÉDIO": "mean", "VENDAS": "sum", "Estoque_Count": "sum", "Absorcao": "mean", "VGV_Rate": "mean", "Distancia_km": "first"
    }).reset_index().sort_values("Distancia_km", ascending=True)
    
    st.dataframe(df_tab.rename(columns={
        "EMPREENDIMENTO": "Produto", "PREÇO MÉDIO": "R$/m²", "Estoque_Count": "Estoque Real", "Absorcao": "Absorção", "VGV_Rate": "VGV Rate", "Distancia_km": "Distância"
    }).style.format({
        "R$/m²": "R$ {:.2f}", "Absorção": "{:.1%}", "VGV Rate": "{:.1%}", "Distância": "{:.1f} km"
    }), use_container_width=True, hide_index=True)

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.82rem;">Direcional Engenharia · Inteligência de Mercado · Estudo de Cluster Georreferenciado</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
