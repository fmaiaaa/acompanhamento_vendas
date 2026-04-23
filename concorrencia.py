# -*- coding: utf-8 -*-
"""
Análise de Concorrência — Inteligência Competitiva e Performance.
Foco: Estudo Geográfico de Precisão (15km) e Performance Real.
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
        f'<p class="ficha-sub">Análise Realizado vs Projetado em Cluster Regional.</p>'
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
    s = str(v).replace("R$", "").replace(".", "").replace(",", ".").replace(" ", "").strip()
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

@st.cache_data(ttl=86400, show_spinner=False)
def geocode_address_precision(address: str) -> tuple[Optional[float], Optional[float]]:
    """Geocodificação com parâmetros de precisão para Rio de Janeiro."""
    if not address or str(address).lower() == "nan": return None, None
    try:
        geolocator = Nominatim(user_agent="direcional_precision_v2")
        clean_addr = str(address).split("(")[0].strip() # Remove "Aprox." ou termos entre parênteses
        location = geolocator.geocode(clean_addr + ", Rio de Janeiro, RJ, Brasil", timeout=15)
        if location:
            return location.latitude, location.longitude
    except:
        pass
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

def process_pipeline(mes_selecionado: str):
    sheets = load_raw_sheets()
    
    df_det = sheets.get("BD DETALHADA", pd.DataFrame())
    df_men = sheets.get(mes_selecionado, pd.DataFrame())
    df_procv = sheets.get("PROCV", pd.DataFrame())
    df_ref_dir = sheets.get("DADOS DIRECIONAL", pd.DataFrame())
    df_est_dir = sheets.get("Automação Estudos Concorrentes - Lucas", pd.DataFrame())
    df_ven_dir = sheets.get("Controle de Vendas RJ - Periodo Integral", pd.DataFrame())
    
    # 1. Cruzamento Mercado (Preço/m2 <= 10k)
    df_det["Preço_Float"] = df_det["PREÇO"].apply(parse_val)
    df_det["Metragem_Float"] = df_det["METRAGEM"].apply(parse_val)
    df_det["Preço_m2"] = np.where(df_det["Metragem_Float"] > 0, df_det["Preço_Float"] / df_det["Metragem_Float"], 0)
    df_det = df_det[df_det["Preço_m2"] <= 10000].copy()
    
    # Merge com performance e endereços específicos da PROCV
    df_master = df_det.merge(df_men[["CHAVE", "Vendas (Qnt.)", "Estoque (Qnt.)"]], on="CHAVE", how="left")
    df_procv_sub = df_procv[["Chave", "Endereço"]].drop_duplicates().rename(columns={"Chave": "CHAVE", "Endereço": "Endereço_Conc"})
    df_master = df_master.merge(df_procv_sub, on="CHAVE", how="left")
    
    # 2. Processamento Real Direcional
    # Unidades de Estoque: Disponível, Fora de venda, Fora de venda - Comercial, Mirror
    status_venda = ["Disponível", "Fora de venda", "Fora de venda - Comercial", "Mirror"]
    df_est_dir = df_est_dir[df_est_dir["Status da unidade"].isin(status_venda)].copy()
    
    # Valor Líquido = Valor Final Com Kit - Folga de Tabela - Bônus Adimplência
    df_est_dir["Valor_Liquido"] = df_est_dir.apply(
        lambda r: parse_val(r.get("Valor Final Com Kit", 0)) - 
                  parse_val(r.get("Folga de Tabela", 0)) - 
                  parse_val(r.get("Bônus Adimplência", 0)), axis=1
    )
    df_est_dir["Area_Float"] = df_est_dir["Área privativa total"].apply(parse_val)
    df_est_dir["Preço_m2_Real"] = np.where(df_est_dir["Area_Float"] > 0, df_est_dir["Valor_Liquido"] / df_est_dir["Area_Float"], 0)
    
    # Vendas: Filtrar Venda Comercial? == 1
    df_ven_dir = df_ven_dir[pd.to_numeric(df_ven_dir["Venda Comercial?"], errors='coerce') == 1].copy()
    df_ven_dir["Data_Venda_DT"] = pd.to_datetime(df_ven_dir["Data da venda"], dayfirst=True, errors='coerce')
    
    return df_master, df_est_dir, df_ven_dir, df_ref_dir

# -----------------------------------------------------------------------------
# Main Application Logic
# -----------------------------------------------------------------------------
def main():
    aplicar_estilo()
    _cabecalho_pagina()
    
    # Coletar abas de performance
    import gspread
    from google.oauth2.service_account import Credentials
    sec = st.secrets["connections"]["gsheets"]
    info = {k: v for k, v in sec.items() if v}
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID_CONC)
    abas_mensais = [ws.title for ws in sh.worksheets() if "/" in ws.title]

    # Filtros de Topo
    st.markdown("<div style='text-align: center; font-weight: bold; margin-bottom: 1rem;'>Configuração do Cluster de Estudo</div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([2, 1, 1])
    
    with c1:
        df_ref_temp = pd.DataFrame(sh.worksheet("DADOS DIRECIONAL").get_all_records())
        df_ref_temp.columns = [str(c).strip() for c in df_ref_temp.columns]
        lista_alvos = sorted(df_ref_temp["Nome do Empreendimento (Chave)"].unique())
        alvo_selecionado = st.selectbox("Escolha o Produto Direcional para análise", lista_alvos)
        
    with c2:
        mes_estudo = st.selectbox("Mês Referência Performance", abas_mensais, index=len(abas_mensais)-1 if abas_mensais else 0)
        
    with c3:
        ano_estudo = st.selectbox("Ano Referência Vendas (Real)", [2024, 2025, 2026], index=2)

    try:
        df_mercado, df_est_dir, df_ven_dir, df_ref_dir = process_pipeline(mes_estudo)
    except Exception as e:
        st.error(f"Erro no processamento das bases: {e}")
        return

    # -------------------------------------------------------------------------
    # Geocodificação e Cálculo de Raio (15km)
    # -------------------------------------------------------------------------
    info_alvo = df_ref_dir[df_ref_dir["Nome do Empreendimento (Chave)"] == alvo_selecionado].iloc[0]
    endereco_alvo = info_alvo["Endereço"]
    regiao_alvo = info_alvo["Região"]
    
    with st.spinner("Calculando cluster georreferenciado..."):
        lat_alvo, lon_alvo = geocode_address_precision(endereco_alvo)
        
        if not lat_alvo:
            st.error("Falha ao localizar centro do cluster. Verifique o endereço na aba DADOS DIRECIONAL.")
            return

        # Filtrar mercado regional primeiro (performance)
        df_cluster = df_mercado[df_mercado["REGIÃO"] == regiao_alvo].copy()
        
        # Geocodificar vizinhos únicos
        addrs_conc = df_cluster["Endereço_Conc"].dropna().unique()
        coords_cache = {addr: geocode_address_precision(addr) for addr in addrs_conc}
        
        df_cluster["lat"] = df_cluster["Endereço_Conc"].map(lambda x: coords_cache.get(x, (None, None))[0])
        df_cluster["lon"] = df_cluster["Endereço_Conc"].map(lambda x: coords_cache.get(x, (None, None))[1])
        
        # Calcular Distância
        df_cluster["Distancia_km"] = df_cluster.apply(
            lambda r: geodesic((lat_alvo, lon_alvo), (r["lat"], r["lon"])).km if pd.notna(r["lat"]) else 999, axis=1
        )
        
        # Filtrar Cluster de 15km
        df_final = df_cluster[df_cluster["Distancia_km"] <= 15].copy()

    # -------------------------------------------------------------------------
    # Consolidação Direcional (Alvo) para o Mapa e KPIs
    # -------------------------------------------------------------------------
    v_alvo_qtd = len(df_ven_dir[(df_ven_dir["Empreendimento"] == alvo_selecionado) & (df_ven_dir["Data_Venda_DT"].dt.year == ano_estudo)])
    e_alvo_data = df_est_dir[df_est_dir["Nome do Empreendimento"] == alvo_selecionado]
    e_alvo_qtd = len(e_alvo_data)
    m2_alvo_real = e_alvo_data["Preço_m2_Real"].mean() if not e_alvo_data.empty else 0
    
    # Preparar DataFrame Unificado para o Mapa
    # Precisamos que o Alvo apareça como uma linha no mercado para o plot
    linha_alvo = pd.DataFrame([{
        "EMPREENDIMENTO": alvo_selecionado,
        "CONCORRENTE": "DIRECIONAL",
        "Preço_m2": m2_alvo_real,
        "Estoque (Qnt.)": e_alvo_qtd,
        "Vendas (Qnt.)": v_alvo_qtd,
        "lat": lat_alvo,
        "lon": lon_alvo,
        "Distancia_km": 0.0,
        "TIPOLOGIA": "Direcional Alvo"
    }])
    
    df_plot = pd.concat([df_final, linha_alvo], ignore_index=True)
    
    # -------------------------------------------------------------------------
    # KPIs Comparativos
    # -------------------------------------------------------------------------
    avg_m2_cluster = df_final["Preço_m2"].mean()
    abs_real_alvo = (v_alvo_qtd / (v_alvo_qtd + e_alvo_qtd)) * 100 if (v_alvo_qtd + e_alvo_qtd) > 0 else 0

    st.markdown(f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Preço m² Médio Cluster</div><div class="val">R$ {avg_m2_cluster:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Preço m² Real (Direcional)</div><div class="val val--red">R$ {m2_alvo_real:,.2f}</div></div>
            <div class="vel-kpi"><div class="lbl">Vendas Direcional ({ano_estudo})</div><div class="val val--red">{v_alvo_qtd} unid.</div></div>
            <div class="vel-kpi"><div class="lbl">Estoque Direcional Atual</div><div class="val">{e_alvo_qtd} unid.</div></div>
            <div class="vel-kpi"><div class="lbl">Absorção Real Direcional</div><div class="val val--red">{abs_real_alvo:.1f}%</div></div>
        </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # MAPA GEOGRÁFICO (REQUISITOS ESPECÍFICOS)
    # -------------------------------------------------------------------------
    st.subheader(f"Mapa do Cluster Competitivo (Raio 15km)")
    
    # Agrupar para o mapa por empreendimento
    df_map_points = df_plot.groupby("EMPREENDIMENTO").agg({
        "CONCORRENTE": "first", "lat": "first", "lon": "first", "Preço_m2": "mean", "Estoque (Qnt.)": "first", "Vendas (Qnt.)": "first"
    }).reset_index()
    
    # Ajuste de tamanho da bolha: proporcional ao estoque
    df_map_points["Marker_Size"] = pd.to_numeric(df_map_points["Estoque (Qnt.)"], errors='coerce').fillna(0) + 10
    
    # Cor: Tão mais vermelha quanto maior o Preço/m² (Escala Sequencial Vermelha)
    fig_map = px.scatter_mapbox(df_map_points, lat="lat", lon="lon", 
                                color="Preço_m2", 
                                size="Marker_Size",
                                hover_name="EMPREENDIMENTO", 
                                hover_data={"lat": False, "lon": False, "Preço_m2": ":.2f", "CONCORRENTE": True, "Estoque (Qnt.)": True},
                                color_continuous_scale="Reds", # Escala Vermelha (mais caro = mais escuro)
                                zoom=12, height=600, center={"lat": lat_alvo, "lon": lon_alvo})
    
    fig_map.update_layout(mapbox_style="carto-positron", margin={"r":0,"t":0,"l":0,"b":0}, paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_map, use_container_width=True)
    st.caption(f"Mapa centralizado em: {endereco_alvo}")

    # -------------------------------------------------------------------------
    # TABELA DE BENCHMARKING
    # -------------------------------------------------------------------------
    st.markdown("### Benchmarking de Performance (Raio 15km)")
    
    df_tab = df_plot.groupby("EMPREENDIMENTO").agg({
        "CONCORRENTE": "first",
        "TIPOLOGIA": "first",
        "Metragem_Float": "mean",
        "Preço_m2": "mean",
        "Vendas (Qnt.)": "first",
        "Distancia_km": "first"
    }).reset_index().sort_values("Distancia_km", ascending=True)
    
    st.dataframe(df_tab.rename(columns={
        "EMPREENDIMENTO": "Produto",
        "Metragem_Float": "m² Médio",
        "Preço_m2": "R$/m²",
        "Vendas (Qnt.)": f"Vendas ({mes_estudo})",
        "Distancia_km": "Distância (km)"
    }).style.format({
        "R$/m²": "R$ {:.2f}", 
        "m² Médio": "{:.1f}",
        "Distância (km)": "{:.1f} km"
    }), use_container_width=True, hide_index=True)

    # Diagnóstico Técnico
    st.markdown("<br>### Diagnóstico de Cluster", unsafe_allow_html=True)
    c_d1, c_d2 = st.columns(2)
    with c_d1:
        if m2_alvo_real < avg_m2_cluster:
            st.info(f"O preço médio real do produto está abaixo da média do cluster regional. Existe margem técnica para reposicionamento.")
        else:
            st.warning(f"O ticket médio real do produto está acima da média do cluster regional. Focar no valor agregado da marca e diferenciais do projeto.")
    
    with c_d2:
        if abs_real_alvo > 10:
            st.info(f"Performance de vendas acima de 10% no ano de referência. Velocidade de escoamento saudável.")
        else:
            st.error(f"Velocidade de vendas abaixo do esperado. Reavaliar competitividade frente aos novos lançamentos do raio de 15km.")

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.82rem;">Direcional Engenharia · Inteligência de Mercado · Estudo Georreferenciado sem Sidebar</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
