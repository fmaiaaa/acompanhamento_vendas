# -*- coding: utf-8 -*-
"""
Acompanhamento de vendas — metas vs realizado (Direcional).
Focado em Premiações: Coordenadores IMOB, Comerciais e Grandes Contas.
Design atualizado baseado no app de Gaps.
"""
from __future__ import annotations

import base64
import html
import math
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# -----------------------------------------------------------------------------
# Identificação da planilha e Arquivos Visuais
# -----------------------------------------------------------------------------
SPREADSHEET_ID = "1wpuNQvksot9CLhGgQRe7JlyDeRISEh_sc3-6VRDyQYk"

WS_VENDAS = "BD Vendas Completa"
WS_METAS_IMOB = "Metas Coordenadores IMOB"
WS_METAS_COMERCIAIS = "Metas Coordenadores Comerciais"
WS_METAS_GC = "Metas Grandes Contas"
WS_DICIONARIO = "Dicionário Coordenadores"

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"

# Paleta alinhada à Ficha de Credenciamento / Vendas RJ
COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_BORDA = "#eef2f6"
COR_TEXTO_MUTED = "#64748b"
COR_TEXTO_LABEL = "#1e293b"
COR_INPUT_BG = "#f0f2f6"

MESES_TEXTO_MAP = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
}

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
        if p.is_file(): return p
    return None

def _css_url_fundo_cadastro() -> str:
    p = _resolver_png_raiz(FUNDO_CADASTRO_ARQUIVO)
    if p and p.is_file():
        try:
            b64 = base64.b64encode(p.read_bytes()).decode("ascii")
            return f"data:image/jpeg;base64,{b64}"
        except: pass
    return "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?auto=format&fit=crop&w=1920&q=80"

def _exibir_logo_topo() -> None:
    path = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
    if path:
        try:
            b64 = base64.b64encode(path.read_bytes()).decode("ascii")
            st.markdown(f'<div class="ficha-logo-wrap"><img src="data:image/png;base64,{b64}" alt="Direcional" /></div>', unsafe_allow_html=True)
        except: pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Acompanhamento de Metas e Premiações</p>'
        f'<p class="ficha-sub">Realizado X Projetado — <strong>Vendas RJ</strong>.</p>'
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
        
        [data-testid="stSidebar"] {{ display: block !important; }}
        
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

        /* Centralização das Abas */
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {{
            display: flex;
            justify-content: center;
            width: 100%;
            gap: 12px;
            border-bottom: none !important;
        }}
        div[data-testid="stTabs"] button[data-baseweb="tab"] {{
            background-color: rgba(255,255,255,0.4);
            border-radius: 12px 12px 0 0;
            padding: 12px 24px;
            font-family: 'Montserrat', sans-serif;
            font-weight: 700;
            color: {COR_AZUL_ESC};
            border: 1px solid rgba(255,255,255,0.3);
        }}
        div[data-testid="stTabs"] button[aria-selected="true"] {{
            background-color: {COR_AZUL_ESC} !important;
            color: white !important;
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
        
        .debug-box {{
            background: rgba(15, 23, 42, 0.9);
            color: #10b981;
            padding: 1rem;
            border-radius: 8px;
            font-family: 'Courier New', Courier, monospace;
            font-size: 0.85rem;
            margin: 10px 0;
            border-left: 5px solid #10b981;
            overflow-x: auto;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

# -----------------------------------------------------------------------------
# Lógicas de Tratamento de Dados
# -----------------------------------------------------------------------------
def _secrets_connections_gsheets() -> Dict[str, Any]:
    try: return dict(st.secrets["connections"]["gsheets"])
    except: return {}

def montar_service_account_info(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not raw: return None
    pk = str(raw.get("private_key", "")).replace("\\n", "\n")
    return {
        "type": "service_account",
        "project_id": raw.get("project_id"),
        "private_key": pk,
        "client_email": raw.get("client_email"),
        "token_uri": "https://oauth2.googleapis.com/token",
    }

@st.cache_data(ttl=300)
def ler_aba_df(sid: str, worksheet: str) -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials
    raw = _secrets_connections_gsheets()
    info = montar_service_account_info(raw)
    creds = Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sid)
    ws = sh.worksheet(worksheet)
    data = ws.get_all_values()
    if not data: return pd.DataFrame()
    df = pd.DataFrame(data[1:], columns=[str(c).strip() for c in data[0]])
    return df

def parse_valor_br(val: Any) -> float:
    if pd.isna(val) or val == "" or str(val).lower() == "none": return 0.0
    s = str(val).replace("R$", "").replace(".", "").replace(",", ".").strip()
    try: return float(s)
    except: return 0.0

def extrair_lista_coords(val: str) -> List[str]:
    s = str(val).strip()
    if s.startswith("{") and s.endswith("}"): s = s[1:-1]
    return [c.strip() for c in s.split(",") if c.strip()]

def normalizar_mes_para_int(val: Any) -> int:
    v = str(val).strip().lower()
    if not v: return 0
    if v in MESES_TEXTO_MAP: return MESES_TEXTO_MAP[v]
    try:
        return int(float(v))
    except:
        return 0

def normalizar_ano_para_int(val: Any) -> int:
    s = str(val).strip()
    if not s: return 0
    s_limpo = s.replace(".", "").replace(",", "")
    try:
        return int(float(s_limpo))
    except:
        return 0

def get_vendedores_do_coordenador(df_dic: pd.DataFrame, nome_coord_tabela: str) -> List[str]:
    if df_dic.empty: return []
    col_coord = df_dic.columns[0]
    col_prop = df_dic.columns[1]
    mask = df_dic[col_coord].astype(str).str.strip().str.lower() == nome_coord_tabela.strip().lower()
    return df_dic[mask][col_prop].astype(str).str.strip().unique().tolist()

def calcular_realizado(df_vendas: pd.DataFrame, vendedores: List[str] = None, emp: Optional[str] = None, ignora_vendedor: bool = False) -> int:
    if df_vendas.empty: return 0
    
    mask = pd.Series([True] * len(df_vendas), index=df_vendas.index)
    col_prop = "Proprietário da oportunidade"
    col_emp = "Empreendimento"
    
    if not ignora_vendedor:
        if vendedores:
            vendedores_reais = [v.strip().lower() for v in vendedores]
            mask &= df_vendas[col_prop].astype(str).str.strip().str.lower().isin(vendedores_reais)
        else:
            return 0
    
    if emp:
        mask &= (df_vendas[col_emp].astype(str).str.strip().str.lower() == str(emp).strip().lower())
    
    return int(mask.sum())

# -----------------------------------------------------------------------------
# Main App
# -----------------------------------------------------------------------------
def main():
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(page_title="Premiações Direcional", page_icon=str(fav) if fav else None, layout="wide")
    aplicar_estilo()
    _cabecalho_pagina()
    
    sid = _secrets_connections_gsheets().get("spreadsheet_id", SPREADSHEET_ID)

    st.sidebar.title("Opções de Visibilidade")
    modo_debug = st.sidebar.toggle("Ativar Modo Debug", value=False)

    try:
        df_vendas_raw = ler_aba_df(sid, WS_VENDAS)
        df_dic = ler_aba_df(sid, WS_DICIONARIO)
        df_metas_imob = ler_aba_df(sid, WS_METAS_IMOB)
        df_metas_comerciais = ler_aba_df(sid, WS_METAS_COMERCIAIS)
        df_metas_gc = ler_aba_df(sid, WS_METAS_GC)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return

    # --- Pré-processamento Vendas ---
    df_vendas = df_vendas_raw.copy()
    
    # Filtro Comercial Global: Venda Comercial? == 1
    col_comercial = "Venda Comercial?"
    if col_comercial in df_vendas.columns:
        df_vendas["_V_COM_VAL"] = pd.to_numeric(df_vendas[col_comercial], errors='coerce').fillna(0)
        df_vendas = df_vendas[df_vendas["_V_COM_VAL"] == 1]
    
    # Normalização de datas no BD
    df_vendas["_MES_NUM"] = df_vendas["Mês Venda"].apply(normalizar_mes_para_int)
    df_vendas["_ANO_NUM"] = df_vendas["Ano da Venda"].apply(normalizar_ano_para_int)

    # -------------------------------------------------------------------------
    # Filtros de Período e Filtro Extra de Vendas Facilitadas
    # -------------------------------------------------------------------------
    st.markdown("<div style='margin-bottom:1rem; text-align: center;'><strong>Filtros de Análise</strong></div>", unsafe_allow_html=True)
    c_f1, c_f2, c_f3 = st.columns(3)
    with c_f1:
        anos_disponiveis = sorted([str(int(x)) for x in df_vendas["_ANO_NUM"].unique() if x > 0], reverse=True)
        if not anos_disponiveis: anos_disponiveis = [str(datetime.now().year)]
        ano_sel = st.selectbox("Selecione o Ano", anos_disponiveis)
    with c_f2:
        meses_nomes = [
            ("Janeiro", 1), ("Fevereiro", 2), ("Março", 3), ("Abril", 4), 
            ("Maio", 5), ("Junho", 6), ("Julho", 7), ("Agosto", 8), 
            ("Setembro", 9), ("Outubro", 10), ("Novembro", 11), ("Dezembro", 12)
        ]
        mes_atual_idx = datetime.now().month - 1
        mes_sel_nome, mes_sel_val = st.selectbox("Selecione o Mês", meses_nomes, index=mes_atual_idx, format_func=lambda x: x[0])
    with c_f3:
        facilitada_opcoes = ["Todas", "Sim", "Não"]
        facilitada_sel = st.selectbox("Venda Facilitada", facilitada_opcoes)

    # Chave para buscar na aba Metas
    data_filtro_meta = f"{int(mes_sel_val):02d}/{ano_sel}"
    
    # Filtrar BD Vendas para o período selecionado
    vendas_periodo = df_vendas[
        (df_vendas["_ANO_NUM"] == int(ano_sel)) & 
        (df_vendas["_MES_NUM"] == int(mes_sel_val))
    ]

    # Aplicar Filtro de Venda Facilitada se selecionado
    col_facilitada = "Venda facilitada"
    if col_facilitada in vendas_periodo.columns and facilitada_sel != "Todas":
        # Consideramos 1 ou Sim como facilitada
        if facilitada_sel == "Sim":
            vendas_periodo = vendas_periodo[vendas_periodo[col_facilitada].astype(str).str.strip().isin(["1", "1.0", "Sim", "SIM", "True", "TRUE"])]
        else:
            vendas_periodo = vendas_periodo[~vendas_periodo[col_facilitada].astype(str).str.strip().isin(["1", "1.0", "Sim", "SIM", "True", "TRUE"])]

    # --- DEBUG GERAL ---
    if modo_debug:
        with st.expander("🛠️ DEBUG: Auditoria de Filtros", expanded=True):
            st.write(f"**Data Filtro Metas:** `{data_filtro_meta}`")
            st.write(f"**Total Vendas Comerciais no Mês:** `{len(vendas_periodo)}`")
            st.write(f"**Filtro Facilitada:** `{facilitada_sel}`")
            if not vendas_periodo.empty:
                st.write("**Amostra do período filtrado:**")
                st.dataframe(vendas_periodo[["Empreendimento", "Proprietário da oportunidade", "Mês Venda", "Ano da Venda", col_facilitada]].head(3))

    # -------------------------------------------------------------------------
    # Abas Centralizadas
    # -------------------------------------------------------------------------
    tabs = st.tabs(["Coordenadores IMOB", "Coordenadores Comerciais", "Grandes Contas"])

    # --- ABA 1: IMOB ---
    with tabs[0]:
        st.subheader(f"Premiação IMOB — {mes_sel_nome}/{ano_sel}")
        df_imob_metas = df_metas_imob[df_metas_imob["DATA"] == data_filtro_meta]
        
        if df_imob_metas.empty:
            st.info(f"Nenhuma meta encontrada para {data_filtro_meta} na aba IMOB.")
        else:
            regioes = sorted(df_imob_metas["REGIÃO"].unique())
            for reg in regioes:
                with st.expander(f"Região: {reg}", expanded=True):
                    df_reg = df_imob_metas[df_imob_metas["REGIÃO"] == reg]
                    rows_imob = []
                    
                    for _, r in df_reg.iterrows():
                        coords_meta = extrair_lista_coords(r["COORDENADORES"])
                        emp_nome = r["EMPREENDIMENTO"]
                        for c_meta in coords_meta:
                            vendedores_reais = get_vendedores_do_coordenador(df_dic, c_meta)
                            realizado = calcular_realizado(vendas_periodo, vendedores_reais, emp_nome)
                            
                            m_dir = int(parse_valor_br(r.get("META DIRECIONAL", 0)))
                            m_imob = int(parse_valor_br(r.get("META IMOB", 0)))
                            m_imob2 = int(parse_valor_br(r.get("META IMOB 2", 0)))
                            
                            status = "NÃO BATEU"
                            if realizado >= m_imob2 and m_imob2 > 0: status = "META IMOB 2 ✅"
                            elif realizado >= m_imob and m_imob > 0: status = "META IMOB ✅"
                            elif realizado >= m_dir and m_dir > 0: status = "META DIRECIONAL ✅"
                            
                            rows_imob.append({
                                "Coordenador": c_meta, "Empreendimento": emp_nome,
                                "Meta Dir.": m_dir, "Meta IMOB": m_imob, "Meta IMOB 2": m_imob2,
                                "Realizado": realizado, "Resultado": status
                            })
                    if rows_imob:
                        st.table(pd.DataFrame(rows_imob))

    # --- ABA 2: COMERCIAIS ---
    with tabs[1]:
        st.subheader(f"Premiação Comerciais — {mes_sel_nome}/{ano_sel}")
        df_com_metas = df_metas_comerciais[df_metas_comerciais["DATA"] == data_filtro_meta]
        
        if df_com_metas.empty:
            st.info(f"Nenhuma meta encontrada para {data_filtro_meta} na aba Comerciais.")
        else:
            all_coords = []
            for _, r in df_com_metas.iterrows():
                all_coords.extend(extrair_lista_coords(r["COORDENADORES"]))
            unique_coords = sorted(list(set(all_coords)))
            
            for c_name in unique_coords:
                st.markdown(f"#### Coordenador: {c_name}")
                rows_com = []
                mask_coord = df_com_metas["COORDENADORES"].str.contains(c_name, case=False, na=False)
                df_c_metas = df_com_metas[mask_coord]
                
                for _, r in df_c_metas.iterrows():
                    emp_nome = r["EMPREENDIMENTO"]
                    realizado = calcular_realizado(vendas_periodo, emp=emp_nome, ignora_vendedor=True)
                    
                    m_desafio = int(parse_valor_br(r.get("META DESAFIO VENDAS", 0)))
                    m_bp = int(parse_valor_br(r.get("META BP", 0)))
                    m_bp70 = int(parse_valor_br(r.get("META BP 70%", 0)))
                    
                    status = "NÃO BATEU"
                    if realizado >= m_desafio and m_desafio > 0: status = "DESAFIO ✅"
                    elif realizado >= m_bp and m_bp > 0: status = "BP ✅"
                    elif realizado >= m_bp70 and m_bp70 > 0: status = "BP 70% ✅"
                    
                    rows_com.append({
                        "Empreendimento": emp_nome, "Meta Desafio": m_desafio, "Meta BP": m_bp,
                        "Meta BP 70%": m_bp70, "Realizado": realizado, "Resultado": status
                    })
                if rows_com:
                    st.dataframe(pd.DataFrame(rows_com), use_container_width=True, hide_index=True)

    # --- ABA 3: GRANDES CONTAS ---
    with tabs[2]:
        st.subheader(f"Premiação Grandes Contas — {mes_sel_nome}/{ano_sel}")
        df_gc_metas = df_metas_gc[df_metas_gc["DATA"] == data_filtro_meta]
        
        if df_gc_metas.empty:
            st.info(f"Nenhuma meta encontrada para {data_filtro_meta} na aba Grandes Contas.")
        else:
            rows_gc = []
            for _, r in df_gc_metas.iterrows():
                coords_meta = extrair_lista_coords(r["COORDENADORES"])
                m1 = int(parse_valor_br(r.get("META 1", 0)))
                m2 = int(parse_valor_br(r.get("META 2", 0)))
                focos = extrair_lista_coords(r.get("PRODUTOS FOCO", ""))
                
                for c_meta in coords_meta:
                    vendedores_reais = get_vendedores_do_coordenador(df_dic, c_meta)
                    real_total = calcular_realizado(vendas_periodo, vendedores_reais)
                    
                    real_foco = 0
                    for p_foco in focos:
                        real_foco += calcular_realizado(vendas_periodo, vendedores_reais, p_foco)
                    
                    status = "NÃO BATEU"
                    if real_total >= m2 and m2 > 0: status = "META 2 ✅"
                    elif real_total >= m1 and m1 > 0: status = "META 1 ✅"
                    
                    rows_gc.append({
                        "Coordenador": c_meta, "Meta 1": m1, "Meta 2": m2,
                        "Realizado Total": real_total, "Realizado Foco": real_foco, "Resultado": status
                    })
            if rows_gc:
                st.dataframe(pd.DataFrame(rows_gc), use_container_width=True, hide_index=True)

    st.markdown(f'<div style="text-align:center; color:{COR_TEXTO_MUTED}; font-size:0.85rem; margin-top:3rem;">Direcional Engenharia • Vendas RJ — Premiações</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
