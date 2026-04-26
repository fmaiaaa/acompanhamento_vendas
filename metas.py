# -*- coding: utf-8 -*-
"""
Acompanhamento de vendas — metas vs realizado (Direcional).
Focado em Premiações: Coordenadores IMOB, Comerciais e Grandes Contas.
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

COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_BORDA = "#eef2f6"
COR_TEXTO_MUTED = "#64748b"
COR_TEXTO_LABEL = "#1e293b"
COR_INPUT_BG = "#f0f2f6"

MESES_TEXTO_MAP = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
}

def _hex_rgb_triplet(hex_color: str) -> str:
    x = (hex_color or "").strip().lstrip("#")
    if len(x) != 6: return "0, 0, 0"
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}"

RGB_AZUL_CSS = _hex_rgb_triplet(COR_AZUL_ESC)
RGB_VERMELHO_CSS = _hex_rgb_triplet(COR_VERMELHO)

# -----------------------------------------------------------------------------
# Funções de Design e Auxiliares
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

def aplicar_estilo() -> None:
    bg_url = _css_url_fundo_cadastro()
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;800;900&family=Inter:wght@400;500;600;700&display=swap');
        
        /* Centralização das Abas */
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {{
            display: flex;
            justify-content: center;
            width: 100%;
            gap: 10px;
        }}
        div[data-testid="stTabs"] button[data-baseweb="tab"] {{
            background-color: rgba(255,255,255,0.5);
            border-radius: 8px 8px 0 0;
            padding: 10px 20px;
        }}

        html, body, [data-testid="stApp"] {{
            color-scheme: light !important;
            font-family: 'Inter', sans-serif;
        }}
        .stApp {{
            background: linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.85) 0%, rgba({RGB_VERMELHO_CSS}, 0.2) 100%), url("{bg_url}") center / cover no-repeat !important;
            background-attachment: fixed !important;
        }}
        .block-container {{
            max-width: 1600px !important;
            padding: 2rem !important;
            background: rgba(255, 255, 255, 0.82) !important;
            backdrop-filter: blur(15px);
            border-radius: 24px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.2);
            margin-top: 2vh !important;
        }}
        h1, h2, h3 {{ font-family: 'Montserrat', sans-serif !important; color: {COR_AZUL_ESC} !important; text-align: center; }}
        .ficha-logo-wrap {{ text-align: center; padding-bottom: 1rem; }}
        .ficha-logo-wrap img {{ max-height: 80px; }}
        
        .vel-kpi-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin: 1rem 0; justify-content: center; }}
        .vel-kpi {{
            flex: 1 1 200px;
            max-width: 300px;
            background: white;
            border: 1px solid {COR_BORDA};
            border-radius: 12px;
            padding: 15px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        }}
        .vel-kpi .lbl {{ font-size: 0.75rem; font-weight: 700; color: {COR_TEXTO_MUTED}; text-transform: uppercase; }}
        .vel-kpi .val {{ font-size: 1.5rem; font-weight: 800; color: {COR_AZUL_ESC}; margin-top: 5px; }}
        
        .status-tag {{
            padding: 4px 10px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.85rem;
        }}
        .status-bateu {{ background: #dcfce7; color: #166534; }}
        .status-nao {{ background: #fee2e2; color: #991b1b; }}
        </style>
        """,
        unsafe_allow_html=True,
    )

# -----------------------------------------------------------------------------
# Integração GSheets
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
    if pd.isna(val) or val == "": return 0.0
    s = str(val).replace("R$", "").replace(".", "").replace(",", ".").strip()
    try: return float(s)
    except: return 0.0

def extrair_lista_coords(val: str) -> List[str]:
    s = str(val).strip()
    if s.startswith("{") and s.endswith("}"): s = s[1:-1]
    return [c.strip() for c in s.split(",") if c.strip()]

# -----------------------------------------------------------------------------
# Lógica do Dicionário e Realizado
# -----------------------------------------------------------------------------
def get_vendedores_do_coordenador(df_dic: pd.DataFrame, nome_coord_tabela: str) -> List[str]:
    """Retorna lista de 'Proprietários' associados ao 'Coordenador' no Dicionário."""
    # Col 0: Coordenador (Tabela), Col 1: Proprietário (BD)
    df_dic.columns = ["COORDENADOR", "PROPRIETARIO"]
    mask = df_dic["COORDENADOR"].astype(str).str.strip().str.lower() == nome_coord_tabela.strip().lower()
    return df_dic[mask]["PROPRIETARIO"].astype(str).str.strip().unique().tolist()

def calcular_realizado(df_vendas: pd.DataFrame, vendedores: List[str], emp: Optional[str] = None) -> float:
    """Soma as vendas no BD onde o Proprietário está na lista e opcionalmente filtra por empreendimento."""
    if not vendedores: return 0.0
    
    col_prop = "Proprietário da oportunidade"
    col_emp = "Empreendimento" # Normalizado no main
    
    mask = df_vendas[col_prop].astype(str).str.strip().isin(vendedores)
    if emp:
        mask &= (df_vendas[col_emp].astype(str).str.strip().str.lower() == str(emp).strip().lower())
    
    return float(len(df_vendas[mask]))

# -----------------------------------------------------------------------------
# Main App
# -----------------------------------------------------------------------------
def main():
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(page_title="Premiações Direcional", page_icon=str(fav) if fav else None, layout="wide")
    aplicar_estilo()
    _exibir_logo_topo()

    sid = _secrets_connections_gsheets().get("spreadsheet_id", SPREADSHEET_ID)

    try:
        df_vendas_raw = ler_aba_df(sid, WS_VENDAS)
        df_dic = ler_aba_df(sid, WS_DICIONARIO)
        df_metas_imob = ler_aba_df(sid, WS_METAS_IMOB)
        df_metas_comerciais = ler_aba_df(sid, WS_METAS_COMERCIAIS)
        df_metas_gc = ler_aba_df(sid, WS_METAS_GC)
    except Exception as e:
        st.error(f"Erro na conexão: {e}")
        return

    # Normalização de colunas Vendas
    df_vendas = df_vendas_raw.copy()
    col_mes = "Mês da Venda" if "Mês da Venda" in df_vendas.columns else df_vendas.columns[1] # Fallback
    col_ano = "Ano da Venda" if "Ano da Venda" in df_vendas.columns else df_vendas.columns[0]
    
    # Normalização de Nomes de Coluna para Vendas
    def find_and_rename(df, aliases, target):
        for a in aliases:
            for c in df.columns:
                if a.lower() in c.lower():
                    df.rename(columns={c: target}, inplace=True)
                    return target
        return None

    find_and_rename(df_vendas, ["Proprietário", "Proprietario"], "Proprietário da oportunidade")
    find_and_rename(df_vendas, ["Empreendimento", "Obra"], "Empreendimento")
    find_and_rename(df_vendas, ["Mês", "Mes"], "MES")
    find_and_rename(df_vendas, ["Ano"], "ANO")

    # -------------------------------------------------------------------------
    # Filtros Centrais (Apenas Ano e Mês)
    # -------------------------------------------------------------------------
    st.markdown("### Filtros de Período")
    c_f1, c_f2 = st.columns(2)
    with c_f1:
        anos_list = sorted(df_vendas["ANO"].unique(), reverse=True)
        ano_sel = st.selectbox("Selecione o Ano", anos_list)
    with c_f2:
        # Extrair meses e formatar
        meses_num = sorted([m for m in [1,2,3,4,5,6,7,8,9,10,11,12]], reverse=False)
        mes_sel = st.selectbox("Selecione o Mês", meses_num, index=datetime.now().month-1)

    data_filtro = f"{mes_sel:02d}/{ano_sel}"
    
    # Filtrar Base de Vendas para o Mês/Ano selecionado
    # Tentativa de match flexível para data (MM/AAAA ou apenas Mês e Ano separados)
    mask_vendas = (df_vendas["ANO"].astype(str) == str(ano_sel))
    # Se a coluna MES for texto "jan", "fev"...
    try:
        if df_vendas["MES"].iloc[0].lower() in MESES_TEXTO_MAP:
             mask_vendas &= (df_vendas["MES"].str.lower().map(MESES_TEXTO_MAP) == mes_sel)
        else:
             mask_vendas &= (df_vendas["MES"].astype(float).astype(int) == mes_sel)
    except: pass
    
    vendas_f = df_vendas[mask_vendas]

    # -------------------------------------------------------------------------
    # Interface de Abas Centralizada
    # -------------------------------------------------------------------------
    tabs = st.tabs(["Coordenadores IMOB", "Coordenadores Comerciais", "Grandes Contas"])

    # --- ABA 1: IMOB ---
    with tabs[0]:
        st.subheader(f"Premiação IMOB — Competência {data_filtro}")
        df_imob = df_metas_imob[df_metas_imob["DATA"] == data_filtro]
        
        if df_imob.empty:
            st.info("Nenhuma meta cadastrada para este período na aba IMOB.")
        else:
            regioes = df_imob["REGIÃO"].unique()
            for reg in regioes:
                with st.expander(f"Região: {reg}", expanded=True):
                    df_reg = df_imob[df_imob["REGIÃO"] == reg]
                    rows = []
                    for _, r in df_reg.iterrows():
                        coords = extrair_lista_coords(r["COORDENADORES"])
                        emp = r["EMPREENDIMENTO"]
                        for c in coords:
                            vendedores = get_vendedores_do_coordenador(df_dic, c)
                            realizado = calcular_realizado(vendas_f, vendedores, emp)
                            
                            m_dir = parse_valor_br(r.get("META DIRECIONAL", 0))
                            m_imob = parse_valor_br(r.get("META IMOB", 0))
                            m_imob2 = parse_valor_br(r.get("META IMOB 2", 0))
                            
                            status = "NÃO BATEU META"
                            if realizado >= m_imob2 and m_imob2 > 0: status = "BATEU META IMOB 2"
                            elif realizado >= m_imob and m_imob > 0: status = "BATEU META IMOB"
                            elif realizado >= m_dir and m_dir > 0: status = "BATEU META DIRECIONAL"
                            
                            rows.append({
                                "Coordenador": c, "Empreendimento": emp,
                                "M. Direcional": m_dir, "M. IMOB": m_imob, "M. IMOB 2": m_imob2,
                                "Realizado": realizado, "Resultado": status
                            })
                    st.table(pd.DataFrame(rows))

    # --- ABA 2: COMERCIAIS ---
    with tabs[1]:
        st.subheader(f"Premiação Comerciais — Competência {data_filtro}")
        df_com = df_metas_comerciais[df_metas_comerciais["DATA"] == data_filtro]
        
        if df_com.empty:
            st.info("Nenhuma meta cadastrada para este período na aba Comerciais.")
        else:
            coords_com = []
            for _, r in df_com.iterrows():
                for c in extrair_lista_coords(r["COORDENADORES"]): coords_com.append(c)
            
            coords_com = sorted(list(set(coords_com)))
            for c_name in coords_com:
                st.markdown(f"#### Coordenador: {c_name}")
                rows_c = []
                # Encontrar todas as metas onde este coordenador aparece
                df_c_metas = df_com[df_com["COORDENADORES"].str.contains(c_name, case=False)]
                for _, r in df_c_metas.iterrows():
                    emp = r["EMPREENDIMENTO"]
                    vendedores = get_vendedores_do_coordenador(df_dic, c_name)
                    realizado = calcular_realizado(vendas_f, vendedores, emp)
                    
                    m_desafio = parse_valor_br(r.get("META DESAFIO VENDAS", 0))
                    m_bp = parse_valor_br(r.get("META BP", 0))
                    m_bp70 = parse_valor_br(r.get("META BP 70%", 0))
                    
                    status = "NÃO BATEU META"
                    if realizado >= m_desafio and m_desafio > 0: status = "BATEU DESAFIO"
                    elif realizado >= m_bp and m_bp > 0: status = "BATEU BP"
                    elif realizado >= m_bp70 and m_bp70 > 0: status = "BATEU BP 70%"
                    
                    rows_c.append({
                        "Empreendimento": emp, "Meta Desafio": m_desafio, "Meta BP": m_bp,
                        "Meta BP 70%": m_bp70, "Realizado": realizado, "Resultado": status
                    })
                st.dataframe(pd.DataFrame(rows_c), use_container_width=True, hide_index=True)

    # --- ABA 3: GRANDES CONTAS ---
    with tabs[2]:
        st.subheader(f"Premiação Grandes Contas — Competência {data_filtro}")
        df_gc_periodo = df_metas_gc[df_metas_gc["DATA"] == data_filtro]
        
        if df_gc_periodo.empty:
            st.info("Nenhuma meta cadastrada para este período na aba Grandes Contas.")
        else:
            rows_gc = []
            for _, r in df_gc_periodo.iterrows():
                coords = extrair_lista_coords(r["COORDENADORES"])
                m1 = parse_valor_br(r.get("META 1", 0))
                m2 = parse_valor_br(r.get("META 2", 0))
                foco_list = extrair_lista_coords(r.get("PRODUTOS FOCO", ""))
                
                for c in coords:
                    vendedores = get_vendedores_do_coordenador(df_dic, c)
                    real_total = calcular_realizado(vendas_f, vendedores)
                    
                    # Realizado Foco
                    real_foco = 0
                    for p_foco in foco_list:
                        real_foco += calcular_realizado(vendas_f, vendedores, p_foco)
                    
                    res = "NÃO BATEU"
                    if real_total >= m2 and m2 > 0: res = "BATEU META 2"
                    elif real_total >= m1 and m1 > 0: res = "BATEU META 1"
                    
                    rows_gc.append({
                        "Coordenador": c, "Meta 1": m1, "Meta 2": m2,
                        "Realizado Total": real_total, "Realizado Foco": real_foco, "Resultado": res
                    })
            st.dataframe(pd.DataFrame(rows_gc), use_container_width=True, hide_index=True)

    st.markdown(f'<div style="text-align:center; color:{COR_TEXTO_MUTED}; font-size:0.8rem; margin-top:2rem;">Direcional Engenharia • Vendas RJ</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
