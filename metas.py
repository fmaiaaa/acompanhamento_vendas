# -*- coding: utf-8 -*-
"""
Acompanhamento de vendas — metas vs realizado (Direcional).
Focado em Premiações: Coordenadores IMOB, Comerciais e Grandes Contas.
Design sofisticado (revertido ao original) com filtros simplificados.
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

# Paleta Direcional
COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_BORDA = "#eef2f6"
COR_TEXTO_MUTED = "#64748b"
COR_TEXTO_LABEL = "#1e293b"
COR_INPUT_BG = "#f0f2f6"

def _hex_rgb_triplet(hex_color: str) -> str:
    x = (hex_color or "").strip().lstrip("#")
    if len(x) != 6: return "0, 0, 0"
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}"

RGB_AZUL_CSS = _hex_rgb_triplet(COR_AZUL_ESC)
RGB_VERMELHO_CSS = _hex_rgb_triplet(COR_VERMELHO)

# -----------------------------------------------------------------------------
# Funções de Design (Revertido ao Estilo Inicial)
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
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap');
        
        @keyframes fichaFadeIn {{
            from {{ opacity: 0; transform: translateY(18px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        @keyframes fichaShimmer {{
            0% {{ background-position: 0% 50%; }}
            100% {{ background-position: 200% 50%; }}
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

        html, body, [data-testid="stApp"] {{
            color-scheme: light !important;
        }}
        .stApp {{
            background: 
                linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%),
                url("{bg_url}") center / cover no-repeat !important;
            background-attachment: fixed !important;
        }}
        .block-container {{
            max-width: 1600px !important;
            margin-top: 1vh !important;
            padding: 2.5rem !important;
            background: rgba(255, 255, 255, 0.78) !important;
            backdrop-filter: blur(18px) saturate(1.15);
            -webkit-backdrop-filter: blur(18px) saturate(1.15);
            border-radius: 24px !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            box-shadow: 0 24px 48px -12px rgba({RGB_AZUL_CSS}, 0.18);
            animation: fichaFadeIn 0.7s cubic-bezier(0.22, 1, 0.36, 1) both;
        }}
        
        .ficha-logo-wrap {{ text-align: center; padding: 0.5rem 0 1rem 0; }}
        .ficha-logo-wrap img {{ max-height: 75px; width: auto; }}
        
        .ficha-hero-stack {{ text-align: center; margin-bottom: 1.5rem; }}
        .ficha-hero-bar {{
            height: 4px; width: 100%; border-radius: 999px;
            background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC});
            background-size: 200% 100%;
            animation: fichaShimmer 4s ease-in-out infinite alternate;
            margin-top: 10px;
        }}
        
        h1, h2, h3 {{ 
            font-family: 'Montserrat', sans-serif !important; 
            color: {COR_AZUL_ESC} !important; 
            font-weight: 800 !important;
        }}

        /* Tabelas */
        [data-testid="stDataFrame"] {{
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid {COR_BORDA};
        }}
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
    if pd.isna(val) or val == "" or val == "None": return 0.0
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
    """Busca no dicionário os proprietários vinculados ao coordenador."""
    # Garante que as colunas sejam as corretas
    # Dicionário: Coordenador | Proprietário da Oportunidade
    df_dic.columns = ["COORDENADOR", "PROPRIETARIO"]
    mask = df_dic["COORDENADOR"].astype(str).str.strip().str.lower() == nome_coord_tabela.strip().lower()
    return df_dic[mask]["PROPRIETARIO"].astype(str).str.strip().unique().tolist()

def calcular_realizado(df_vendas: pd.DataFrame, vendedores: List[str] = None, emp: Optional[str] = None, ignora_vendedor: bool = False) -> int:
    """Soma as vendas filtradas pelo dicionário ou apenas por empreendimento."""
    mask = pd.Series([True] * len(df_vendas), index=df_vendas.index)
    
    col_prop = "Proprietário da oportunidade"
    col_emp = "Empreendimento"
    
    # Filtro por vendedores (apenas se não for para ignorar)
    if not ignora_vendedor:
        if vendedores:
            mask &= df_vendas[col_prop].astype(str).str.strip().isin(vendedores)
        else:
            return 0
    
    # Filtra pelo empreendimento se houver
    if emp:
        mask &= (df_vendas[col_emp].astype(str).str.strip().str.lower() == str(emp).strip().lower())
    
    return int(len(df_vendas[mask]))

# -----------------------------------------------------------------------------
# Main App
# -----------------------------------------------------------------------------
def main():
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(page_title="Premiações Direcional", page_icon=str(fav) if fav else None, layout="wide")
    aplicar_estilo()
    _exibir_logo_topo()
    
    st.markdown('<div class="ficha-hero-stack"><h2>Painel de Premiações de Vendas</h2><div class="ficha-hero-bar"></div></div>', unsafe_allow_html=True)

    sid = _secrets_connections_gsheets().get("spreadsheet_id", SPREADSHEET_ID)

    try:
        df_vendas_raw = ler_aba_df(sid, WS_VENDAS)
        df_dic = ler_aba_df(sid, WS_DICIONARIO)
        df_metas_imob = ler_aba_df(sid, WS_METAS_IMOB)
        df_metas_comerciais = ler_aba_df(sid, WS_METAS_COMERCIAIS)
        df_metas_gc = ler_aba_df(sid, WS_METAS_GC)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return

    # --- Limpeza e Filtro Inicial (Venda Comercial? == 1) ---
    df_vendas = df_vendas_raw.copy()
    
    col_venda_comercial = "Venda Comercial?"
    if col_venda_comercial in df_vendas.columns:
        # Filtra apenas Vendas Comerciais iguais a 1
        df_vendas = df_vendas[df_vendas[col_venda_comercial].astype(str).str.strip() == "1"]
    
    # -------------------------------------------------------------------------
    # Filtros de Período (Ano e Mês)
    # -------------------------------------------------------------------------
    st.markdown("### Filtros de Período")
    c_f1, c_f2 = st.columns(2)
    with c_f1:
        anos_list = sorted(df_vendas["Ano da Venda"].astype(str).unique(), reverse=True)
        ano_sel = st.selectbox("Selecione o Ano", anos_list if anos_list else [str(datetime.now().year)])
    with c_f2:
        meses_list = [
            ("Janeiro", "01"), ("Fevereiro", "02"), ("Março", "03"), ("Abril", "04"), 
            ("Maio", "05"), ("Junho", "06"), ("Julho", "07"), ("Agosto", "08"), 
            ("Setembro", "09"), ("Outubro", "10"), ("Novembro", "11"), ("Dezembro", "12")
        ]
        mes_atual_idx = datetime.now().month - 1
        mes_sel_nome, mes_sel_val = st.selectbox("Selecione o Mês", meses_list, index=mes_atual_idx, format_func=lambda x: x[0])

    data_filtro = f"{mes_sel_val}/{ano_sel}"
    
    # Filtra vendas pelo período selecionado
    vendas_f = df_vendas[
        (df_vendas["Ano da Venda"].astype(str) == str(ano_sel)) & 
        (df_vendas["Mês Venda"].astype(str).str.zfill(2) == str(mes_sel_val))
    ]

    # -------------------------------------------------------------------------
    # Interface de Abas Centralizada
    # -------------------------------------------------------------------------
    tabs = st.tabs(["Coordenadores IMOB", "Coordenadores Comerciais", "Grandes Contas"])

    # --- ABA 1: IMOB (Filtrado por Coordenador + Empreendimento) ---
    with tabs[0]:
        st.subheader(f"Premiação IMOB — {mes_sel_nome}/{ano_sel}")
        df_imob = df_metas_imob[df_metas_imob["DATA"] == data_filtro]
        
        if df_imob.empty:
            st.info("Nenhuma meta cadastrada para este período na aba IMOB.")
        else:
            regioes = sorted(df_imob["REGIÃO"].unique())
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
                            
                            m_dir = int(parse_valor_br(r.get("META DIRECIONAL", 0)))
                            m_imob = int(parse_valor_br(r.get("META IMOB", 0)))
                            m_imob2 = int(parse_valor_br(r.get("META IMOB 2", 0)))
                            
                            res = "NÃO BATEU"
                            if realizado >= m_imob2 and m_imob2 > 0: res = "META IMOB 2 ✅"
                            elif realizado >= m_imob and m_imob > 0: res = "META IMOB ✅"
                            elif realizado >= m_dir and m_dir > 0: res = "META DIRECIONAL ✅"
                            
                            rows.append({
                                "Coordenador": c, "Empreendimento": emp,
                                "Meta Dir.": m_dir, "Meta IMOB": m_imob, "Meta IMOB 2": m_imob2,
                                "Realizado": realizado, "Resultado": res
                            })
                    if rows:
                        st.table(pd.DataFrame(rows))

    # --- ABA 2: COMERCIAIS (Realizado é o Total do Empreendimento) ---
    with tabs[1]:
        st.subheader(f"Premiação Comerciais — {mes_sel_nome}/{ano_sel}")
        df_com = df_metas_comerciais[df_metas_comerciais["DATA"] == data_filtro]
        
        if df_com.empty:
            st.info("Nenhuma meta cadastrada para este período na aba Comerciais.")
        else:
            # Lista única de coordenadores presentes nas metas do mês para separação visual
            all_coords_in_month = []
            for _, r in df_com.iterrows():
                all_coords_in_month.extend(extrair_lista_coords(r["COORDENADORES"]))
            unique_coords = sorted(list(set(all_coords_in_month)))
            
            for c_name in unique_coords:
                st.markdown(f"#### Coordenador: {c_name}")
                rows_c = []
                # Filtrar metas onde este coordenador está incluído
                df_c_metas = df_com[df_com["COORDENADORES"].str.contains(c_name, case=False, na=False)]
                for _, r in df_c_metas.iterrows():
                    emp = r["EMPREENDIMENTO"]
                    
                    # Conforme solicitado: Coordenadores não importam para o cômputo.
                    # Deve ser computado o valor TOTAL de cada empreendimento no período.
                    realizado = calcular_realizado(vendas_f, emp=emp, ignora_vendedor=True)
                    
                    m_desafio = int(parse_valor_br(r.get("META DESAFIO VENDAS", 0)))
                    m_bp = int(parse_valor_br(r.get("META BP", 0)))
                    m_bp70 = int(parse_valor_br(r.get("META BP 70%", 0)))
                    
                    res = "NÃO BATEU"
                    if realizado >= m_desafio and m_desafio > 0: res = "DESAFIO ✅"
                    elif realizado >= m_bp and m_bp > 0: res = "BP ✅"
                    elif realizado >= m_bp70 and m_bp70 > 0: res = "BP 70% ✅"
                    
                    rows_c.append({
                        "Empreendimento": emp, 
                        "Meta Desafio": m_desafio, 
                        "Meta BP": m_bp,
                        "Meta BP 70%": m_bp70, 
                        "Realizado": realizado, 
                        "Resultado": res
                    })
                if rows_c:
                    st.dataframe(pd.DataFrame(rows_c), use_container_width=True, hide_index=True)

    # --- ABA 3: GRANDES CONTAS (Filtrado por Coordenador) ---
    with tabs[2]:
        st.subheader(f"Premiação Grandes Contas — {mes_sel_nome}/{ano_sel}")
        df_gc_periodo = df_metas_gc[df_metas_gc["DATA"] == data_filtro]
        
        if df_gc_periodo.empty:
            st.info("Nenhuma meta cadastrada para este período na aba Grandes Contas.")
        else:
            rows_gc = []
            for _, r in df_gc_periodo.iterrows():
                coords = extrair_lista_coords(r["COORDENADORES"])
                m1 = int(parse_valor_br(r.get("META 1", 0)))
                m2 = int(parse_valor_br(r.get("META 2", 0)))
                foco_list = extrair_lista_coords(r.get("PRODUTOS FOCO", ""))
                
                for c in coords:
                    vendedores = get_vendedores_do_coordenador(df_dic, c)
                    # Realizado total do coordenador (sem filtro de empreendimento)
                    real_total = calcular_realizado(vendas_f, vendedores)
                    
                    # Realizado Foco (apenas nos produtos da lista)
                    real_foco = 0
                    for p_foco in foco_list:
                        real_foco += calcular_realizado(vendas_f, vendedores, p_foco)
                    
                    res = "NÃO BATEU"
                    if real_total >= m2 and m2 > 0: res = "META 2 ✅"
                    elif real_total >= m1 and m1 > 0: res = "META 1 ✅"
                    
                    rows_gc.append({
                        "Coordenador": c, "Meta 1": m1, "Meta 2": m2,
                        "Realizado Total": real_total, "Realizado Foco": real_foco, "Resultado": res
                    })
            if rows_gc:
                st.dataframe(pd.DataFrame(rows_gc), use_container_width=True, hide_index=True)

    st.markdown(f'<div style="text-align:center; color:{COR_TEXTO_MUTED}; font-size:0.85rem; margin-top:3rem;">Direcional Engenharia • Vendas RJ — Premiações</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
