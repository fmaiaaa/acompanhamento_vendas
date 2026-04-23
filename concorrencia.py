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
# Funções de Design (Padrão Ficha/Gaps)
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
        html, body {{ font-family: 'Inter', sans-serif; color: {COR_TEXTO_LABEL}; }}
        .stApp {{ background: linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%), url("{bg_url}") center / cover no-repeat !important; }}
        .block-container {{ max-width: 1700px !important; padding: 1.45rem 2.25rem; background: rgba(255, 255, 255, 0.78) !important; backdrop-filter: blur(18px); border-radius: 24px; border: 1px solid rgba(255, 255, 255, 0.45); box-shadow: 0 4px 6px -1px rgba({RGB_AZUL_CSS}, 0.06), 0 24px 48px -12px rgba({RGB_AZUL_CSS}, 0.18); animation: fichaFadeIn 0.7s both; }}
        h1, h2, h3, h4 {{ font-family: 'Montserrat', sans-serif; color: {COR_AZUL_ESC}; font-weight: 800; text-align: center; }}
        .vel-kpi-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 1.25rem; }}
        .vel-kpi {{ flex: 1 1 18%; background: white; border: 1px solid {COR_BORDA}; border-radius: 14px; padding: 14px; text-align: center; transition: transform 0.3s ease; }}
        .vel-kpi:hover {{ transform: translateY(-4px); box-shadow: 0 10px 20px -5px rgba({RGB_AZUL_CSS}, 0.15); }}
        .vel-kpi .lbl {{ font-size: 0.72rem; font-weight: 700; text-transform: uppercase; color: {COR_AZUL_ESC}; }}
        .vel-kpi .val {{ font-family: 'Montserrat', sans-serif; font-size: 1.35rem; font-weight: 800; color: {COR_AZUL_ESC}; }}
        </style>
        """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Lógicas de Dados
# -----------------------------------------------------------------------------
def _secrets_connections_gsheets() -> Dict[str, Any]:
    try: return dict(st.secrets["connections"]["gsheets"])
    except: return {}

def montar_service_account_info(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not raw: return None
    out = {k: v for k, v in raw.items() if v}
    if "private_key" in out: out["private_key"] = out["private_key"].replace("\\n", "\n")
    return out

@st.cache_data(ttl=300, show_spinner=False)
def ler_aba(spreadsheet_id: str, worksheet: str) -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials
    info = montar_service_account_info(_secrets_connections_gsheets())
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
    st.markdown("#### Análise de Concorrência")
    st.markdown("### Inteligência Competitiva e Performance de Mercado")

    with st.spinner("Consolidando bases de mercado..."):
        # 1. Carregando Bases
        try:
            df_detalhada = ler_aba(SPREADSHEET_ID_CONC, "BD DETALHADA")
            df_geral = ler_aba(SPREADSHEET_ID_CONC, "BD GERAL")
            df_mensal = ler_aba(SPREADSHEET_ID_CONC, "Abr/2026") # Mês base
            df_dados_dir = ler_aba(SPREADSHEET_ID_CONC, "DADOS DIRECIONAL")
            df_vendas_dir = ler_aba(SPREADSHEET_ID_CONC, "Controle de Vendas RJ - Periodo Integral")
            df_estoque_dir = ler_aba(SPREADSHEET_ID_CONC, "Automação Estudos Concorrentes - Lucas")
        except Exception as e:
            st.error(f"Erro ao carregar abas: {e}")
            return

    # 2. Processamento e Enriquecimento Concorrentes
    df_detalhada["Preço_Float"] = df_detalhada["PREÇO"].apply(parse_val)
    df_detalhada["Preço_m2"] = df_detalhada["PREÇO_M2"].apply(parse_val)
    
    # Merge Master
    df_master = df_detalhada.merge(df_geral[["Empreendimento", "Venda a Partir", "Previsão"]], left_on="EMPREENDIMENTOVendas", right_on="Empreendimento", how="left", suffixes=('', '_geral'))
    df_master = df_master.merge(df_mensal[["CHAVE", "Vendas (Qnt.)", "Estoque (Qnt.)", "VGV (R$)"]], on="CHAVE", how="left")
    
    # KPI: Taxa de Absorção
    df_master["Vendas_Num"] = pd.to_numeric(df_master["Vendas (Qnt.)"], errors="coerce").fillna(0)
    df_master["Estoque_Num"] = pd.to_numeric(df_master["Estoque (Qnt.)"], errors="coerce").fillna(0)
    df_master["Absorcao"] = df_master["Vendas_Num"] / (df_master["Vendas_Num"] + df_master["Estoque_Num"])
    df_master["Absorcao"] = df_master["Absorcao"].fillna(0)

    # 3. Processamento Direcional (Filtrado por Endereço)
    allowed_addresses = df_dados_dir["Endereço"].unique()
    
    # Limpeza e Filtro Direcional
    df_vendas_dir["Valor_Float"] = df_vendas_dir["Valor Real de Venda"].apply(parse_val)
    # Nota: No gsheets da Direcional, assumimos que há uma coluna Identificador ou similar para bater com o endereço
    # Para este MVP, usaremos o Nome do Empreendimento mapeado
    allowed_emps = df_dados_dir["Nome do Empreendimento (Chave)"].unique()
    df_v_dir_f = df_vendas_dir[df_vendas_dir["Empreendimento"].isin(allowed_emps)].copy()
    
    # -------------------------------------------------------------------------
    # Filtros
    # -------------------------------------------------------------------------
    st.markdown("<div style='text-align:center; font-weight:bold;'>Filtros Estratégicos</div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: f_regiao = st.multiselect("Região", sorted(df_master["REGIÃO"].unique()))
    with c2: f_bairro = st.multiselect("Bairro", sorted(df_master["BAIRRO"].unique()))
    with c3: f_conc = st.multiselect("Concorrente", sorted(df_master["CONCORRENTE"].unique()))
    with c4: f_tipo = st.multiselect("Tipologia", sorted(df_master["TIPOLOGIAS"].unique() if "TIPOLOGIAS" in df_master.columns else []))

    df_f = df_master.copy()
    if f_regiao: df_f = df_f[df_f["REGIÃO"].isin(f_regiao)]
    if f_bairro: df_f = df_f[df_f["BAIRRO"].isin(f_bairro)]
    if f_conc: df_f = df_f[df_f["CONCORRENTE"].isin(f_conc)]

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
            <div class="vel-kpi"><div class="lbl">Vendas no Mês (Mercado)</div><div class="val">{int(total_vendas)}</div></div>
            <div class="vel-kpi"><div class="lbl">Nº Empreendimentos</div><div class="val">{df_f['CHAVE'].nunique()}</div></div>
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
            text=df_f["EMPREENDIMENTOVendas"],
            marker=dict(size=12, color=df_f["Preço_Float"], colorscale='Viridis', showscale=True)
        ))
        fig_demand.update_layout(xaxis_title="R$ / m²", yaxis_title="Absorção (%)", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_demand, use_container_width=True)

    with col_g2:
        st.subheader("Eficiência por Concorrente")
        df_eff = df_f.groupby("CONCORRENTE").agg({"Absorcao": "mean", "Vendas_Num": "sum"}).sort_values("Absorcao", ascending=False)
        fig_eff = go.Figure(go.Bar(
            x=df_eff.index, y=df_eff["Absorcao"] * 100,
            marker_color=COR_AZUL_ESC
        ))
        fig_eff.update_layout(yaxis_title="Absorção Média (%)", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_eff, use_container_width=True)

    # -------------------------------------------------------------------------
    # Oportunidades e Gaps de Mercado
    # -------------------------------------------------------------------------
    st.subheader("Gap de Mercado: Análise por Bairro")
    df_gap = df_f.groupby("BAIRRO").agg(
        Oferta_Ativa=("CHAVE", "nunique"),
        Preco_m2_Medio=("Preço_m2", "mean"),
        Absorcao_Media=("Absorcao", "mean")
    ).reset_index()
    
    def identificar_status(row):
        if row['Absorcao_Media'] > avg_abs/100 and row['Oferta_Ativa'] < df_gap['Oferta_Ativa'].median():
            return "🔥 Oportunidade: Alta Demanda / Baixa Oferta"
        if row['Absorcao_Media'] < avg_abs/100 and row['Oferta_Ativa'] > df_gap['Oferta_Ativa'].median():
            return "⚠️ Atenção: Mercado Saturado"
        return "Alinhado ao Mercado"

    df_gap["Status"] = df_gap.apply(identificar_status, axis=1)
    df_gap["Absorcao_Media"] = (df_gap["Absorcao_Media"] * 100).map("{:.1f}%".format)
    df_gap["Preco_m2_Medio"] = df_gap["Preco_m2_Medio"].map("R$ {:,.2f}".format)
    
    st.dataframe(df_gap.sort_values("Oferta_Ativa", ascending=False), use_container_width=True, hide_index=True)

    st.markdown(f'<div style="text-align:center; padding:1rem; color:{COR_TEXTO_MUTED}; font-size:0.8rem;">Direcional Engenharia · Inteligência de Mercado</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
