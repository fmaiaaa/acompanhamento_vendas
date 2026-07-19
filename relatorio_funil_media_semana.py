# -*- coding: utf-8 -*-
"""
Relatório 1 — Funil: médias (1a / 6m / 3m / 1m) × semana fixa
por Regional, Gerente de vendas e Corretor.

Hospedagem Streamlit Cloud (mesma pasta / secrets do velocímetro).
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# =============================================================================
# Módulo embutido (funil_pessoas_comum) — evita ModuleNotFoundError no Cloud
# =============================================================================

import copy
import os
import re
import unicodedata
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# Relatórios Salesforce (mesmos do velocímetro + dicionário gerente→regional)
SF_REPORT_AGENDAMENTOS_ID = "00OU600000AcFGPMA3"
SF_REPORT_PASTAS_ID = "00OU600000FEOoDMAX"
SF_REPORT_VENDAS_ID = "00O3Z000005ZsPmUAK"
SF_REPORT_DIC_REGIONAL_ID = "00OU600000FIH6bMAH"

FUNIL_ETAPAS = ("agendamentos", "visitas", "pastas", "pastas_aprovadas", "vendas")
FUNIL_LABELS = {
    "agendamentos": "Agendamentos",
    "visitas": "Visitas",
    "pastas": "Pastas",
    "pastas_aprovadas": "Pastas aprovadas",
    "vendas": "Vendas",
}
DIMENSOES = ("regional", "gerente", "corretor")
DIM_LABELS = {
    "regional": "Regional",
    "gerente": "Gerente de vendas",
    "corretor": "Corretor",
}

COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_TEXTO_PRETO = "#000000"
COR_FUNDO_CARD = "rgba(255, 255, 255, 0.78)"
COR_BORDA = "#eef2f6"

COLUNAS_PASTAS_ALIASES = [
    "Data Primeiro Envio Análise", "Data Primeiro Envio Analise",
    "Data do Primeiro Envio Análise", "Data do Primeiro Envio Analise",
    "Primeiro Envio Análise", "Primeiro Envio Analise",
    "Data 1º Envio Análise", "Data 1o Envio Analise",
    "Data da Análise", "Data da Analise", "Data Análise", "Data Analise",
]
COLUNAS_PASTAS_APROV_ALIASES = [
    "Data Aprovação SAFI", "Data Aprovacao SAFI",
    "Data da Aprovação SAFI", "Data da Aprovacao SAFI",
    "Data de Aprovação SAFI", "Data de Aprovacao SAFI",
    "Aprovação SAFI", "Aprovacao SAFI",
    "Data Aprov. SAFI", "Data Aprov SAFI",
]
ALIASES_VENDA_COMERCIAL = [
    "Venda Comercial?", "Venda Comercial", "Venda comercial?",
    "Venda comercial", "Comercial?",
]
ALIASES_ID_OPORTUNIDADE = [
    "ID da Oportunidade", "Id da Oportunidade", "Opportunity ID", "Opportunity Id",
    "ID Oportunidade", "Id Oportunidade",
]
ALIASES_CONTRATO_GERADO = [
    "Contrato gerado em", "Contrato Gerado em", "Contrato gerado",
    "Data do Contrato", "Data Contrato", "Close Date", "Data da venda", "Data Venda",
]
ALIASES_NOME_AVALIACAO_CREDITO = [
    "Nome da Avaliação de crédito", "Nome da Avaliacao de credito",
    "Nome da Avaliação de Crédito", "Nome da Avaliacao de Credito",
]
ALIASES_CODIGO_AGENDAMENTO = [
    "Código do agendamento", "Codigo do agendamento",
    "Código do Agendamento", "Codigo do Agendamento",
]
ALIASES_DATA_CRIACAO = [
    "Data de criação", "Data de criacao", "Data Criação", "Data Criacao",
    "Created Date", "Criado em",
]
ALIASES_DATA_VISITA = [
    "Data da visita", "Data da Visita", "Data visita", "Data Visita",
    "Activity Date", "Data da Atividade", "Data do agendamento",
    "Data Agendamento", "Start Date Time", "Data/Hora",
]

# Hierarquia — pastas
ALIASES_REGIONAL_PASTAS = [
    "Gerente Regional", "Gerente regional", "Regional",
]
ALIASES_REGIONAL_PASTAS_FALLBACK = [
    "Avaliação de crédito : Oportunidade : Gerente regional",
    "Avaliacao de credito : Oportunidade : Gerente regional",
    "Oportunidade : Gerente regional", "Oportunidade: Gerente regional",
]
ALIASES_GERENTE_PASTAS = [
    "Gerente Vendas", "Gerente de Vendas", "Gerente vendas", "Gerente de vendas",
]
ALIASES_CORRETOR_PASTAS = ["Corretor", "Nome do Corretor", "Corretor Nome"]

# Hierarquia — agendamentos
ALIASES_CORRETOR_AG = [
    "Corretor: Nome completo", "Corretor : Nome completo",
    "Corretor Nome completo", "Nome completo Corretor", "Corretor",
]
ALIASES_GERENTE_AG = [
    "Gerente de Vendas", "Gerente Vendas", "Gerente de vendas", "Gerente vendas",
]
ALIASES_REGIONAL_AG = [
    "Gerente Regional", "Gerente regional", "Regional",
]

# Hierarquia — vendas
ALIASES_GERENTE_VENDAS = [
    "Proprietário da oportunidade", "Proprietario da oportunidade",
    "Proprietário da Oportunidade", "Opportunity Owner",
]
ALIASES_CORRETOR_VENDAS = [
    "Contato Corretor Proprietário", "Contato Corretor Proprietario",
    "Corretor Proprietário", "Corretor Proprietario", "Contato Corretor",
]
ALIASES_REGIONAL_DIC = [
    "Gerente regional", "Gerente Regional", "Regional",
]


# -----------------------------------------------------------------------------
# Design (mesmo padrão do velocímetro / Ficha Direcional)
# -----------------------------------------------------------------------------
from pathlib import Path
import base64
import html as _html_mod

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"
COR_INPUT_BG = "#f0f2f6"
COR_FUNDO_CARD = "rgba(255, 255, 255, 0.78)"
COR_BORDA = "#eef2f6"


def _hex_rgb_triplet(hex_color: str) -> str:
    x = (hex_color or "").strip().lstrip("#")
    if len(x) != 6:
        return "0, 0, 0"
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}"


RGB_AZUL_CSS = _hex_rgb_triplet(COR_AZUL_ESC)
RGB_VERMELHO_CSS = _hex_rgb_triplet(COR_VERMELHO)


def _resolver_png_raiz(nome: str):
    for base in (_DIR_APP, _DIR_APP.parent):
        p = base / nome
        if p.is_file():
            return p
    return None


def _resolver_imagem_fundo_local(nome: str):
    for base in (_DIR_APP, _DIR_APP.parent):
        for ext in (".jpg", ".jpeg", ".JPG", ".JPEG", ".png", ".PNG"):
            p = base / f"{Path(nome).stem}{ext}"
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
    return (
        "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab"
        "?auto=format&fit=crop&w=1920&q=80"
    )


def _logo_arquivo_local():
    p_topo = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
    if p_topo:
        return str(p_topo)
    for name in ("logo_direcional.png", "logo_direcional.jpg", "logo.png"):
        p = _DIR_APP / "assets" / name
        if p.is_file():
            return str(p)
    return None


def _logo_url_secrets():
    try:
        if hasattr(st, "secrets"):
            b = st.secrets.get("branding")
            if isinstance(b, dict):
                u = (b.get("LOGO_URL") or "").strip()
                if u:
                    return u
    except Exception:
        pass
    return None


def _exibir_logo_topo() -> None:
    path = _logo_arquivo_local()
    url = _logo_url_secrets()
    try:
        if path:
            ext = Path(path).suffix.lower().lstrip(".")
            mime = "image/png" if ext == "png" else "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            st.markdown(
                f'<div class="ficha-logo-wrap"><img src="data:{mime};base64,{b64}" alt="Direcional" /></div>',
                unsafe_allow_html=True,
            )
            return
        if url:
            st.markdown(
                f'<div class="ficha-logo-wrap"><img src="{_html_mod.escape(url)}" alt="Direcional" /></div>',
                unsafe_allow_html=True,
            )
    except Exception:
        pass


def _cabecalho_pagina(titulo: str) -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">{_html_mod.escape(titulo)}</p>'
        f"</div>"
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true">'
        f'<div class="ficha-hero-bar"></div>'
        f"</div>"
        f"</div>",
        unsafe_allow_html=True,
    )


def aplicar_estilo() -> None:
    """Mesmo visual do velocímetro (fundo, glass card, tipografia) — sem sidebar."""
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
        html, body, :root, [data-testid="stApp"] {{
            color-scheme: light !important;
        }}
        /* Tabelas sempre claras, mesmo com Windows em modo escuro */
        .stDataFrame,
        [data-testid="stDataFrame"],
        [data-testid="stDataFrame"] > div,
        [data-testid="stDataFrameResizable"],
        [data-testid="stTable"],
        [data-testid="stTable"] table {{
            color-scheme: light !important;
            background-color: #ffffff !important;
        }}
        [data-testid="stTable"] th,
        [data-testid="stTable"] td {{
            background-color: #ffffff !important;
            color: #000000 !important;
        }}
        html, body {{
            font-family: 'Inter', sans-serif;
            color: {COR_TEXTO_PRETO};
            background: transparent !important;
        }}
        .stApp,
        [data-testid="stApp"] {{
            background:
                linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%),
                url("{bg_url}") center / cover no-repeat !important;
            background-attachment: scroll !important;
            background-color: transparent !important;
        }}
        [data-testid="stAppViewContainer"] {{
            background: transparent !important;
        }}
        header[data-testid="stHeader"],
        [data-testid="stHeader"] {{
            background: transparent !important;
            border: none !important;
            box-shadow: none !important;
        }}
        [data-testid="stToolbar"] {{
            background: transparent !important;
            color: rgba(255, 255, 255, 0.92) !important;
        }}
        [data-testid="stToolbar"] button,
        [data-testid="stToolbar"] a,
        [data-testid="stHeader"] button {{
            color: rgba(255, 255, 255, 0.92) !important;
            background: transparent !important;
        }}
        [data-testid="stMain"] {{
            padding-left: clamp(14px, 5vw, 56px) !important;
            padding-right: clamp(14px, 5vw, 56px) !important;
            padding-top: clamp(12px, 3.5vh, 40px) !important;
            padding-bottom: clamp(14px, 4vh, 44px) !important;
            box-sizing: border-box !important;
        }}
        .block-container {{
            max-width: 1700px !important;
            margin-left: auto !important;
            margin-right: auto !important;
            padding: 1.45rem 2.25rem 1.55rem 2.25rem !important;
            background: rgba(255, 255, 255, 0.96) !important;
            backdrop-filter: blur(18px) saturate(1.15);
            -webkit-backdrop-filter: blur(18px) saturate(1.15);
            border-radius: 24px !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            box-shadow:
                0 4px 6px -1px rgba({RGB_AZUL_CSS}, 0.06),
                0 24px 48px -12px rgba({RGB_AZUL_CSS}, 0.18),
                inset 0 1px 0 rgba(255, 255, 255, 0.55) !important;
            animation: fichaFadeIn 0.7s cubic-bezier(0.22, 1, 0.36, 1) both;
        }}
        /* Sem sidebar + forçar tema claro */
        [data-testid="stSidebar"],
        [data-testid="stSidebarNav"],
        section[data-testid="stSidebar"],
        button[kind="header"] {{
            display: none !important;
        }}
        [data-testid="stAppViewContainer"] > .main {{
            margin-left: 0 !important;
        }}
        [data-testid="stApp"],
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"],
        .main,
        .block-container,
        [data-testid="stMarkdownContainer"],
        [data-testid="stVerticalBlock"],
        [data-baseweb="base-input"],
        [data-baseweb="input"],
        [data-baseweb="select"] > div,
        [data-baseweb="popover"],
        .stSelectbox div[data-baseweb="select"] > div,
        .stMultiSelect div[data-baseweb="select"] > div,
        .stDateInput > div > div,
        .stNumberInput > div > div,
        .stRadio,
        .stTabs [data-baseweb="tab-list"],
        .stDataFrame,
        [data-testid="stDataFrame"] {{
            color-scheme: light !important;
        }}
        .stButton > button {{
            width: 100% !important;
            background: linear-gradient(135deg, {COR_AZUL_ESC} 0%, #0a5bb8 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 10px !important;
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 700 !important;
            box-shadow: 0 4px 12px rgba({RGB_AZUL_CSS}, 0.22) !important;
        }}
        .stButton > button:hover {{
            background: linear-gradient(135deg, {COR_VERMELHO} 0%, #e11d48 100%) !important;
            color: #ffffff !important;
        }}
        .stButton > button[kind="secondary"],
        .stButton > button[data-testid="baseButton-secondary"] {{
            background: #ffffff !important;
            color: {COR_AZUL_ESC} !important;
            border: 1.5px solid {COR_AZUL_ESC} !important;
        }}
        div[data-baseweb="select"] > div,
        .stMultiSelect [data-baseweb="tag"] {{
            background-color: #ffffff !important;
            color: {COR_TEXTO_PRETO} !important;
        }}
        .stCaption, [data-testid="stCaptionContainer"] {{
            text-align: center !important;
        }}
        .stTabs [data-baseweb="tab-list"] {{
            justify-content: center !important;
            gap: 0.35rem !important;
        }}
        .stTabs [data-baseweb="tab"] {{
            color: {COR_AZUL_ESC} !important;
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 700 !important;
        }}
        .stTabs [aria-selected="true"] {{
            color: {COR_AZUL_ESC} !important;
            border-bottom-color: {COR_VERMELHO} !important;
        }}
        .bloco-pessoa .nome,
        .ficha-hero .ficha-title {{
            color: {COR_AZUL_ESC} !important;
        }}
        h1, h2, h3, h4,
        h1 *, h2 *, h3 *, h4 *,
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        [data-testid="stMarkdownContainer"] h4,
        .stHeading, .stHeading * {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_AZUL_ESC} !important;
            font-weight: 800 !important;
        }}
        h1, h2, h3, h4, .stHeading {{
            text-align: center !important;
        }}
        h5, h6,
        h5 *, h6 *,
        [data-testid="stMarkdownContainer"] h5,
        [data-testid="stMarkdownContainer"] h6 {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_TEXTO_PRETO} !important;
            font-weight: 700 !important;
            text-align: center !important;
        }}
        .block-container > div p,
        .block-container label,
        [data-testid="stMarkdownContainer"] > p,
        [data-testid="stCaption"],
        [data-testid="stCaptionContainer"] p,
        [data-testid="stWidgetLabel"] p {{
            color: {COR_TEXTO_PRETO} !important;
        }}
        .ficha-logo-wrap {{
            text-align: center;
            padding: 0.1rem 0 0.45rem 0;
        }}
        .ficha-logo-wrap img {{
            max-height: 72px;
            width: auto;
            max-width: min(280px, 85vw);
            height: auto;
            object-fit: contain;
            display: inline-block;
        }}
        .ficha-hero-stack {{
            width: 100%;
            margin-bottom: 0.35rem;
        }}
        .ficha-hero {{
            text-align: center;
            padding: 0.5rem 0 0 0;
            margin: 0 auto;
            max-width: 720px;
            animation: fichaFadeIn 0.85s cubic-bezier(0.22, 1, 0.36, 1) 0.1s both;
        }}
        .ficha-hero .ficha-title {{
            font-family: 'Montserrat', sans-serif;
            font-size: clamp(1.25rem, 3.2vw, 1.65rem);
            font-weight: 900;
            color: {COR_AZUL_ESC};
            margin: 0;
            line-height: 1.25;
            letter-spacing: -0.02em;
        }}
        .ficha-hero-bar-wrap {{
            width: 100%;
            margin: clamp(0.85rem, 2.4vw, 1.2rem) 0;
        }}
        .ficha-hero-bar {{
            height: 4px;
            width: 100%;
            border-radius: 999px;
            background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC});
            background-size: 200% 100%;
            animation: fichaShimmer 4s ease-in-out infinite alternate;
        }}
        .bloco-pessoa {{
            background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.9) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 14px;
            padding: 0.85rem 1rem;
            margin-bottom: 0.9rem;
            box-shadow: 0 2px 8px rgba({RGB_AZUL_CSS}, 0.06);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        .bloco-pessoa:hover {{
            transform: translateY(-2px);
            box-shadow: 0 10px 20px -5px rgba({RGB_AZUL_CSS}, 0.12);
        }}
        .bloco-pessoa .nome {{
            font-family: 'Montserrat', sans-serif;
            font-weight: 800;
            font-size: 1.05rem;
            color: {COR_AZUL_ESC} !important;
            margin-bottom: 0.45rem;
            padding-bottom: 0.35rem;
            border-bottom: 1px solid {COR_BORDA};
        }}
        .footer {{
            text-align: center;
            padding: 1rem 0 0.25rem 0;
            color: {COR_TEXTO_PRETO};
            font-size: 0.82rem;
            opacity: 0.85;
        }}
        div[data-baseweb="input"] {{
            border-radius: 10px !important;
            border: 1px solid #e2e8f0 !important;
            background-color: {COR_INPUT_BG} !important;
        }}
        div[data-baseweb="input"]:focus-within {{
            border-color: rgba({RGB_AZUL_CSS}, 0.35) !important;
            box-shadow: 0 0 0 3px rgba({RGB_AZUL_CSS}, 0.08) !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def aplicar_estilo_basico() -> None:
    """Compatibilidade: encaminha para o design completo."""
    aplicar_estilo()


def _norm_txt_col(s: Any) -> str:
    t = str(s or "").strip().lower()
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def achar_coluna(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    cols_norm = {_norm_txt_col(c): c for c in cols}
    for a in aliases:
        al = str(a).strip().lower()
        for c in cols:
            if al == str(c).strip().lower():
                return c
        an = _norm_txt_col(a)
        if an in cols_norm:
            return cols_norm[an]
    for a in aliases:
        al = str(a).strip().lower()
        an = _norm_txt_col(a)
        for c in cols:
            cl = str(c).strip().lower()
            cn = _norm_txt_col(c)
            if al and (al in cl or an in cn):
                return c
    return None


def achar_coluna_aprovacao_safi(df: pd.DataFrame) -> Optional[str]:
    col = achar_coluna(df, COLUNAS_PASTAS_APROV_ALIASES)
    if col:
        return col
    if df is None or df.empty:
        return None
    for c in df.columns:
        cn = _norm_txt_col(c)
        if "safi" in cn and ("aprov" in cn or "data" in cn):
            return c
    return None


def achar_coluna_primeiro_envio_analise(df: pd.DataFrame) -> Optional[str]:
    col = achar_coluna(df, COLUNAS_PASTAS_ALIASES)
    if col and "primeiro" in _norm_txt_col(col) and "envio" in _norm_txt_col(col):
        return col
    if df is None or df.empty:
        return col
    for c in df.columns:
        cn = _norm_txt_col(c)
        if "primeiro" in cn and "envio" in cn:
            return c
    return col


def limpar_nome(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s.lower() in ("", "nan", "none", "null", "-", "n/a", "na", "não informado", "nao informado"):
        return ""
    return s


def parse_data_serie(serie: pd.Series) -> pd.Series:
    """
    Converte datas Salesforce/relatórios para datetime64[ns] naive.
    Não usa dayfirst em ISO SF — isso troca 2025-07-01 por 2025-01-07.
    """
    if serie is None:
        return pd.Series(dtype="datetime64[ns]")
    raw = serie
    as_str = raw.map(
        lambda x: ""
        if x is None or (isinstance(x, float) and pd.isna(x))
        else str(x).strip()
    )
    out = pd.Series(pd.NaT, index=raw.index, dtype="datetime64[ns]")
    mask_iso = as_str.str.match(r"^\d{4}-\d{2}-\d{2}", na=False)
    if mask_iso.any():
        has_time = mask_iso & as_str.str.contains("T", na=False)
        date_only = mask_iso & ~has_time
        if has_time.any():
            ts = pd.to_datetime(raw.loc[has_time], errors="coerce", utc=True)
            out.loc[has_time] = ts.dt.tz_convert("America/Sao_Paulo").dt.tz_localize(None)
        if date_only.any():
            out.loc[date_only] = pd.to_datetime(
                as_str.loc[date_only], format="%Y-%m-%d", errors="coerce"
            )
    vazios = {"", "nan", "none", "nat", "null", "na", "n/a", "-"}
    mask_rest = out.isna() & ~as_str.str.lower().isin(vazios)
    if mask_rest.any():
        out.loc[mask_rest] = pd.to_datetime(
            raw.loc[mask_rest], dayfirst=True, errors="coerce"
        )
    if out.isna().all():
        nums = pd.to_numeric(raw, errors="coerce")
        if nums.notna().any():
            med = float(nums.dropna().median())
            if med > 1e12:
                out = pd.to_datetime(nums, unit="ms", errors="coerce")
            elif med > 1e9:
                out = pd.to_datetime(nums, unit="s", errors="coerce")
            else:
                out = pd.to_datetime(
                    nums, unit="D", origin="1899-12-30", errors="coerce"
                )
    return out


def filtrar_vendas_comerciais(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    col = achar_coluna(df, ALIASES_VENDA_COMERCIAL)
    if not col:
        return df
    mask = (
        (pd.to_numeric(df[col], errors="coerce") == 1)
        | (df[col].astype(str).str.strip().str.upper().isin(["SIM", "TRUE", "1", "1.0"]))
    )
    return df.loc[mask].copy()


def deduplicar_por_chave_mais_recente(
    df: pd.DataFrame,
    aliases_chave: List[str],
    aliases_data: List[str],
) -> pd.DataFrame:
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    col_chave = achar_coluna(df, aliases_chave)
    col_data = achar_coluna(df, aliases_data) if isinstance(aliases_data, list) else None
    if isinstance(aliases_data, list) and len(aliases_data) == 1 and aliases_data[0] in (df.columns if df is not None else []):
        col_data = aliases_data[0]
    if not col_chave or not col_data:
        return df
    out = df.copy()
    out["_dedup_dt"] = parse_data_serie(out[col_data])
    out["_dedup_key"] = out[col_chave].astype(str).str.strip()
    mask_ok = out["_dedup_key"].ne("") & out["_dedup_key"].str.lower().ne("nan")
    ok = out.loc[mask_ok].sort_values("_dedup_dt", ascending=False, na_position="last")
    ok = ok.drop_duplicates(subset=["_dedup_key"], keep="first")
    resto = out.loc[~mask_ok]
    out = pd.concat([ok, resto], ignore_index=True)
    return out.drop(columns=["_dedup_dt", "_dedup_key"], errors="ignore")


def _aplicar_secrets_salesforce() -> None:
    try:
        if hasattr(st, "secrets") and "salesforce" in st.secrets:
            sec = st.secrets["salesforce"]
            for k_src, k_env in (
                ("USER", "SALESFORCE_USER"),
                ("USERNAME", "SALESFORCE_USER"),
                ("PASSWORD", "SALESFORCE_PASSWORD"),
                ("TOKEN", "SALESFORCE_TOKEN"),
                ("SECURITY_TOKEN", "SALESFORCE_TOKEN"),
                ("DOMAIN", "SALESFORCE_DOMAIN"),
            ):
                if k_src in sec and sec[k_src]:
                    os.environ[k_env] = str(sec[k_src]).strip()
    except Exception:
        pass


def conectar_salesforce_app() -> Tuple[Any, Optional[str]]:
    _aplicar_secrets_salesforce()
    try:
        from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
    except ImportError:
        return None, "Pacote simple-salesforce não instalado."

    username = (os.environ.get("SALESFORCE_USER") or "").strip()
    password = (os.environ.get("SALESFORCE_PASSWORD") or "").strip()
    token = (os.environ.get("SALESFORCE_TOKEN") or "").strip()
    domain = (os.environ.get("SALESFORCE_DOMAIN") or "login").strip() or "login"
    if not username or not password:
        return None, "Credenciais Salesforce ausentes ([salesforce] USER/PASSWORD/TOKEN)."
    try:
        kwargs: Dict[str, Any] = {"username": username, "password": password, "domain": domain}
        if token:
            kwargs["security_token"] = token
        return Salesforce(**kwargs), None
    except SalesforceAuthenticationFailed as e:
        return None, f"Autenticação Salesforce recusada: {e}"
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


# Analytics API: teto duro de detalhe por chamada. CSV Details: ~100k.
SF_ANALYTICS_ROW_CAP = 2000
SF_CSV_SOFT_CAP = 95_000


def _sf_session_bits(sf):
    """Retorna (base_url, session_id, headers) para export/API."""
    base = (getattr(sf, "sf_instance", None) or "").rstrip("/")
    if base and not base.startswith("http"):
        base = f"https://{base}"
    sid = str(getattr(sf, "session_id", "") or "")
    headers = dict(getattr(sf, "headers", {}) or {})
    if sid and "Authorization" not in headers:
        headers["Authorization"] = f"Bearer {sid}"
    return base, sid, headers


def _relatorio_sf_via_csv(sf, report_id: str):
    """
    Export CSV do relatório (Details). Contorna o teto de 2k da Analytics API,
    mas o Salesforce ainda pode cortar ~100k linhas em exports grandes.
    """
    import requests
    from io import StringIO

    rid = (report_id or "").strip()
    if not rid:
        return pd.DataFrame()

    base, sid, headers = _sf_session_bits(sf)
    if not base or not sid:
        raise ValueError("Sessão Salesforce sem instance/session_id.")

    urls = [
        f"{base}/{rid}?isdtp=p1&export=1&enc=UTF-8&xf=csv",
        f"{base}/{rid}?isdtp=p1&export=1&enc=UTF-8&xf=csv&detailsOnly=1",
        f"{base}/servlet/PrintableViewDownloadServlet?isdtp=p1&reportid={rid}",
    ]
    erros = []
    for url in urls:
        try:
            resp = requests.get(
                url,
                headers=headers,
                cookies={"sid": sid},
                timeout=900,
                allow_redirects=True,
            )
            resp.raise_for_status()
            text = resp.content.decode("utf-8-sig", errors="replace")
            sample = text.lstrip()[:300].lower()
            if (
                sample.startswith("<!doctype")
                or sample.startswith("<html")
                or "<table" in sample[:800]
                or ("login" in sample[:400] and "password" in sample[:800])
            ):
                erros.append(f"HTML em {url.split('?')[0][-40:]}")
                continue
            df = pd.read_csv(StringIO(text), low_memory=False)
            if df is None or df.empty:
                erros.append("CSV vazio")
                continue
            return df
        except Exception as e:
            erros.append(f"{type(e).__name__}: {e}")
    raise ValueError("Export CSV falhou: " + " | ".join(erros[:4]))


def _analytics_raw_to_df(raw):
    meta = (raw.get("reportMetadata") or {})
    cols_meta = meta.get("detailColumns") or []
    ext = ((raw.get("reportExtendedMetadata") or {}).get("detailColumnInfo") or {})
    headers = []
    for c in cols_meta:
        info = ext.get(c) or {}
        headers.append(str(info.get("label") or c))

    def _cell_scalar(cell):
        if not isinstance(cell, dict):
            return cell
        v = cell.get("value")
        if isinstance(v, (dict, list, tuple)):
            v = cell.get("label")
        elif v is None:
            v = cell.get("label")
        return v

    rows_out = []
    fact = (raw.get("factMap") or {}).get("T!T") or {}
    for row in fact.get("rows") or []:
        cells = row.get("dataCells") or []
        vals = [_cell_scalar(cell) for cell in cells]
        if len(vals) < len(headers):
            vals = vals + [None] * (len(headers) - len(vals))
        rows_out.append(vals[: len(headers)])
    if not headers:
        return pd.DataFrame()
    return pd.DataFrame(rows_out, columns=headers)


def _analytics_run(sf, report_id: str, report_metadata=None, tentativas: int = 8):
    """Executa relatório; em rate-limit (500/hora) espera e tenta de novo."""
    import time as _time

    rid = (report_id or "").strip()
    ultimo = None
    for i in range(max(1, int(tentativas))):
        try:
            if report_metadata is None:
                return sf.restful(f"analytics/reports/{rid}", params={"includeDetails": "true"})
            return sf.restful(
                f"analytics/reports/{rid}",
                method="POST",
                json={"reportMetadata": report_metadata},
            )
        except Exception as e:
            ultimo = e
            msg = str(e).lower()
            rate = (
                ("500" in msg and "relat" in msg)
                or ("forbidden" in msg and ("60 minuto" in msg or "60 minute" in msg))
                or ("não é possível executar mais de 500" in msg)
                or ("nao e possivel executar mais de 500" in msg)
            )
            if not rate or i >= tentativas - 1:
                raise
            espera = min(90, 15 * (i + 1))
            _time.sleep(espera)
    raise ultimo  # pragma: no cover



def _analytics_pick_date_column(meta, ext):
    sdf = meta.get("standardDateFilter") or {}
    col = sdf.get("column")
    if col:
        return str(col)
    for c in meta.get("detailColumns") or []:
        info = ext.get(c) or {}
        data_type = str(info.get("dataType") or "").lower()
        label = str(info.get("label") or c).lower()
        api = str(c).lower()
        if data_type in ("date", "datetime") or "date" in api or "data" in label:
            return str(c)
    return None


def _analytics_pick_id_column(meta, ext):
    """Escolhe coluna estável p/ keyset — evita falso positivo (ex.: 'credito' contém 'id')."""
    cols = list(meta.get("detailColumns") or [])
    scored = []
    for c in cols:
        info = ext.get(c) or {}
        api = str(c)
        label = str(info.get("label") or "")
        api_l = api.lower()
        lab_l = label.lower()
        score = 0
        if api_l == "id" or api_l.endswith(".id") or api_l.endswith("_id") or api_l.endswith("id__c"):
            score += 100
        if re.search(r"(^|[._])id($|[._])", api_l):
            score += 80
        if "identificador" in api_l or "identificador" in lab_l:
            score += 90
        if "codigo_do_agendamento" in api_l or "código do agendamento" in lab_l or "codigo do agendamento" in lab_l:
            score += 95
        if "código" in lab_l or "codigo" in lab_l:
            score += 50
        if "nome da avalia" in lab_l or "nome_da_avalia" in api_l.replace("__", "_").lower():
            score += 70
        if any(x in lab_l for x in ("status", "contrato", "valor", "telefone", "email")):
            score -= 60
        if score > 0:
            scored.append((score, api))
    if scored:
        scored.sort(key=lambda t: (t[0], t[1]), reverse=True)
        return scored[0][1]
    return cols[0] if cols else None


def _analytics_apply_date_filter(meta, date_col: str, ini, fim):
    m = copy.deepcopy(meta)
    m["standardDateFilter"] = {
        "column": date_col,
        "durationValue": "CUSTOM",
        "startDate": ini.isoformat(),
        "endDate": fim.isoformat(),
    }
    return m


def _analytics_keyset_pages(sf, report_id: str, base_meta, date_col: str, id_col: str, ini, fim):
    """Pagina um intervalo (ex.: 1 dia) com filtro greaterThan no Id ordenado."""
    partes = []
    last_val = None
    vistos = set()
    for _ in range(500):
        meta = _analytics_apply_date_filter(base_meta, date_col, ini, fim)
        meta["sortBy"] = [{"sortColumn": id_col, "sortOrder": "Asc"}]
        filtros = [
            f for f in (meta.get("reportFilters") or [])
            if not (f.get("column") == id_col and f.get("operator") == "greaterThan")
        ]
        if last_val is not None:
            filtros.append({"column": id_col, "operator": "greaterThan", "value": str(last_val)})
        meta["reportFilters"] = filtros
        orig_bool = str(base_meta.get("reportBooleanFilter") or "").strip()
        if last_val is not None:
            idx = len(filtros)
            meta["reportBooleanFilter"] = f"({orig_bool}) AND {idx}" if orig_bool else str(idx)
        else:
            meta["reportBooleanFilter"] = orig_bool or None

        raw = _analytics_run(sf, report_id, meta)
        df = _analytics_raw_to_df(raw)
        if df.empty:
            break
        ext = ((raw.get("reportExtendedMetadata") or {}).get("detailColumnInfo") or {})
        id_label = str((ext.get(id_col) or {}).get("label") or id_col)
        col_id = id_label if id_label in df.columns else (id_col if id_col in df.columns else df.columns[0])
        novos = df[~df[col_id].astype(str).isin(vistos)] if col_id in df.columns else df
        if novos.empty:
            break
        partes.append(novos)
        vals = novos[col_id].astype(str).tolist()
        vistos.update(vals)
        last_val = vals[-1]
        all_data = bool(raw.get("allData", False))
        if all_data or len(df) < SF_ANALYTICS_ROW_CAP:
            break
    if not partes:
        return pd.DataFrame()
    return pd.concat(partes, ignore_index=True)


def _analytics_fetch_range(sf, report_id: str, base_meta, date_col: str, id_col, ini, fim, profundidade: int = 0):
    """Baixa um intervalo; se bater 2k, divide a janela (ou usa keyset no dia)."""
    if ini > fim:
        return pd.DataFrame()
    meta = _analytics_apply_date_filter(base_meta, date_col, ini, fim)
    meta.pop("sortBy", None)

    raw = _analytics_run(sf, report_id, meta)
    df = _analytics_raw_to_df(raw)
    n = len(df)
    all_data = bool(raw.get("allData", True)) if n < SF_ANALYTICS_ROW_CAP else bool(raw.get("allData", False))

    if n == 0:
        return df
    if n < SF_ANALYTICS_ROW_CAP or all_data:
        return df

    if ini == fim:
        if id_col:
            return _analytics_keyset_pages(sf, report_id, base_meta, date_col, id_col, ini, fim)
        return df

    if profundidade > 40:
        if id_col:
            return _analytics_keyset_pages(sf, report_id, base_meta, date_col, id_col, ini, fim)
        return df

    mid = ini + timedelta(days=(fim - ini).days // 2)
    if mid >= fim:
        mid = fim - timedelta(days=1)
    if mid < ini:
        mid = ini

    esq = _analytics_fetch_range(sf, report_id, base_meta, date_col, id_col, ini, mid, profundidade + 1)
    dir_ = _analytics_fetch_range(
        sf, report_id, base_meta, date_col, id_col, mid + timedelta(days=1), fim, profundidade + 1
    )
    if esq.empty:
        return dir_
    if dir_.empty:
        return esq
    return pd.concat([esq, dir_], ignore_index=True)


def _relatorio_sf_via_analytics(sf, report_id: str):
    """Uma chamada Analytics (máx. ~2k) — só para metadados / fallback mínimo."""
    raw = _analytics_run(sf, report_id)
    return _analytics_raw_to_df(raw)


def _relatorio_sf_via_analytics_chunked(sf, report_id: str, anos_historico: int = 4, chunk_dias: int = 14):
    """
    Contorna o limite de 2.000 linhas da Analytics API:
    fatia por data (standardDateFilter) e, se um dia ainda passar de 2k, pagina por Id.
    """
    rid = (report_id or "").strip()
    raw0 = _analytics_run(sf, rid)
    base_meta = copy.deepcopy(raw0.get("reportMetadata") or {})
    ext = ((raw0.get("reportExtendedMetadata") or {}).get("detailColumnInfo") or {})
    date_col = _analytics_pick_date_column(base_meta, ext)
    if not date_col:
        return _analytics_raw_to_df(raw0)

    id_col = _analytics_pick_id_column(base_meta, ext)
    hoje = date.today()
    ini_global = date(hoje.year - int(anos_historico), 1, 1)
    sdf = base_meta.get("standardDateFilter") or {}
    try:
        sd = sdf.get("startDate")
        if sd:
            ini_meta = date.fromisoformat(str(sd)[:10])
            if ini_meta < ini_global:
                ini_global = ini_meta
    except Exception:
        pass

    partes = []
    cursor = ini_global
    while cursor <= hoje:
        fim_chunk = min(cursor + timedelta(days=chunk_dias - 1), hoje)
        parte = _analytics_fetch_range(sf, rid, base_meta, date_col, id_col, cursor, fim_chunk)
        if not parte.empty:
            partes.append(parte)
        cursor = fim_chunk + timedelta(days=1)

    if not partes:
        return pd.DataFrame()
    out = pd.concat(partes, ignore_index=True)
    try:
        return out.drop_duplicates().reset_index(drop=True)
    except TypeError:
        return out.astype(str).drop_duplicates().reset_index(drop=True)


def _sf_rel_name(val):
    if isinstance(val, dict):
        return val.get("Name") or val.get("name")
    return None


def _sf_janela_12_meses_fechados(ref: Optional[date] = None) -> Tuple[date, date]:
    """12 meses-calendário anteriores, excluindo o mês atual."""
    ref = ref or date.today()
    inicio = date(ref.year - 1, ref.month, 1)
    fim = date(ref.year, ref.month, 1) - timedelta(days=1)
    return inicio, fim


def _sf_soql_desde() -> str:
    """Início dos 12 meses fechados; mantém mês atual só para o realizado."""
    ini, _ = _sf_janela_12_meses_fechados()
    return f"{ini.isoformat()}T00:00:00Z"


def _sf_soql_agendamentos(sf) -> pd.DataFrame:
    """
    Extrai agendamentos/visitas via SOQL em Event (queryMore).
    Contorna Analytics 2k e cota de 500 relatórios/hora.
    """
    desde = _sf_soql_desde()
    soql = (
        "SELECT Id, Codigo_do_agendamento__c, CreatedDate, Data_da_Visita__c, "
        "Gerente_Regional__c, Gerente_de_Vendas__c, Corretor__r.Name, Regional__c, "
        "Unidade_de_negocio__c, Empreendimento_de_interesse__c "
        "FROM Event "
        "WHERE Unidade_de_negocio__c = 'Direcional' "
        "AND Regional__c = 'RJ' "
        "AND PDV__r.Regional_Comercial__c = 'RJ' "
        "AND PDV__r.UnidadeDeNegocio__c = 'Direcional' "
        "AND Empreendimento_de_interesse__c != null "
        "AND Account.Regional_Comercial__c = 'RJ' "
        f"AND CreatedDate >= {desde}"
    )
    res = sf.query_all(soql)
    rows = []
    for r in res.get("records") or []:
        rows.append(
            {
                "Código do agendamento": r.get("Codigo_do_agendamento__c"),
                "Data de criação": r.get("CreatedDate"),
                "Data da visita": r.get("Data_da_Visita__c"),
                "Gerente Regional": r.get("Gerente_Regional__c"),
                "Gerente de Vendas": r.get("Gerente_de_Vendas__c"),
                "Corretor: Nome completo": _sf_rel_name(r.get("Corretor__r")),
                "Regional": r.get("Regional__c"),
            }
        )
    return pd.DataFrame(rows)


def _sf_soql_pastas(sf) -> pd.DataFrame:
    """Extrai avaliações de crédito (pastas) via SOQL."""
    desde = _sf_soql_desde()
    soql = (
        "SELECT Id, Name, CreatedDate, dataPrimeiroEnvioAnalise__c, dataAprovacaoSAFI__c, "
        "Gerente_Regional__c, Gerente_Vendas__c, Corretor__r.Name, "
        "Oportunidade__r.Gerente_regional__c "
        "FROM Avaliacao_credito__c "
        "WHERE Empreendimento__r.Regional__c = 'RJ' "
        "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND CreatedDate >= {desde}"
    )
    res = sf.query_all(soql)
    rows = []
    for r in res.get("records") or []:
        opp = r.get("Oportunidade__r") if isinstance(r.get("Oportunidade__r"), dict) else {}
        rows.append(
            {
                "Nome da Avaliação de crédito": r.get("Name"),
                "Data de criação": r.get("CreatedDate"),
                "Data Primeiro Envio Análise": r.get("dataPrimeiroEnvioAnalise__c"),
                "Data Aprovação SAFI": r.get("dataAprovacaoSAFI__c"),
                "Gerente Regional": r.get("Gerente_Regional__c"),
                "Gerente Vendas": r.get("Gerente_Vendas__c"),
                "Corretor": _sf_rel_name(r.get("Corretor__r")),
                "Avaliação de crédito : Oportunidade : Gerente regional": (
                    opp.get("Gerente_regional__c") if opp else None
                ),
            }
        )
    return pd.DataFrame(rows)


def _sf_soql_dicionario_regional(sf) -> pd.DataFrame:
    """Replica o relatório gerente→regional no trimestre fiscal atual."""
    soql = (
        "SELECT Owner.Name, Gerente_regional__c "
        "FROM Opportunity "
        "WHERE CloseDate = THIS_FISCAL_QUARTER "
        "AND Gerente_regional__c != null"
    )
    res = sf.query_all(soql)
    rows = []
    for r in res.get("records") or []:
        rows.append(
            {
                "Proprietário da oportunidade": _sf_rel_name(r.get("Owner")),
                "Gerente regional": r.get("Gerente_regional__c"),
            }
        )
    return pd.DataFrame(rows).drop_duplicates()


def _sf_soql_por_relatorio(sf, report_id: str, rotulo: str):
    """Tenta SOQL para relatórios grandes. Retorna (df, origem) ou (None, None)."""
    rid = (report_id or "").strip()
    rotulo_l = (rotulo or "").lower()
    try:
        if rid == SF_REPORT_AGENDAMENTOS_ID or "agendamento" in rotulo_l:
            df = _sf_soql_agendamentos(sf)
            if df is not None and not df.empty:
                origem = (
                    f"Salesforce SOQL · Event · {rotulo} · "
                    f"{len(df):,} linhas".replace(",", ".")
                )
                return df, origem
        if rid == SF_REPORT_PASTAS_ID or "pasta" in rotulo_l:
            df = _sf_soql_pastas(sf)
            if df is not None and not df.empty:
                origem = (
                    f"Salesforce SOQL · Avaliacao_credito__c · {rotulo} · "
                    f"{len(df):,} linhas".replace(",", ".")
                )
                return df, origem
        if rid == SF_REPORT_DIC_REGIONAL_ID or "dicionário" in rotulo_l:
            df = _sf_soql_dicionario_regional(sf)
            if df is not None and not df.empty:
                return df, (
                    "Salesforce SOQL · Opportunity gerente→regional · "
                    f"{len(df):,} linhas".replace(",", ".")
                )
    except Exception:
        return None, None
    return None, None


@st.cache_data(ttl=3600, show_spinner="Baixando dados Salesforce (SOQL)…")
def carregar_relatorio_salesforce(report_id: str, rotulo: str = "relatório"):
    """
    Baixa dados Salesforce sem teto de 2k:
    1) SOQL query_all (Event / Avaliacao_credito) — preferencial p/ volumes grandes
    2) CSV do relatório
    3) Analytics fatiada (fallback; cota 500/h)
    """
    sf, err = conectar_salesforce_app()
    if sf is None:
        raise RuntimeError(err or "Falha ao conectar no Salesforce.")

    rid = (report_id or "").strip()
    if not rid:
        raise RuntimeError(f"Report ID vazio ({rotulo}).")

    tentativas = []

    # 1) SOQL direto (melhor caminho para ~180k)
    df_soql, origem_soql = _sf_soql_por_relatorio(sf, rid, rotulo)
    if df_soql is not None and not df_soql.empty:
        df = normalizar_colunas(df_soql)
        return df, origem_soql or f"Salesforce SOQL · {rotulo} · {len(df)} linhas"

    # 2) CSV
    df_csv = None
    try:
        df_csv = _relatorio_sf_via_csv(sf, rid)
    except Exception as e_csv:
        tentativas.append(f"CSV: {e_csv}")

    rotulo_l = (rotulo or "").lower()
    precisa_chunk = False
    if df_csv is None or df_csv.empty:
        precisa_chunk = True
    else:
        n_csv = len(df_csv)
        if SF_ANALYTICS_ROW_CAP - 5 <= n_csv <= SF_ANALYTICS_ROW_CAP + 50:
            precisa_chunk = True
        if n_csv >= SF_CSV_SOFT_CAP:
            precisa_chunk = True
        if "agendamento" in rotulo_l and n_csv < 120_000:
            precisa_chunk = True
        if "pasta" in rotulo_l and n_csv <= SF_ANALYTICS_ROW_CAP + 50:
            precisa_chunk = True

    df = df_csv
    origem = f"Salesforce CSV · {rotulo} · {rid}" if df_csv is not None and not df_csv.empty else ""

    # 3) Analytics fatiada
    if precisa_chunk:
        try:
            df_chunk = _relatorio_sf_via_analytics_chunked(sf, rid)
            if df_chunk is not None and not df_chunk.empty:
                if df is None or df.empty or len(df_chunk) > len(df):
                    df = df_chunk
                    origem = (
                        f"Salesforce Analytics fatiada · {rotulo} · {rid} · "
                        f"{len(df_chunk):,} linhas".replace(",", ".")
                    )
                else:
                    origem = (
                        f"Salesforce CSV · {rotulo} · {rid} · "
                        f"{len(df):,} linhas (chunk≤CSV)".replace(",", ".")
                    )
        except Exception as e_an:
            tentativas.append(f"Analytics fatiada: {e_an}")
            if df is None or df.empty:
                try:
                    df = _relatorio_sf_via_analytics(sf, rid)
                    origem = f"Salesforce Analytics · {rotulo} · {rid} (truncado ~2k)"
                except Exception as e_an2:
                    tentativas.append(f"Analytics: {e_an2}")
                    raise RuntimeError(
                        f"Não foi possível baixar o {rotulo} ({rid}). "
                        + " | ".join(tentativas)
                    ) from e_an2

    if df is None or df.empty:
        raise RuntimeError(
            f"Relatório {rotulo} ({rid}) retornou vazio. " + " | ".join(tentativas)
        )

    df = normalizar_colunas(df)
    if rid == SF_REPORT_VENDAS_ID:
        col_data = achar_coluna(df, ALIASES_CONTRATO_GERADO)
        if col_data:
            inicio_hist, _ = _sf_janela_12_meses_fechados()
            dt = parse_data_serie(df[col_data])
            df = df.loc[dt.notna() & (dt.dt.date >= inicio_hist)].copy()
    if not origem:
        origem = f"Salesforce · {rotulo} · {rid} · {len(df):,} linhas".replace(",", ".")
    elif "linhas" not in origem.lower():
        origem = f"{origem} · {len(df):,} linhas".replace(",", ".")
    return df, origem

def segunda_da_semana(d: date) -> date:
    return d - timedelta(days=d.weekday())


def domingo_da_semana(d: date) -> date:
    return segunda_da_semana(d) + timedelta(days=6)


def semana_iso_atual(hoje: Optional[date] = None) -> Tuple[date, date]:
    hoje = hoje or date.today()
    ini = segunda_da_semana(hoje)
    return ini, ini + timedelta(days=6)


def n_dias_periodo(inicio: date, fim: date) -> int:
    if fim < inicio:
        return 0
    return (fim - inicio).days + 1


def _serie_hierarquia(df: pd.DataFrame, aliases: List[str], fallback: Optional[List[str]] = None) -> pd.Series:
    col = achar_coluna(df, aliases)
    s = df[col].map(limpar_nome) if col else pd.Series([""] * len(df), index=df.index)
    if fallback:
        col_fb = achar_coluna(df, fallback)
        if col_fb:
            fb = df[col_fb].map(limpar_nome)
            s = s.where(s.ne(""), fb)
    return s


def _eventos_de_df(
    df: pd.DataFrame,
    col_data: Optional[str],
    etapa: str,
    regional: pd.Series,
    gerente: pd.Series,
    corretor: pd.Series,
) -> pd.DataFrame:
    if not col_data or col_data not in df.columns or df.empty:
        return pd.DataFrame(columns=["data", "etapa", "regional", "gerente", "corretor"])
    dt = parse_data_serie(df[col_data])
    mask = dt.notna()
    if not mask.any():
        return pd.DataFrame(columns=["data", "etapa", "regional", "gerente", "corretor"])
    out = pd.DataFrame({
        "data": dt.loc[mask].dt.normalize().dt.date,
        "etapa": etapa,
        "regional": regional.loc[mask].values,
        "gerente": gerente.loc[mask].values,
        "corretor": corretor.loc[mask].values,
    })
    return out


def montar_mapa_gerente_regional(df_dic: pd.DataFrame) -> Dict[str, str]:
    """Dicionário Proprietário da oportunidade → Gerente regional."""
    col_ger = achar_coluna(df_dic, ALIASES_GERENTE_VENDAS)
    col_reg = achar_coluna(df_dic, ALIASES_REGIONAL_DIC)
    if not col_ger or not col_reg or df_dic.empty:
        return {}
    mapa: Dict[str, str] = {}
    for _, r in df_dic.iterrows():
        g = limpar_nome(r.get(col_ger))
        reg = limpar_nome(r.get(col_reg))
        if g and reg and g not in mapa:
            mapa[g] = reg
    return mapa


@st.cache_data(ttl=3600, show_spinner="Montando base de funil por pessoa…")
def carregar_eventos_funil_pessoas() -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Retorna DataFrame de eventos (data, etapa, regional, gerente, corretor)
    e dict de origens dos relatórios.
    """
    origens: Dict[str, str] = {}

    df_ag, origens["agendamentos"] = carregar_relatorio_salesforce(
        SF_REPORT_AGENDAMENTOS_ID, rotulo="agendamentos/visitas"
    )
    df_ag = deduplicar_por_chave_mais_recente(df_ag, ALIASES_CODIGO_AGENDAMENTO, ALIASES_DATA_CRIACAO)
    reg_ag = _serie_hierarquia(df_ag, ALIASES_REGIONAL_AG)
    ger_ag = _serie_hierarquia(df_ag, ALIASES_GERENTE_AG)
    cor_ag = _serie_hierarquia(df_ag, ALIASES_CORRETOR_AG)
    col_criacao = achar_coluna(df_ag, ALIASES_DATA_CRIACAO)
    col_visita = achar_coluna(df_ag, ALIASES_DATA_VISITA)
    ev_ag = _eventos_de_df(df_ag, col_criacao, "agendamentos", reg_ag, ger_ag, cor_ag)
    ev_vis = _eventos_de_df(df_ag, col_visita, "visitas", reg_ag, ger_ag, cor_ag)

    df_pas, origens["pastas"] = carregar_relatorio_salesforce(
        SF_REPORT_PASTAS_ID, rotulo="pastas"
    )
    col_envio = achar_coluna_primeiro_envio_analise(df_pas)
    col_safi = achar_coluna_aprovacao_safi(df_pas)
    df_pas_envio = (
        deduplicar_por_chave_mais_recente(df_pas, ALIASES_NOME_AVALIACAO_CREDITO, [col_envio])
        if col_envio else df_pas
    )
    df_pas_aprov = (
        deduplicar_por_chave_mais_recente(df_pas, ALIASES_NOME_AVALIACAO_CREDITO, [col_safi])
        if col_safi else df_pas
    )
    reg_pas = _serie_hierarquia(df_pas_envio, ALIASES_REGIONAL_PASTAS, ALIASES_REGIONAL_PASTAS_FALLBACK)
    ger_pas = _serie_hierarquia(df_pas_envio, ALIASES_GERENTE_PASTAS)
    cor_pas = _serie_hierarquia(df_pas_envio, ALIASES_CORRETOR_PASTAS)
    reg_aprov = _serie_hierarquia(df_pas_aprov, ALIASES_REGIONAL_PASTAS, ALIASES_REGIONAL_PASTAS_FALLBACK)
    ger_aprov = _serie_hierarquia(df_pas_aprov, ALIASES_GERENTE_PASTAS)
    cor_aprov = _serie_hierarquia(df_pas_aprov, ALIASES_CORRETOR_PASTAS)
    ev_pas = _eventos_de_df(df_pas_envio, col_envio, "pastas", reg_pas, ger_pas, cor_pas)
    ev_aprov = _eventos_de_df(df_pas_aprov, col_safi, "pastas_aprovadas", reg_aprov, ger_aprov, cor_aprov)

    df_ven, origens["vendas"] = carregar_relatorio_salesforce(
        SF_REPORT_VENDAS_ID, rotulo="vendas"
    )
    df_ven = filtrar_vendas_comerciais(df_ven)
    df_ven = deduplicar_por_chave_mais_recente(df_ven, ALIASES_ID_OPORTUNIDADE, ALIASES_CONTRATO_GERADO)
    ger_ven = _serie_hierarquia(df_ven, ALIASES_GERENTE_VENDAS)
    cor_ven = _serie_hierarquia(df_ven, ALIASES_CORRETOR_VENDAS)

    try:
        df_dic, origens["dicionario_regional"] = carregar_relatorio_salesforce(
            SF_REPORT_DIC_REGIONAL_ID, rotulo="dicionário gerente→regional"
        )
        mapa_reg = montar_mapa_gerente_regional(df_dic)
    except Exception as e:
        origens["dicionario_regional"] = f"indisponível ({e})"
        mapa_reg = {}

    reg_ven = ger_ven.map(lambda g: mapa_reg.get(g, ""))
    # se o relatório de vendas já trouxer regional, preenche vazios
    col_reg_ven = achar_coluna(df_ven, ALIASES_REGIONAL_DIC + ALIASES_REGIONAL_AG)
    if col_reg_ven:
        reg_direct = df_ven[col_reg_ven].map(limpar_nome)
        reg_ven = reg_ven.where(reg_ven.ne(""), reg_direct)

    col_contrato = achar_coluna(df_ven, ALIASES_CONTRATO_GERADO)
    ev_ven = _eventos_de_df(df_ven, col_contrato, "vendas", reg_ven, ger_ven, cor_ven)

    eventos = pd.concat([ev_ag, ev_vis, ev_pas, ev_aprov, ev_ven], ignore_index=True)
    if eventos.empty:
        return eventos, origens
    eventos["data"] = pd.to_datetime(eventos["data"], errors="coerce").dt.date
    eventos = eventos.dropna(subset=["data"]).reset_index(drop=True)
    return eventos, origens


def filtrar_periodo(eventos: pd.DataFrame, inicio: date, fim: date) -> pd.DataFrame:
    if eventos is None or eventos.empty:
        return pd.DataFrame(columns=eventos.columns if eventos is not None else [])
    mask = (eventos["data"] >= inicio) & (eventos["data"] <= fim)
    return eventos.loc[mask].copy()


def agregar_funil_por_dimensao(
    eventos: pd.DataFrame,
    dimensao: str,
) -> pd.DataFrame:
    """Uma linha por pessoa com contagens por etapa do funil."""
    if dimensao not in DIMENSOES:
        raise ValueError(f"Dimensão inválida: {dimensao}")
    if eventos is None or eventos.empty:
        cols = [dimensao] + list(FUNIL_ETAPAS)
        return pd.DataFrame(columns=cols)

    base = eventos.copy()
    base[dimensao] = base[dimensao].map(limpar_nome)
    base = base[base[dimensao].ne("")]
    if base.empty:
        cols = [dimensao] + list(FUNIL_ETAPAS)
        return pd.DataFrame(columns=cols)

    pivot = (
        base.groupby([dimensao, "etapa"], as_index=False)
        .size()
        .rename(columns={"size": "qtd"})
    )
    wide = pivot.pivot(index=dimensao, columns="etapa", values="qtd").fillna(0.0)
    for e in FUNIL_ETAPAS:
        if e not in wide.columns:
            wide[e] = 0.0
    wide = wide[list(FUNIL_ETAPAS)].reset_index()
    return wide


def escalar_media_para_periodo(
    totais_base: pd.DataFrame,
    dias_base: int,
    dias_alvo: int,
    dimensao: str,
) -> pd.DataFrame:
    """
    Média diária da base × quantidade de dias do período alvo.
    (total_base / dias_base) * dias_alvo
    """
    if totais_base is None or totais_base.empty or dias_base <= 0 or dias_alvo <= 0:
        cols = [dimensao] + list(FUNIL_ETAPAS)
        return pd.DataFrame(columns=cols)
    out = totais_base.copy()
    fator = float(dias_alvo) / float(dias_base)
    for e in FUNIL_ETAPAS:
        out[e] = out[e].astype(float) * fator
    return out


def fmt_num(v: float, casas: int = 1) -> str:
    v = float(v)
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v)):,}".replace(",", ".")
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v: Optional[float]) -> str:
    if v is None:
        return "—"
    return f"{float(v):.0f}%"

# =============================================================================
# Fim do módulo embutido
# =============================================================================

# Janelas históricas em meses-calendário fechados (mês atual excluído)
JANELAS_MEDIA: Tuple[Tuple[str, str, int], ...] = (
    ("12_meses", "Média 12 meses", 12),
    ("6_meses", "Média 6 meses", 6),
    ("3_meses", "Média 3 meses", 3),
    ("1_mes", "Média 1 mês", 1),
)

OPCOES_SEMANA = {
    "atual": "Semana atual",
    "passada": "Semana passada",
    "retrasada": "Semana retrasada",
}


def semana_por_offset(hoje: date, semanas_atras: int) -> Tuple[date, date]:
    """Semana ISO (seg→dom) deslocada: 0=atual, 1=passada, 2=retrasada."""
    ini_atual = segunda_da_semana(hoje)
    ini = ini_atual - timedelta(days=7 * int(semanas_atras))
    return ini, domingo_da_semana(ini)


def filtrar_hierarquia(
    eventos: pd.DataFrame,
    regionais: Optional[List[str]] = None,
    gerentes: Optional[List[str]] = None,
    corretores: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Filtros opcionais: lista vazia/None = não filtra aquela dimensão."""
    if eventos is None or eventos.empty:
        return eventos if eventos is not None else pd.DataFrame()
    out = eventos
    if regionais:
        regs = {limpar_nome(r) for r in regionais if limpar_nome(r)}
        out = out.loc[out["regional"].map(limpar_nome).isin(regs)]
    if gerentes:
        gers = {limpar_nome(g) for g in gerentes if limpar_nome(g)}
        out = out.loc[out["gerente"].map(limpar_nome).isin(gers)]
    if corretores:
        cors = {limpar_nome(c) for c in corretores if limpar_nome(c)}
        out = out.loc[out["corretor"].map(limpar_nome).isin(cors)]
    return out.copy()


def filtrar_regionais(eventos: pd.DataFrame, regionais: List[str]) -> pd.DataFrame:
    return filtrar_hierarquia(eventos, regionais=regionais)


def _inicio_mes_deslocado(inicio_mes: date, meses: int) -> date:
    total = inicio_mes.year * 12 + (inicio_mes.month - 1) + int(meses)
    return date(total // 12, total % 12 + 1, 1)


def _janela_meses_fechados(fim_base: date, meses: int) -> Tuple[date, date]:
    primeiro_mes_seguinte = date(
        (fim_base + timedelta(days=1)).year,
        (fim_base + timedelta(days=1)).month,
        1,
    )
    return _inicio_mes_deslocado(primeiro_mes_seguinte, -int(meses)), fim_base


def media_escalada_pessoa_etapa(
    datas: List[date],
    ini_base: date,
    fim_base: date,
    dias_alvo: int,
) -> float:
    """
    Média diária no período disponível da pessoa × dias_alvo.
    Início efetivo = max(início da janela, menor data do indicador).
    """
    if not datas or ini_base > fim_base or dias_alvo <= 0:
        return 0.0
    datas_ok = [d for d in datas if d is not None and d <= fim_base]
    if not datas_ok:
        return 0.0
    min_ind = min(datas_ok)
    ini_eff = max(ini_base, min_ind)
    if ini_eff > fim_base:
        return 0.0
    dias_disp = (fim_base - ini_eff).days + 1
    if dias_disp <= 0:
        return 0.0
    qtd = sum(1 for d in datas_ok if ini_eff <= d <= fim_base)
    return (float(qtd) / float(dias_disp)) * float(dias_alvo)


def medias_historicas_pessoa(
    eventos_pessoa: pd.DataFrame,
    fim_base: date,
    dias_alvo: int,
) -> Dict[str, Dict[str, float]]:
    """Para cada janela e etapa, média escalada ao período-alvo (ex.: 7 dias)."""
    por_etapa: Dict[str, List[date]] = {e: [] for e in FUNIL_ETAPAS}
    if eventos_pessoa is not None and not eventos_pessoa.empty:
        for _, r in eventos_pessoa.iterrows():
            e = str(r.get("etapa", ""))
            d = r.get("data")
            if e in por_etapa and d is not None:
                por_etapa[e].append(d if isinstance(d, date) else pd.to_datetime(d).date())

    out: Dict[str, Dict[str, float]] = {}
    for chave, _rotulo, meses in JANELAS_MEDIA:
        ini_base, _ = _janela_meses_fechados(fim_base, meses)
        out[chave] = {
            e: media_escalada_pessoa_etapa(
                por_etapa[e], ini_base, fim_base, dias_alvo
            )
            for e in FUNIL_ETAPAS
        }
    return out


def _montar_tabela_pessoa(
    realizado: Dict[str, float],
    medias: Dict[str, Dict[str, float]],
) -> pd.DataFrame:
    rows = []
    for chave, rotulo, _ in JANELAS_MEDIA:
        row = {"Linha": rotulo}
        fonte = medias.get(chave, {})
        for e in FUNIL_ETAPAS:
            row[FUNIL_LABELS[e]] = float(fonte.get(e, 0.0))
        rows.append(row)

    row_real = {"Linha": "Realizado da semana"}
    for e in FUNIL_ETAPAS:
        row_real[FUNIL_LABELS[e]] = float(realizado.get(e, 0.0))
    rows.append(row_real)

    for chave, rotulo, _ in JANELAS_MEDIA:
        row_pct = {"Linha": f"% vs {rotulo.replace('Média ', '')}"}
        fonte = medias.get(chave, {})
        for e in FUNIL_ETAPAS:
            m = float(fonte.get(e, 0.0))
            r = float(realizado.get(e, 0.0))
            row_pct[FUNIL_LABELS[e]] = (100.0 * r / m) if m > 1e-9 else None
        rows.append(row_pct)
    return pd.DataFrame(rows)


def _taxa_conversao(origem: float, destino: float) -> float:
    origem = float(origem or 0.0)
    destino = float(destino or 0.0)
    # Sem origem (ou ambos zero): 0%, nunca infinito / "—"
    if origem <= 0:
        return 0.0
    return 100.0 * destino / origem


def _montar_tabela_conversoes(
    realizado: Dict[str, float],
    medias: Dict[str, Dict[str, float]],
) -> pd.DataFrame:
    """Conversões etapa→etapa e diretas→venda nas cinco linhas de tempo."""
    fontes = [
        (rotulo, medias.get(chave, {}))
        for chave, rotulo, _meses in JANELAS_MEDIA
    ]
    fontes.append(("Realizado da semana", realizado))

    rows = []
    for rotulo, fonte in fontes:
        row: Dict[str, Any] = {"Linha": rotulo}
        row["Ag. → Vis."] = _taxa_conversao(
            fonte.get("agendamentos", 0.0), fonte.get("visitas", 0.0)
        )
        row["Vis. → Pastas"] = _taxa_conversao(
            fonte.get("visitas", 0.0), fonte.get("pastas", 0.0)
        )
        row["Pastas → Aprov."] = _taxa_conversao(
            fonte.get("pastas", 0.0), fonte.get("pastas_aprovadas", 0.0)
        )
        row["Aprov. → Vendas"] = _taxa_conversao(
            fonte.get("pastas_aprovadas", 0.0), fonte.get("vendas", 0.0)
        )
        row["Direta Ag. → Vendas"] = _taxa_conversao(
            fonte.get("agendamentos", 0.0), fonte.get("vendas", 0.0)
        )
        row["Direta Vis. → Vendas"] = _taxa_conversao(
            fonte.get("visitas", 0.0), fonte.get("vendas", 0.0)
        )
        row["Direta Pastas → Vendas"] = _taxa_conversao(
            fonte.get("pastas", 0.0), fonte.get("vendas", 0.0)
        )
        row["Direta Aprov. → Vendas"] = _taxa_conversao(
            fonte.get("pastas_aprovadas", 0.0), fonte.get("vendas", 0.0)
        )
        rows.append(row)
    return pd.DataFrame(rows)


# Farol realizado × média: verde ≥ 100%, amarelo 75–100%, vermelho < 75%
_CSS_VERDE = "background-color: #ecfdf5; color: #065f46; font-weight: 600;"
_CSS_AMARELO = "background-color: #fffbeb; color: #92400e; font-weight: 600;"
_CSS_VERMELHO = f"background-color: #fef2f2; color: {COR_VERMELHO}; font-weight: 600;"
_ROTULO_MEDIA_REF = JANELAS_MEDIA[0][1]  # "Média 12 meses"


def _css_farol_pct(pct: Optional[float]) -> str:
    if pct is None or (isinstance(pct, float) and pd.isna(pct)):
        return ""
    if float(pct) >= 100.0:
        return _CSS_VERDE
    if float(pct) >= 75.0:
        return _CSS_AMARELO
    return _CSS_VERMELHO


def _css_realizado_vs_media(realizado: Any, media: Any) -> str:
    if realizado is None or (isinstance(realizado, float) and pd.isna(realizado)):
        return ""
    r = float(realizado)
    if media is None or (isinstance(media, float) and pd.isna(media)) or float(media) <= 1e-9:
        # sem média de referência: acima de zero conta como ≥ média
        return _CSS_VERDE if r > 0 else ""
    return _css_farol_pct(100.0 * r / float(media))


def _estilo_conversoes(df: pd.DataFrame):
    df_fmt = df.copy()
    for c in df_fmt.columns:
        if c == "Linha":
            continue
        df_fmt[c] = df_fmt[c].map(
            lambda v: "—" if v is None or pd.isna(v) else fmt_pct(float(v))
        )

    ref = df.loc[df["Linha"] == _ROTULO_MEDIA_REF]
    medias_ref = ref.iloc[0] if not ref.empty else None

    def farol(row):
        if str(row.get("Linha", "")) != "Realizado da semana" or medias_ref is None:
            return [""] * len(df.columns)
        orig = df.loc[row.name]
        css = []
        for c in df.columns:
            if c == "Linha":
                css.append(
                    f"background-color: #f8fafc; font-weight: 600; color: {COR_TEXTO_PRETO};"
                )
                continue
            css.append(_css_realizado_vs_media(orig.get(c), medias_ref.get(c)))
        return css

    return (
        df_fmt.style.set_properties(
            **{"text-align": "center", "color": COR_TEXTO_PRETO}
        )
        .apply(farol, axis=1)
    )


def _estilo_tabela(df: pd.DataFrame):
    df_fmt = df.copy().astype(object)
    for i, row in df.iterrows():
        linha = str(row.get("Linha", ""))
        for c in df.columns:
            if c == "Linha":
                continue
            val = row[c]
            if val is None or (isinstance(val, float) and pd.isna(val)):
                df_fmt.at[i, c] = "—"
            elif linha.startswith("%"):
                df_fmt.at[i, c] = fmt_pct(float(val))
            else:
                df_fmt.at[i, c] = fmt_num(float(val), 1)

    ref = df.loc[df["Linha"] == _ROTULO_MEDIA_REF]
    medias_ref = ref.iloc[0] if not ref.empty else None

    def highlight(row):
        linha = str(row.get("Linha", ""))
        eh_pct = linha.startswith("%")
        eh_real = linha == "Realizado da semana"
        if not eh_pct and not eh_real:
            return [""] * len(df.columns)
        cores = []
        orig = df.loc[row.name]
        for c in df.columns:
            if c == "Linha":
                cores.append(
                    f"background-color: #f8fafc; font-weight: 600; color: {COR_TEXTO_PRETO};"
                )
                continue
            v = orig.get(c)
            if eh_pct:
                cores.append(_css_farol_pct(v))
            else:
                m = medias_ref.get(c) if medias_ref is not None else None
                cores.append(_css_realizado_vs_media(v, m))
        return cores

    return df_fmt.style.apply(highlight, axis=1)


def nomes_ativos_30d(
    eventos: pd.DataFrame,
    dimensao: str,
    ref: Optional[date] = None,
    min_agendamentos: int = 4,
    min_vendas: int = 1,
    janela_dias: int = 30,
) -> set:
    """
    Pessoas com agendamentos ≥ min_agendamentos OU vendas ≥ min_vendas
    nos últimos `janela_dias` dias (em relação a `ref`, padrão = hoje).
    """
    if eventos is None or eventos.empty or dimensao not in eventos.columns:
        return set()
    fim = ref or date.today()
    ini = fim - timedelta(days=janela_dias - 1)
    ev = filtrar_periodo(eventos, ini, fim)
    if ev.empty:
        return set()

    nomes: set = set()
    for etapa, minimo in (("agendamentos", min_agendamentos), ("vendas", min_vendas)):
        sub = ev[ev["etapa"] == etapa]
        if sub.empty:
            continue
        counts = (
            sub.assign(_n=sub[dimensao].map(limpar_nome))
            .loc[lambda d: d["_n"].ne("")]
            .groupby("_n")
            .size()
        )
        nomes.update(counts[counts >= minimo].index.tolist())
    return nomes


def _render_aba_dimensao(
    eventos: pd.DataFrame,
    dimensao: str,
    ini_semana: date,
    fim_semana: date,
    fim_base: date,
) -> None:
    dias_alvo = n_dias_periodo(ini_semana, fim_semana)

    ev_semana = filtrar_periodo(eventos, ini_semana, fim_semana)
    real = agregar_funil_por_dimensao(ev_semana, dimensao)

    if real.empty:
        return

    mask_pos = real[list(FUNIL_ETAPAS)].sum(axis=1) > 0
    real = real.loc[mask_pos].copy()
    if real.empty:
        return

    # Só quem agendou 4x+ ou vendeu 1x+ nos últimos 30 dias
    elegiveis = nomes_ativos_30d(eventos, dimensao)
    if not elegiveis:
        return
    real["_nome"] = real[dimensao].map(limpar_nome)
    real = real.loc[real["_nome"].isin(elegiveis)].drop(columns=["_nome"])
    if real.empty:
        return

    real = real.sort_values("vendas", ascending=False)
    ev_hist = eventos[eventos["data"] <= fim_base] if not eventos.empty else eventos

    for _, row in real.iterrows():
        nome = limpar_nome(row[dimensao])
        if not nome:
            continue
        realizado = {e: float(row.get(e, 0.0)) for e in FUNIL_ETAPAS}
        ev_p = ev_hist[ev_hist[dimensao].map(limpar_nome) == nome] if not ev_hist.empty else ev_hist
        medias = medias_historicas_pessoa(ev_p, fim_base, dias_alvo)
        df_t = _montar_tabela_pessoa(realizado, medias)
        df_conv = _montar_tabela_conversoes(realizado, medias)
        st.markdown(
            f'<div class="bloco-pessoa"><div class="nome">{nome}</div></div>',
            unsafe_allow_html=True,
        )
        st.dataframe(_estilo_tabela(df_t), use_container_width=True, hide_index=True)
        st.dataframe(
            _estilo_conversoes(df_conv), use_container_width=True, hide_index=True
        )


def main() -> None:
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(
        page_title="Funil · Média × Semana | Direcional",
        layout="wide",
        page_icon=str(fav) if fav else None,
        initial_sidebar_state="collapsed",
    )
    aplicar_estilo()
    _cabecalho_pagina("Funil por pessoa — média × semana")

    hoje = date.today()
    if "sem_escolha" not in st.session_state:
        st.session_state["sem_escolha"] = "atual"

    def _set_semana(chave: str):
        def _cb():
            st.session_state["sem_escolha"] = chave
        return _cb

    b1, b2, b3 = st.columns(3)
    with b1:
        st.button(
            "Semana atual",
            on_click=_set_semana("atual"),
            type="primary" if st.session_state["sem_escolha"] == "atual" else "secondary",
            use_container_width=True,
        )
    with b2:
        st.button(
            "Semana passada",
            on_click=_set_semana("passada"),
            type="primary" if st.session_state["sem_escolha"] == "passada" else "secondary",
            use_container_width=True,
        )
    with b3:
        st.button(
            "Semana retrasada",
            on_click=_set_semana("retrasada"),
            type="primary" if st.session_state["sem_escolha"] == "retrasada" else "secondary",
            use_container_width=True,
        )

    offset = {"atual": 0, "passada": 1, "retrasada": 2}[st.session_state["sem_escolha"]]
    ini_semana, fim_semana = semana_por_offset(hoje, offset)
    # Médias sempre terminam no último dia do mês anterior, independentemente
    # da semana escolhida. O mês atual entra somente no realizado semanal.
    _, fim_base = _sf_janela_12_meses_fechados(hoje)

    try:
        eventos, _origens = carregar_eventos_funil_pessoas()
    except Exception as e:
        st.error(f"Falha ao carregar bases Salesforce: {e}")
        return

    if eventos is None or eventos.empty:
        return

    regionais_disp = sorted({limpar_nome(r) for r in eventos["regional"].tolist() if limpar_nome(r)})
    regionais_sel = st.multiselect(
        "Regionais",
        options=regionais_disp,
        default=[],
        key="regionais_filtro",
        placeholder="Regionais",
    )

    ev_pos_reg = filtrar_hierarquia(eventos, regionais=regionais_sel or None)
    gerentes_disp = sorted({limpar_nome(g) for g in ev_pos_reg["gerente"].tolist() if limpar_nome(g)})
    gerentes_sel = st.multiselect(
        "Gerentes de vendas",
        options=gerentes_disp,
        default=[],
        key="gerentes_filtro",
        placeholder="Gerentes de vendas",
    )

    ev_pos_ger = filtrar_hierarquia(ev_pos_reg, gerentes=gerentes_sel or None)
    corretores_disp = sorted({limpar_nome(c) for c in ev_pos_ger["corretor"].tolist() if limpar_nome(c)})
    corretores_sel = st.multiselect(
        "Corretores",
        options=corretores_disp,
        default=[],
        key="corretores_filtro",
        placeholder="Corretores",
    )

    eventos = filtrar_hierarquia(
        eventos,
        regionais=regionais_sel or None,
        gerentes=gerentes_sel or None,
        corretores=corretores_sel or None,
    )

    tabs = st.tabs([DIM_LABELS[d] for d in DIMENSOES])
    for tab, dim in zip(tabs, DIMENSOES):
        with tab:
            _render_aba_dimensao(eventos, dim, ini_semana, fim_semana, fim_base)


if __name__ == "__main__":
    main()
