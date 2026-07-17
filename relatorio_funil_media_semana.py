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
    dt = pd.to_datetime(serie, dayfirst=True, errors="coerce")
    if dt.isna().all():
        nums = pd.to_numeric(serie, errors="coerce")
        if nums.notna().any():
            med = float(nums.dropna().median())
            if med > 1e9:
                dt = pd.to_datetime(nums, unit="s", errors="coerce")
            elif med > 1e12:
                dt = pd.to_datetime(nums, unit="ms", errors="coerce")
            else:
                dt = pd.to_datetime(nums, unit="D", origin="1899-12-30", errors="coerce")
    return dt


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


def _relatorio_sf_via_csv(sf: Any, report_id: str) -> pd.DataFrame:
    import requests
    from io import StringIO

    rid = (report_id or "").strip()
    base = (getattr(sf, "sf_instance", None) or "").rstrip("/")
    if not base.startswith("http"):
        base = f"https://{base}"
    url = f"{base}/{rid}?isdtp=p1&export=1&enc=UTF-8&xf=csv"
    resp = requests.get(
        url,
        headers=dict(getattr(sf, "headers", {}) or {}),
        cookies={"sid": getattr(sf, "session_id", "")},
        timeout=600,
    )
    resp.raise_for_status()
    text = resp.content.decode("utf-8", errors="replace")
    sample = text.lstrip()[:200].lower()
    if sample.startswith("<!doctype") or sample.startswith("<html"):
        raise ValueError("Export CSV retornou HTML (sessão/permissão).")
    return pd.read_csv(StringIO(text))


def _relatorio_sf_via_analytics(sf: Any, report_id: str) -> pd.DataFrame:
    rid = (report_id or "").strip()
    raw = sf.restful(f"analytics/reports/{rid}", params={"includeDetails": "true"})
    meta = (raw.get("reportMetadata") or {})
    cols_meta = meta.get("detailColumns") or []
    ext = ((raw.get("reportExtendedMetadata") or {}).get("detailColumnInfo") or {})
    headers: List[str] = []
    for c in cols_meta:
        info = ext.get(c) or {}
        headers.append(str(info.get("label") or c))
    fact = (raw.get("factMap") or {}).get("T!T") or {}
    rows = fact.get("rows") or []
    rows_out: List[List[Any]] = []
    for row in rows:
        cells = row.get("dataCells") or []
        vals = []
        for cell in cells:
            v = cell.get("label")
            if v is None:
                v = cell.get("value")
            vals.append(v)
        if len(vals) < len(headers):
            vals = vals + [None] * (len(headers) - len(vals))
        rows_out.append(vals[: len(headers)])
    if not headers:
        return pd.DataFrame()
    return pd.DataFrame(rows_out, columns=headers)


@st.cache_data(ttl=3600, show_spinner="Baixando relatório do Salesforce…")
def carregar_relatorio_salesforce(report_id: str, rotulo: str = "relatório") -> Tuple[pd.DataFrame, str]:
    sf, err = conectar_salesforce_app()
    if sf is None:
        raise RuntimeError(err or "Falha ao conectar no Salesforce.")
    rid = (report_id or "").strip()
    if not rid:
        raise RuntimeError(f"Report ID vazio ({rotulo}).")
    tentativas: List[str] = []
    try:
        df = _relatorio_sf_via_csv(sf, rid)
        origem = f"Salesforce CSV · {rotulo} · {rid}"
    except Exception as e_csv:
        tentativas.append(f"CSV: {e_csv}")
        try:
            df = _relatorio_sf_via_analytics(sf, rid)
            origem = f"Salesforce Analytics · {rotulo} · {rid}"
        except Exception as e_an:
            tentativas.append(f"Analytics: {e_an}")
            raise RuntimeError(
                f"Não foi possível baixar o {rotulo} ({rid}). " + " | ".join(tentativas)
            ) from e_an
    df = normalizar_colunas(df)
    if df.empty:
        raise RuntimeError(f"Relatório {rotulo} ({rid}) retornou vazio.")
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

# Janelas históricas (sempre terminam no dia anterior à semana escolhida)
JANELAS_MEDIA: Tuple[Tuple[str, str, int], ...] = (
    ("1_ano", "Média 1 ano", 365),
    ("6_meses", "Média 6 meses", 182),
    ("3_meses", "Média 3 meses", 91),
    ("1_mes", "Média 1 mês", 30),
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


def media_escalada_pessoa_etapa(
    datas: List[date],
    fim_base: date,
    dias_janela: int,
    dias_alvo: int,
) -> float:
    """
    Média diária no período disponível da pessoa × dias_alvo.
    Início efetivo = max(fim_base - janela + 1, menor data do indicador).
    """
    if not datas or dias_janela <= 0 or dias_alvo <= 0:
        return 0.0
    datas_ok = [d for d in datas if d is not None and d <= fim_base]
    if not datas_ok:
        return 0.0
    min_ind = min(datas_ok)
    ini_ideal = fim_base - timedelta(days=dias_janela - 1)
    ini_eff = max(ini_ideal, min_ind)
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
    for chave, _rotulo, dias_janela in JANELAS_MEDIA:
        out[chave] = {
            e: media_escalada_pessoa_etapa(por_etapa[e], fim_base, dias_janela, dias_alvo)
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


def _estilo_tabela(df: pd.DataFrame):
    df_fmt = df.copy()
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

    def highlight_pct(row):
        if not str(row.get("Linha", "")).startswith("%"):
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
            if v is None or (isinstance(v, float) and pd.isna(v)):
                cores.append("")
            elif float(v) >= 100:
                cores.append("background-color: #ecfdf5; color: #065f46; font-weight: 600;")
            elif float(v) >= 80:
                cores.append("background-color: #fffbeb; color: #92400e; font-weight: 600;")
            else:
                cores.append(f"background-color: #fef2f2; color: {COR_VERMELHO}; font-weight: 600;")
        return cores

    return df_fmt.style.apply(highlight_pct, axis=1)


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
        st.markdown(
            f'<div class="bloco-pessoa"><div class="nome">{nome}</div></div>',
            unsafe_allow_html=True,
        )
        st.dataframe(_estilo_tabela(df_t), use_container_width=True, hide_index=True)


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
    fim_base = ini_semana - timedelta(days=1)

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
