# -*- coding: utf-8 -*-
"""
Relatório 1 — Funil: média histórica × período selecionado
por Regional, Gerente de vendas e Corretor.

Hospedagem Streamlit Cloud (mesma pasta / secrets do velocímetro).
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Dict, Optional, Tuple

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
            background: rgba(255, 255, 255, 0.78) !important;
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
        /* Sem sidebar nestes relatórios */
        [data-testid="stSidebar"],
        [data-testid="stSidebarNav"],
        section[data-testid="stSidebar"],
        button[kind="header"] {{
            display: none !important;
        }}
        [data-testid="stAppViewContainer"] > .main {{
            margin-left: 0 !important;
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

def _default_base_periodo(semana_ini: date, semana_fim: date) -> Tuple[date, date]:
    """Base padrão: 12 semanas anteriores à semana selecionada (sem incluí-la)."""
    fim_base = semana_ini - timedelta(days=1)
    ini_base = segunda_da_semana(fim_base - timedelta(days=7 * 11))
    if ini_base > fim_base:
        ini_base = fim_base - timedelta(days=83)
    return ini_base, fim_base


def _montar_tabela_pessoa(
    nome: str,
    media: Dict[str, float],
    realizado: Dict[str, float],
) -> pd.DataFrame:
    rows = []
    for rotulo, fonte in (
        ("Média (equivalente ao período)", media),
        ("Realizado do período", realizado),
    ):
        row = {"Linha": rotulo}
        for e in FUNIL_ETAPAS:
            row[FUNIL_LABELS[e]] = float(fonte.get(e, 0.0))
        rows.append(row)

    row_pct = {"Linha": "% Realizado / Média"}
    for e in FUNIL_ETAPAS:
        m = float(media.get(e, 0.0))
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


def _render_aba_dimensao(
    eventos: pd.DataFrame,
    dimensao: str,
    ini_periodo: date,
    fim_periodo: date,
    ini_base: date,
    fim_base: date,
) -> None:
    label_dim = DIM_LABELS.get(dimensao, dimensao)
    dias_periodo = n_dias_periodo(ini_periodo, fim_periodo)
    dias_base = n_dias_periodo(ini_base, fim_base)

    ev_periodo = filtrar_periodo(eventos, ini_periodo, fim_periodo)
    ev_base = filtrar_periodo(eventos, ini_base, fim_base)

    real = agregar_funil_por_dimensao(ev_periodo, dimensao)
    base = agregar_funil_por_dimensao(ev_base, dimensao)
    media = escalar_media_para_periodo(base, dias_base, dias_periodo, dimensao)

    # Somente quem tem ≥1 indicador positivo no período escolhido
    if real.empty:
        st.info(f"Nenhum {label_dim.lower()} com indicadores no período selecionado.")
        return

    mask_pos = real[list(FUNIL_ETAPAS)].sum(axis=1) > 0
    real = real.loc[mask_pos].copy()
    if real.empty:
        st.info(f"Nenhum {label_dim.lower()} com indicador positivo no período.")
        return

    media_idx = media.set_index(dimensao) if not media.empty else pd.DataFrame()
    real = real.sort_values("vendas", ascending=False)

    st.caption(
        f"{len(real)} {label_dim.lower()}(is) com pelo menos 1 indicador no período · "
        f"média = (total da base ÷ {dias_base} dias) × {dias_periodo} dias do período."
    )

    for _, row in real.iterrows():
        nome = limpar_nome(row[dimensao])
        if not nome:
            continue
        realizado = {e: float(row.get(e, 0.0)) for e in FUNIL_ETAPAS}
        if nome in media_idx.index:
            media_row = media_idx.loc[nome]
            media_vals = {e: float(media_row.get(e, 0.0)) for e in FUNIL_ETAPAS}
        else:
            media_vals = {e: 0.0 for e in FUNIL_ETAPAS}

        df_t = _montar_tabela_pessoa(nome, media_vals, realizado)
        # Exibe % já como número; formatação trata a linha
        st.markdown(
            f'<div class="bloco-pessoa"><div class="nome">{nome}</div></div>',
            unsafe_allow_html=True,
        )
        # Converte linha de % para exibição amigável numa cópia formatada
        df_show = df_t.copy()
        # Mantém numérico; o styler formata
        st.dataframe(_estilo_tabela(df_show), use_container_width=True, hide_index=True)


def main() -> None:
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(
        page_title="Funil · Média × Período | Direcional",
        layout="wide",
        page_icon=str(fav) if fav else None,
        initial_sidebar_state="collapsed",
    )
    aplicar_estilo()
    _cabecalho_pagina("Funil por pessoa — média × período")
    st.caption(
        "Compara o funil Agendamentos → Visitas → Pastas → Pastas aprovadas → Vendas "
        "da média histórica (escalada pelos dias) com o período escolhido. "
        "Abas: Regional, Gerente de vendas e Corretor."
    )

    hoje = date.today()
    sem_ini_padrao, sem_fim_padrao = semana_iso_atual(hoje)

    # Defaults no session_state ANTES dos widgets (evita StreamlitAPIException)
    if "ini_per" not in st.session_state:
        st.session_state["ini_per"] = sem_ini_padrao
    if "fim_per" not in st.session_state:
        st.session_state["fim_per"] = sem_fim_padrao

    def _btn_semana_atual() -> None:
        st.session_state["ini_per"] = sem_ini_padrao
        st.session_state["fim_per"] = sem_fim_padrao

    def _btn_ajustar_seg_dom() -> None:
        d0 = segunda_da_semana(st.session_state.get("ini_per", sem_ini_padrao))
        st.session_state["ini_per"] = d0
        st.session_state["fim_per"] = domingo_da_semana(d0)

    st.markdown("##### Período de análise")
    st.caption("Padrão: semana atual (segunda → domingo).")
    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.3, 1.5])
    with c1:
        ini_periodo = st.date_input("Início do período", key="ini_per")
    with c2:
        fim_periodo = st.date_input("Fim do período", key="fim_per")
    with c3:
        st.write("")
        st.button("Usar semana atual (seg–dom)", on_click=_btn_semana_atual, use_container_width=True)
    with c4:
        st.write("")
        st.button(
            "Ajustar para seg–dom da data início",
            on_click=_btn_ajustar_seg_dom,
            use_container_width=True,
        )

    st.markdown("##### Base da média")
    st.caption(
        "Período usado para calcular a média diária "
        "(padrão: 12 semanas anteriores, sem o período selecionado)."
    )
    ini_b_pad, fim_b_pad = _default_base_periodo(ini_periodo, fim_periodo)
    if "ini_base" not in st.session_state:
        st.session_state["ini_base"] = ini_b_pad
    if "fim_base" not in st.session_state:
        st.session_state["fim_base"] = fim_b_pad
    b1, b2, _ = st.columns([1.2, 1.2, 2.8])
    with b1:
        ini_base = st.date_input("Início da base", key="ini_base")
    with b2:
        fim_base = st.date_input("Fim da base", key="fim_base")

    if fim_periodo < ini_periodo:
        st.error("O fim do período deve ser ≥ início.")
        return
    if fim_base < ini_base:
        st.error("O fim da base deve ser ≥ início.")
        return

    # Evita vazamento: remove interseção período ∩ base
    if not (fim_base < ini_periodo or ini_base > fim_periodo):
        st.warning(
            "A base da média intersecta o período selecionado. "
            "A média ficará enviesada — ajuste os filtros se quiser excluir o período."
        )

    try:
        eventos, origens = carregar_eventos_funil_pessoas()
    except Exception as e:
        st.error(f"Falha ao carregar bases Salesforce: {e}")
        return

    with st.expander("Origem dos dados", expanded=False):
        for k, v in origens.items():
            st.caption(f"**{k}:** {v}")
        st.caption(f"Eventos carregados: {len(eventos):,}".replace(",", "."))

    st.markdown(
        f"**Período:** {ini_periodo.strftime('%d/%m/%Y')} → {fim_periodo.strftime('%d/%m/%Y')} "
        f"({n_dias_periodo(ini_periodo, fim_periodo)} dias) · "
        f"**Base da média:** {ini_base.strftime('%d/%m/%Y')} → {fim_base.strftime('%d/%m/%Y')} "
        f"({n_dias_periodo(ini_base, fim_base)} dias)"
    )

    tabs = st.tabs([DIM_LABELS[d] for d in DIMENSOES])
    for tab, dim in zip(tabs, DIMENSOES):
        with tab:
            _render_aba_dimensao(
                eventos, dim, ini_periodo, fim_periodo, ini_base, fim_base
            )

    st.markdown(
        '<div class="footer">Direcional Engenharia · Vendas — Relatório de funil</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
