# -*- coding: utf-8 -*-
"""
Acompanhamento de vendas — metas vs realizado (Direcional).
Planilha: BD Vendas Completa + Metas.
Design: Gaps Style (Transparência, Blur, Inter/Montserrat).
Funcionalidade: Engenharia Reversa, Comparativo MTD e Pesos de Coordenadores.
"""
from __future__ import annotations

import base64
import calendar
import copy
import html
import math
import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# -----------------------------------------------------------------------------
# Identificação da planilha e Arquivos Visuais
# -----------------------------------------------------------------------------
SPREADSHEET_ID = "1wpuNQvksot9CLhGgQRe7JlyDeRISEh_sc3-6VRDyQYk"

WS_VENDAS = "BD Vendas Completa"
WS_METAS = "Metas"

# Funil comercial
SPREADSHEET_FUNIL_ID = "1ckdfpUr7qhr9YHnlfJ_rYrs6c6ZFugqYG0zrZDI9bck"
# Pastas / pastas aprovadas (aba BASE)
SPREADSHEET_PASTAS_ID: Optional[str] = "1wnJgJyrVM2k9SfQ8PCxFRr9odh75DEOg0tjtkPfyroc"
# Agendamentos/visitas, pastas e vendas: relatórios Salesforce (evita limite do Sheets)
SF_REPORT_AGENDAMENTOS_ID = "00OU600000AcFGPMA3"
SF_REPORT_PASTAS_ID = "00OU600000FEOoDMAX"
SF_REPORT_VENDAS_ID = "00O3Z000005ZsPmUAK"
ABA_AGENDAMENTOS_VISITAS = "Dados Únicos"  # fallback Sheets se SF falhar
ABA_PASTAS_CANDIDATAS = (
    "BASE",
    "Base",
    "Pastas",
    "Pastas e Pastas Aprovadas",
    "Pastas Aprovadas",
    "BD Pastas",
)
FUNIL_ETAPAS = ("agendamentos", "visitas", "pastas", "pastas_aprovadas", "vendas")
FUNIL_DRIVERS = ("agendamentos", "visitas", "pastas", "pastas_aprovadas")
FUNIL_LABELS = {
    "agendamentos": "Agendamentos",
    "visitas": "Visitas",
    "pastas": "Pastas",
    "pastas_aprovadas": "Pastas aprovadas",
    "vendas": "Vendas",
}
# Conversões etapa → etapa seguinte
FUNIL_PARES_ETAPA = (
    ("agendamentos", "visitas"),
    ("visitas", "pastas"),
    ("pastas", "pastas_aprovadas"),
    ("pastas_aprovadas", "vendas"),
)
FUNIL_LAGS = tuple(range(1, 31))  # lags 1..30
FUNIL_LAGS_PERFIL = tuple(range(1, 31))  # perfil de tempo até efeito: só lags 1..30
FUNIL_JANELA_CONV = 30  # janela móvel para conversão → vendas (sem vazamento do dia)
FUNIL_JANELA_FORCA = 7  # janela da força de trabalho (atividade cruzada)
FUNIL_ITERS_CRUZADAS = 3  # iterações na projeção (efeitos contemporâneos cruzados)
FUNIL_CORES_DRIVER = {
    "agendamentos": "#04428f",
    "visitas": "#cb0935",
    "pastas": "#0f766e",
    "pastas_aprovadas": "#b45309",
}
FUNIL_CORES_NIVEIS = ["#022654", "#04428f", "#1e60b3", "#cb0935", "#9e0828"]
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
    "SAFI Approval Date", "Approval Date SAFI",
]

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"
URL_LOGO_DIRECIONAL_EMAIL = "https://logodownload.org/wp-content/uploads/2021/04/direcional-engenharia-logo.png"

# Paleta alinhada à Ficha Credenciamento / Vendas RJ
COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_VERMELHO_ESCURO = "#9e0828"
COR_FUNDO_CARD = "rgba(255, 255, 255, 0.78)"
COR_BORDA = "#eef2f6"
COR_TEXTO_PRETO = "#000000"
COR_TEXTO_MUTED = "#000000"
COR_TEXTO_LABEL = "#000000"
COR_INPUT_BG = "#f0f2f6"

MESES_TEXTO_MAP = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
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
# Funções de Design (Ficha Direcional)
# -----------------------------------------------------------------------------
def _resolver_png_raiz(nome: str) -> Path | None:
    """Procura o PNG na pasta do app e na pasta pai."""
    for base in (_DIR_APP, _DIR_APP.parent):
        p = base / nome
        if p.is_file():
            return p
    return None

def _resolver_imagem_fundo_local(nome: str) -> Path | None:
    """Imagem JPG/PNG na pasta do app ou na pasta pai."""
    for base in (_DIR_APP, _DIR_APP.parent):
        for ext in (".jpg", ".jpeg", ".JPG", ".JPEG", ".png", ".PNG"):
            stem = Path(nome).stem
            p = base / f"{stem}{ext}"
            if p.is_file():
                return p
        p = base / nome
        if p.is_file():
            return p
    return None

def _css_url_fundo_cadastro() -> str:
    """String para `url(...)` no CSS: data-URL or URL https (fallback)."""
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

def _logo_arquivo_local() -> str | None:
    p_topo = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
    if p_topo:
        return str(p_topo)
    for name in ("logo_direcional.png", "logo_direcional.jpg", "logo_direcional.jpeg", "logo.png"):
        p = _DIR_APP / "assets" / name
        if p.is_file():
            return str(p)
    return None

def _logo_url_secrets() -> str | None:
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

def _logo_url_drive_por_id_arquivo() -> str | None:
    fid = (os.environ.get("DIrecIONAL_LOGO_FILE_ID") or "").strip()
    if len(fid) < 10:
        return None
    return f"https://drive.google.com/uc?export=view&id={fid}"

def _exibir_logo_topo() -> None:
    """Logo centralizada no topo."""
    path = _logo_arquivo_local()
    url = _logo_url_secrets() or _logo_url_drive_por_id_arquivo()
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
                f'<div class="ficha-logo-wrap"><img src="{html.escape(url)}" alt="Direcional" /></div>',
                unsafe_allow_html=True,
            )
    except Exception:
        pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Acompanhamento de metas de vendas</p>'
        f"</div>"
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true">'
        f'<div class="ficha-hero-bar"></div>'
        f"</div>"
        f"</div>",
        unsafe_allow_html=True,
    )


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
        html, body, :root, [data-testid="stApp"] {{
            color-scheme: light !important;
        }}
        html, body {{
            font-family: 'Inter', sans-serif;
            color: {COR_TEXTO_LABEL};
            background: transparent !important;
            background-color: transparent !important;
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
            background-color: transparent !important;
        }}
        header[data-testid="stHeader"],
        [data-testid="stHeader"] {{
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            border: none !important;
            box-shadow: none !important;
            backdrop-filter: none !important;
            -webkit-backdrop-filter: none !important;
        }}
        [data-testid="stHeader"] > div,
        [data-testid="stHeader"] header {{
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            box-shadow: none !important;
        }}
        [data-testid="stDecoration"] {{
            background: transparent !important;
            background-color: transparent !important;
        }}
        [data-testid="stSidebar"] {{ display: none !important; }}
        [data-testid="stSidebarCollapsedControl"] {{ display: none !important; }}
        [data-testid="stToolbar"] {{
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            border: none !important;
            border-radius: 0 !important;
            box-shadow: none !important;
            color: rgba(255, 255, 255, 0.92) !important;
        }}
        [data-testid="stToolbar"] button,
        [data-testid="stToolbar"] a {{
            color: rgba(255, 255, 255, 0.92) !important;
            background: transparent !important;
            background-color: transparent !important;
        }}
        [data-testid="stHeader"] button {{
            background: transparent !important;
            background-color: transparent !important;
        }}
        [data-testid="stToolbar"] svg {{
            fill: currentColor !important;
            color: inherit !important;
        }}
        [data-testid="stToolbar"] svg path[stroke] {{
            stroke: currentColor !important;
        }}
        [data-testid="stToolbar"] button:hover,
        [data-testid="stToolbar"] a:hover,
        [data-testid="stHeader"] button:hover {{
            background: rgba(255, 255, 255, 0.12) !important;
        }}
        [data-testid="stMain"] {{
            padding-left: clamp(14px, 5vw, 56px) !important;
            padding-right: clamp(14px, 5vw, 56px) !important;
            padding-top: clamp(12px, 3.5vh, 40px) !important;
            padding-bottom: clamp(14px, 4vh, 44px) !important;
            box-sizing: border-box !important;
        }}
        section.main > div {{
            padding-top: 0.25rem !important;
            padding-bottom: 0.35rem !important;
        }}
        .block-container {{
            max-width: 1700px !important;
            margin-left: auto !important;
            margin-right: auto !important;
            margin-top: clamp(4px, 1vh, 14px) !important;
            margin-bottom: clamp(4px, 1vh, 14px) !important;
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
        /* Títulos de seção do dashboard: azuis (inclui spans internos do Streamlit) */
        h1, h2, h3, h4,
        h1 *, h2 *, h3 *, h4 *,
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        [data-testid="stMarkdownContainer"] h4,
        [data-testid="stMarkdownContainer"] h1 *,
        [data-testid="stMarkdownContainer"] h2 *,
        [data-testid="stMarkdownContainer"] h3 *,
        [data-testid="stMarkdownContainer"] h4 *,
        [data-testid="stHeadingWithAction"],
        [data-testid="stHeadingWithAction"] *,
        .stHeading, .stHeading * {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_AZUL_ESC} !important;
            font-weight: 800 !important;
        }}
        h1, h2, h3, h4,
        [data-testid="stHeadingWithAction"],
        .stHeading {{
            text-align: center !important;
        }}
        /* Títulos de gráficos (#####): pretos */
        h5, h6,
        h5 *, h6 *,
        [data-testid="stMarkdownContainer"] h5,
        [data-testid="stMarkdownContainer"] h6,
        [data-testid="stMarkdownContainer"] h5 *,
        [data-testid="stMarkdownContainer"] h6 * {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_TEXTO_PRETO} !important;
            font-weight: 700 !important;
            text-align: center !important;
        }}
        /* Texto geral do dashboard: preto (não sobrescreve títulos) */
        .block-container,
        .block-container > div p,
        .block-container label,
        .block-container li,
        [data-testid="stMarkdownContainer"] > p,
        [data-testid="stCaption"],
        [data-testid="stCaptionContainer"],
        [data-testid="stCaptionContainer"] p,
        [data-testid="stWidgetLabel"] p,
        [data-testid="stWidgetLabel"] label,
        div[data-baseweb="select"] span,
        .stSelectbox label,
        .stMultiSelect label {{
            color: {COR_TEXTO_PRETO} !important;
        }}
        .block-container span:not(h1 span):not(h2 span):not(h3 span):not(h4 span) {{
            color: {COR_TEXTO_PRETO};
        }}
        h1 span, h2 span, h3 span, h4 span,
        [data-testid="stHeadingWithAction"] span {{
            color: {COR_AZUL_ESC} !important;
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
            vertical-align: middle;
        }}
        .ficha-hero-stack {{
            width: 100%;
            max-width: 100%;
            margin-bottom: 0.35rem;
            box-sizing: border-box;
        }}
        .ficha-hero {{
            text-align: center;
            padding: 0.5rem 0 0 0;
            margin: 0 auto 0 auto;
            max-width: 640px;
            animation: fichaFadeIn 0.85s cubic-bezier(0.22, 1, 0.36, 1) 0.1s both;
        }}
        .ficha-hero .ficha-title {{
            font-family: 'Montserrat', sans-serif;
            font-size: clamp(1.35rem, 3.5vw, 1.75rem);
            font-weight: 900;
            color: {COR_AZUL_ESC};
            margin: 0;
            line-height: 1.25;
            letter-spacing: -0.02em;
        }}
        .ficha-hero-bar-wrap {{
            width: 100%;
            max-width: 100%;
            margin: clamp(0.85rem, 2.4vw, 1.2rem) 0;
            padding: 0;
            box-sizing: border-box;
        }}
        .ficha-hero-bar {{
            height: 4px;
            width: 100%;
            margin: 0;
            border-radius: 999px;
            background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC});
            background-size: 200% 100%;
            animation: fichaShimmer 4s ease-in-out infinite alternate;
        }}
        .vel-kpi-row {{
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            margin-bottom: 1.25rem;
        }}
        .vel-kpi {{
            flex: 1 1 20%;
            background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.9) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 14px;
            padding: 14px 16px;
            text-align: center;
            box-shadow: 0 2px 8px rgba({RGB_AZUL_CSS}, 0.06);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        .vel-kpi:hover {{
            transform: translateY(-4px);
            box-shadow: 0 10px 20px -5px rgba({RGB_AZUL_CSS}, 0.15);
        }}
        .vel-kpi .lbl {{
            font-size: 0.72rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: {COR_TEXTO_PRETO};
            opacity: 0.85;
        }}
        /* Valores dos boxes (rótulos de dados): mantêm azul / vermelho */
        .vel-kpi .val {{
            font-family: 'Montserrat', sans-serif;
            font-size: 1.35rem;
            font-weight: 800;
            color: {COR_AZUL_ESC} !important;
            margin-top: 6px;
        }}
        .vel-kpi .val--red {{ color: {COR_VERMELHO} !important; }}
        div[data-testid="stMetric"] {{
            background: rgba(255,255,256,0.6);
            padding: 12px;
            border-radius: 12px;
            border: 1px solid {COR_BORDA};
        }}
        div[data-baseweb="input"] {{
            border-radius: 10px !important;
            border: 1px solid #e2e8f0 !important;
            background-color: {COR_INPUT_BG} !important;
            transition: border-color 0.2s ease, box-shadow 0.2s ease;
        }}
        div[data-baseweb="input"]:focus-within {{
            border-color: rgba({RGB_AZUL_CSS}, 0.35) !important;
            box-shadow: 0 0 0 3px rgba({RGB_AZUL_CSS}, 0.08) !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# Lógicas de Integração e Gsheets
# -----------------------------------------------------------------------------

def _secrets_connections_gsheets() -> Dict[str, Any]:
    try:
        sec = st.secrets
        if hasattr(sec, "get") and sec.get("connections"):
            g = sec["connections"].get("gsheets")
            if g is not None:
                return dict(g)
    except Exception:
        pass
    return {}


def _normalizar_private_key_toml(pk: str) -> str:
    s = (pk or "").strip()
    if not s:
        return s
    if "\\n" in s and "\n" not in s:
        s = s.replace("\\n", "\n")
    return s


def montar_service_account_info(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not raw:
        return None
    chaves = (
        "type", "project_id", "private_key_id", "private_key",
        "client_email", "client_id", "auth_uri", "token_uri",
        "auth_provider_x509_cert_url", "client_x509_cert_url",
    )
    out: Dict[str, Any] = {}
    for k in chaves:
        v = raw.get(k)
        if v is None: continue
        if isinstance(v, str): v = v.strip()
        if v == "": continue
        out[k] = v
    if "private_key" in out:
        out["private_key"] = _normalizar_private_key_toml(str(out["private_key"]))
    if "private_key" not in out or "client_email" not in out:
        return None
    typ = str(out.get("type") or "").strip()
    if not typ:
        out["type"] = "service_account"
    if "token_uri" not in out:
        out["token_uri"] = "https://oauth2.googleapis.com/token"
    if "auth_uri" not in out:
        out["auth_uri"] = "https://accounts.google.com/o/oauth2/auth"
    return out


def spreadsheet_id_de_secrets(cfg: Dict[str, Any]) -> str:
    for k in ("spreadsheet_id", "SPREADSHEET_ID", "spreadsheet", "planilha_id"):
        v = str(cfg.get(k) or "").strip()
        if v: return v
    return SPREADSHEET_ID


def valores_para_dataframe(rows: List[List[str]]) -> pd.DataFrame:
    if not rows: return pd.DataFrame()
    header = [str(c).strip() for c in rows[0]]
    w = len(header)
    if w == 0: return pd.DataFrame()
    body = rows[1:]
    if not body: return pd.DataFrame(columns=header)
    norm: List[List[str]] = []
    for r in body:
        cells = [str(c) for c in r]
        if len(cells) < w: cells = cells + [""] * (w - len(cells))
        else: cells = cells[:w]
        norm.append(cells)
    return pd.DataFrame(norm, columns=header)


def ler_aba_gsheets(
    service_account_info: Dict[str, Any],
    spreadsheet_id: str,
    worksheet: str,
) -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id.strip())
    nome = worksheet.strip()

    def _abrir() -> Any:
        try:
            return sh.worksheet(nome)
        except gspread.WorksheetNotFound:
            for w in sh.worksheets():
                if w.title.strip() == nome: return w
            for w in sh.worksheets():
                if w.title.strip().lower() == nome.lower(): return w
            titulos = [w.title for w in sh.worksheets()]
            raise gspread.WorksheetNotFound(
                f"Aba {nome!r} não encontrada. Abas: {titulos}"
            ) from None

    ws = _abrir()
    return valores_para_dataframe(ws.get_all_values())


def _fingerprint_credenciais(info: Dict[str, Any]) -> str:
    pk = str(info.get("private_key") or "")
    return str(hash(pk))[-12:] if pk else "0"


@st.cache_data(ttl=300, show_spinner=False)
def ler_planilha_aba_df(spreadsheet_id: str, worksheet: str, _cred_fp: str) -> pd.DataFrame:
    raw = _secrets_connections_gsheets()
    info = montar_service_account_info(raw)
    if not info: raise ValueError("Credenciais [connections.gsheets] ausentes ou incompleta.")
    return ler_aba_gsheets(info, spreadsheet_id, worksheet)


def _cabecalho_tem_coluna(header: List[str], aliases: List[str]) -> bool:
    lows = [str(h).strip().lower() for h in header]
    for a in aliases:
        al = a.strip().lower()
        if any(h == al or al in h for h in lows):
            return True
    return False


def _df_parece_pastas(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    return bool(
        achar_coluna(df, COLUNAS_PASTAS_ALIASES)
        or achar_coluna(df, COLUNAS_PASTAS_APROV_ALIASES)
    )


@st.cache_data(ttl=300, show_spinner=False)
def carregar_df_pastas_funil(
    spreadsheet_funil_id: str,
    spreadsheet_principal_id: str,
    spreadsheet_pastas_id: str,
    _cred_fp: str,
) -> Tuple[pd.DataFrame, str]:
    """
    Localiza a base de pastas:
      1) abas candidatas (BASE, Pastas, ...)
      2) varredura de abas com colunas Data Criação Pasta / Data Aprovação SAFI
    Procura na planilha de pastas (se informada), na de agendamentos e na principal.
    Retorna (df, origem descritiva).
    """
    import gspread
    from google.oauth2.service_account import Credentials

    raw = _secrets_connections_gsheets()
    info = montar_service_account_info(raw)
    if not info:
        raise ValueError("Credenciais [connections.gsheets] ausentes ou incompleta.")

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)

    ids: List[str] = []
    for sid in (spreadsheet_pastas_id, spreadsheet_funil_id, spreadsheet_principal_id):
        s = (sid or "").strip()
        if s and s not in ids:
            ids.append(s)

    # 1) candidatos por nome
    for sid in ids:
        try:
            sh = gc.open_by_key(sid)
        except Exception:
            continue
        titulos = {w.title.strip().lower(): w.title for w in sh.worksheets()}
        for cand in ABA_PASTAS_CANDIDATAS:
            real = titulos.get(cand.strip().lower())
            if not real:
                continue
            try:
                df = valores_para_dataframe(sh.worksheet(real).get_all_values())
                df = normalizar_colunas(df)
                if _df_parece_pastas(df):
                    return df, f"{real} ({sid[:8]}…)"
            except Exception:
                continue

    # 2) varredura por colunas no cabeçalho
    for sid in ids:
        try:
            sh = gc.open_by_key(sid)
        except Exception:
            continue
        for w in sh.worksheets():
            try:
                header = [str(c) for c in (w.row_values(1) or [])]
            except Exception:
                continue
            if not (
                _cabecalho_tem_coluna(header, COLUNAS_PASTAS_ALIASES)
                or _cabecalho_tem_coluna(header, COLUNAS_PASTAS_APROV_ALIASES)
            ):
                continue
            try:
                df = valores_para_dataframe(w.get_all_values())
                df = normalizar_colunas(df)
                if _df_parece_pastas(df):
                    return df, f"{w.title} ({sid[:8]}…)"
            except Exception:
                continue

    return pd.DataFrame(), ""


def parse_valor_br(val: Any) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace("R$", "").replace(" ", "").replace("\xa0", "")
    if not s or s.lower() == "nan" or s.lower() == "null": return 0.0
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s: return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0


def extrair_mes_da_data_venda(val: Any) -> Optional[int]:
    if val is None or pd.isna(val): return None
    s = str(val).strip()
    if not s or s in ("nan", "null", ""): return None
    
    if "/" in s:
        parts = s.split("/")
        if len(parts) >= 2 and parts[1].isdigit():
            m = int(parts[1])
            if 1 <= m <= 12: return m
            
    for k, v in MESES_TEXTO_MAP.items():
        if k in s.lower(): return v
    return None


def extrair_ano_da_data_venda(val: Any) -> Optional[int]:
    if val is None or pd.isna(val): return None
    s = str(val).strip()
    if not s or s in ("nan", "null", ""): return None
    
    if "/" in s:
        parts = s.split("/")
        if len(parts) >= 3:
            ano_str = re.sub(r"[^\d]", "", parts[2].split()[0])
            if len(ano_str) == 4 and ano_str.isdigit():
                return int(ano_str)
                
    cleaned = re.sub(r"[^\d]", "", s)
    if len(cleaned) >= 4:
        ano = int(cleaned[-4:])
        if ano > 2000: return ano
    return None


def _norm_txt_col(s: Any) -> str:
    """Normaliza texto de coluna para match sem acento/caixa."""
    import unicodedata
    t = str(s or "").strip().lower()
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


def achar_coluna(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    cols_norm = {_norm_txt_col(c): c for c in cols}
    # 1) match exato (com e sem acento)
    for a in aliases:
        al = str(a).strip().lower()
        for c in cols:
            if al == str(c).strip().lower():
                return c
        an = _norm_txt_col(a)
        if an in cols_norm:
            return cols_norm[an]
    # 2) substring
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
    """Localiza a coluna Data Aprovação SAFI (prioridade: contém 'safi' + 'aprov')."""
    col = achar_coluna(df, COLUNAS_PASTAS_APROV_ALIASES)
    if col:
        return col
    if df is None or df.empty:
        return None
    candidatas: List[str] = []
    for c in df.columns:
        cn = _norm_txt_col(c)
        if "safi" in cn and "aprov" in cn:
            candidatas.append(c)
        elif "safi" in cn and "data" in cn:
            candidatas.append(c)
    return candidatas[0] if candidatas else None


def achar_coluna_primeiro_envio_analise(df: pd.DataFrame) -> Optional[str]:
    """Localiza Data Primeiro Envio Análise (prioridade: 'primeiro' + 'envio')."""
    col = achar_coluna(df, COLUNAS_PASTAS_ALIASES)
    if col and "primeiro" in _norm_txt_col(col) and "envio" in _norm_txt_col(col):
        return col
    if df is None or df.empty:
        return None
    for c in df.columns:
        cn = _norm_txt_col(c)
        if "primeiro" in cn and "envio" in cn:
            return c
        if "1o" in cn and "envio" in cn:
            return c
        if "1" in cn and "envio" in cn and "analise" in cn:
            return c
    # fallback: aliases genéricos (ex.: Data da Análise) se não houver coluna específica
    return col


ALIASES_VENDA_COMERCIAL = [
    "Venda Comercial?", "Venda Comercial", "Venda comercial?",
    "Venda comercial", "Comercial?",
]


def filtrar_vendas_comerciais(df: pd.DataFrame) -> pd.DataFrame:
    """Mantém apenas vendas comerciais (Venda Comercial? = 1 / SIM / TRUE)."""
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


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def coluna_existe(df: pd.DataFrame, nome: str) -> bool:
    return nome in df.columns


def melt_metas(df_metas_raw: pd.DataFrame) -> pd.DataFrame:
    df = normalizar_colunas(df_metas_raw)
    df["_row_id"] = range(len(df))
    
    id_vars = [c for c in df.columns if str(c).lower() in ["empreendimento", "região", "regiao", "obra", "coordenador"]]
    id_vars_merge = id_vars + ["_row_id"]
    
    for c in id_vars:
        df[c] = df[c].fillna("Não Informado").astype(str).str.strip()
        df.loc[df[c] == "", c] = "Não Informado"
        df.loc[df[c].str.lower() == "nan", c] = "Não Informado"

    if "Coordenador" not in df.columns:
        df["Coordenador"] = "Não Informado"
        id_vars_merge.append("Coordenador")

    cols_qtd = [c for c in df.columns if re.match(r'^qtd\s*(1[0-2]|[1-9])$', str(c).lower().strip())]
    cols_vgv = [c for c in df.columns if re.match(r'^vgv\s*(1[0-2]|[1-9])$', str(c).lower().strip())]

    if not [c for c in df.columns if str(c).lower() in ["empreendimento", "obra"]] or not cols_qtd:
        return pd.DataFrame(columns=["Empreendimento", "Região", "Coordenador", "Mes_Num", "Meta_Qtd", "Meta_VGV"])

    df_qtd = df.melt(id_vars=id_vars_merge, value_vars=cols_qtd, var_name="Mes_Str", value_name="Meta_Qtd")
    df_qtd["Mes_Num"] = df_qtd["Mes_Str"].str.extract(r'(\d+)')[0].astype(int)
    df_qtd.drop(columns=["Mes_Str"], inplace=True)
    df_qtd["Meta_Qtd"] = pd.to_numeric(df_qtd["Meta_Qtd"], errors="coerce").fillna(0)

    if cols_vgv:
        df_vgv = df.melt(id_vars=id_vars_merge, value_vars=cols_vgv, var_name="Mes_Str", value_name="Meta_VGV")
        df_vgv["Mes_Num"] = df_vgv["Mes_Str"].str.extract(r'(\d+)')[0].astype(int)
        df_vgv.drop(columns=["Mes_Str"], inplace=True)
        df_vgv["Meta_VGV"] = df_vgv["Meta_VGV"].apply(parse_valor_br)
        out = pd.merge(df_qtd, df_vgv, on=id_vars_merge + ["Mes_Num"], how="outer").fillna(0)
    else:
        out = df_qtd.copy()
        out["Meta_VGV"] = 0.0

    out.drop(columns=["_row_id"], inplace=True, errors="ignore")

    for c in out.columns:
        if str(c).lower() in ["empreendimento", "obra"]:
            out.rename(columns={c: "Empreendimento"}, inplace=True)
        elif str(c).lower() in ["região", "regiao"]:
            out.rename(columns={c: "Região"}, inplace=True)
            
    return out


def fmt_br_milhoes(v: float) -> str:
    if v >= 1e6: return f"R$ {v / 1e6:.2f} mi"
    if v >= 1e3: return f"R$ {v / 1e3:.1f} mil"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_qtd(v: float) -> str:
    """Retorna int se .0, ou .1f se fracionado."""
    return f"{v:.1f}" if abs(v % 1) > 0.01 else str(int(v))


def fmt_funil_valor(v: float) -> str:
    """Valor exibido no funil: sempre inteiro (já arredondado para cima)."""
    return str(int(math.ceil(max(0.0, float(v)))))


def ceil_funil_totais(totais: Dict[str, float]) -> Dict[str, float]:
    """Arredonda para cima todos os volumes do funil."""
    return {e: float(math.ceil(max(0.0, float((totais or {}).get(e, 0.0))))) for e in FUNIL_ETAPAS}


FUNIL_FONTE_TAMANHO = 18
FUNIL_TEXTO_BRANCO = "#ffffff"


def _criar_fig_funil(
    labels: List[str],
    valores: List[float],
    titulo: str = "",
    cores: Optional[List[str]] = None,
    altura: int = 380,
) -> go.Figure:
    """
    Funil Plotly com texto branco dentro do bloco.
    Volumes sempre arredondados para cima (ceil).
    """
    vals = [float(math.ceil(max(0.0, float(v)))) for v in valores]
    textos = [fmt_funil_valor(v) for v in vals]
    fig = go.Figure(go.Funnel(
        y=labels,
        x=vals,
        text=textos,
        textinfo="text",
        textposition="auto",
        insidetextfont=dict(
            size=FUNIL_FONTE_TAMANHO,
            color=FUNIL_TEXTO_BRANCO,
            family="Inter",
        ),
        outsidetextfont=dict(
            size=FUNIL_FONTE_TAMANHO,
            color=COR_TEXTO_PRETO,
            family="Inter",
        ),
        marker={"color": cores or FUNIL_CORES_NIVEIS},
        connector={"fillcolor": "rgba(4, 66, 143, 0.15)"},
    ))
    layout: Dict[str, Any] = dict(
        margin=dict(l=10, r=90, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=altura,
        font=dict(family="Inter", color=COR_TEXTO_PRETO, size=16),
    )
    if titulo:
        layout["title"] = dict(
            text=titulo,
            font=dict(family="Inter", color=COR_TEXTO_PRETO, size=16),
        )
    fig.update_layout(**layout)
    return fig


DIAS_SEMANA_PT = {
    0: "segunda",
    1: "terça",
    2: "quarta",
    3: "quinta",
    4: "sexta",
    5: "sábado",
    6: "domingo",
}
MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}


def janela_treino_meses_exatos(hoje: Optional[date] = None) -> Tuple[date, date]:
    """
    Janela de treino em meses calendário exatos, excluindo o mês atual.
    Ex.: se hoje = julho/2026 → 01/07/2025 a 30/06/2026.
    """
    hoje = hoje or date.today()
    inicio = date(hoje.year - 1, hoje.month, 1)
    if hoje.month == 1:
        fim = date(hoje.year - 1, 12, 31)
    else:
        ano_fim, mes_fim = hoje.year, hoje.month - 1
        fim = date(ano_fim, mes_fim, calendar.monthrange(ano_fim, mes_fim)[1])
    return inicio, fim


def janela_treino_52_semanas(hoje: Optional[date] = None) -> Tuple[date, date]:
    """Compatibilidade: encaminha para janela de meses exatos."""
    return janela_treino_meses_exatos(hoje)


_TZ_BR = "America/Sao_Paulo"


def parse_data_serie(serie: pd.Series) -> pd.Series:
    """
    Converte datas Salesforce/relatórios para datetime64[ns] naive.

    Datas ISO do SF (`2025-07-01T12:14:25.000+0000` ou `2025-07-01`) NÃO podem
    usar dayfirst=True — o pandas troca 2025-07-01 por 2025-01-07 e invalida
    dias > 12, o que derruba agendamentos/visitas do mês atual.
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
            out.loc[has_time] = ts.dt.tz_convert(_TZ_BR).dt.tz_localize(None)
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


def _as_date_funil(val: Any) -> Optional[date]:
    """Normaliza valor para date (evita falha de lookup Timestamp vs date)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, pd.Timestamp):
        return val.date()
    try:
        dt = parse_data_serie(pd.Series([val])).iloc[0]
        if pd.isna(dt):
            return None
        return pd.Timestamp(dt).date()
    except Exception:
        return None


def _indice_por_data_cal(cal: pd.DataFrame) -> Dict[date, int]:
    """Mapa data → posição no calendário (índice resetado)."""
    out: Dict[date, int] = {}
    for i, r in cal.reset_index(drop=True).iterrows():
        d = _as_date_funil(r["data"])
        if d is not None:
            out[d] = int(i)
    return out


def _media_etapa_dia_semana(treino: pd.DataFrame, etapa: str, d: date) -> float:
    if treino.empty or etapa not in treino.columns:
        return 0.0
    ds = DIAS_SEMANA_PT[d.weekday()]
    sub = treino[treino["dia_semana"] == ds]
    if not sub.empty:
        return max(float(sub[etapa].mean()), 0.0)
    return max(float(treino[etapa].mean()), 0.0)


def _garantir_previsoes_futuras_funil(
    cal_reg: pd.DataFrame,
    cal_med: pd.DataFrame,
    idx_por_data: Dict[date, int],
    dias_futuros: List[date],
    treino: pd.DataFrame,
    medias: Dict[str, Any],
    coefs: Dict[str, np.ndarray],
    incluir_mes: bool,
    lags: Tuple[int, ...],
) -> None:
    """Evita projeção zerada no restante do mês (fallback: médias / histórico)."""
    for d in dias_futuros:
        i = idx_por_data.get(d)
        if i is None:
            continue
        for etapa in FUNIL_ETAPAS:
            v_reg = float(cal_reg.at[i, etapa])
            if v_reg < 1e-6:
                row_reg = cal_reg.loc[[i]]
                v_reg = _prever_linha_reg_funil(coefs[etapa], row_reg, incluir_mes, lags, etapa)
                if v_reg < 1e-6:
                    v_med = _prever_linha_medias_funil(
                        cal_med.iloc[i], d, etapa, medias, incluir_mes
                    )
                    v_reg = v_med if v_med > 1e-6 else _media_etapa_dia_semana(treino, etapa, d)
                cal_reg.at[i, etapa] = v_reg
            v_med = float(cal_med.at[i, etapa])
            if v_med < 1e-6:
                v_med = _prever_linha_medias_funil(
                    cal_med.iloc[i], d, etapa, medias, incluir_mes
                )
                if v_med < 1e-6:
                    v_med = _media_etapa_dia_semana(treino, etapa, d)
                cal_med.at[i, etapa] = v_med


def serie_diaria_contratos(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    col_qtd: str = "_qtd_venda",
    col_vgv: str = "_vgv_venda",
) -> pd.DataFrame:
    """Agrega vendas comerciais por data de 'Contrato gerado em'."""
    base = df_vendas.copy()
    base["_dt_contrato"] = parse_data_serie(base[col_contrato])
    base = base.dropna(subset=["_dt_contrato"])
    if base.empty:
        return pd.DataFrame(columns=["data", "qtd", "vgv"])

    base["_data"] = base["_dt_contrato"].dt.normalize()
    qtd_col = col_qtd if col_qtd in base.columns else None
    vgv_col = col_vgv if col_vgv in base.columns else None

    if qtd_col:
        base["_q"] = pd.to_numeric(base[qtd_col], errors="coerce").fillna(0.0)
    else:
        base["_q"] = 1.0
    if vgv_col:
        base["_v"] = pd.to_numeric(base[vgv_col], errors="coerce").fillna(0.0)
    else:
        base["_v"] = 0.0

    agg = (
        base.groupby("_data", as_index=False)
        .agg(qtd=("_q", "sum"), vgv=("_v", "sum"))
        .rename(columns={"_data": "data"})
    )
    agg["data"] = pd.to_datetime(agg["data"]).dt.date
    return agg


def calendario_diario(inicio: date, fim: date, serie: pd.DataFrame) -> pd.DataFrame:
    """Calendário completo com zeros nos dias sem venda (necessário para a regressão)."""
    idx = pd.date_range(inicio, fim, freq="D")
    cal = pd.DataFrame({"data": [d.date() for d in idx]})
    mapa = {r["data"]: (float(r["qtd"]), float(r["vgv"])) for _, r in serie.iterrows()}
    cal["qtd"] = cal["data"].map(lambda d: mapa.get(d, (0.0, 0.0))[0])
    cal["vgv"] = cal["data"].map(lambda d: mapa.get(d, (0.0, 0.0))[1])
    cal["dia_mes"] = cal["data"].map(lambda d: d.day)
    cal["dia_semana"] = cal["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    cal["mes"] = cal["data"].map(lambda d: MESES_PT[d.month])
    return cal


def _matriz_explicativas(df: pd.DataFrame, incluir_mes: bool = True) -> np.ndarray:
    """One-hot (numpy) de dia do mês + dia da semana (+ mês opcional) + intercepto."""
    n = len(df)
    n_cols = 31 + 7 + (12 if incluir_mes else 0) + 1
    X = np.zeros((n, n_cols), dtype=float)
    X[:, -1] = 1.0  # intercepto

    dias_semana_idx = {nome: i for i, nome in DIAS_SEMANA_PT.items()}
    meses_idx = {nome: i for i, nome in MESES_PT.items()}

    for i, row in enumerate(df.itertuples(index=False)):
        dia = int(row.dia_mes)
        if 1 <= dia <= 31:
            X[i, dia - 1] = 1.0
        ds = dias_semana_idx.get(str(row.dia_semana), None)
        if ds is not None:
            X[i, 31 + ds] = 1.0
        if incluir_mes:
            ms = meses_idx.get(str(row.mes), None)
            if ms is not None:
                X[i, 31 + 7 + (ms - 1)] = 1.0
    return X


def treinar_regressao_vendas_diarias(
    treino: pd.DataFrame,
    incluir_mes: bool = True,
) -> np.ndarray:
    """OLS via numpy.linalg.lstsq. Retorna vetor de coeficientes."""
    X = _matriz_explicativas(treino, incluir_mes=incluir_mes)
    y = treino["qtd"].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    return coef


def prever_qtd_dias(
    coef: np.ndarray,
    datas: List[date],
    incluir_mes: bool = True,
) -> np.ndarray:
    if not datas:
        return np.array([])
    df = pd.DataFrame({"data": datas})
    df["dia_mes"] = df["data"].map(lambda d: d.day)
    df["dia_semana"] = df["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    df["mes"] = df["data"].map(lambda d: MESES_PT[d.month])
    X = _matriz_explicativas(df, incluir_mes=incluir_mes)
    pred = X @ coef
    return np.maximum(pred, 0.0)


def _r2_treino(
    treino: pd.DataFrame,
    coef: np.ndarray,
    incluir_mes: bool = True,
) -> float:
    X = _matriz_explicativas(treino, incluir_mes=incluir_mes)
    y = treino["qtd"].astype(float).values
    y_hat = X @ coef
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    if ss_tot <= 0:
        return 0.0
    return 1.0 - (ss_res / ss_tot)


def calcular_medias_sazonais(
    treino: pd.DataFrame,
    incluir_mes: bool = True,
) -> Dict[str, Any]:
    """Médias históricas por dia da semana, dia do mês e (opcionalmente) mês."""
    mu = float(treino["qtd"].mean()) if len(treino) else 0.0
    if mu <= 0:
        mu = 1e-9
    out: Dict[str, Any] = {
        "mu": mu,
        "incluir_mes": incluir_mes,
        "media_dia_semana": {
            k: float(v) for k, v in treino.groupby("dia_semana")["qtd"].mean().items()
        },
        "media_dia_mes": {
            int(k): float(v) for k, v in treino.groupby("dia_mes")["qtd"].mean().items()
        },
    }
    if incluir_mes:
        out["media_mes"] = {
            k: float(v) for k, v in treino.groupby("mes")["qtd"].mean().items()
        }
    return out


def prever_qtd_medias(
    datas: List[date],
    medias: Dict[str, Any],
    incluir_mes: bool = True,
) -> np.ndarray:
    """
    Combinação multiplicativa das médias.
    Com mês:   pred = m_ds × m_dm × m_mes / μ²
    Sem mês:   pred = m_ds × m_dm / μ
    """
    if not datas:
        return np.array([])
    mu = float(medias["mu"])
    m_ds = medias["media_dia_semana"]
    m_dm = medias["media_dia_mes"]
    m_mes = medias.get("media_mes") or {}
    usar_mes = incluir_mes and bool(m_mes)
    out: List[float] = []
    for d in datas:
        a = float(m_ds.get(DIAS_SEMANA_PT[d.weekday()], mu))
        b = float(m_dm.get(d.day, mu))
        if usar_mes:
            c = float(m_mes.get(MESES_PT[d.month], mu))
            out.append(max((a * b * c) / (mu * mu), 0.0))
        else:
            out.append(max((a * b) / mu, 0.0))
    return np.asarray(out, dtype=float)


def _matriz_explicativas_relativa(df: pd.DataFrame) -> np.ndarray:
    """
    One-hot com categorias de referência omitidas:
      dia do mês → dia 1
      dia da semana → segunda
      mês → janeiro
    + intercepto (= nível esperado em segunda + dia 1 + janeiro).
    """
    n = len(df)
    # 30 dias (2–31) + 6 dias semana (terça–domingo) + 11 meses (fev–dez) + intercepto
    X = np.zeros((n, 30 + 6 + 11 + 1), dtype=float)
    X[:, -1] = 1.0

    dias_semana_idx = {nome: i for i, nome in DIAS_SEMANA_PT.items()}  # 0=segunda
    meses_idx = {nome: i for i, nome in MESES_PT.items()}  # 1=janeiro

    for i, row in enumerate(df.itertuples(index=False)):
        dia = int(row.dia_mes)
        if 2 <= dia <= 31:
            X[i, dia - 2] = 1.0  # colunas 0..29
        ds = dias_semana_idx.get(str(row.dia_semana), None)
        if ds is not None and ds >= 1:  # terça=1 .. domingo=6
            X[i, 30 + (ds - 1)] = 1.0
        ms = meses_idx.get(str(row.mes), None)
        if ms is not None and ms >= 2:  # fevereiro=2 .. dezembro=12
            X[i, 30 + 6 + (ms - 2)] = 1.0
    return X


def estimar_efeitos_sazonais(treino: pd.DataFrame) -> Optional[Dict[str, Any]]:
    """
    OLS com referência em segunda-feira, dia 1 e janeiro.
    Retorna efeitos aditivos (qtd) e índices relativos (baseline = 1.0).
    """
    if treino.empty or float(treino["qtd"].sum()) <= 0 or len(treino) < 30:
        return None

    X = _matriz_explicativas_relativa(treino)
    y = treino["qtd"].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    intercepto = float(coef[-1])
    base = intercepto if abs(intercepto) > 1e-9 else 1.0

    # dia do mês
    e_dm = [0.0]  # dia 1
    for d in range(2, 32):
        e_dm.append(float(coef[d - 2]))
    df_dm = pd.DataFrame({
        "categoria": [str(d) for d in range(1, 32)],
        "efeito": e_dm,
        "indice": [(base + e) / base for e in e_dm],
    })

    # dia da semana (ordem: segunda → domingo)
    nomes_ds = [DIAS_SEMANA_PT[i] for i in range(7)]
    e_ds = [0.0]  # segunda
    for i in range(1, 7):
        e_ds.append(float(coef[30 + (i - 1)]))
    df_ds = pd.DataFrame({
        "categoria": nomes_ds,
        "efeito": e_ds,
        "indice": [(base + e) / base for e in e_ds],
    })

    # mês (ordem: janeiro → dezembro)
    nomes_mes = [MESES_PT[i] for i in range(1, 13)]
    e_mes = [0.0]  # janeiro
    for m in range(2, 13):
        e_mes.append(float(coef[30 + 6 + (m - 2)]))
    df_mes = pd.DataFrame({
        "categoria": nomes_mes,
        "efeito": e_mes,
        "indice": [(base + e) / base for e in e_mes],
    })

    y_hat = X @ coef
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    r2 = (1.0 - ss_res / ss_tot) if ss_tot > 0 else 0.0

    return {
        "intercepto": intercepto,
        "r2": r2,
        "dia_mes": df_dm,
        "dia_semana": df_ds,
        "mes": df_mes,
    }


def _distribuir_gap_por_pesos(
    gap: float,
    pesos: np.ndarray,
    dias: List[date],
    qtd_mtd: float,
    dia_hoje: int,
    arredondar_cima: bool = False,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    ritmo_diario: List[Dict[str, Any]] = []
    ritmo_acum: List[Dict[str, Any]] = [{"dia": dia_hoje, "acum": qtd_mtd}]
    if not dias:
        return pd.DataFrame(ritmo_diario), pd.DataFrame(ritmo_acum)
    w = np.maximum(np.asarray(pesos, dtype=float), 0.0)
    soma = float(w.sum())
    if soma <= 1e-9:
        w = np.ones(len(dias), dtype=float)
        soma = float(len(dias))
    nec = gap * (w / soma)
    if arredondar_cima and gap > 0:
        nec = np.ceil(nec)
    running = qtd_mtd
    for i, d in enumerate(dias):
        v = float(nec[i])
        ritmo_diario.append({"dia": d.day, "qtd": v})
        running += v
        ritmo_acum.append({"dia": d.day, "acum": running})
    return pd.DataFrame(ritmo_diario), pd.DataFrame(ritmo_acum)


JANELAS_FIM_MES = (15, 10, 5)


def calcular_intensidade_fim_mes(treino: pd.DataFrame, janela: int = 15) -> float:
    """
    Razão empírica: média dos últimos `janela` dias do mês / média do restante.
    """
    if treino.empty:
        return 1.4
    limiar = max(1, 31 - janela + 1)
    early = treino[treino["dia_mes"] < limiar]["qtd"]
    late = treino[treino["dia_mes"] >= limiar]["qtd"]
    m_early = float(early.mean()) if len(early) else 0.0
    m_late = float(late.mean()) if len(late) else 0.0
    if m_early <= 1e-9:
        return 1.4
    ratio = m_late / m_early
    return float(np.clip(ratio, 1.10, 2.8))


def calcular_intensidades_fim_mes(
    treino: pd.DataFrame,
    janelas: Tuple[int, ...] = JANELAS_FIM_MES,
) -> Dict[int, float]:
    return {int(j): calcular_intensidade_fim_mes(treino, janela=int(j)) for j in janelas}


def fatores_sazonalidade_fim_mes(
    dias: List[date],
    ultimo_dia: int,
    intensidades: Dict[int, float],
    janelas: Tuple[int, ...] = JANELAS_FIM_MES,
) -> np.ndarray:
    """
    Combina sazonalidades de 15, 10 e 5 dias (produto de rampas).
    Quanto mais perto do fim do mês, maior o fator.
    """
    if not dias:
        return np.array([])
    out: List[float] = []
    for d in dias:
        f = 1.0
        for janela in janelas:
            intens = float(intensidades.get(int(janela), 1.0))
            inicio_boost = max(1, ultimo_dia - int(janela) + 1)
            if d.day >= inicio_boost:
                span = max(1, ultimo_dia - inicio_boost)
                t = (d.day - inicio_boost) / span  # 0 → 1
                f *= 1.0 + (intens - 1.0) * t
        out.append(f)
    return np.asarray(out, dtype=float)


def aplicar_sazonalidade_fim_mes(
    pesos: np.ndarray,
    dias: List[date],
    ultimo_dia: int,
    intensidades: Dict[int, float],
    janelas: Tuple[int, ...] = JANELAS_FIM_MES,
) -> np.ndarray:
    """Multiplica pesos pela sazonalidade combinada (15/10/5 dias)."""
    if len(pesos) == 0:
        return pesos
    fat = fatores_sazonalidade_fim_mes(dias, ultimo_dia, intensidades, janelas=janelas)
    return np.maximum(np.asarray(pesos, dtype=float), 0.0) * fat


def _series_atingido_projetado(
    dias_passados: List[date],
    dias_futuros: List[date],
    mapa_real: Dict[Any, float],
    pred_futuro: np.ndarray,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    diaria: List[Dict[str, Any]] = []
    acum: List[Dict[str, Any]] = []
    running = 0.0
    for d in dias_passados:
        q = float(mapa_real.get(d, 0.0))
        diaria.append({"dia": d.day, "qtd": q, "tipo": "Atingida"})
        running += q
        acum.append({"dia": d.day, "acum": running, "tipo": "Atingida"})
    for i, d in enumerate(dias_futuros):
        q = float(pred_futuro[i]) if i < len(pred_futuro) else 0.0
        diaria.append({"dia": d.day, "qtd": q, "tipo": "Projetada"})
        running += q
        acum.append({"dia": d.day, "acum": running, "tipo": "Projetada"})
    return pd.DataFrame(diaria), pd.DataFrame(acum)


def projetar_vendas_mes_atual(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    meta_vgv_mes: float,
    meta_qtd_mes: float = 0.0,
    hoje: Optional[date] = None,
    incluir_mes: bool = True,
) -> Optional[Dict[str, Any]]:
    """
    Projeção do mês corrente com duas lógicas (mesma janela de meses exatos):
      1) Regressão OLS (dia_mês + dia_semana [+ mês])
      2) Médias sazonais combinadas (multiplicativo)
    Com incluir_mes=False, remove o efeito/beta de mês (só dia do mês e dia da semana).
    Treino: do mês atual − 1 ano até o fim do mês anterior
    (ex.: jul/2026 → 01/07/2025 a 30/06/2026).
    """
    hoje = hoje or date.today()
    inicio, fim_treino = janela_treino_meses_exatos(hoje)
    serie = serie_diaria_contratos(df_vendas, col_contrato)
    if serie.empty:
        return None

    treino = calendario_diario(inicio, fim_treino, serie)
    if treino["qtd"].sum() <= 0 or len(treino) < 30:
        return None

    modelo = treinar_regressao_vendas_diarias(treino, incluir_mes=incluir_mes)
    medias = calcular_medias_sazonais(treino, incluir_mes=incluir_mes)

    ano, mes = hoje.year, hoje.month
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    dias_mes = [date(ano, mes, d) for d in range(1, ultimo_dia + 1)]
    dias_passados = [d for d in dias_mes if d <= hoje]
    dias_futuros = [d for d in dias_mes if d > hoje]

    mapa_real = {r["data"]: float(r["qtd"]) for _, r in serie.iterrows()}
    mapa_vgv = {r["data"]: float(r["vgv"]) for _, r in serie.iterrows()}

    qtd_mtd = float(sum(mapa_real.get(d, 0.0) for d in dias_passados))
    vgv_mtd = float(sum(mapa_vgv.get(d, 0.0) for d in dias_passados))

    pred_reg_mes = prever_qtd_dias(modelo, dias_mes, incluir_mes=incluir_mes)
    pred_med_mes = prever_qtd_medias(dias_mes, medias, incluir_mes=incluir_mes)

    # Reforço explícito de fim de mês: sazonalidade em 15, 10 e 5 dias
    intensidades_fim = calcular_intensidades_fim_mes(treino, janelas=JANELAS_FIM_MES)
    pred_reg_mes = aplicar_sazonalidade_fim_mes(pred_reg_mes, dias_mes, ultimo_dia, intensidades_fim)
    pred_med_mes = aplicar_sazonalidade_fim_mes(pred_med_mes, dias_mes, ultimo_dia, intensidades_fim)

    n_passados = len(dias_passados)
    pred_reg = pred_reg_mes[n_passados:] if len(pred_reg_mes) else np.array([])
    pred_med = pred_med_mes[n_passados:] if len(pred_med_mes) else np.array([])

    mu_dia = max(float(treino["qtd"].mean()), 0.0)
    if len(pred_reg) > 0 and float(np.sum(pred_reg)) < 1e-6 and mu_dia > 0:
        pred_reg = np.full(len(pred_reg), mu_dia)
        pred_reg_mes = np.concatenate([pred_reg_mes[:n_passados], pred_reg])
    if len(pred_med) > 0 and float(np.sum(pred_med)) < 1e-6 and mu_dia > 0:
        pred_med = np.full(len(pred_med), mu_dia)
        pred_med_mes = np.concatenate([pred_med_mes[:n_passados], pred_med])

    # Comparativo dia a dia: projetado (modelo) × realizado (mês atual)
    comp_reg: List[Dict[str, Any]] = []
    comp_med: List[Dict[str, Any]] = []
    for i, d in enumerate(dias_mes):
        real = float(mapa_real.get(d, 0.0)) if d <= hoje else None
        p_reg = float(pred_reg_mes[i]) if i < len(pred_reg_mes) else 0.0
        p_med = float(pred_med_mes[i]) if i < len(pred_med_mes) else 0.0
        comp_reg.append({"dia": d.day, "realizado": real, "projetado": p_reg})
        comp_med.append({"dia": d.day, "realizado": real, "projetado": p_med})

    qtd_proj_reg = qtd_mtd + float(pred_reg.sum() if len(pred_reg) else 0.0)
    qtd_proj_med = qtd_mtd + float(pred_med.sum() if len(pred_med) else 0.0)

    # Ticket médio dos 30 dias anteriores à projeção (estoque muda constantemente)
    fim_ticket = hoje - timedelta(days=1)
    inicio_ticket = hoje - timedelta(days=30)
    qtd_30 = float(sum(mapa_real.get(inicio_ticket + timedelta(days=i), 0.0) for i in range(30)))
    vgv_30 = float(sum(mapa_vgv.get(inicio_ticket + timedelta(days=i), 0.0) for i in range(30)))
    ticket_medio = (vgv_30 / qtd_30) if qtd_30 > 0 else 0.0
    if ticket_medio <= 0:
        q_tr = float(treino["qtd"].sum())
        v_tr = float(treino["vgv"].sum())
        ticket_medio = (v_tr / q_tr) if q_tr > 0 else 0.0

    rest_reg = qtd_proj_reg - qtd_mtd
    rest_med = qtd_proj_med - qtd_mtd
    vgv_proj_reg = vgv_mtd + rest_reg * ticket_medio
    vgv_proj_med = vgv_mtd + rest_med * ticket_medio
    pct_vgv_reg = (vgv_proj_reg / meta_vgv_mes * 100.0) if meta_vgv_mes and meta_vgv_mes > 0 else 0.0
    pct_vgv_med = (vgv_proj_med / meta_vgv_mes * 100.0) if meta_vgv_mes and meta_vgv_mes > 0 else 0.0

    meta_qtd = float(meta_qtd_mes or 0.0)
    gap_qtd = max(0.0, meta_qtd - qtd_mtd)

    ritmo_reg_d, ritmo_reg_a = _distribuir_gap_por_pesos(
        gap_qtd, pred_reg, dias_futuros, qtd_mtd, hoje.day, arredondar_cima=True
    )
    ritmo_med_d, ritmo_med_a = _distribuir_gap_por_pesos(
        gap_qtd, pred_med, dias_futuros, qtd_mtd, hoje.day, arredondar_cima=True
    )

    diaria_reg, acum_reg = _series_atingido_projetado(
        dias_passados, dias_futuros, mapa_real, pred_reg
    )
    diaria_med, acum_med = _series_atingido_projetado(
        dias_passados, dias_futuros, mapa_real, pred_med
    )

    intensidade_resumo = float(np.mean(list(intensidades_fim.values()))) if intensidades_fim else 1.0

    return {
        "hoje": hoje,
        "inicio_treino": inicio,
        "fim_treino": fim_treino,
        "incluir_mes": incluir_mes,
        "qtd_mtd": qtd_mtd,
        "vgv_mtd": vgv_mtd,
        "ticket_medio": ticket_medio,
        "inicio_ticket_30d": inicio_ticket,
        "fim_ticket_30d": fim_ticket,
        "ultimo_dia": ultimo_dia,
        "meta_vgv_mes": meta_vgv_mes,
        "meta_qtd_mes": meta_qtd,
        "gap_qtd_meta": gap_qtd,
        "r2_treino": _r2_treino(treino, modelo, incluir_mes=incluir_mes),
        "medias": medias,
        "intensidades_fim_mes": intensidades_fim,
        "intensidade_fim_mes": intensidade_resumo,
        "qtd_projetada_mes": qtd_proj_reg,
        "vgv_projetado": vgv_proj_reg,
        "pct_vgv_meta": pct_vgv_reg,
        "diaria": diaria_reg,
        "acumulado": acum_reg,
        "ritmo_meta_diario": ritmo_reg_d,
        "ritmo_meta_acum": ritmo_reg_a,
        "qtd_projetada_medias": qtd_proj_med,
        "vgv_projetado_medias": vgv_proj_med,
        "pct_vgv_medias": pct_vgv_med,
        "diaria_medias": diaria_med,
        "acumulado_medias": acum_med,
        "ritmo_meta_diario_medias": ritmo_med_d,
        "ritmo_meta_acum_medias": ritmo_med_a,
        "comparativo_diario_reg": pd.DataFrame(comp_reg),
        "comparativo_diario_medias": pd.DataFrame(comp_med),
    }


def _plot_projecao_mes(
    titulo: str,
    caption: str,
    proj: Dict[str, Any],
    diaria: pd.DataFrame,
    acumulado: pd.DataFrame,
) -> None:
    """Gráfico de projeção: realizado até hoje + projetado a partir de amanhã (sem meta)."""
    st.markdown(f"##### {titulo}")
    if caption:
        st.caption(caption)

    dia_hoje = proj["hoje"].day
    ating_d = diaria[diaria["tipo"] == "Atingida"] if not diaria.empty else diaria
    proj_d = diaria[diaria["tipo"] == "Projetada"] if not diaria.empty else diaria
    ating_a = acumulado[acumulado["tipo"] == "Atingida"] if not acumulado.empty else acumulado
    proj_a = acumulado[acumulado["tipo"] == "Projetada"] if not acumulado.empty else acumulado

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    if not ating_d.empty:
        fig.add_trace(
            go.Bar(
                x=ating_d["dia"], y=ating_d["qtd"], name="Realizado (dia)",
                marker_color=COR_AZUL_ESC, opacity=0.85,
            ),
            secondary_y=False,
        )
    if not proj_d.empty:
        fig.add_trace(
            go.Bar(
                x=proj_d["dia"], y=proj_d["qtd"], name="Projetado (dia)",
                marker_color="rgba(203, 9, 53, 0.45)",
            ),
            secondary_y=False,
        )

    if not ating_a.empty:
        fig.add_trace(
            go.Scatter(
                x=ating_a["dia"], y=ating_a["acum"], mode="lines+markers",
                name="Acumulado realizado",
                line=dict(color=COR_AZUL_ESC, width=3), marker=dict(size=7),
            ),
            secondary_y=True,
        )
    if not proj_a.empty:
        x_proj = [dia_hoje] + proj_a["dia"].tolist()
        y_proj = [float(ating_a["acum"].iloc[-1]) if not ating_a.empty else 0.0] + proj_a["acum"].tolist()
        fig.add_trace(
            go.Scatter(
                x=x_proj, y=y_proj, mode="lines+markers",
                name="Acumulado projetado",
                line=dict(color=COR_VERMELHO, width=3, dash="dash"),
                marker=dict(size=7, color=COR_VERMELHO),
            ),
            secondary_y=True,
        )

    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        barmode="group",
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=420,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, proj["ultimo_dia"] + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        secondary_y=False,
        showgrid=False,
    )
    fig.update_yaxes(
        title_text="Qtd. acumulada",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        secondary_y=True,
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _plot_meta_diaria_integrada(proj: Dict[str, Any]) -> None:
    """
    Um único gráfico com 3 linhas (somente diário, sem acumulado):
      1) Realizado até hoje
      2) Meta diária (regressão) a partir de amanhã
      3) Meta diária (médias) a partir de amanhã
    """
    ints = proj.get("intensidades_fim_mes") or {}
    st.caption(
        " · ".join(f"Sazonalidade {j}d: {float(ints.get(j, 1)):.2f}x" for j in JANELAS_FIM_MES)
    )

    dia_hoje = proj["hoje"].day
    diaria = proj.get("diaria", pd.DataFrame())
    ating_d = (
        diaria[diaria["tipo"] == "Atingida"]
        if (not diaria.empty and "tipo" in diaria.columns)
        else pd.DataFrame()
    )
    ritmo_reg = proj.get("ritmo_meta_diario", pd.DataFrame())
    ritmo_med = proj.get("ritmo_meta_diario_medias", pd.DataFrame())

    fig = go.Figure()

    if not ating_d.empty:
        textos_real = [fmt_qtd(float(v)) for v in ating_d["qtd"]]
        fig.add_trace(
            go.Scatter(
                x=ating_d["dia"],
                y=ating_d["qtd"],
                mode="lines+markers+text",
                name="Realizado até hoje",
                text=textos_real,
                textposition="top center",
                textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
                line=dict(color=COR_AZUL_ESC, width=3),
                marker=dict(size=8, color=COR_AZUL_ESC),
            )
        )

    if ritmo_reg is not None and not ritmo_reg.empty:
        textos_reg = [fmt_qtd(float(v)) for v in ritmo_reg["qtd"]]
        fig.add_trace(
            go.Scatter(
                x=ritmo_reg["dia"],
                y=ritmo_reg["qtd"],
                mode="lines+markers+text",
                name="Meta diária (regressão)",
                text=textos_reg,
                textposition="top center",
                textfont=dict(size=10, color=COR_VERMELHO, family="Inter"),
                line=dict(color=COR_VERMELHO, width=3),
                marker=dict(size=8, color=COR_VERMELHO),
            )
        )

    if ritmo_med is not None and not ritmo_med.empty:
        textos_med = [fmt_qtd(float(v)) for v in ritmo_med["qtd"]]
        fig.add_trace(
            go.Scatter(
                x=ritmo_med["dia"],
                y=ritmo_med["qtd"],
                mode="lines+markers+text",
                name="Meta diária (médias)",
                text=textos_med,
                textposition="bottom center",
                textfont=dict(size=10, color="#0f766e", family="Inter"),
                line=dict(color="#0f766e", width=3, dash="dash"),
                marker=dict(size=8, symbol="diamond", color="#0f766e"),
            )
        )

    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=460,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, proj["ultimo_dia"] + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _plot_comparativo_realizado_projetado(proj: Dict[str, Any]) -> None:
    """Gráfico: projetado × realizado por dia no mês atual (regressão e médias)."""
    st.markdown("##### Projetado × realizado por dia (mês atual)")

    dia_hoje = proj["hoje"].day
    df_reg = proj.get("comparativo_diario_reg", pd.DataFrame())
    df_med = proj.get("comparativo_diario_medias", pd.DataFrame())
    if (df_reg is None or df_reg.empty) and (df_med is None or df_med.empty):
        st.info("Sem dados para o comparativo diário.")
        return

    fig = go.Figure()

    # Realizado (uma série; igual em reg e médias)
    base = df_reg if (df_reg is not None and not df_reg.empty) else df_med
    real = base.dropna(subset=["realizado"]) if "realizado" in base.columns else pd.DataFrame()
    if not real.empty:
        textos_real = [fmt_qtd(float(v)) for v in real["realizado"]]
        fig.add_trace(
            go.Scatter(
                x=real["dia"],
                y=real["realizado"],
                mode="lines+markers+text",
                name="Realizado",
                text=textos_real,
                textposition="top center",
                textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
                line=dict(color=COR_AZUL_ESC, width=3),
                marker=dict(size=8, color=COR_AZUL_ESC),
            )
        )

    if df_reg is not None and not df_reg.empty:
        textos_reg = [fmt_qtd(float(v)) for v in df_reg["projetado"]]
        fig.add_trace(
            go.Scatter(
                x=df_reg["dia"],
                y=df_reg["projetado"],
                mode="lines+markers+text",
                name="Projetado (regressão)",
                text=textos_reg,
                textposition="bottom center",
                textfont=dict(size=10, color=COR_VERMELHO, family="Inter"),
                line=dict(color=COR_VERMELHO, width=3, dash="dash"),
                marker=dict(size=8, color=COR_VERMELHO),
            )
        )

    if df_med is not None and not df_med.empty:
        textos_med = [fmt_qtd(float(v)) for v in df_med["projetado"]]
        fig.add_trace(
            go.Scatter(
                x=df_med["dia"],
                y=df_med["projetado"],
                mode="lines+markers+text",
                name="Projetado (médias)",
                text=textos_med,
                textposition="top center",
                textfont=dict(size=10, color="#0f766e", family="Inter"),
                line=dict(color="#0f766e", width=3, dash="dot"),
                marker=dict(size=8, symbol="diamond", color="#0f766e"),
            )
        )

    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=460,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, proj["ultimo_dia"] + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_projecao_vendas(
    proj: Dict[str, Any],
    titulo: Optional[str] = None,
) -> None:
    """Seção Streamlit: cartões + gráficos de projeção e gráfico único de meta diária."""
    incluir_mes = bool(proj.get("incluir_mes", True))
    if titulo is None:
        titulo = (
            "Projeção de Vendas"
            if incluir_mes
            else "Projeção de Vendas (sem efeito de mês)"
        )

    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader(titulo)

    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi">
                <div class="lbl">Realizado MTD</div>
                <div class="val">{fmt_qtd(proj['qtd_mtd'])}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">Qtd. (regressão)</div>
                <div class="val">{fmt_qtd(proj['qtd_projetada_mes'])}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">Qtd. (médias)</div>
                <div class="val">{fmt_qtd(proj.get('qtd_projetada_medias', 0))}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">VGV (regressão)</div>
                <div class="val val--red">{fmt_br_milhoes(proj['vgv_projetado'])}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">VGV (médias)</div>
                <div class="val val--red">{fmt_br_milhoes(proj.get('vgv_projetado_medias', 0))}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">% VGV reg. / médias</div>
                <div class="val">{proj['pct_vgv_meta']:.1f}% / {proj.get('pct_vgv_medias', 0):.1f}%</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _plot_projecao_mes(
        "Projeção por regressão",
        f"R² treino: {proj['r2_treino']:.2f}",
        proj,
        proj["diaria"],
        proj["acumulado"],
    )

    med = proj.get("medias") or {}
    mu = float(med.get("mu") or 0.0)
    _plot_projecao_mes(
        "Projeção por médias sazonais",
        f"μ: {mu:.2f}",
        proj,
        proj.get("diaria_medias", pd.DataFrame()),
        proj.get("acumulado_medias", pd.DataFrame()),
    )

    st.markdown("<br/>", unsafe_allow_html=True)
    _plot_comparativo_realizado_projetado(proj)

    st.markdown("<br/>", unsafe_allow_html=True)
    st.subheader(
        "Meta diária para bater a quantidade"
        if incluir_mes
        else "Meta diária para bater a quantidade (sem efeito de mês)"
    )
    _plot_meta_diaria_integrada(proj)


def _plot_barra_efeitos(
    titulo: str,
    df: pd.DataFrame,
    cor: str,
    referencia: str,
) -> None:
    """Barra de efeitos relativos (índice; referência = 1.0)."""
    st.markdown(f"##### {titulo}")
    st.caption(f"Referência: {referencia} = 1,00")
    if df is None or df.empty:
        return

    textos = [f"{float(v):.2f}" for v in df["indice"]]
    cores = [COR_TEXTO_PRETO if abs(float(v) - 1.0) < 1e-9 else cor for v in df["indice"]]

    fig = go.Figure(
        go.Bar(
            x=df["categoria"],
            y=df["indice"],
            text=textos,
            textposition="outside",
            textfont=dict(size=10, color=COR_TEXTO_PRETO, family="Inter"),
            marker_color=cores,
            name="Índice",
            hovertemplate="%{x}<br>Índice: %{y:.2f}<br>Efeito (qtd): %{customdata:.2f}<extra></extra>",
            customdata=df["efeito"],
        )
    )
    fig.add_hline(
        y=1.0, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Ref. = 1", annotation_position="top left",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=10),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        showlegend=False,
        height=360,
        bargap=0.25,
    )
    fig.update_xaxes(
        title_text="",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter", size=11),
    )
    fig.update_yaxes(
        title_text="Índice (ref. = 1)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_efeitos_sazonais(efeitos: Dict[str, Any]) -> None:
    """Seção: efeitos de dia da semana, dia do mês e mês (relativos à referência)."""
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Efeitos sazonais (regressão)")
    st.caption(
        f"R²: {float(efeitos.get('r2', 0)):.2f} · "
        f"Baseline (segunda + dia 1 + janeiro): {float(efeitos.get('intercepto', 0)):.2f} vendas/dia"
    )

    _plot_barra_efeitos(
        "Efeito de dia da semana",
        efeitos.get("dia_semana", pd.DataFrame()),
        COR_AZUL_ESC,
        "segunda-feira",
    )
    _plot_barra_efeitos(
        "Efeito de dia do mês",
        efeitos.get("dia_mes", pd.DataFrame()),
        COR_VERMELHO,
        "dia 1",
    )
    _plot_barra_efeitos(
        "Efeito de mês",
        efeitos.get("mes", pd.DataFrame()),
        "#0f766e",
        "janeiro",
    )


# -----------------------------------------------------------------------------
# Funil comercial: agendamentos → visitas → pastas → pastas aprovadas → vendas
# -----------------------------------------------------------------------------

def _aplicar_secrets_salesforce() -> None:
    """Copia [salesforce] dos secrets Streamlit para variáveis de ambiente."""
    try:
        if hasattr(st, "secrets") and "salesforce" in st.secrets:
            sec = st.secrets["salesforce"]
            if sec.get("USER"):
                os.environ["SALESFORCE_USER"] = str(sec["USER"]).strip()
            if sec.get("PASSWORD"):
                os.environ["SALESFORCE_PASSWORD"] = str(sec["PASSWORD"]).strip()
            if sec.get("TOKEN"):
                os.environ["SALESFORCE_TOKEN"] = str(sec["TOKEN"]).strip()
            dom = str(sec.get("DOMAIN") or sec.get("domain") or "").strip()
            if dom:
                os.environ["SALESFORCE_DOMAIN"] = dom
    except Exception:
        pass


def conectar_salesforce_app() -> Tuple[Any, Optional[str]]:
    """Conecta ao Salesforce via simple_salesforce. Retorna (cliente, erro)."""
    _aplicar_secrets_salesforce()
    try:
        from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
    except ImportError:
        return None, "Pacote simple-salesforce não instalado (pip install simple-salesforce)."

    username = (os.environ.get("SALESFORCE_USER") or "").strip()
    password = (os.environ.get("SALESFORCE_PASSWORD") or "").strip()
    token = (os.environ.get("SALESFORCE_TOKEN") or "").strip()
    domain = (os.environ.get("SALESFORCE_DOMAIN") or "login").strip() or "login"
    if not username or not password:
        return None, "Credenciais Salesforce ausentes ([salesforce] USER/PASSWORD/TOKEN nos secrets)."
    try:
        kwargs: Dict[str, Any] = {
            "username": username,
            "password": password,
            "domain": domain,
        }
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
        # value às vezes vem como dict/OrderedDict (lookup) — usa label
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
                "500" in msg and "relat" in msg
            ) or "forbidden" in msg and ("60 minuto" in msg or "60 minute" in msg or "rate" in msg)
            rate = rate or ("não é possível executar mais de 500" in msg) or (
                "nao e possivel executar mais de 500" in msg
            )
            if not rate or i >= tentativas - 1:
                raise
            # backoff crescente (rate limit SF: 500 sync / 60 min)
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
        # Id clássico / lookup Id
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
        # Nome da avaliação de crédito (pastas) — chave de dedup
        if "nome da avalia" in lab_l or "nome_da_avalia" in api_l.replace("__", "_").lower():
            score += 70
        # Evita campos de status/texto genéricos
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
    # Dedup seguro (células podem ter tipos mistos)
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
    """
    Início dos 12 meses fechados. A consulta também traz o mês atual para
    realizado MTD/semanal; treino e médias o excluem explicitamente.
    """
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

def carregar_agendamentos_visitas_salesforce(
    report_id: str = SF_REPORT_AGENDAMENTOS_ID,
) -> Tuple[pd.DataFrame, str]:
    """Compatibilidade: delega para carregar_relatorio_salesforce."""
    return carregar_relatorio_salesforce(report_id, rotulo="agendamentos/visitas")


def contar_eventos_por_dia(df: pd.DataFrame, aliases: List[str]) -> Dict[date, float]:
    """Conta ocorrências por dia a partir de uma coluna de data."""
    col = achar_coluna(df, aliases)
    if not col or df.empty:
        return {}
    return contar_eventos_por_coluna(df, col)


def contar_eventos_por_coluna(df: pd.DataFrame, col: str) -> Dict[date, float]:
    """Conta ocorrências por dia a partir de uma coluna já resolvida."""
    if not col or col not in df.columns or df.empty:
        return {}
    dt = parse_data_serie(df[col]).dropna()
    if dt.empty:
        return {}
    vc = dt.dt.normalize().value_counts()
    return {pd.Timestamp(k).date(): float(v) for k, v in vc.items()}


def deduplicar_pastas_aprovadas_funil(df: pd.DataFrame) -> pd.DataFrame:
    """
    Uma linha por Nome da Avaliação de crédito, mantendo a Data Aprovação SAFI
    mais recente.
    """
    col_safi = achar_coluna_aprovacao_safi(df)
    if not col_safi:
        return df if df is not None else pd.DataFrame()
    return deduplicar_por_chave_mais_recente(
        df, ALIASES_NOME_AVALIACAO_CREDITO, [col_safi]
    )


def deduplicar_por_chave_mais_recente(
    df: pd.DataFrame,
    aliases_chave: List[str],
    aliases_data: List[str],
) -> pd.DataFrame:
    """
    Remove duplicatas pela chave, mantendo a linha com a data mais recente.
    Se chave ou data não existirem, devolve o DataFrame original.
    """
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    col_chave = achar_coluna(df, aliases_chave)
    col_data = achar_coluna(df, aliases_data)
    if not col_chave or not col_data:
        return df
    out = df.copy()
    out["_dedup_dt"] = parse_data_serie(out[col_data])
    out["_dedup_key"] = out[col_chave].astype(str).str.strip()
    # Chaves vazias/NaN não entram na deduplicação entre si
    mask_ok = out["_dedup_key"].ne("") & out["_dedup_key"].str.lower().ne("nan")
    ok = out.loc[mask_ok].sort_values("_dedup_dt", ascending=False, na_position="last")
    ok = ok.drop_duplicates(subset=["_dedup_key"], keep="first")
    resto = out.loc[~mask_ok]
    out = pd.concat([ok, resto], ignore_index=True)
    return out.drop(columns=["_dedup_dt", "_dedup_key"], errors="ignore")


ALIASES_ID_OPORTUNIDADE = [
    "ID da Oportunidade", "Id da Oportunidade", "Opportunity ID", "Opportunity Id",
    "ID Oportunidade", "Id Oportunidade",
]
ALIASES_CONTRATO_GERADO = [
    "Contrato gerado em", "Contrato Gerado em", "Contrato gerado",
    "Data do Contrato", "Data Contrato", "Close Date", "Data da venda",
    "Data Venda",
]
ALIASES_NOME_AVALIACAO_CREDITO = [
    "Nome da Avaliação de crédito", "Nome da Avaliacao de credito",
    "Nome da Avaliação de Crédito", "Nome da Avaliacao de Credito",
    "Nome Avaliação de crédito", "Nome Avaliacao de credito",
]
ALIASES_CODIGO_AGENDAMENTO = [
    "Código do agendamento", "Codigo do agendamento",
    "Código do Agendamento", "Codigo do Agendamento",
    "Código Agendamento", "Codigo Agendamento",
]
ALIASES_DATA_CRIACAO = [
    "Data de criação", "Data de criacao", "Data Criação", "Data Criacao",
    "Created Date", "Data de Criação", "Criado em", "Data criação",
]


def deduplicar_vendas_funil(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por ID da Oportunidade (Contrato gerado em mais recente)."""
    return deduplicar_por_chave_mais_recente(df, ALIASES_ID_OPORTUNIDADE, ALIASES_CONTRATO_GERADO)


def deduplicar_pastas_funil(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por Nome da Avaliação (Data Primeiro Envio Análise mais recente)."""
    col_envio = achar_coluna_primeiro_envio_analise(df)
    if not col_envio:
        return df if df is not None else pd.DataFrame()
    return deduplicar_por_chave_mais_recente(
        df, ALIASES_NOME_AVALIACAO_CREDITO, [col_envio]
    )


def deduplicar_agendamentos_funil(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por Código do agendamento (Data de criação mais recente)."""
    return deduplicar_por_chave_mais_recente(
        df, ALIASES_CODIGO_AGENDAMENTO, ALIASES_DATA_CRIACAO
    )


def montar_mapa_funil_diario(
    df_ag_vis: pd.DataFrame,
    df_pastas: pd.DataFrame,
    serie_vendas: Optional[pd.DataFrame] = None,
    df_vendas: Optional[pd.DataFrame] = None,
) -> Dict[str, Dict[date, float]]:
    """
    Séries diárias do funil:
      agendamentos ← Data de criação
      visitas ← Data da visita
      pastas ← Data Primeiro Envio Análise (relatório SF 00OU600000FEOoDMAX)
      pastas_aprovadas ← Data Aprovação SAFI
      vendas ← Contrato gerado em, somente vendas comerciais

    Deduplicação antes da contagem:
      vendas → ID da Oportunidade (Contrato gerado em mais recente)
      pastas → Nome da Avaliação (Data Primeiro Envio Análise mais recente)
      pastas aprovadas → Nome da Avaliação (Data Aprovação SAFI mais recente)
      agendamentos → Código do agendamento (Data de criação mais recente)
    """
    df_ag = deduplicar_agendamentos_funil(
        df_ag_vis if df_ag_vis is not None else pd.DataFrame()
    )
    df_pas_raw = df_pastas if df_pastas is not None else pd.DataFrame()
    df_pas = deduplicar_pastas_funil(df_pas_raw)
    df_pas_aprov = deduplicar_pastas_aprovadas_funil(df_pas_raw)

    mapa_vendas: Dict[date, float] = {}
    if df_vendas is not None and not df_vendas.empty:
        df_ven = deduplicar_vendas_funil(filtrar_vendas_comerciais(df_vendas))
        mapa_vendas = contar_eventos_por_dia(df_ven, ALIASES_CONTRATO_GERADO)
    elif serie_vendas is not None and not serie_vendas.empty:
        for _, r in serie_vendas.iterrows():
            mapa_vendas[r["data"]] = float(r["qtd"])

    col_envio = achar_coluna_primeiro_envio_analise(df_pas)
    col_safi_aprov = achar_coluna_aprovacao_safi(df_pas_aprov)
    mapa_pastas = (
        contar_eventos_por_coluna(df_pas, col_envio) if col_envio else {}
    )
    mapa_aprov = (
        contar_eventos_por_coluna(df_pas_aprov, col_safi_aprov) if col_safi_aprov else {}
    )

    return {
        "agendamentos": contar_eventos_por_dia(df_ag, ALIASES_DATA_CRIACAO),
        "visitas": contar_eventos_por_dia(
            df_ag,
            [
                "Data da visita", "Data da Visita", "Data visita", "Data Visita",
                "Activity Date", "Data da Atividade", "Data do agendamento",
                "Data Agendamento", "Start Date Time", "Data/Hora",
            ],
        ),
        "pastas": mapa_pastas,
        "pastas_aprovadas": mapa_aprov,
        "vendas": mapa_vendas,
    }


def calendario_funil_diario(
    inicio: date,
    fim: date,
    mapas: Dict[str, Dict[date, float]],
    lags: Tuple[int, ...] = FUNIL_LAGS,
) -> pd.DataFrame:
    """Calendário diário do funil com dummies, lags, conversões e força de trabalho."""
    idx = pd.date_range(inicio, fim, freq="D")
    cal = pd.DataFrame({"data": [d.date() for d in idx]})
    for etapa in FUNIL_ETAPAS:
        m = mapas.get(etapa) or {}
        cal[etapa] = cal["data"].map(lambda d, _m=m: float(_m.get(d, 0.0)))
    cal["dia_mes"] = cal["data"].map(lambda d: d.day)
    cal["dia_semana"] = cal["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    cal["mes"] = cal["data"].map(lambda d: MESES_PT[d.month])
    for etapa in FUNIL_ETAPAS:
        for L in lags:
            cal[f"{etapa}_lag{L}"] = cal[etapa].shift(L).fillna(0.0)
    cal = adicionar_conversoes_funil(cal)
    return adicionar_forca_trabalho(cal)


def adicionar_forca_trabalho(
    cal: pd.DataFrame,
    janela: int = FUNIL_JANELA_FORCA,
) -> pd.DataFrame:
    """
    Indicador de força de trabalho: atividade recente em TODAS as etapas do funil.
    Captura que o time trabalha em paralelo (pastas podem 'vazar' para visitas etc.).
    Usa média móvel das etapas (drivers + vendas) deslocada 1 dia (sem vazamento do dia).
    """
    out = cal.copy()
    partes: List[pd.Series] = []
    for e in FUNIL_ETAPAS:
        if e not in out.columns:
            continue
        s = pd.to_numeric(out[e], errors="coerce").fillna(0.0).astype(float)
        mu = float(s.mean()) if len(s) else 0.0
        sd = float(s.std()) if len(s) else 0.0
        if sd < 1e-9:
            z = s * 0.0
        else:
            z = (s - mu) / sd
        partes.append(z)
    if not partes:
        out["forca_trabalho"] = 0.0
        out["atividade_bruta"] = 0.0
        return out
    z_mean = sum(partes) / float(len(partes))
    cols_drv = [e for e in FUNIL_DRIVERS if e in out.columns]
    out["atividade_bruta"] = out[cols_drv].sum(axis=1) if cols_drv else 0.0
    # força = média móvel da atividade padronizada, sem o dia corrente
    out["forca_trabalho"] = (
        z_mean.rolling(janela, min_periods=1).mean().shift(1).fillna(0.0)
    )
    return out


def _atualizar_forca_trabalho_linha(
    cal: pd.DataFrame,
    i: int,
    janela: int = FUNIL_JANELA_FORCA,
) -> None:
    """Recalcula força de trabalho na linha i a partir do histórico até i-1."""
    try:
        pos = int(cal.index.get_loc(i))
    except Exception:
        pos = int(i)
    i0 = max(0, pos - janela)
    hist = cal.iloc[i0:pos]
    if hist.empty:
        cal.at[i, "forca_trabalho"] = 0.0
        cal.at[i, "atividade_bruta"] = 0.0
        return
    z_vals: List[float] = []
    for e in FUNIL_ETAPAS:
        if e not in hist.columns:
            continue
        s_all = pd.to_numeric(cal[e], errors="coerce").fillna(0.0)
        mu = float(s_all.mean()) if len(s_all) else 0.0
        sd = float(s_all.std()) if len(s_all) else 0.0
        s_h = pd.to_numeric(hist[e], errors="coerce").fillna(0.0)
        if sd < 1e-9:
            z_vals.append(0.0)
        else:
            z_vals.append(float(((s_h - mu) / sd).mean()))
    cal.at[i, "forca_trabalho"] = float(np.mean(z_vals)) if z_vals else 0.0
    cal.at[i, "atividade_bruta"] = float(
        sum(float(cal.at[i, e]) for e in FUNIL_DRIVERS if e in cal.columns)
    )


def _safe_ratio(num: pd.Series, den: pd.Series) -> pd.Series:
    """Razão segura; 0 quando denominador ≤ 0. Clip para estabilidade do modelo."""
    n = pd.to_numeric(num, errors="coerce").fillna(0.0).astype(float)
    d = pd.to_numeric(den, errors="coerce").fillna(0.0).astype(float)
    out = np.where(d > 0, n / d, 0.0)
    return pd.Series(np.clip(out, 0.0, 5.0), index=num.index)


def col_conv_etapa(origem: str, destino: str) -> str:
    return f"conv_{origem}_{destino}"


def col_conv_venda(origem: str) -> str:
    return f"conv_{origem}_vendas"


def cols_conversoes_funil() -> List[str]:
    """Colunas de conversão usadas no modelo de vendas."""
    cols = [col_conv_etapa(a, b) for a, b in FUNIL_PARES_ETAPA if b != "vendas"]
    cols += [col_conv_venda(e) for e in FUNIL_DRIVERS]
    return cols


def adicionar_conversoes_funil(
    cal: pd.DataFrame,
    janela: int = FUNIL_JANELA_CONV,
) -> pd.DataFrame:
    """
    Conversões diárias:
      - etapa→seguinte: razão no mesmo dia (sem usar vendas)
      - indicador→venda: razão móvel (janela) deslocada 1 dia (evita vazamento do y)
    """
    out = cal.copy()
    for a, b in FUNIL_PARES_ETAPA:
        if b == "vendas":
            continue
        if a in out.columns and b in out.columns:
            out[col_conv_etapa(a, b)] = _safe_ratio(out[b], out[a])

    for etapa in FUNIL_DRIVERS:
        if etapa not in out.columns or "vendas" not in out.columns:
            out[col_conv_venda(etapa)] = 0.0
            continue
        roll_v = out["vendas"].rolling(janela, min_periods=5).sum()
        roll_e = out[etapa].rolling(janela, min_periods=5).sum()
        out[col_conv_venda(etapa)] = _safe_ratio(roll_v, roll_e).shift(1).fillna(0.0)
    return out


def _atualizar_conversoes_linha(
    cal: pd.DataFrame,
    i: int,
    janela: int = FUNIL_JANELA_CONV,
) -> None:
    """Recalcula conversões na linha i (após atualizar níveis do funil)."""
    for a, b in FUNIL_PARES_ETAPA:
        if b == "vendas":
            continue
        den = float(cal.at[i, a]) if a in cal.columns else 0.0
        num = float(cal.at[i, b]) if b in cal.columns else 0.0
        cal.at[i, col_conv_etapa(a, b)] = float(np.clip(num / den, 0.0, 5.0)) if den > 0 else 0.0

    # Posição inteira no índice (RangeIndex 0..n-1)
    try:
        pos = int(cal.index.get_loc(i))
    except Exception:
        pos = int(i)
    i0 = max(0, pos - janela)
    hist = cal.iloc[i0:pos]
    if hist.empty or "vendas" not in cal.columns:
        for etapa in FUNIL_DRIVERS:
            cal.at[i, col_conv_venda(etapa)] = 0.0
        return
    sum_v = float(hist["vendas"].sum())
    for etapa in FUNIL_DRIVERS:
        sum_e = float(hist[etapa].sum()) if etapa in hist.columns else 0.0
        cal.at[i, col_conv_venda(etapa)] = (
            float(np.clip(sum_v / sum_e, 0.0, 5.0)) if sum_e > 0 else 0.0
        )


def taxa_conversao(origem: float, destino: float) -> Optional[float]:
    """Taxa percentual origem→destino; None se origem ≤ 0."""
    if origem is None or destino is None:
        return None
    o = float(origem)
    if o <= 0:
        return None
    return 100.0 * float(destino) / o


def calcular_conversoes_totais(totais: Dict[str, float]) -> Dict[str, Any]:
    """Conversões a partir de totais (MTD ou projetado do mês)."""
    etapa_a_etapa: List[Dict[str, Any]] = []
    for a, b in FUNIL_PARES_ETAPA:
        taxa = taxa_conversao(totais.get(a, 0.0), totais.get(b, 0.0))
        etapa_a_etapa.append({
            "origem": a,
            "destino": b,
            "label": f"{FUNIL_LABELS.get(a, a)} → {FUNIL_LABELS.get(b, b)}",
            "taxa": taxa,
            "origem_qtd": float(totais.get(a, 0.0)),
            "destino_qtd": float(totais.get(b, 0.0)),
        })
    para_venda: List[Dict[str, Any]] = []
    for a in FUNIL_DRIVERS:
        taxa = taxa_conversao(totais.get(a, 0.0), totais.get("vendas", 0.0))
        para_venda.append({
            "origem": a,
            "destino": "vendas",
            "label": f"{FUNIL_LABELS.get(a, a)} → Vendas",
            "taxa": taxa,
            "origem_qtd": float(totais.get(a, 0.0)),
            "destino_qtd": float(totais.get("vendas", 0.0)),
        })
    return {"etapa_a_etapa": etapa_a_etapa, "para_venda": para_venda}


def _funil_gap_vendas(
    gap_vendas: float,
    taxas_hist_frac: Dict[str, Optional[float]],
    totais_hist: Dict[str, float],
    gap_vendas_mes: Optional[float] = None,
    funil_necessario: Optional[Dict[str, float]] = None,
    totais_mtd: Optional[Dict[str, float]] = None,
) -> Dict[str, float]:
    """
    Funil por etapa a partir de um gap de vendas (restante para a meta).
    Usa conversão histórica indicador→venda; fallback proporcional ao gap do mês.
    """
    gap_v = max(0.0, float(gap_vendas))
    gap_mes = max(0.0, float(gap_vendas_mes if gap_vendas_mes is not None else gap_v))
    funil: Dict[str, float] = {"vendas": gap_v}
    for etapa in FUNIL_DRIVERS:
        t = taxas_hist_frac.get(etapa)
        if gap_v > 0 and t is not None and t > 1e-9:
            funil[etapa] = float(math.ceil(gap_v / t))
            continue
        v_h = float(totais_hist.get("vendas", 0.0))
        i_h = float(totais_hist.get(etapa, 0.0))
        if gap_v > 0 and v_h > 0 and i_h > 0:
            funil[etapa] = float(math.ceil(gap_v * (i_h / v_h)))
            continue
        if (
            gap_v > 0
            and gap_mes > 1e-9
            and funil_necessario is not None
            and totais_mtd is not None
        ):
            gap_etapa_mes = max(
                0.0,
                float(funil_necessario.get(etapa, 0.0)) - float(totais_mtd.get(etapa, 0.0)),
            )
            funil[etapa] = float(math.ceil(gap_etapa_mes * (gap_v / gap_mes)))
        else:
            funil[etapa] = 0.0
    return funil


def _pesos_distribuicao_gap(
    dias_distrib: List[date],
    resultados: Dict[str, Any],
) -> np.ndarray:
    """Pesos diários para distribuir o gap de vendas (projeção reg. de vendas)."""
    df_ven = (resultados.get("vendas") or {}).get("diaria", pd.DataFrame())
    pesos: List[float] = []
    for d in dias_distrib:
        p = 1.0
        if df_ven is not None and not df_ven.empty and "dia" in df_ven.columns:
            row = df_ven.loc[df_ven["dia"] == d.day]
            if not row.empty and "projetado_reg" in row.columns:
                p = max(float(row["projetado_reg"].iloc[0]), 0.0)
        pesos.append(p if p > 0 else 1.0)
    return np.asarray(pesos, dtype=float)


def _meta_vendas_por_pesos(
    gap_vendas: float,
    dias_alvo: List[date],
    dias_distrib: List[date],
    pesos: np.ndarray,
) -> float:
    """Parcela do gap de vendas alocada a um conjunto de dias (pesos normalizados)."""
    if gap_vendas <= 0 or not dias_alvo or not dias_distrib:
        return 0.0
    w = np.maximum(np.asarray(pesos, dtype=float), 0.0)
    soma = float(w.sum())
    if soma <= 1e-9:
        w = np.ones(len(dias_distrib), dtype=float)
        soma = float(len(dias_distrib))
    mapa = {dias_distrib[i]: float(w[i] / soma) for i in range(len(dias_distrib))}
    return float(gap_vendas) * sum(mapa.get(d, 0.0) for d in dias_alvo)


def _razoes_entre_funis(
    numerador: Dict[str, float],
    denominador: Dict[str, float],
) -> Dict[str, Optional[float]]:
    """Razão numerador/denominador por etapa (None se denominador ≤ 0)."""
    out: Dict[str, Optional[float]] = {}
    for etapa in FUNIL_ETAPAS:
        den = float(denominador.get(etapa, 0.0))
        num = float(numerador.get(etapa, 0.0))
        out[etapa] = (num / den) if den > 1e-9 else None
    return out


def _somar_funis(*funis: Dict[str, float]) -> Dict[str, float]:
    out: Dict[str, float] = {e: 0.0 for e in FUNIL_ETAPAS}
    for f in funis:
        for e in FUNIL_ETAPAS:
            out[e] += float((f or {}).get(e, 0.0))
    return out


def _domingo_semana_iso(d: date) -> date:
    """Último dia (domingo) da semana ISO que contém d."""
    return d + timedelta(days=(7 - d.isoweekday()))


def _cols_lag_funil(
    lags: Tuple[int, ...] = FUNIL_LAGS,
    alvo: Optional[str] = None,
    etapas: Optional[Tuple[str, ...]] = None,
) -> List[str]:
    """Lags das etapas; exclui lag 0 do próprio alvo (colinear com y)."""
    etapas_uso = etapas if etapas is not None else FUNIL_ETAPAS
    cols: List[str] = []
    for e in etapas_uso:
        for L in lags:
            if L == 0 and alvo is not None and e == alvo:
                continue
            cols.append(f"{e}_lag{L}")
    return cols


def _matriz_funil_explicativas(
    df: pd.DataFrame,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
    alvo: Optional[str] = None,
    etapas_lag: Optional[Tuple[str, ...]] = None,
    modelo_vendas_completo: bool = False,
    efeitos_cruzados: bool = True,
) -> Tuple[np.ndarray, List[str]]:
    """
    Calendário + lags de todas as etapas + efeitos cruzados contemporâneos
    + força de trabalho (+ conversões no modelo de vendas).

    Efeitos cruzados: o trabalho 'vaza' entre etapas (ex.: pastas influenciam
    visitas), não só a cascata agendamentos→visitas→pastas→aprovadas→vendas.
    """
    X_cal = _matriz_explicativas(df, incluir_mes=incluir_mes)
    lag_etapas = etapas_lag if etapas_lag is not None else FUNIL_ETAPAS
    lag_cols = _cols_lag_funil(lags, alvo=alvo, etapas=lag_etapas)

    extra_cols: List[str] = list(lag_cols)

    if efeitos_cruzados or modelo_vendas_completo:
        # Níveis contemporâneos das OUTRAS etapas (não o próprio alvo)
        for e in FUNIL_ETAPAS:
            if e == alvo:
                continue
            if e in df.columns:
                extra_cols.append(e)
        if "forca_trabalho" in df.columns:
            extra_cols.append("forca_trabalho")
        if modelo_vendas_completo or alvo == "vendas":
            for c in cols_conversoes_funil():
                if c in df.columns:
                    extra_cols.append(c)

    # remove duplicatas preservando ordem
    seen = set()
    extra_unique: List[str] = []
    for c in extra_cols:
        if c not in seen:
            seen.add(c)
            extra_unique.append(c)
    extra_cols = extra_unique

    X_extra = np.zeros((len(df), len(extra_cols)), dtype=float)
    for j, c in enumerate(extra_cols):
        if c in df.columns:
            X_extra[:, j] = pd.to_numeric(df[c], errors="coerce").fillna(0.0).values
    X = np.hstack([X_cal[:, :-1], X_extra, X_cal[:, -1:]])
    return X, extra_cols


def treinar_regressao_funil(
    treino: pd.DataFrame,
    alvo: str,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
) -> np.ndarray:
    X, _ = _matriz_funil_explicativas(
        treino,
        incluir_mes=incluir_mes,
        lags=lags,
        alvo=alvo,
        modelo_vendas_completo=(alvo == "vendas"),
        efeitos_cruzados=True,
    )
    y = treino[alvo].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    return coef


def _r2_funil(
    treino: pd.DataFrame,
    coef: np.ndarray,
    alvo: str,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
) -> float:
    X, _ = _matriz_funil_explicativas(
        treino,
        incluir_mes=incluir_mes,
        lags=lags,
        alvo=alvo,
        modelo_vendas_completo=(alvo == "vendas"),
        efeitos_cruzados=True,
    )
    y = treino[alvo].astype(float).values
    y_hat = X @ coef
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    if ss_tot <= 0:
        return 0.0
    return 1.0 - (ss_res / ss_tot)


def calcular_medias_funil(
    treino: pd.DataFrame,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
) -> Dict[str, Any]:
    """Médias sazonais por etapa + lags cruzados + força de trabalho."""
    forca_mu = (
        float(treino["forca_trabalho"].mean())
        if "forca_trabalho" in treino.columns and len(treino)
        else 1.0
    )
    if abs(forca_mu) < 1e-9:
        forca_mu = 1e-9
    out: Dict[str, Any] = {
        "incluir_mes": incluir_mes,
        "lags": lags,
        "forca_mu": forca_mu,
        "etapas": {},
    }
    for etapa in FUNIL_ETAPAS:
        mu = float(treino[etapa].mean()) if len(treino) else 0.0
        if mu <= 0:
            mu = 1e-9
        etapa_info: Dict[str, Any] = {
            "mu": mu,
            "media_dia_semana": {
                k: float(v) for k, v in treino.groupby("dia_semana")[etapa].mean().items()
            },
            "media_dia_mes": {
                int(k): float(v) for k, v in treino.groupby("dia_mes")[etapa].mean().items()
            },
        }
        if incluir_mes:
            etapa_info["media_mes"] = {
                k: float(v) for k, v in treino.groupby("mes")[etapa].mean().items()
            }
        # lags de TODAS as etapas (efeitos cruzados na intensidade)
        lag_mu: Dict[str, float] = {}
        for e2 in FUNIL_ETAPAS:
            for L in lags:
                col = f"{e2}_lag{L}"
                if col in treino.columns:
                    m = float(treino[col].mean())
                    lag_mu[col] = m if m > 0 else mu
        etapa_info["lag_mu"] = lag_mu
        out["etapas"][etapa] = etapa_info
    return out


def _pred_sazonal_etapa(d: date, info: Dict[str, Any], incluir_mes: bool) -> float:
    mu = float(info["mu"])
    a = float(info["media_dia_semana"].get(DIAS_SEMANA_PT[d.weekday()], mu))
    b = float(info["media_dia_mes"].get(d.day, mu))
    if incluir_mes and info.get("media_mes"):
        c = float(info["media_mes"].get(MESES_PT[d.month], mu))
        return max((a * b * c) / (mu * mu), 0.0)
    return max((a * b) / mu, 0.0)


def _intensidade_lags_linha(row: pd.Series, info: Dict[str, Any]) -> float:
    """Fator multiplicativo pela intensidade dos lags (próprios + cruzados)."""
    lag_mu = info.get("lag_mu") or {}
    if not lag_mu:
        return 1.0
    ratios: List[float] = []
    for col, mu_l in lag_mu.items():
        v = float(row.get(col, 0.0) or 0.0)
        mu_l = float(mu_l) if float(mu_l) > 1e-9 else 1e-9
        ratios.append(v / mu_l)
    if not ratios:
        return 1.0
    ratios = [max(r, 0.05) for r in ratios]
    geo = float(np.exp(np.mean(np.log(ratios))))
    return float(np.clip(geo, 0.25, 3.0))


def _intensidade_forca_linha(row: pd.Series, medias: Dict[str, Any]) -> float:
    """Fator pela força de trabalho vs média histórica."""
    mu = float(medias.get("forca_mu") or 0.0)
    v = float(row.get("forca_trabalho", 0.0) or 0.0)
    fat = 1.0 + 0.35 * (v - mu)
    return float(np.clip(fat, 0.35, 2.5))


def _prever_linha_reg_funil(
    coef: np.ndarray,
    row_df: pd.DataFrame,
    incluir_mes: bool,
    lags: Tuple[int, ...],
    alvo: str,
) -> float:
    X, _ = _matriz_funil_explicativas(
        row_df,
        incluir_mes=incluir_mes,
        lags=lags,
        alvo=alvo,
        modelo_vendas_completo=(alvo == "vendas"),
        efeitos_cruzados=True,
    )
    return float(max((X @ coef)[0], 0.0))


def _prever_linha_medias_funil(
    row: pd.Series,
    d: date,
    etapa: str,
    medias: Dict[str, Any],
    incluir_mes: bool,
) -> float:
    info = medias["etapas"][etapa]
    saz = _pred_sazonal_etapa(d, info, incluir_mes)
    intens = _intensidade_lags_linha(row, info)
    forca = _intensidade_forca_linha(row, medias)
    return max(saz * intens * forca, 0.0)


def _r2_medias_funil(
    treino: pd.DataFrame,
    medias: Dict[str, Any],
    alvo: str,
    incluir_mes: bool = True,
) -> float:
    if treino.empty or alvo not in treino.columns:
        return 0.0
    y = treino[alvo].astype(float).values
    y_hat = np.array([
        _prever_linha_medias_funil(row, row["data"], alvo, medias, incluir_mes)
        for _, row in treino.iterrows()
    ], dtype=float)
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    if ss_tot <= 0:
        return 0.0
    return 1.0 - (ss_res / ss_tot)


def _atualizar_lags_linha(cal: pd.DataFrame, i: int, lags: Tuple[int, ...]) -> None:
    for etapa in FUNIL_ETAPAS:
        for L in lags:
            if L == 0:
                cal.at[i, f"{etapa}_lag0"] = float(cal.at[i, etapa])
            else:
                j = i - L
                cal.at[i, f"{etapa}_lag{L}"] = float(cal.at[j, etapa]) if j >= 0 else 0.0


def _resumir_perfil_lag_vendas(df_p: pd.DataFrame, etapa: str) -> Dict[str, Any]:
    """
    Resume o perfil de lags 1..30 de uma etapa sobre vendas.
    Pico = lag com maior efeito acumulado até aquele dia (cumsum dos coeficientes).
    Meia-vida ≈ 50% do efeito acumulado total no horizonte 1..30d.
    """
    df_p = df_p[df_p["lag"] >= 1].copy()
    if df_p.empty:
        return {
            "etapa": etapa,
            "label": FUNIL_LABELS.get(etapa, etapa),
            "lag_pico": 1,
            "efeito_pico": 0.0,
            "lag_meia_vida": 1,
            "efeito_lag1": 0.0,
            "efeito_acum": 0.0,
        }

    df_p = df_p.sort_values("lag").reset_index(drop=True)
    df_p["acumulado"] = df_p["efeito"].cumsum()

    i_peak = int(df_p["acumulado"].idxmax())
    lag_pico = int(df_p.loc[i_peak, "lag"])
    efeito_pico = float(df_p.loc[i_peak, "acumulado"])

    acum_final = float(df_p["acumulado"].iloc[-1])
    lag_meia = lag_pico
    if abs(acum_final) > 1e-9:
        alvo = 0.5 * acum_final
        lag_meia = int(df_p["lag"].iloc[-1])
        for _, r in df_p.iterrows():
            ac = float(r["acumulado"])
            if acum_final > 0 and ac >= alvo:
                lag_meia = int(r["lag"])
                break
            if acum_final < 0 and ac <= alvo:
                lag_meia = int(r["lag"])
                break
    else:
        lag_meia = max(1, lag_pico)

    efeito_lag1 = float(df_p.loc[df_p["lag"] == 1, "efeito"].iloc[0]) if (df_p["lag"] == 1).any() else 0.0
    return {
        "etapa": etapa,
        "label": FUNIL_LABELS.get(etapa, etapa),
        "lag_pico": lag_pico,
        "efeito_pico": efeito_pico,
        "lag_meia_vida": lag_meia,
        "efeito_lag1": efeito_lag1,
        "efeito_acum": acum_final,
    }


def estimar_efeitos_lags_sobre_vendas(
    treino: pd.DataFrame,
    lags: Tuple[int, ...] = FUNIL_LAGS_PERFIL,
    incluir_mes: bool = True,
) -> Optional[Dict[str, Any]]:
    """
    Perfil de tempo até efeito nas vendas (lags 1..30).

    Um modelo OLS por etapa: vendas ~ calendário + lags 1..30 da etapa.
    Evita que lag 0 / multicolinearidade entre etapas zere o tempo de efeito.
    """
    if treino.empty or float(treino["vendas"].sum()) <= 0 or len(treino) < 40:
        return None

    lags_uso = tuple(L for L in lags if L >= 1)
    if not lags_uso:
        lags_uso = FUNIL_LAGS_PERFIL

    cal = treino.copy()
    for etapa in FUNIL_DRIVERS:
        if etapa not in cal.columns:
            return None
        for L in lags_uso:
            col = f"{etapa}_lag{L}"
            if col not in cal.columns:
                cal[col] = cal[etapa].shift(L).fillna(0.0)

    y = cal["vendas"].astype(float).values
    perfis: Dict[str, pd.DataFrame] = {}
    resumo: List[Dict[str, Any]] = []
    r2s_etapa: List[float] = []

    for etapa in FUNIL_DRIVERS:
        lag_cols = [f"{etapa}_lag{L}" for L in lags_uso]
        X_cal = _matriz_explicativas(cal, incluir_mes=incluir_mes)
        X_lag = np.zeros((len(cal), len(lag_cols)), dtype=float)
        for j, c in enumerate(lag_cols):
            X_lag[:, j] = pd.to_numeric(cal[c], errors="coerce").fillna(0.0).values
        X = np.hstack([X_cal[:, :-1], X_lag, X_cal[:, -1:]])

        coef, *_ = np.linalg.lstsq(X, y, rcond=None)
        n_cal = X_cal.shape[1] - 1
        coef_lags = coef[n_cal:n_cal + len(lag_cols)]

        efeitos = [float(coef_lags[j]) for j in range(len(lag_cols))]
        df_p = pd.DataFrame({"lag": list(lags_uso), "efeito": efeitos})
        df_p["acumulado"] = df_p["efeito"].cumsum()
        perfis[etapa] = df_p
        resumo.append(_resumir_perfil_lag_vendas(df_p, etapa))

        y_hat = X @ coef
        ss_res = float(np.sum((y - y_hat) ** 2))
        ss_tot = float(np.sum((y - y.mean()) ** 2))
        r2s_etapa.append((1.0 - ss_res / ss_tot) if ss_tot > 0 else 0.0)

    r2 = float(np.mean(r2s_etapa)) if r2s_etapa else 0.0

    return {
        "r2": r2,
        "r2s_etapa": {e: r2s_etapa[i] for i, e in enumerate(FUNIL_DRIVERS)},
        "lags": lags_uso,
        "perfis": perfis,
        "resumo": resumo,
    }


def projetar_funil_mes_atual(
    mapas: Dict[str, Dict[date, float]],
    hoje: Optional[date] = None,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
    meta_qtd_mes: float = 0.0,
) -> Optional[Dict[str, Any]]:
    """
    Projeta cada etapa do funil no mês corrente.
    Modelos (reg e médias) com efeitos cruzados: calendário + lags de todas as
    etapas + níveis contemporâneos das outras etapas + força de trabalho
    (+ conversões no modelo de vendas). O trabalho 'vaza' entre etapas.
    meta_qtd_mes: meta de vendas do mês (para funil necessário e metas diárias).
    """
    hoje = hoje or date.today()
    inicio, fim_treino = janela_treino_meses_exatos(hoje)
    max_lag = max(lags) if lags else 0
    # buffer de lags antes do treino
    inicio_cal = inicio - timedelta(days=max(max_lag, FUNIL_JANELA_CONV, FUNIL_JANELA_FORCA))

    ano, mes = hoje.year, hoje.month
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    fim_mes = date(ano, mes, ultimo_dia)
    # Se a semana ISO atravessa o fim do mês, estende o calendário até o domingo
    # para completar a meta da semana com o projetado do mês seguinte.
    fim_semana = _domingo_semana_iso(hoje)
    fim_cal = max(fim_mes, fim_semana)

    cal = calendario_funil_diario(inicio_cal, fim_cal, mapas, lags=lags)
    cal = cal.reset_index(drop=True)
    cal["data"] = cal["data"].map(_as_date_funil)
    cal = cal.dropna(subset=["data"]).reset_index(drop=True)
    if cal.empty or cal.loc[cal["data"].between(inicio, fim_treino), "vendas"].sum() <= 0:
        # ainda permite se houver outras etapas
        if cal.loc[cal["data"].between(inicio, fim_treino), list(FUNIL_ETAPAS)].sum().sum() <= 0:
            return None

    mask_treino = (cal["data"] >= inicio) & (cal["data"] <= fim_treino)
    treino = cal.loc[mask_treino].copy()
    if len(treino) < 30:
        return None

    medias = calcular_medias_funil(treino, incluir_mes=incluir_mes, lags=lags)
    coefs: Dict[str, np.ndarray] = {}
    r2s: Dict[str, float] = {}
    r2s_medias: Dict[str, float] = {}
    for etapa in FUNIL_ETAPAS:
        coefs[etapa] = treinar_regressao_funil(treino, etapa, incluir_mes=incluir_mes, lags=lags)
        r2s[etapa] = _r2_funil(treino, coefs[etapa], etapa, incluir_mes=incluir_mes, lags=lags)
        r2s_medias[etapa] = _r2_medias_funil(treino, medias, etapa, incluir_mes=incluir_mes)

    # Cópias para projeção (reg e médias)
    cal_reg = cal.copy()
    cal_med = cal.copy()

    dias_mes = [date(ano, mes, d) for d in range(1, ultimo_dia + 1)]
    # Projeta dias futuros do mês atual + dias da semana no mês seguinte (se houver)
    dias_a_projetar = []
    d_cursor = hoje + timedelta(days=1)
    while d_cursor <= fim_cal:
        dias_a_projetar.append(d_cursor)
        d_cursor += timedelta(days=1)
    idx_por_data = _indice_por_data_cal(cal_reg)

    for d in dias_a_projetar:
        i = idx_por_data.get(d)
        if i is None:
            continue
        _atualizar_lags_linha(cal_reg, i, lags)
        _atualizar_lags_linha(cal_med, i, lags)
        _atualizar_forca_trabalho_linha(cal_reg, i)
        _atualizar_forca_trabalho_linha(cal_med, i)

        # Iterações: efeitos contemporâneos cruzados (trabalho paralelo)
        for _ in range(FUNIL_ITERS_CRUZADAS):
            for etapa in FUNIL_DRIVERS:
                row_reg = cal_reg.loc[[i]]
                row_med = cal_med.loc[[i]]
                pred_r = _prever_linha_reg_funil(coefs[etapa], row_reg, incluir_mes, lags, etapa)
                pred_m = _prever_linha_medias_funil(
                    row_med.iloc[0], d, etapa, medias, incluir_mes
                )
                cal_reg.at[i, etapa] = pred_r
                cal_med.at[i, etapa] = pred_m
                if 0 in lags:
                    cal_reg.at[i, f"{etapa}_lag0"] = pred_r
                    cal_med.at[i, f"{etapa}_lag0"] = pred_m
            _atualizar_conversoes_linha(cal_reg, i)
            _atualizar_conversoes_linha(cal_med, i)
            # vendas no fim de cada iteração (feedback cruzado)
            row_reg = cal_reg.loc[[i]]
            row_med = cal_med.loc[[i]]
            pred_r = _prever_linha_reg_funil(coefs["vendas"], row_reg, incluir_mes, lags, "vendas")
            pred_m = _prever_linha_medias_funil(
                row_med.iloc[0], d, "vendas", medias, incluir_mes
            )
            cal_reg.at[i, "vendas"] = pred_r
            cal_med.at[i, "vendas"] = pred_m
            if 0 in lags:
                cal_reg.at[i, "vendas_lag0"] = pred_r
                cal_med.at[i, "vendas_lag0"] = pred_m

    dias_futuros_mes = [d for d in dias_mes if d > hoje]
    _garantir_previsoes_futuras_funil(
        cal_reg, cal_med, idx_por_data, dias_futuros_mes, treino, medias, coefs, incluir_mes, lags
    )

    # montar resultados por etapa
    resultados: Dict[str, Any] = {}
    for etapa in FUNIL_ETAPAS:
        mtd = 0.0
        proj_mtd_reg = 0.0
        proj_mtd_med = 0.0
        rest_reg = 0.0
        rest_med = 0.0
        diaria: List[Dict[str, Any]] = []
        for d in dias_mes:
            i = idx_por_data.get(d)
            if i is None:
                continue
            real = float((mapas.get(etapa) or {}).get(d, 0.0)) if d <= hoje else None
            if d <= hoje:
                mtd += real or 0.0
                row_reg = cal_reg.loc[[i]]
                pred_reg_dia = _prever_linha_reg_funil(
                    coefs[etapa], row_reg, incluir_mes, lags, etapa
                )
                pred_med_dia = _prever_linha_medias_funil(
                    cal_med.iloc[i], d, etapa, medias, incluir_mes
                )
                proj_mtd_reg += pred_reg_dia
                proj_mtd_med += pred_med_dia
            else:
                pred_reg_dia = float(cal_reg.at[i, etapa])
                pred_med_dia = float(cal_med.at[i, etapa])
                rest_reg += pred_reg_dia
                rest_med += pred_med_dia
            diaria.append({
                "dia": d.day,
                "realizado": real,
                "projetado_reg": pred_reg_dia,
                "projetado_med": pred_med_dia,
            })
        resultados[etapa] = {
            "mtd": mtd,
            "projetado_mtd_reg": proj_mtd_reg,
            "projetado_mtd_med": proj_mtd_med,
            "projetado_reg": mtd + rest_reg,
            "projetado_med": mtd + rest_med,
            "r2": r2s[etapa],
            "r2_medias": r2s_medias[etapa],
            "diaria": pd.DataFrame(diaria),
        }

    totais_mtd = {e: float((resultados.get(e) or {}).get("mtd", 0)) for e in FUNIL_ETAPAS}
    totais_proj_mtd = {
        e: float((resultados.get(e) or {}).get("projetado_mtd_reg", 0)) for e in FUNIL_ETAPAS
    }
    totais_proj = {e: float((resultados.get(e) or {}).get("projetado_reg", 0)) for e in FUNIL_ETAPAS}
    totais_hist = {e: float(treino[e].sum()) for e in FUNIL_ETAPAS}

    # Conversões do mês (MTD / projetado) × histórico (treino: sem mês atual, até 1 ano)
    conversoes = {
        "realizado_mtd": calcular_conversoes_totais(totais_mtd),
        "projetado_mtd": calcular_conversoes_totais(totais_proj_mtd),
        "projetado_mes": calcular_conversoes_totais(totais_proj),
        "historico": calcular_conversoes_totais(totais_hist),
        "inicio_hist": inicio,
        "fim_hist": fim_treino,
    }

    # Funil necessário para bater a meta de vendas (via conversões históricas indicador→venda)
    meta_qtd = float(meta_qtd_mes or 0.0)
    gap_vendas = max(0.0, meta_qtd - totais_mtd.get("vendas", 0.0))
    taxas_hist_frac: Dict[str, Optional[float]] = {}
    for item in (conversoes["historico"].get("para_venda") or []):
        orig = str(item.get("origem") or "")
        taxa = item.get("taxa")
        taxas_hist_frac[orig] = (float(taxa) / 100.0) if taxa is not None else None

    funil_necessario: Dict[str, float] = {"vendas": meta_qtd if meta_qtd > 0 else totais_mtd.get("vendas", 0.0)}
    for etapa in FUNIL_DRIVERS:
        t = taxas_hist_frac.get(etapa)
        if meta_qtd > 0 and t is not None and t > 1e-9:
            funil_necessario[etapa] = float(math.ceil(meta_qtd / t))
        else:
            # fallback: proporção histórica vs vendas
            v_h = totais_hist.get("vendas", 0.0)
            i_h = totais_hist.get(etapa, 0.0)
            if meta_qtd > 0 and v_h > 0 and i_h > 0:
                funil_necessario[etapa] = float(math.ceil(meta_qtd * (i_h / v_h)))
            else:
                funil_necessario[etapa] = float(totais_proj.get(etapa, 0.0))

    # Gap restante por indicador (conversão histórica indicador→venda)
    gap_indicadores = _funil_gap_vendas(
        gap_vendas,
        taxas_hist_frac,
        totais_hist,
        gap_vendas_mes=gap_vendas,
        funil_necessario=funil_necessario,
        totais_mtd=totais_mtd,
    )

    # Distribuição do gap só nos dias restantes do mês atual
    dias_distrib = [d for d in dias_mes if d >= hoje]
    pesos_gap = _pesos_distribuicao_gap(dias_distrib, resultados)
    meta_vendas_dia = _meta_vendas_por_pesos(gap_vendas, [hoje], dias_distrib, pesos_gap)

    semana_iso = hoje.isocalendar()[:2]
    dias_semana_rest = []
    d_sem = hoje
    while d_sem <= fim_semana and d_sem.isocalendar()[:2] == semana_iso:
        dias_semana_rest.append(d_sem)
        d_sem += timedelta(days=1)
    dias_semana_mes = [d for d in dias_semana_rest if d <= fim_mes]
    dias_semana_prox = [d for d in dias_semana_rest if d > fim_mes]

    meta_vendas_semana_mes = _meta_vendas_por_pesos(
        gap_vendas, dias_semana_mes, dias_distrib, pesos_gap
    )
    funil_meta_dia = _funil_gap_vendas(
        meta_vendas_dia,
        taxas_hist_frac,
        totais_hist,
        gap_vendas_mes=gap_vendas,
        funil_necessario=funil_necessario,
        totais_mtd=totais_mtd,
    )
    funil_semana_mes = _funil_gap_vendas(
        meta_vendas_semana_mes,
        taxas_hist_frac,
        totais_hist,
        gap_vendas_mes=gap_vendas,
        funil_necessario=funil_necessario,
        totais_mtd=totais_mtd,
    )
    # Completa a semana com o projetado do mês seguinte (dias após fim do mês)
    funil_semana_prox: Dict[str, float] = {e: 0.0 for e in FUNIL_ETAPAS}
    for d in dias_semana_prox:
        i = idx_por_data.get(d)
        if i is None:
            continue
        for etapa in FUNIL_ETAPAS:
            funil_semana_prox[etapa] += float(cal_reg.at[i, etapa])
    funil_meta_semana = _somar_funis(funil_semana_mes, funil_semana_prox)
    meta_vendas_semana = float(funil_meta_semana.get("vendas", 0.0))

    razoes_real_vs_proj_mtd = _razoes_entre_funis(totais_mtd, totais_proj_mtd)
    razoes_proj_vs_nec = _razoes_entre_funis(totais_proj, funil_necessario)

    dias_futuros = [d for d in dias_mes if d > hoje]
    metas_diarias: Dict[str, Any] = {}
    for etapa in FUNIL_ETAPAS:
        df_d = (resultados.get(etapa) or {}).get("diaria", pd.DataFrame())
        if df_d is None or df_d.empty or not dias_futuros:
            metas_diarias[etapa] = {
                "gap": gap_indicadores.get(etapa, 0.0),
                "ritmo_reg": pd.DataFrame(columns=["dia", "qtd"]),
                "ritmo_med": pd.DataFrame(columns=["dia", "qtd"]),
                "realizado": pd.DataFrame(columns=["dia", "qtd"]),
            }
            continue
        fut = df_d[df_d["dia"] > hoje.day]
        pesos_reg = fut["projetado_reg"].astype(float).values if not fut.empty else np.ones(len(dias_futuros))
        pesos_med = fut["projetado_med"].astype(float).values if not fut.empty else np.ones(len(dias_futuros))
        # Alinha comprimento
        if len(pesos_reg) != len(dias_futuros):
            pesos_reg = np.ones(len(dias_futuros), dtype=float)
        if len(pesos_med) != len(dias_futuros):
            pesos_med = np.ones(len(dias_futuros), dtype=float)
        gap_e = float(gap_indicadores.get(etapa, 0.0))
        ritmo_reg, _ = _distribuir_gap_por_pesos(
            gap_e, pesos_reg, dias_futuros, totais_mtd.get(etapa, 0.0), hoje.day, arredondar_cima=True
        )
        ritmo_med, _ = _distribuir_gap_por_pesos(
            gap_e, pesos_med, dias_futuros, totais_mtd.get(etapa, 0.0), hoje.day, arredondar_cima=True
        )
        real_rows = []
        for d in dias_mes:
            if d > hoje:
                break
            real_rows.append({
                "dia": d.day,
                "qtd": float((mapas.get(etapa) or {}).get(d, 0.0)),
            })
        metas_diarias[etapa] = {
            "gap": gap_e,
            "necessario_mes": float(funil_necessario.get(etapa, 0.0)),
            "ritmo_reg": ritmo_reg,
            "ritmo_med": ritmo_med,
            "realizado": pd.DataFrame(real_rows),
        }

    # Perfil de lags 0..14 sobre vendas (tempo até o efeito)
    max_perfil = max(FUNIL_LAGS_PERFIL) if FUNIL_LAGS_PERFIL else 0
    cal_perfil = calendario_funil_diario(
        inicio - timedelta(days=max_perfil),
        fim_treino,
        mapas,
        lags=FUNIL_LAGS_PERFIL,
    )
    treino_perfil = cal_perfil.loc[
        (cal_perfil["data"] >= inicio) & (cal_perfil["data"] <= fim_treino)
    ].copy()
    efeitos_lags = estimar_efeitos_lags_sobre_vendas(
        treino_perfil, lags=FUNIL_LAGS_PERFIL, incluir_mes=incluir_mes
    )

    return {
        "hoje": hoje,
        "inicio_treino": inicio,
        "fim_treino": fim_treino,
        "incluir_mes": incluir_mes,
        "lags": lags,
        "ultimo_dia": ultimo_dia,
        "fim_mes": fim_mes,
        "fim_semana": fim_semana,
        "meta_qtd_mes": meta_qtd,
        "gap_vendas": gap_vendas,
        "r2s": r2s,
        "r2s_medias": r2s_medias,
        "medias": medias,
        "etapas": resultados,
        "totais_mtd": totais_mtd,
        "totais_proj_mtd": totais_proj_mtd,
        "totais_proj": totais_proj,
        "totais_hist": totais_hist,
        "funil_necessario": funil_necessario,
        "funil_meta_dia": funil_meta_dia,
        "funil_meta_semana": funil_meta_semana,
        "funil_semana_mes": funil_semana_mes,
        "funil_semana_prox": funil_semana_prox,
        "meta_vendas_dia": meta_vendas_dia,
        "meta_vendas_semana": meta_vendas_semana,
        "meta_vendas_semana_mes": meta_vendas_semana_mes,
        "dias_semana_prox": dias_semana_prox,
        "razoes_real_vs_proj_mtd": razoes_real_vs_proj_mtd,
        "razoes_proj_vs_nec": razoes_proj_vs_nec,
        "conversoes": conversoes,
        "metas_diarias": metas_diarias,
        "efeitos_lags_vendas": efeitos_lags,
    }


def _fmt_taxa_pct(taxa: Optional[float]) -> str:
    if taxa is None:
        return "—"
    return f"{float(taxa):.1f}%"


def _fmt_razao(r: Optional[float]) -> str:
    if r is None:
        return "—"
    return f"{float(r):.2f}×"


def _render_razoes_funil(
    razoes: Dict[str, Optional[float]],
    titulo: str,
    caption: str = "",
) -> None:
    """Cards com razão entre dois funis por etapa."""
    st.markdown(f"###### {titulo}")
    if caption:
        st.caption(caption)
    _render_kpi_cards([
        {
            "lbl": FUNIL_LABELS.get(etapa, etapa),
            "val": _fmt_razao(razoes.get(etapa)),
            "sub": (
                "acima" if (razoes.get(etapa) or 0) > 1.02
                else ("abaixo" if (razoes.get(etapa) is not None and float(razoes.get(etapa)) < 0.98) else "no ritmo")
            ),
        }
        for etapa in FUNIL_ETAPAS
    ])


def _plot_funil_etapa_comparativo(etapa: str, df: pd.DataFrame, ultimo_dia: int, dia_hoje: int) -> None:
    label = FUNIL_LABELS.get(etapa, etapa)
    st.markdown(f"##### {label}: projetado × realizado")
    if df is None or df.empty:
        return
    fig = go.Figure()
    real = df.dropna(subset=["realizado"])
    if not real.empty:
        fig.add_trace(
            go.Scatter(
                x=real["dia"], y=real["realizado"],
                mode="lines+markers+text",
                name="Realizado",
                text=[fmt_qtd(float(v)) for v in real["realizado"]],
                textposition="top center",
                textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
                line=dict(color=COR_AZUL_ESC, width=3),
                marker=dict(size=7, color=COR_AZUL_ESC),
            )
        )
    fig.add_trace(
        go.Scatter(
            x=df["dia"], y=df["projetado_reg"],
            mode="lines+markers+text",
            name="Projetado (regressão)",
            text=[fmt_qtd(float(v)) for v in df["projetado_reg"]],
            textposition="bottom center",
            textfont=dict(size=10, color=COR_VERMELHO, family="Inter"),
            line=dict(color=COR_VERMELHO, width=3, dash="dash"),
            marker=dict(size=7, color=COR_VERMELHO),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=df["dia"], y=df["projetado_med"],
            mode="lines+markers+text",
            name="Projetado (médias)",
            text=[fmt_qtd(float(v)) for v in df["projetado_med"]],
            textposition="top center",
            textfont=dict(size=10, color="#0f766e", family="Inter"),
            line=dict(color="#0f766e", width=3, dash="dot"),
            marker=dict(size=7, symbol="diamond", color="#0f766e"),
        )
    )
    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=11),
        ),
        hovermode="x unified",
        height=380,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, ultimo_dia + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _render_kpi_cards(items: List[Dict[str, str]]) -> None:
    """
    Cards KPI em colunas Streamlit (evita HTML indentado virar bloco de código no markdown).
    items: [{"lbl": ..., "val": ..., "sub": ...}, ...]
    """
    if not items:
        return
    n = len(items)
    cols = st.columns(n)
    for col, it in zip(cols, items):
        lbl = html.escape(str(it.get("lbl", "")))
        val = html.escape(str(it.get("val", "")))
        sub = html.escape(str(it.get("sub", "")))
        with col:
            st.markdown(
                (
                    f'<div class="vel-kpi" style="width:100%;box-sizing:border-box;">'
                    f'<div class="lbl">{lbl}</div>'
                    f'<div class="val">{val}</div>'
                    f'<div class="lbl" style="margin-top:6px;opacity:0.75;">{sub}</div>'
                    f"</div>"
                ),
                unsafe_allow_html=True,
            )


def _render_conversoes_funil(conversoes: Dict[str, Any]) -> None:
    """Conversões do mês atual × histórico (sem mês atual, até 1 ano)."""
    if not conversoes:
        return
    mes = conversoes.get("realizado_mtd") or {}
    hist = conversoes.get("historico") or {}
    ini = conversoes.get("inicio_hist")
    fim = conversoes.get("fim_hist")
    periodo = ""
    if ini and fim:
        periodo = f"{ini.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"

    st.markdown("##### Referência: conversões MTD × histórico")
    st.caption(f"Mês atual (MTD) × histórico ({periodo or 'janela de treino'})")
    pares_m = mes.get("etapa_a_etapa") or []
    pares_h = {r["label"]: r for r in (hist.get("etapa_a_etapa") or [])}
    st.caption("Conversões etapa → etapa")
    _render_kpi_cards([
        {
            "lbl": str(r.get("label", "")),
            "val": _fmt_taxa_pct(r.get("taxa")),
            "sub": f"Hist. {_fmt_taxa_pct((pares_h.get(r.get('label')) or {}).get('taxa'))}",
        }
        for r in pares_m
    ])

    st.caption("Conversões diretas (indicador → venda)")
    fins_m = mes.get("para_venda") or []
    fins_h = {r["label"]: r for r in (hist.get("para_venda") or [])}
    _render_kpi_cards([
        {
            "lbl": str(r.get("label", "")),
            "val": _fmt_taxa_pct(r.get("taxa")),
            "sub": f"Hist. {_fmt_taxa_pct((fins_h.get(r.get('label')) or {}).get('taxa'))}",
        }
        for r in fins_m
    ])

    labels_e = [str(r.get("label", "")) for r in pares_m]
    y_mes_e = [float(r["taxa"]) if r.get("taxa") is not None else 0.0 for r in pares_m]
    y_hist_e = [float((pares_h.get(lbl) or {}).get("taxa") or 0.0) for lbl in labels_e]
    labels_f = [str(r.get("label", "")) for r in fins_m]
    y_mes_f = [float(r["taxa"]) if r.get("taxa") is not None else 0.0 for r in fins_m]
    y_hist_f = [float((fins_h.get(lbl) or {}).get("taxa") or 0.0) for lbl in labels_f]

    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=("Etapa → etapa", "Indicador → venda"),
        horizontal_spacing=0.08,
    )
    fig.add_trace(
        go.Bar(
            name="Mês atual (MTD)", x=labels_e, y=y_mes_e,
            text=[_fmt_taxa_pct(v) for v in y_mes_e], textposition="outside",
            marker_color=COR_AZUL_ESC, textfont=dict(color=COR_TEXTO_PRETO, size=10),
            legendgroup="mes",
        ),
        row=1, col=1,
    )
    fig.add_trace(
        go.Bar(
            name="Histórico (sem mês atual)", x=labels_e, y=y_hist_e,
            text=[_fmt_taxa_pct(v) for v in y_hist_e], textposition="outside",
            marker_color="#0f766e", textfont=dict(color=COR_TEXTO_PRETO, size=10),
            legendgroup="hist",
        ),
        row=1, col=1,
    )
    fig.add_trace(
        go.Bar(
            name="Mês atual (MTD)", x=labels_f, y=y_mes_f,
            text=[_fmt_taxa_pct(v) for v in y_mes_f], textposition="outside",
            marker_color=COR_AZUL_ESC, textfont=dict(color=COR_TEXTO_PRETO, size=10),
            legendgroup="mes", showlegend=False,
        ),
        row=1, col=2,
    )
    fig.add_trace(
        go.Bar(
            name="Histórico (sem mês atual)", x=labels_f, y=y_hist_f,
            text=[_fmt_taxa_pct(v) for v in y_hist_f], textposition="outside",
            marker_color="#0f766e", textfont=dict(color=COR_TEXTO_PRETO, size=10),
            legendgroup="hist", showlegend=False,
        ),
        row=1, col=2,
    )
    fig.update_layout(
        barmode="group",
        margin=dict(l=20, r=20, t=50, b=80),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.12, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        height=420,
    )
    fig.update_yaxes(
        title_text="Taxa (%)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    fig.update_xaxes(tickfont=dict(color=COR_TEXTO_PRETO, family="Inter", size=10), tickangle=-25)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _render_conversoes_de_totais(
    totais: Dict[str, float],
    titulo: str = "",
) -> None:
    """Conversões etapa→etapa e diretas (indicador→venda) de um funil."""
    conv = calcular_conversoes_totais(ceil_funil_totais(totais))
    if titulo:
        st.markdown(f"###### {titulo}")
    pares = conv.get("etapa_a_etapa") or []
    diretas = conv.get("para_venda") or []
    st.caption("Conversões etapa → etapa")
    _render_kpi_cards([
        {
            "lbl": str(r.get("label", "")),
            "val": _fmt_taxa_pct(r.get("taxa")),
            "sub": (
                f"{fmt_qtd(float(r.get('destino_qtd', 0)))} / "
                f"{fmt_qtd(float(r.get('origem_qtd', 0)))}"
            ),
        }
        for r in pares
    ])
    st.caption("Conversões diretas (indicador → venda)")
    _render_kpi_cards([
        {
            "lbl": str(r.get("label", "")),
            "val": _fmt_taxa_pct(r.get("taxa")),
            "sub": (
                f"{fmt_qtd(float(r.get('destino_qtd', 0)))} / "
                f"{fmt_qtd(float(r.get('origem_qtd', 0)))}"
            ),
        }
        for r in diretas
    ])


def _render_conversoes_par_funis(
    totais_a: Dict[str, float],
    label_a: str,
    totais_b: Dict[str, float],
    label_b: str,
) -> None:
    """Exibe conversões etapa→etapa e diretas dos dois funis da seção."""
    st.markdown("###### Conversões etapa → etapa e diretas")
    ca, cb = st.columns(2)
    with ca:
        _render_conversoes_de_totais(totais_a, titulo=label_a)
    with cb:
        _render_conversoes_de_totais(totais_b, titulo=label_b)


def _plot_funil_go(titulo: str, totais: Dict[str, float], altura: int = 350) -> None:
    """Funil estilo marketing (go.Funnel); volumes arredondados para cima."""
    labels = [FUNIL_LABELS[e] for e in FUNIL_ETAPAS]
    ceil_tot = ceil_funil_totais(totais)
    vals = [float(ceil_tot.get(e, 0.0)) for e in FUNIL_ETAPAS]
    fig = _criar_fig_funil(labels, vals, titulo=titulo, altura=altura)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _plot_meta_diaria_indicador(
    etapa: str,
    meta_info: Dict[str, Any],
    ultimo_dia: int,
    dia_hoje: int,
) -> None:
    """Realizado + meta diária (reg/médias) para bater a necessidade do indicador."""
    label = FUNIL_LABELS.get(etapa, etapa)
    gap = float(meta_info.get("gap", 0.0))
    nec = float(meta_info.get("necessario_mes", 0.0))
    st.markdown(f"##### Meta diária — {label}")
    st.caption(
        f"Necessário no mês: {fmt_qtd(nec)} · faltam {fmt_qtd(gap)} · "
        "pesos = projeção diária (dia da semana / dia do mês / mês)"
    )
    real = meta_info.get("realizado", pd.DataFrame())
    ritmo_reg = meta_info.get("ritmo_reg", pd.DataFrame())
    ritmo_med = meta_info.get("ritmo_med", pd.DataFrame())

    fig = go.Figure()
    if real is not None and not real.empty:
        fig.add_trace(go.Scatter(
            x=real["dia"], y=real["qtd"],
            mode="lines+markers+text",
            name="Realizado",
            text=[fmt_qtd(float(v)) for v in real["qtd"]],
            textposition="top center",
            textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
            line=dict(color=COR_AZUL_ESC, width=3),
            marker=dict(size=7, color=COR_AZUL_ESC),
        ))
    if ritmo_reg is not None and not ritmo_reg.empty:
        fig.add_trace(go.Scatter(
            x=ritmo_reg["dia"], y=ritmo_reg["qtd"],
            mode="lines+markers+text",
            name="Meta diária (reg.)",
            text=[fmt_qtd(float(v)) for v in ritmo_reg["qtd"]],
            textposition="top center",
            textfont=dict(size=10, color=COR_VERMELHO, family="Inter"),
            line=dict(color=COR_VERMELHO, width=3),
            marker=dict(size=7, color=COR_VERMELHO),
        ))
    if ritmo_med is not None and not ritmo_med.empty:
        fig.add_trace(go.Scatter(
            x=ritmo_med["dia"], y=ritmo_med["qtd"],
            mode="lines+markers+text",
            name="Meta diária (médias)",
            text=[fmt_qtd(float(v)) for v in ritmo_med["qtd"]],
            textposition="bottom center",
            textfont=dict(size=10, color="#0f766e", family="Inter"),
            line=dict(color="#0f766e", width=3, dash="dash"),
            marker=dict(size=7, symbol="diamond", color="#0f766e"),
        ))
    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=11),
        ),
        hovermode="x unified",
        height=360,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, ultimo_dia + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_projecao_funil(proj: Dict[str, Any]) -> None:
    """Seção Streamlit: 3 funis, conversões mês×histórico, metas diárias e projeções."""
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Projeção do Funil Comercial")
    lags = proj.get("lags") or FUNIL_LAGS
    meta_qtd = float(proj.get("meta_qtd_mes") or 0.0)
    st.caption(
        f"Modelos com efeitos cruzados + força de trabalho "
        f"(o trabalho vaza entre etapas; não só cascata) · "
        f"lags 1–{max(lags)}d"
        + (f" · Meta vendas: {fmt_qtd(meta_qtd)}" if meta_qtd > 0 else "")
    )

    etapas = proj.get("etapas") or {}
    r2s = proj.get("r2s") or {}
    r2s_med = proj.get("r2s_medias") or {}

    st.markdown("##### R² dos modelos (treino sem o mês atual)")
    _render_kpi_cards([
        {
            "lbl": FUNIL_LABELS.get(etapa, etapa),
            "val": f"Reg {float(r2s.get(etapa, (etapas.get(etapa) or {}).get('r2', 0))):.2f}",
            "sub": f"Médias {float(r2s_med.get(etapa, (etapas.get(etapa) or {}).get('r2_medias', 0))):.2f}",
        }
        for etapa in FUNIL_ETAPAS
    ])

    _render_kpi_cards([
        {
            "lbl": FUNIL_LABELS.get(etapa, etapa),
            "val": fmt_qtd(float((etapas.get(etapa) or {}).get("mtd", 0))),
            "sub": (
                f"Proj. reg {fmt_qtd(float((etapas.get(etapa) or {}).get('projetado_reg', 0)))}"
                f" (+{fmt_qtd(max(0.0, float((etapas.get(etapa) or {}).get('projetado_reg', 0)) - float((etapas.get(etapa) or {}).get('mtd', 0))))})"
                f" · méd {fmt_qtd(float((etapas.get(etapa) or {}).get('projetado_med', 0)))}"
            ),
        }
        for etapa in FUNIL_ETAPAS
    ])

    efeitos = proj.get("efeitos_lags_vendas")
    if efeitos:
        st.markdown("##### Tempo até o efeito nas vendas")
        st.caption(
            f"R² médio: {float(efeitos.get('r2', 0)):.2f} · "
            "lags 1–30d por etapa · pico = maior efeito acumulado · "
            "meia-vida ≈ 50% do efeito acumulado total"
        )
        resumo = efeitos.get("resumo") or []
        _render_kpi_cards([
            {
                "lbl": str(r.get("label", "")),
                "val": f"{int(r.get('lag_pico', 0))}d",
                "sub": (
                    f"Meia-vida {int(r.get('lag_meia_vida', 0))}d"
                    f" · lag1 {float(r.get('efeito_lag1', 0)):.2f}"
                    f" · acum {float(r.get('efeito_acum', 0)):.2f}"
                ),
            }
            for r in resumo
        ])

    # -------------------------------------------------------------------------
    # Seção 1 — Realizado até agora × Projetado até agora
    # -------------------------------------------------------------------------
    totais_mtd = proj.get("totais_mtd") or {
        e: float((etapas.get(e) or {}).get("mtd", 0)) for e in FUNIL_ETAPAS
    }
    totais_proj_mtd = proj.get("totais_proj_mtd") or {
        e: float((etapas.get(e) or {}).get("projetado_mtd_reg", 0)) for e in FUNIL_ETAPAS
    }
    totais_proj = proj.get("totais_proj") or {
        e: float((etapas.get(e) or {}).get("projetado_reg", 0)) for e in FUNIL_ETAPAS
    }
    funil_nec = proj.get("funil_necessario") or totais_proj
    razoes_1 = proj.get("razoes_real_vs_proj_mtd") or _razoes_entre_funis(totais_mtd, totais_proj_mtd)
    razoes_2 = proj.get("razoes_proj_vs_nec") or _razoes_entre_funis(totais_proj, funil_nec)

    st.markdown("##### 1) Realizado até agora × Projetado até agora")
    st.caption(
        "Compara o funil realizado (MTD) com o que o modelo previa até hoje. "
        "Razão > 1 = acima do projetado; < 1 = abaixo. Volumes arredondados para cima."
    )
    c1, c2 = st.columns(2)
    with c1:
        _plot_funil_go("Realizado até agora (MTD)", totais_mtd, altura=380)
    with c2:
        _plot_funil_go("Projetado até agora (modelo)", totais_proj_mtd, altura=380)
    _render_razoes_funil(
        razoes_1,
        "Razões realizado / projetado até agora",
        "Por etapa: realizado ÷ projetado MTD do modelo.",
    )
    _render_conversoes_par_funis(
        totais_mtd, "Realizado até agora",
        totais_proj_mtd, "Projetado até agora",
    )

    # -------------------------------------------------------------------------
    # Seção 2 — Projetado do mês × Necessário para a meta
    # -------------------------------------------------------------------------
    st.markdown("##### 2) Projetado do mês × Necessário para a meta")
    st.caption(
        "Projetado do mês = realizado até agora + projeção dos dias restantes. "
        "Necessário = volumes para bater a meta de vendas (conversões históricas). "
        "Razão ≥ 1 sugere que, mantendo o previsto, a meta é atingível."
    )
    c3, c4 = st.columns(2)
    with c3:
        _plot_funil_go("Projetado do mês", totais_proj, altura=380)
    with c4:
        _plot_funil_go("Necessário p/ meta", funil_nec, altura=380)
    _render_razoes_funil(
        razoes_2,
        "Razões projetado do mês / necessário para a meta",
        "Por etapa: projetado do mês ÷ necessário.",
    )
    _render_conversoes_par_funis(
        totais_proj, "Projetado do mês",
        funil_nec, "Necessário p/ meta",
    )

    # -------------------------------------------------------------------------
    # Seção 3 — Meta do dia × Meta da semana
    # -------------------------------------------------------------------------
    funil_dia = proj.get("funil_meta_dia") or {}
    funil_sem = proj.get("funil_meta_semana") or {}
    meta_v_dia = float(math.ceil(float(proj.get("meta_vendas_dia") or 0.0)))
    meta_v_sem = float(math.ceil(float(proj.get("meta_vendas_semana") or 0.0)))
    meta_v_sem_mes = float(math.ceil(float(proj.get("meta_vendas_semana_mes") or 0.0)))
    gap_v = float(math.ceil(float(proj.get("gap_vendas") or 0.0)))
    dias_prox = proj.get("dias_semana_prox") or []
    fim_mes = proj.get("fim_mes")
    fim_semana = proj.get("fim_semana")
    nota_prox = ""
    if dias_prox:
        d0 = min(dias_prox)
        d1 = max(dias_prox)
        fim_mes_txt = f" ({fim_mes.strftime('%d/%m')})" if fim_mes else ""
        nota_prox = (
            f" Semana atravessa o fim do mês{fim_mes_txt}: "
            f"dias {d0.strftime('%d/%m')}–{d1.strftime('%d/%m')} "
            f"completados com o projetado do mês seguinte "
            f"(meta do mês na semana: {fmt_qtd(meta_v_sem_mes)} vendas)."
        )

    st.markdown("##### 3) Meta do dia × Meta da semana (para bater a meta de vendas)")
    st.caption(
        f"Restante do mês: {fmt_qtd(gap_v)} vendas · "
        f"hoje: {fmt_qtd(meta_v_dia)} · semana: {fmt_qtd(meta_v_sem)}"
        + (
            f" (até {fim_semana.strftime('%d/%m')})"
            if fim_semana else ""
        )
        + " · gap do mês distribuído pelos pesos da projeção diária."
        + nota_prox
    )
    c5, c6 = st.columns(2)
    with c5:
        _plot_funil_go(f"Meta do dia ({fmt_qtd(meta_v_dia)} vendas)", funil_dia, altura=380)
    with c6:
        _plot_funil_go(
            f"Meta da semana ({fmt_qtd(meta_v_sem)} vendas)", funil_sem, altura=380
        )
    _render_conversoes_par_funis(
        funil_dia, "Meta do dia",
        funil_sem, "Meta da semana",
    )

    # Comparativo histórico de conversões (referência)
    _render_conversoes_funil(proj.get("conversoes") or {})

    # Metas diárias por indicador
    st.markdown("##### Meta diária por indicador (para bater a meta de vendas)")
    st.caption(
        "Vendas faltantes ÷ conversão histórica indicador→venda = gap do indicador; "
        "distribuído nos dias restantes com pesos da projeção (reg / médias)."
    )
    metas_diarias = proj.get("metas_diarias") or {}
    dia_hoje = proj["hoje"].day
    ultimo = int(proj["ultimo_dia"])
    for etapa in FUNIL_ETAPAS:
        info = metas_diarias.get(etapa) or {}
        if info:
            _plot_meta_diaria_indicador(etapa, info, ultimo, dia_hoje)

    # Comparativo diário projetado × realizado (ambos os modelos)
    st.markdown("##### Projeção diária × realizado (modelos treinados sem o mês atual)")
    st.caption(
        "Linhas de projetado = previsão do modelo (reg / médias) em cada dia. "
        "Realizado só até hoje. Total do mês = realizado MTD + previsão dos dias restantes."
    )
    for etapa in FUNIL_ETAPAS:
        df = (etapas.get(etapa) or {}).get("diaria", pd.DataFrame())
        _plot_funil_etapa_comparativo(etapa, df, ultimo, dia_hoje)

    if efeitos:
        render_efeitos_lags_sobre_vendas(efeitos, mostrar_cards=False)


def render_efeitos_lags_sobre_vendas(
    efeitos: Dict[str, Any],
    mostrar_cards: bool = True,
) -> None:
    """Perfil de lags (1–30d) das etapas do funil sobre vendas."""
    st.markdown("<br/>", unsafe_allow_html=True)
    st.subheader("Perfil de lags sobre vendas")
    st.caption(
        f"R² médio: {float(efeitos.get('r2', 0)):.2f} · "
        "modelo por etapa (vendas ~ calendário + lags 1–30 da etapa) · "
        "pico = lag de maior efeito acumulado"
    )

    if mostrar_cards:
        resumo = efeitos.get("resumo") or []
        _render_kpi_cards([
            {
                "lbl": str(r.get("label", "")),
                "val": f"{int(r.get('lag_pico', 0))}d",
                "sub": (
                    f"Meia-vida {int(r.get('lag_meia_vida', 0))}d"
                    f" · lag1 {float(r.get('efeito_lag1', 0)):.2f}"
                    f" · acum {float(r.get('efeito_acum', 0)):.2f}"
                ),
            }
            for r in resumo
        ])

    perfis = efeitos.get("perfis") or {}
    fig = go.Figure()
    for etapa in FUNIL_DRIVERS:
        df = perfis.get(etapa)
        if df is None or df.empty:
            continue
        cor = FUNIL_CORES_DRIVER.get(etapa, COR_AZUL_ESC)
        fig.add_trace(
            go.Scatter(
                x=df["lag"],
                y=df["efeito"],
                mode="lines+markers",
                name=FUNIL_LABELS.get(etapa, etapa),
                line=dict(color=cor, width=3),
                marker=dict(size=7, color=cor),
            )
        )
    fig.add_hline(y=0, line_width=1, line_dash="dot", line_color="#64748b")
    fig.update_layout(
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=420,
    )
    fig.update_xaxes(
        title_text="Lag (dias)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
    )
    fig.update_yaxes(
        title_text="Efeito nas vendas",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.markdown("##### Efeito por lag (1–30 dias)")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    fig_c = go.Figure()
    for etapa in FUNIL_DRIVERS:
        df = perfis.get(etapa)
        if df is None or df.empty:
            continue
        cor = FUNIL_CORES_DRIVER.get(etapa, COR_AZUL_ESC)
        fig_c.add_trace(
            go.Scatter(
                x=df["lag"],
                y=df["acumulado"],
                mode="lines+markers",
                name=FUNIL_LABELS.get(etapa, etapa),
                line=dict(color=cor, width=3),
                marker=dict(size=7, color=cor),
            )
        )
    fig_c.add_hline(y=0, line_width=1, line_dash="dot", line_color="#64748b")
    fig_c.update_layout(
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=380,
    )
    fig_c.update_xaxes(
        title_text="Lag (dias)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
    )
    fig_c.update_yaxes(
        title_text="Efeito acumulado",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.markdown("##### Efeito acumulado até o lag")
    st.plotly_chart(fig_c, use_container_width=True, config={"displayModeBar": False})


def criar_medidor(titulo: str, realizado: float, meta: float, vgv: float, meta_vgv: float, vendas_qtd: float) -> None:
    meta_f = float(meta) if meta and meta > 0 else 0.0
    true_perc = (realizado / meta_f * 100.0) if meta_f > 0 else 0.0
    axis_max = 100
    fill_limit = min(true_perc, axis_max)

    gradient_steps = []
    for i in range(100):
        if i >= fill_limit: break
        ratio = i / 100.0
        r = int(203 + (4 - 203) * ratio)
        g = int(9 + (66 - 9) * ratio)
        b = int(53 + (143 - 53) * ratio)
        end_val = min(i + 1.0, fill_limit)
        gradient_steps.append({"range": [i, end_val], "color": f"rgba({r}, {g}, {b}, 0.9)"})
        
    if fill_limit < 100:
        gradient_steps.append({"range": [fill_limit, 100], "color": "rgba(226, 232, 240, 0.4)"})

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=true_perc,
            number={
                "suffix": "%",
                "font": {"size": 26, "family": "Montserrat", "color": COR_AZUL_ESC},
                "valueformat": ".1f",
            },
            title={
                "text": titulo,
                "font": {"size": 16, "family": "Montserrat", "color": COR_TEXTO_PRETO},
            },
            gauge={
                "axis": {
                    "range": [0, axis_max],
                    "tickwidth": 1,
                    "tickcolor": "#64748b",
                    "tickfont": {"color": COR_TEXTO_PRETO, "family": "Inter"},
                },
                "bar": {"color": "rgba(0,0,0,0)"},
                "bgcolor": "white",
                "borderwidth": 2,
                "bordercolor": COR_BORDA,
                "steps": gradient_steps,
                "threshold": {
                    "line": {"color": COR_AZUL_ESC, "width": 3},
                    "thickness": 0.8,
                    "value": 100,
                },
            },
        )
    )

    fig.update_layout(
        height=300,
        margin=dict(l=24, r=24, t=56, b=24),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown(
        f"""
        <div style="text-align:center;font-size:0.85rem;color:{COR_TEXTO_PRETO};margin-top:-8px;line-height:1.4;">
            <strong>Qtd:</strong> {fmt_qtd(vendas_qtd)} / {fmt_qtd(meta_f)} <br/>
            <strong>VGV:</strong> {fmt_br_milhoes(float(vgv))} / {fmt_br_milhoes(float(meta_vgv))}
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)


def main() -> None:
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(
        page_title="Acompanhamento de Vendas | Direcional",
        page_icon=str(fav) if fav else None,
        layout="wide",
    )
    aplicar_estilo()
    _cabecalho_pagina()

    raw_gs = _secrets_connections_gsheets()
    info = montar_service_account_info(raw_gs)
    if not info:
        st.error("Credenciais Google em **[connections.gsheets]** incompletas. Preencha pelo menos **private_key** e **client_email**.")
        return

    sid = spreadsheet_id_de_secrets(raw_gs)
    cred_fp = _fingerprint_credenciais(info)

    try:
        df_vendas_raw = ler_planilha_aba_df(sid, WS_VENDAS, cred_fp)
        df_metas_raw = ler_planilha_aba_df(sid, WS_METAS, cred_fp)
    except Exception as e:
        st.error(f"Erro ao ler a planilha principal: {str(e)}")
        return

    df_vendas = normalizar_colunas(df_vendas_raw)
    df_metas_melted = melt_metas(df_metas_raw)

    # -------------------------------------------------------------------------
    # Limpeza de Lixo (Mantida Conforme Pedido)
    # -------------------------------------------------------------------------
    if "Região" in df_metas_melted.columns:
        df_metas_melted = df_metas_melted[~df_metas_melted["Região"].astype(str).str.strip().str.lower().isin(["total", "geral", "não informado", "nao informado", "nan", "none", "null", ""])]
    if "Empreendimento" in df_metas_melted.columns:
        df_metas_melted = df_metas_melted[~df_metas_melted["Empreendimento"].astype(str).str.strip().str.lower().isin(["total", "geral", "não informado", "nao informado", "nan", "none", "null", ""])]

    # -------------------------------------------------------------------------
    # Tratamento da Coluna Coordenador
    # -------------------------------------------------------------------------
    if "Coordenador" not in df_metas_melted.columns:
        df_metas_melted["Coordenador"] = "Não Informado"
        
    rows_metas = []
    for _, row in df_metas_melted.iterrows():
        coords = [c.strip() for c in re.split(r'\s+e\s+', str(row.get("Coordenador", "Não Informado"))) if c.strip()]
        if not coords:
            coords = ["Não Informado"]
        n = len(coords)
        for c in coords:
            new_row = row.copy()
            new_row["Coordenador"] = c
            new_row["Regiao_Coord"] = f"{row['Região']} - {c}" if c not in ("Não Informado", "nan", "") else str(row['Região'])
            new_row["Meta_Qtd"] = float(new_row["Meta_Qtd"]) / n
            new_row["Meta_VGV"] = float(new_row["Meta_VGV"]) / n
            new_row["_peso_coord"] = 1.0 / n
            rows_metas.append(new_row)
            
    df_metas = pd.DataFrame(rows_metas)

    # -------------------------------------------------------------------------
    # Colunas Vendas (Alinhado com o mapeamento exato enviado)
    # -------------------------------------------------------------------------
    col_ano = achar_coluna(df_vendas, ["Ano da Venda", "Ano Venda", "Ano"])
    col_mes_looker = achar_coluna(df_vendas, ["Mês da Venda - Looker"])
    col_mes_venda = achar_coluna(df_vendas, ["Mês Venda"])
    col_regiao = achar_coluna(df_vendas, ["Região", "Regiao"])
    col_canal = achar_coluna(df_vendas, ["Canal"])
    col_valor = achar_coluna(df_vendas, ["Valor Real de Venda", "Valor Real", "Valor"])
    col_emp = achar_coluna(df_vendas, ["Empreendimento", "Obra", "Nome do Empreendimento"])
    col_venda_comercial = achar_coluna(df_vendas, ALIASES_VENDA_COMERCIAL)
    col_venda_facilitada = achar_coluna(df_vendas, ["Venda facilitada", "Venda Facilitada", "Venda Facilitada?"])
    col_proprietario = achar_coluna(df_vendas, ["Proprietário da oportunidade", "Proprietario da oportunidade", "Nome da conta", "Proprietario", "Corretor"])
    col_ranking = achar_coluna(df_vendas, ["Ranking"])
    col_data_venda = achar_coluna(df_vendas, ["Data da venda", "Data Venda", "Data de venda", "Data"])
    col_contrato_gerado = achar_coluna(df_vendas, ["Contrato gerado em", "Contrato gerado"])

    if col_emp and col_emp != "Empreendimento":
        df_vendas.rename(columns={col_emp: "Empreendimento"}, inplace=True)
        col_emp = "Empreendimento"
    if col_regiao and col_regiao != "Região":
        df_vendas.rename(columns={col_regiao: "Região"}, inplace=True)
        col_regiao = "Região"

    # Limpeza de Lixo na Base de Vendas (Mantida Conforme Pedido)
    if col_emp:
        df_vendas = df_vendas[~df_vendas[col_emp].astype(str).str.strip().str.lower().isin(["total", "geral", "nan", "none", "null", ""])]

    if col_venda_comercial:
        df_vendas = filtrar_vendas_comerciais(df_vendas)
    else:
        st.warning("Coluna 'Venda Comercial?' não encontrada na base.")

    if col_venda_facilitada:
        def check_facilitada(val: Any) -> str:
            if pd.isna(val):
                return "Normal"
            v_str = str(val).strip().upper()
            if v_str in ("1", "1.0", "SIM", "TRUE"):
                return "Facilitada"
            return "Normal"
        df_vendas["Tipo_Venda"] = df_vendas[col_venda_facilitada].apply(check_facilitada)
    else:
        df_vendas["Tipo_Venda"] = "Normal"

    # -------------------------------------------------------------------------
    # Extração de Data Segura e Corrigida (Suporta 2.026 e 2026 nativos)
    # -------------------------------------------------------------------------
    if col_data_venda:
        df_vendas["_mes"] = df_vendas[col_data_venda].apply(extrair_mes_da_data_venda)
        df_vendas["_ano"] = df_vendas[col_data_venda].apply(extrair_ano_da_data_venda)
    else:
        if col_mes_looker:
            df_vendas["_mes"] = df_vendas[col_mes_looker].apply(extrair_mes_looker)
            df_vendas["_ano"] = df_vendas[col_mes_looker].apply(extrair_ano_looker)
        else:
            df_vendas["_mes"] = df_vendas[col_mes_venda].apply(extrair_mes_looker) if col_mes_venda else None
            df_vendas["_ano"] = df_vendas[col_ano].apply(extrair_ano_looker) if col_ano else None

    df_vendas["_vgv"] = df_vendas[col_valor].map(parse_valor_br) if col_valor else 0.0

    if col_canal:
        def agrupar_canal(c: Any) -> str:
            bytes_str = str(c).strip().upper()
            prefixo = bytes_str.split('-')[0].strip()
            if prefixo in ['RJ', 'RJG'] or bytes_str in ['RJ', 'RJG']:
                return 'IMOB'
            return 'DV RJ'
        df_vendas['Canal_Agrupado'] = df_vendas[col_canal].apply(agrupar_canal)
    else:
        df_vendas['Canal_Agrupado'] = 'DV RJ'

    # -------------------------------------------------------------------------
    # Multiplicação e Distribuição Segura das Vendas com Coordenador
    # -------------------------------------------------------------------------
    map_emp_regiao = df_metas[["Empreendimento", "Regiao_Coord", "_peso_coord"]].drop_duplicates()
    
    lista_com_peso = []
    for _, venda_row in df_vendas.iterrows():
        emp_venda = str(venda_row["Empreendimento"]).strip()
        sub_map = map_emp_regiao[map_emp_regiao["Empreendimento"] == emp_venda]
        
        if sub_map.empty:
            nova_linha = venda_row.copy()
            nova_linha["_peso_coord"] = 1.0
            nova_linha["Regiao_Coord"] = venda_row.get("Região", "Não Informado")
            lista_com_peso.append(nova_linha)
        else:
            for _, map_row in sub_map.iterrows():
                nova_linha = venda_row.copy()
                nova_linha["_peso_coord"] = float(map_row["_peso_coord"])
                nova_linha["Regiao_Coord"] = map_row["Regiao_Coord"]
                lista_com_peso.append(nova_linha)

    df_vendas = pd.DataFrame(lista_com_peso)
    df_vendas["_qtd_venda"] = 1.0 * df_vendas["_peso_coord"]
    df_vendas["_vgv_venda"] = df_vendas["_vgv"] * df_vendas["_peso_coord"]

    # -------------------------------------------------------------------------
    # LINHA ÚNICA DE FILTROS
    # -------------------------------------------------------------------------
    anos_disponiveis = sorted(int(x) for x in df_vendas["_ano"].dropna().unique().tolist() if pd.notna(x) and x > 2000)
    meses_no_ano = list(range(1, 13))
    mes_atual = datetime.now().month
    mes_padrao = mes_atual if mes_atual in meses_no_ano else 1
    regioes_disponiveis = sorted(set(str(x).strip() for x in df_metas["Regiao_Coord"].dropna().unique() if str(x).strip()))
    
    todos_emps_vendas = sorted(list(set(str(x).strip() for x in df_vendas["Empreendimento"].dropna().unique() if str(x).strip())))

    st.markdown("<div style='margin-bottom:1rem; text-align: center;'><strong>Filtros</strong></div>", unsafe_allow_html=True)
    
    col_filtros = st.columns(6) 
    with col_filtros[0]:
        canais_sel = st.multiselect("Canal da Meta", ["RIO", "DIR", "PARC", "RJ"], default=["RIO"])
    with col_filtros[1]:
        anos_sel = st.multiselect("Ano", anos_disponiveis, default=[anos_disponiveis[-1]] if anos_disponiveis else [])
    with col_filtros[2]:
        meses_venda_sel = st.multiselect("Mês da Venda", meses_no_ano, default=[mes_padrao])
    with col_filtros[3]:
        meses_meta_sel = st.multiselect("Mês da Meta", meses_no_ano, default=[mes_padrao])
    with col_filtros[4]:
        regioes_sel = st.multiselect("Região", regioes_disponiveis, default=[])
    with col_filtros[5]:
        emps_sel = st.multiselect("Empreendimento", todos_emps_vendas, default=[])

    # -------------------------------------------------------------------------
    # Aplicação de Filtros
    # -------------------------------------------------------------------------
    vendas_f = df_vendas.copy()
    metas_f = df_metas.copy()

    if anos_sel: vendas_f = vendas_f[vendas_f["_ano"].isin(anos_sel)]
    
    if meses_venda_sel:
        vendas_f = vendas_f[vendas_f["_mes"].isin(meses_venda_sel)]
    if meses_meta_sel:
        metas_f = metas_f[metas_f["Mes_Num"].isin(meses_meta_sel)]
        
    if regioes_sel:
        metas_f = metas_f[metas_f["Regiao_Coord"].isin(regioes_sel)]
        if not emps_sel:
            regioes_base = [r.split(" - ")[0].strip() for r in regioes_sel]
            vendas_f = vendas_f[vendas_f["Regiao_Coord"].isin(regioes_sel) | vendas_f["Região"].isin(regioes_base)]
        else:
            vendas_f = vendas_f[vendas_f["Regiao_Coord"].isin(regioes_sel)]
    
    if emps_sel:
        metas_f = metas_f[metas_f["Empreendimento"].isin(emps_sel)]
        vendas_f = vendas_f[vendas_f["Empreendimento"].isin(emps_sel)]

    fator_meta = 0.0
    mask_vendas = pd.Series(False, index=vendas_f.index)

    if not canais_sel or "RIO" in canais_sel:
        fator_meta = 1.0
        mask_vendas = pd.Series(True, index=vendas_f.index)
    else:
        if "DIR" in canais_sel:
            fator_meta += 0.50
            mask_vendas |= (vendas_f["Canal_Agrupado"] == "DV RJ")
        if "PARC" in canais_sel and col_canal:
            fator_meta += 0.25
            mask_vendas |= vendas_f[col_canal].astype(str).str.upper().str.strip().apply(lambda x: x.split('-')[0].strip() == 'RJG' or x == 'RJG')
        if "RJ" in canais_sel and col_canal:
            fator_meta += 0.25
            mask_vendas |= vendas_f[col_canal].astype(str).str.upper().str.strip().apply(lambda x: x.split('-')[0].strip() == 'RJ' or x == 'RJ')

    fator_meta = min(1.0, factor_meta := fator_meta)
    vendas_f = vendas_f[mask_vendas]

    total_meta_qtd_base = float(metas_f["Meta_Qtd"].sum()) if not metas_f.empty else 0.0
    total_meta_vgv_base = float(metas_f["Meta_VGV"].sum()) if not metas_f.empty else 0.0

    metas_f["Meta_Qtd"] = (metas_f["Meta_Qtd"] * fator_meta).apply(math.floor)
    metas_f["Meta_VGV"] = metas_f["Meta_VGV"] * fator_meta

    total_realizado_qtd = float(vendas_f["_qtd_venda"].sum())
    total_meta_qtd = float(metas_f["Meta_Qtd"].sum()) if not metas_f.empty else 0.0
    total_vgv_realizado = float(vendas_f["_vgv_venda"].sum())
    total_meta_vgv = float(metas_f["Meta_VGV"].sum()) if not metas_f.empty else 0.0

    pct_qtd = (total_realizado_qtd / total_meta_qtd * 100.0) if total_meta_qtd > 0 else 0.0
    pct_vgv = (total_vgv_realizado / total_meta_vgv * 100.0) if total_meta_vgv > 0 else 0.0

    st.markdown(
        f"""
        <div class="vel-kpi-row" style="margin-top: 1rem;">
            <div class="vel-kpi"><div class="lbl">Qtd Meta</div><div class="val">{int(total_meta_qtd)}</div></div>
            <div class="vel-kpi"><div class="lbl">Qtd Realizado</div><div class="val">{fmt_qtd(total_realizado_qtd)}</div></div>
            <div class="vel-kpi"><div class="lbl">% Qtd</div><div class="val">{pct_qtd:.1f}%</div></div>
        </div>
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">VGV Meta</div><div class="val">{fmt_br_milhoes(total_meta_vgv)}</div></div>
            <div class="vel-kpi"><div class="lbl">VGV Realizado</div><div class="val val--red">{fmt_br_milhoes(total_vgv_realizado)}</div></div>
            <div class="vel-kpi"><div class="lbl">% VGV</div><div class="val">{pct_vgv:.1f}%</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Perfil das Vendas")
    qtd_facilitada = float(vendas_f[vendas_f["Tipo_Venda"] == "Facilitada"]["_qtd_venda"].sum())
    qtd_normal = float(vendas_f[vendas_f["Tipo_Venda"] == "Normal"]["_qtd_venda"].sum())
    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Vendas Facilitadas</div><div class="val">{fmt_qtd(qtd_facilitada)}</div></div>
            <div class="vel-kpi"><div class="lbl">Vendas Normais</div><div class="val">{fmt_qtd(qtd_normal)}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Visão geral")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        criar_medidor("Geral — quantidade vs meta", float(total_realizado_qtd), total_meta_qtd, total_vgv_realizado, total_meta_vgv, total_realizado_qtd)

    st.subheader("Por região")
    if "Regiao_Coord" in metas_f.columns:
        regioes_m = sorted(str(x).strip() for x in metas_f["Regiao_Coord"].dropna().unique() if str(x).strip())
        if regioes_m:
            cols = st.columns(min(3, len(regioes_m)) or 1)
            for i, regiao in enumerate(regioes_m):
                with cols[i % len(cols)]:
                    m_reg = metas_f[metas_f["Regiao_Coord"] == regiao]
                    v_reg = vendas_f[vendas_f["Regiao_Coord"] == regiao]
                    criar_medidor(regiao, float(v_reg["_qtd_venda"].sum()), m_reg["Meta_Qtd"].sum(), v_reg["_vgv_venda"].sum(), m_reg["Meta_VGV"].sum(), float(v_reg["_qtd_venda"].sum()))

    # -------------------------------------------------------------------------
    # TABELAS
    # -------------------------------------------------------------------------
    st.subheader("Tabela Resumo: Por Região")
    if "Regiao_Coord" in metas_f.columns:
        vg_reg = vendas_f.groupby("Regiao_Coord", as_index=False).agg(real_qtd=("_qtd_venda", "sum"), real_vgv=("_vgv_venda", "sum")).rename(columns={"Regiao_Coord": "Região"})
        mg_reg = metas_f.groupby("Regiao_Coord", as_index=False).agg(meta_qtd=("Meta_Qtd", "sum"), meta_vgv=("Meta_VGV", "sum")).rename(columns={"Regiao_Coord": "Região"})
        tab_reg = vg_reg.merge(mg_reg, on="Região", how="outer").fillna(0)
        tab_reg["% Qtd"] = tab_reg.apply(lambda r: (r["real_qtd"] / r["meta_qtd"] * 100.0) if r["meta_qtd"] > 0 else 0.0, axis=1)
        tab_reg["% VGV"] = tab_reg.apply(lambda r: (r["real_vgv"] / r["meta_vgv"] * 100.0) if r["meta_vgv"] > 0 else 0.0, axis=1)
        st.dataframe(tab_reg.sort_values("meta_qtd", ascending=False), use_container_width=True, hide_index=True)

    # -------------------------------------------------------------------------
    # FUNIL IDEAL E ENGENHARIA REVERSA
    # -------------------------------------------------------------------------
    st.markdown("<br><hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Engenharia Reversa: Funil Ideal")

    v_meta = math.floor(total_meta_qtd)
    pa_ideal = math.ceil(v_meta / 0.64) if v_meta > 0 else 0
    p_ideal = math.ceil(pa_ideal / 0.64) if pa_ideal > 0 else 0
    vi_ideal = math.ceil(p_ideal / 0.25) if p_ideal > 0 else 0
    a_ideal = math.ceil(vi_ideal / 0.50) if vi_ideal > 0 else 0
    
    meta_global_referencia = (total_meta_qtd / fator_meta) if fator_meta > 0 else 0
    meta_dvrj_ref = meta_global_referencia * 0.5
    vd_ideal = math.ceil(meta_dvrj_ref * 0.40)
    
    od_ideal = math.ceil(vd_ideal / 0.044) if vd_ideal > 0 else 0
    ld_ideal = math.ceil(od_ideal / 0.50) if od_ideal > 0 else 0
    
    corretores_pessimista = math.ceil(v_meta / 0.15) if v_meta > 0 else 0
    corretores_moderado = math.ceil(v_meta / 0.20) if v_meta > 0 else 0
    corretores_otimista = math.ceil(v_meta / 0.25) if v_meta > 0 else 0

    col_f_meta_espaco, col_f_meta, col_f_meta_espaco2 = st.columns([1, 2, 1])
    with col_f_meta:
        fig_ideal = _criar_fig_funil(
            ['Agendamentos', 'Visitas', 'Pastas', 'Past. Aprov.', 'Vendas (Meta)'],
            [a_ideal, vi_ideal, p_ideal, pa_ideal, v_meta],
            cores=["#022654", "#04428f", "#1e60b3", "#cb0935", "#9e0828"],
            altura=350,
        )
        st.plotly_chart(fig_ideal, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br><hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Funil de Marketing Digital")
    
    col_mkt_espaco, col_mkt_grafico, col_mkt_espaco2 = st.columns([1, 2, 1])
    with col_mkt_grafico:
        fig_mkt = _criar_fig_funil(
            ['Leads Digitais', 'Oport. Digitais', 'Vendas Dig. (40% DV RJ)'],
            [ld_ideal, od_ideal, vd_ideal],
            cores=["#022654", "#1e60b3", "#cb0935"],
            altura=300,
        )
        st.plotly_chart(fig_mkt, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br><hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Cenários: Corretores Ativos (Necessários para bater a meta global)")
    st.markdown(
        f"""
        <div class="vel-kpi-row" style="justify-content: center; margin-top: 1rem;">
            <div class="vel-kpi" style="flex: 0 1 300px;"><div class="lbl">Pessimista (15% convert.)</div><div class="val">{corretores_pessimista}</div></div>
            <div class="vel-kpi" style="flex: 0 1 300px;"><div class="lbl">Moderado (20% convert.)</div><div class="val">{corretores_moderado}</div></div>
            <div class="vel-kpi" style="flex: 0 1 300px;"><div class="lbl">Otimista (25% convert.)</div><div class="val">{corretores_otimista}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -------------------------------------------------------------------------
    # Projeção de Vendas (regressão diária — Contrato gerado em)
    # -------------------------------------------------------------------------
    mes_corrente = datetime.now().month
    metas_mes_atual = metas_f[metas_f["Mes_Num"] == mes_corrente] if "Mes_Num" in metas_f.columns else metas_f.iloc[0:0]
    if metas_mes_atual.empty and "Mes_Num" in df_metas.columns:
        # fallback: meta do mês corrente na base completa, com mesmos filtros de região/emp e fator de canal
        metas_mes_atual = df_metas[df_metas["Mes_Num"] == mes_corrente].copy()
        if regioes_sel:
            metas_mes_atual = metas_mes_atual[metas_mes_atual["Regiao_Coord"].isin(regioes_sel)]
        if emps_sel:
            metas_mes_atual = metas_mes_atual[metas_mes_atual["Empreendimento"].isin(emps_sel)]
        metas_mes_atual["Meta_VGV"] = metas_mes_atual["Meta_VGV"] * fator_meta
        if "Meta_Qtd" in metas_mes_atual.columns:
            metas_mes_atual["Meta_Qtd"] = (metas_mes_atual["Meta_Qtd"] * fator_meta).apply(math.floor)
    meta_vgv_proj = float(metas_mes_atual["Meta_VGV"].sum()) if not metas_mes_atual.empty else float(total_meta_vgv)
    meta_qtd_proj = float(metas_mes_atual["Meta_Qtd"].sum()) if (not metas_mes_atual.empty and "Meta_Qtd" in metas_mes_atual.columns) else float(total_meta_qtd)

    if col_contrato_gerado:
        # Base comercial; mesmos filtros de região/emp/canal do painel (RIO = todas as vendas).
        base_proj = df_vendas.copy()
        if regioes_sel:
            if "Região" in base_proj.columns:
                regioes_base = [r.split(" - ")[0].strip() for r in regioes_sel]
                base_proj = base_proj[
                    base_proj["Regiao_Coord"].isin(regioes_sel) | base_proj["Região"].isin(regioes_base)
                ]
            else:
                base_proj = base_proj[base_proj["Regiao_Coord"].isin(regioes_sel)]
        if emps_sel:
            base_proj = base_proj[base_proj["Empreendimento"].isin(emps_sel)]

        # Mesma regra do KPI: RIO (ou vazio) = 100% das vendas comerciais; senão, recorte por canal.
        if canais_sel and "RIO" not in canais_sel:
            mask_p = pd.Series(False, index=base_proj.index)
            if "DIR" in canais_sel:
                mask_p |= (base_proj["Canal_Agrupado"] == "DV RJ")
            if "PARC" in canais_sel and col_canal:
                mask_p |= base_proj[col_canal].astype(str).str.upper().str.strip().apply(
                    lambda x: x.split("-")[0].strip() == "RJG" or x == "RJG"
                )
            if "RJ" in canais_sel and col_canal:
                mask_p |= base_proj[col_canal].astype(str).str.upper().str.strip().apply(
                    lambda x: x.split("-")[0].strip() == "RJ" or x == "RJ"
                )
            base_proj = base_proj[mask_p]

        try:
            proj = projetar_vendas_mes_atual(
                base_proj, col_contrato_gerado, meta_vgv_proj, meta_qtd_mes=meta_qtd_proj
            )
            if proj:
                render_projecao_vendas(proj)
            else:
                st.info("Dados insuficientes para treinar a projeção de vendas (janela de meses exatos).")

            proj_sem_mes = projetar_vendas_mes_atual(
                base_proj,
                col_contrato_gerado,
                meta_vgv_proj,
                meta_qtd_mes=meta_qtd_proj,
                incluir_mes=False,
            )
            if proj_sem_mes:
                render_projecao_vendas(proj_sem_mes)

            # Efeitos sazonais relativos a segunda / dia 1 / janeiro
            hoje_ef = date.today()
            ini_ef, fim_ef = janela_treino_meses_exatos(hoje_ef)
            serie_ef = serie_diaria_contratos(base_proj, col_contrato_gerado)
            if not serie_ef.empty:
                treino_ef = calendario_diario(ini_ef, fim_ef, serie_ef)
                efeitos = estimar_efeitos_sazonais(treino_ef)
                if efeitos:
                    render_efeitos_sazonais(efeitos)

            # Projeção do funil — 3 relatórios Salesforce (agendamentos, pastas, vendas)
            try:
                # 1) Agendamentos / visitas
                try:
                    df_ag_vis, origem_ag = carregar_relatorio_salesforce(
                        SF_REPORT_AGENDAMENTOS_ID, rotulo="agendamentos/visitas"
                    )
                    n_ag_bruto = len(df_ag_vis)
                    df_ag_vis = deduplicar_agendamentos_funil(df_ag_vis)
                    st.caption(
                        f"Agendamentos/visitas: {origem_ag} · "
                        f"{n_ag_bruto:,} → {len(df_ag_vis):,} linhas "
                        f"(dedup Código do agendamento)"
                    )
                except Exception as e_sf:
                    st.warning(
                        f"SF agendamentos indisponível ({e_sf}). Fallback Sheets 'Dados Únicos'."
                    )
                    df_ag_vis = normalizar_colunas(
                        ler_planilha_aba_df(
                            SPREADSHEET_FUNIL_ID, ABA_AGENDAMENTOS_VISITAS, cred_fp
                        )
                    )
                    df_ag_vis = deduplicar_agendamentos_funil(df_ag_vis)

                # 2) Pastas / pastas aprovadas
                try:
                    df_pastas_funil, origem_pastas = carregar_relatorio_salesforce(
                        SF_REPORT_PASTAS_ID, rotulo="pastas"
                    )
                    n_pas_bruto = len(df_pastas_funil)
                    col_envio = achar_coluna_primeiro_envio_analise(df_pastas_funil)
                    col_safi = achar_coluna_aprovacao_safi(df_pastas_funil)
                    n_com_envio = 0
                    n_com_safi = 0
                    if col_envio:
                        n_com_envio = int(
                            parse_data_serie(df_pastas_funil[col_envio]).notna().sum()
                        )
                    if col_safi:
                        n_com_safi = int(
                            parse_data_serie(df_pastas_funil[col_safi]).notna().sum()
                        )
                    n_aprov_filt = len(deduplicar_pastas_aprovadas_funil(df_pastas_funil))
                    df_pastas_funil = deduplicar_pastas_funil(df_pastas_funil)
                    st.caption(
                        f"Pastas: {origem_pastas} · "
                        f"{n_pas_bruto:,} → {len(df_pastas_funil):,} linhas "
                        f"(dedup Nome da Avaliação) · "
                        f"1º envio: '{col_envio or '?'}' ({n_com_envio:,}) · "
                        f"Aprov. SAFI: '{col_safi or '?'}' ({n_com_safi:,}) · "
                        f"aprovadas (dedup): {n_aprov_filt:,}"
                    )
                except Exception as e_sf_p:
                    st.warning(
                        f"SF pastas indisponível ({e_sf_p}). Tentando planilha Sheets…"
                    )
                    sid_pastas = (SPREADSHEET_PASTAS_ID or "").strip()
                    df_pastas_funil, origem_pastas = carregar_df_pastas_funil(
                        SPREADSHEET_FUNIL_ID, sid, sid_pastas, cred_fp
                    )
                    if df_pastas_funil.empty:
                        st.warning("Pastas não carregadas — funil sem essa etapa.")
                    else:
                        df_pastas_funil = deduplicar_pastas_funil(df_pastas_funil)
                        st.caption(f"Pastas (Sheets): {origem_pastas}")

                # 3) Vendas (Contrato gerado em)
                df_vendas_funil = pd.DataFrame()
                serie_vendas_funil = None
                try:
                    df_vendas_funil, origem_vendas = carregar_relatorio_salesforce(
                        SF_REPORT_VENDAS_ID, rotulo="vendas"
                    )
                    n_ven_bruto = len(df_vendas_funil)
                    df_vendas_funil = filtrar_vendas_comerciais(df_vendas_funil)
                    n_ven_comercial = len(df_vendas_funil)
                    df_vendas_funil = deduplicar_vendas_funil(df_vendas_funil)
                    st.caption(
                        f"Vendas: {origem_vendas} · "
                        f"{n_ven_bruto:,} → {n_ven_comercial:,} comerciais → "
                        f"{len(df_vendas_funil):,} linhas "
                        f"(dedup ID da Oportunidade)"
                    )
                except Exception as e_sf_v:
                    st.warning(
                        f"SF vendas indisponível ({e_sf_v}). "
                        "Usando vendas filtradas do painel (Contrato gerado em)."
                    )
                    serie_vendas_funil = serie_diaria_contratos(base_proj, col_contrato_gerado)

                mapas_funil = montar_mapa_funil_diario(
                    df_ag_vis,
                    df_pastas_funil if df_pastas_funil is not None else pd.DataFrame(),
                    serie_vendas=serie_vendas_funil,
                    df_vendas=df_vendas_funil if not df_vendas_funil.empty else None,
                )
                proj_funil = projetar_funil_mes_atual(
                    mapas_funil, incluir_mes=True, meta_qtd_mes=meta_qtd_proj
                )
                if proj_funil:
                    render_projecao_funil(proj_funil)
                else:
                    st.info("Dados insuficientes para a projeção do funil comercial.")
            except Exception as exc_funil:
                st.warning(f"Não foi possível calcular a projeção do funil: {exc_funil}")
        except Exception as exc:
            st.warning(f"Não foi possível calcular a projeção de vendas: {exc}")
    else:
        st.warning("Coluna 'Contrato gerado em' não encontrada — seção de Projeção de Vendas indisponível.")

    # -------------------------------------------------------------------------
    # Comparativo de Vendas Eficiência Isolado (Janela Histórica: Dia 1 ao Dia Atual MTD)
    # -------------------------------------------------------------------------
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    dia_atual_janela = datetime.now().day
    st.subheader(f"Comparativo de Vendas (Dia 01 ao Dia {dia_atual_janela:02d} do Mês)")
    
    if col_contrato_gerado:
        # Base de espelho limpa para o gráfico de eficiência temporal sem interferência de filtros de UI de competência
        df_grafico_eficiencia = df_vendas.copy()
        df_grafico_eficiencia["Data_Contrato_DT"] = parse_data_serie(df_grafico_eficiencia[col_contrato_gerado])
        df_grafico_eficiencia = df_grafico_eficiencia.dropna(subset=["Data_Contrato_DT"])
        
        if not df_grafico_eficiencia.empty:
            # Trava a janela da série histórica estritamente do dia 01 ao dia atual de cada mês para medição justa de ritmo MTD
            df_parcial_janela = df_grafico_eficiencia[df_grafico_eficiencia["Data_Contrato_DT"].dt.day <= dia_atual_janela].copy()
            
            df_parcial_janela["_ano_c"] = df_parcial_janela["Data_Contrato_DT"].dt.year
            df_parcial_janela["_mes_c"] = df_parcial_janela["Data_Contrato_DT"].dt.month
            
            df_comp = df_parcial_janela.groupby(["_ano_c", "_mes_c"], as_index=False).agg(
                QTD=("_qtd_venda", "sum"),
                VGV=("_vgv_venda", "sum")
            ).sort_values(["_ano_c", "_mes_c"])
            
            df_comp["Periodo"] = df_comp["_mes_c"].astype(str).str.zfill(2) + "/" + df_comp["_ano_c"].astype(str)
            df_comp["VGV_Formatado"] = df_comp["VGV"].apply(lambda x: fmt_br_milhoes(x))
            df_comp["QTD_Formatado"] = df_comp["QTD"].apply(lambda x: fmt_qtd(x))
            
            fig_linha = make_subplots(specs=[[{"secondary_y": True}]])
            
            fig_linha.add_trace(
                go.Scatter(
                    x=df_comp["Periodo"], 
                    y=df_comp["QTD"], 
                    mode="lines+markers+text",
                    name="QTD Vendas",
                    line=dict(color=COR_AZUL_ESC, width=3),
                    marker=dict(size=8, color=COR_AZUL_ESC),
                    text=df_comp["QTD_Formatado"],
                    textposition="top center",
                    textfont=dict(color=COR_AZUL_ESC, size=11, family="Inter")
                ),
                secondary_y=False,
            )
            
            fig_linha.add_trace(
                go.Scatter(
                    x=df_comp["Periodo"], 
                    y=df_comp["VGV"], 
                    mode="lines+markers+text",
                    name="VGV Real",
                    line=dict(color=COR_VERMELHO, width=3),
                    marker=dict(size=8, color=COR_VERMELHO),
                    text=df_comp["VGV_Formatado"],
                    textposition="bottom center",
                    textfont=dict(color=COR_VERMELHO, size=11, family="Inter")
                ),
                secondary_y=True,
            )
            
            fig_linha.update_layout(
                margin=dict(l=20, r=20, t=40, b=20),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="Inter", color=COR_TEXTO_PRETO),
                legend=dict(
                    orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
                    font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
                ),
                hovermode="x unified",
            )

            fig_linha.update_xaxes(
                title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
                tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
            )
            fig_linha.update_yaxes(
                title_text="Quantidade (Vendas)",
                title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
                tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
                secondary_y=False,
                showgrid=False,
            )
            fig_linha.update_yaxes(
                title_text="VGV Real (R$)",
                title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
                tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
                secondary_y=True,
                showgrid=True,
                gridcolor="rgba(226, 232, 240, 0.5)",
            )
            st.plotly_chart(fig_linha, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Não há dados de vendas no período acumulado de eficiência para exibir.")
    else:
        st.warning("A coluna de Contrato Gerado em não foi encontrada. Impossível renderizar a linha do tempo.")

    st.markdown(
        f'<div class="footer" style="text-align:center;padding:1rem 0;color:{COR_TEXTO_PRETO};font-size:0.82rem;">'
        f"Direcional Engenharia · Vendas — Acompanhamento de metas</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
