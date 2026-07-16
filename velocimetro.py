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
COR_TEXTO_MUTED = "#64748b"
COR_TEXTO_LABEL = "#1e293b"
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
        f'<p class="ficha-sub">Realizado vs metas por período.</p>'
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
        h1, h2, h3, h4 {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_AZUL_ESC} !important;
            font-weight: 800 !important;
            text-align: center !important;
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
        .ficha-hero .ficha-sub {{
            color: #475569;
            font-size: 0.95rem;
            margin: 0.45rem 0 0 0;
            line-height: 1.45;
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
            color: {COR_AZUL_ESC};
            opacity: 0.85;
        }}
        .vel-kpi .val {{
            font-family: 'Montserrat', sans-serif;
            font-size: 1.35rem;
            font-weight: 800;
            color: {COR_AZUL_ESC};
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


def achar_coluna(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    for a in aliases:
        for c in df.columns:
            if a.lower() == str(c).strip().lower(): return c
    for a in aliases:
        for c in df.columns:
            if a.lower() in str(c).strip().lower(): return c
    return None


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


def janela_treino_52_semanas(hoje: Optional[date] = None) -> Tuple[date, date]:
    """
    Corte de 1 ano: do dia atual menos 52 semanas até o dia da semana
    anterior ao dia atual (ontem).
    """
    hoje = hoje or date.today()
    fim = hoje - timedelta(days=1)
    inicio = hoje - timedelta(weeks=52)
    return inicio, fim


def serie_diaria_contratos(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    col_qtd: str = "_qtd_venda",
    col_vgv: str = "_vgv_venda",
) -> pd.DataFrame:
    """Agrega vendas comerciais por data de 'Contrato gerado em'."""
    base = df_vendas.copy()
    base["_dt_contrato"] = pd.to_datetime(base[col_contrato], dayfirst=True, errors="coerce")
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


def _matriz_explicativas(df: pd.DataFrame) -> np.ndarray:
    """One-hot (numpy) de dia do mês, dia da semana e mês + intercepto."""
    n = len(df)
    # 31 dias + 7 dias da semana + 12 meses + intercepto
    X = np.zeros((n, 31 + 7 + 12 + 1), dtype=float)
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
        ms = meses_idx.get(str(row.mes), None)
        if ms is not None:
            X[i, 31 + 7 + (ms - 1)] = 1.0
    return X


def treinar_regressao_vendas_diarias(treino: pd.DataFrame) -> np.ndarray:
    """OLS via numpy.linalg.lstsq. Retorna vetor de coeficientes."""
    X = _matriz_explicativas(treino)
    y = treino["qtd"].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    return coef


def prever_qtd_dias(coef: np.ndarray, datas: List[date]) -> np.ndarray:
    if not datas:
        return np.array([])
    df = pd.DataFrame({"data": datas})
    df["dia_mes"] = df["data"].map(lambda d: d.day)
    df["dia_semana"] = df["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    df["mes"] = df["data"].map(lambda d: MESES_PT[d.month])
    X = _matriz_explicativas(df)
    pred = X @ coef
    return np.maximum(pred, 0.0)


def _r2_treino(treino: pd.DataFrame, coef: np.ndarray) -> float:
    X = _matriz_explicativas(treino)
    y = treino["qtd"].astype(float).values
    y_hat = X @ coef
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    if ss_tot <= 0:
        return 0.0
    return 1.0 - (ss_res / ss_tot)


def projetar_vendas_mes_atual(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    meta_vgv_mes: float,
    hoje: Optional[date] = None,
) -> Optional[Dict[str, Any]]:
    """
    Regressão de qtd/dia ~ dia_mês + dia_semana + mês, janela de 52 semanas,
    e projeção do restante do mês corrente (somente vendas comerciais na base).
    """
    hoje = hoje or date.today()
    inicio, fim_treino = janela_treino_52_semanas(hoje)
    serie = serie_diaria_contratos(df_vendas, col_contrato)
    if serie.empty:
        return None

    treino = calendario_diario(inicio, fim_treino, serie)
    if treino["qtd"].sum() <= 0 or len(treino) < 30:
        return None

    modelo = treinar_regressao_vendas_diarias(treino)

    ano, mes = hoje.year, hoje.month
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    dias_mes = [date(ano, mes, d) for d in range(1, ultimo_dia + 1)]
    dias_passados = [d for d in dias_mes if d <= hoje]
    dias_futuros = [d for d in dias_mes if d > hoje]

    mapa_real = {r["data"]: float(r["qtd"]) for _, r in serie.iterrows()}
    mapa_vgv = {r["data"]: float(r["vgv"]) for _, r in serie.iterrows()}

    qtd_atingida_por_dia = [mapa_real.get(d, 0.0) for d in dias_passados]
    vgv_mtd = sum(mapa_vgv.get(d, 0.0) for d in dias_passados)
    qtd_mtd = float(sum(qtd_atingida_por_dia))

    pred_futuro = prever_qtd_dias(modelo, dias_futuros)
    qtd_projetada_restante = float(pred_futuro.sum()) if len(pred_futuro) else 0.0
    qtd_projetada_mes = qtd_mtd + qtd_projetada_restante

    ticket_medio = (vgv_mtd / qtd_mtd) if qtd_mtd > 0 else 0.0
    if ticket_medio <= 0:
        # fallback: ticket médio da janela de treino
        q_tr = float(treino["qtd"].sum())
        v_tr = float(treino["vgv"].sum())
        ticket_medio = (v_tr / q_tr) if q_tr > 0 else 0.0

    vgv_projetado = vgv_mtd + (qtd_projetada_restante * ticket_medio)
    pct_vgv = (vgv_projetado / meta_vgv_mes * 100.0) if meta_vgv_mes and meta_vgv_mes > 0 else 0.0

    # Série do gráfico: atingido até hoje + projetado no restante
    qtd_diaria_plot = []
    for d in dias_passados:
        qtd_diaria_plot.append({"dia": d.day, "qtd": mapa_real.get(d, 0.0), "tipo": "Atingida"})
    for i, d in enumerate(dias_futuros):
        qtd_diaria_plot.append({"dia": d.day, "qtd": float(pred_futuro[i]), "tipo": "Projetada"})

    # Acumulado: real até hoje, depois projeta a partir do acumulado MTD
    acum = []
    running = 0.0
    for d in dias_passados:
        running += mapa_real.get(d, 0.0)
        acum.append({"dia": d.day, "acum": running, "tipo": "Atingida"})
    for i, d in enumerate(dias_futuros):
        running += float(pred_futuro[i])
        acum.append({"dia": d.day, "acum": running, "tipo": "Projetada"})

    X_tr_r2 = _r2_treino(treino, modelo)

    return {
        "hoje": hoje,
        "inicio_treino": inicio,
        "fim_treino": fim_treino,
        "qtd_mtd": qtd_mtd,
        "vgv_mtd": vgv_mtd,
        "qtd_projetada_mes": qtd_projetada_mes,
        "vgv_projetado": vgv_projetado,
        "pct_vgv_meta": pct_vgv,
        "meta_vgv_mes": meta_vgv_mes,
        "ticket_medio": ticket_medio,
        "ultimo_dia": ultimo_dia,
        "diaria": pd.DataFrame(qtd_diaria_plot),
        "acumulado": pd.DataFrame(acum),
        "r2_treino": X_tr_r2,
    }


def render_projecao_vendas(proj: Dict[str, Any]) -> None:
    """Seção Streamlit: cartões + gráfico atingido vs projetado."""
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Projeção de Vendas")
    st.caption(
        f"Regressão diária (dia do mês + dia da semana + mês) · treino "
        f"{proj['inicio_treino'].strftime('%d/%m/%Y')} a {proj['fim_treino'].strftime('%d/%m/%Y')} "
        f"(52 semanas, até o dia anterior) · R² treino: {proj['r2_treino']:.2f} · "
        f"MTD atingido (Contrato gerado em): {fmt_qtd(proj['qtd_mtd'])} vendas / {fmt_br_milhoes(proj['vgv_mtd'])}"
    )

    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi">
                <div class="lbl">Qtd. vendas (projetada)</div>
                <div class="val">{fmt_qtd(proj['qtd_projetada_mes'])}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">VGV vendas (projetado)</div>
                <div class="val val--red">{fmt_br_milhoes(proj['vgv_projetado'])}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">% VGV atingido (proj. vs meta)</div>
                <div class="val">{proj['pct_vgv_meta']:.1f}%</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    df_d = proj["diaria"]
    df_a = proj["acumulado"]
    dia_hoje = proj["hoje"].day

    ating_d = df_d[df_d["tipo"] == "Atingida"]
    proj_d = df_d[df_d["tipo"] == "Projetada"]
    ating_a = df_a[df_a["tipo"] == "Atingida"]
    proj_a = df_a[df_a["tipo"] == "Projetada"]

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=ating_d["dia"],
            y=ating_d["qtd"],
            name="Qtd atingida (dia)",
            marker_color=COR_AZUL_ESC,
            opacity=0.85,
        ),
        secondary_y=False,
    )
    if not proj_d.empty:
        fig.add_trace(
            go.Bar(
                x=proj_d["dia"],
                y=proj_d["qtd"],
                name="Qtd projetada (dia)",
                marker_color="rgba(203, 9, 53, 0.45)",
            ),
            secondary_y=False,
        )

    fig.add_trace(
        go.Scatter(
            x=ating_a["dia"],
            y=ating_a["acum"],
            mode="lines+markers",
            name="Acumulado atingido",
            line=dict(color=COR_AZUL_ESC, width=3),
            marker=dict(size=7),
        ),
        secondary_y=True,
    )
    if not proj_a.empty:
        # Liga o último ponto atingido ao primeiro projetado
        x_proj = [dia_hoje] + proj_a["dia"].tolist()
        y_proj = [float(ating_a["acum"].iloc[-1]) if not ating_a.empty else 0.0] + proj_a["acum"].tolist()
        fig.add_trace(
            go.Scatter(
                x=x_proj,
                y=y_proj,
                mode="lines+markers",
                name="Acumulado projetado",
                line=dict(color=COR_VERMELHO, width=3, dash="dash"),
                marker=dict(size=7, color=COR_VERMELHO),
            ),
            secondary_y=True,
        )

    fig.add_vline(
        x=dia_hoje,
        line_width=1,
        line_dash="dot",
        line_color=COR_TEXTO_MUTED,
        annotation_text="Hoje",
        annotation_position="top",
    )

    fig.update_layout(
        barmode="overlay",
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_LABEL),
        legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5),
        hovermode="x unified",
        height=420,
    )
    fig.update_xaxes(title_text="Dia do mês", dtick=1, range=[0.5, proj["ultimo_dia"] + 0.5])
    fig.update_yaxes(title_text="Qtd. no dia", secondary_y=False, showgrid=False)
    fig.update_yaxes(title_text="Qtd. acumulada", secondary_y=True, showgrid=True, gridcolor="rgba(226,232,240,0.5)")

    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


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
                "font": {"size": 16, "family": "Montserrat", "color": COR_AZUL_ESC},
            },
            gauge={
                "axis": {"range": [0, axis_max], "tickwidth": 1, "tickcolor": COR_TEXTO_MUTED},
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
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown(
        f"""
        <div style="text-align:center;font-size:0.85rem;color:{COR_TEXTO_LABEL};margin-top:-8px;line-height:1.4;">
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
    col_venda_comercial = achar_coluna(df_vendas, ["Venda Comercial?", "Venda Comercial"])
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
        mask_venda = (
            (pd.to_numeric(df_vendas[col_venda_comercial], errors='coerce') == 1) |
            (df_vendas[col_venda_comercial].astype(str).str.strip().str.upper() == 'SIM') |
            (df_vendas[col_venda_comercial].astype(str).str.strip().str.upper() == 'TRUE')
        )
        df_vendas = df_vendas[mask_venda]
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
        fig_ideal = go.Figure(go.Funnel(
            y=['Agendamentos', 'Visitas', 'Pastas', 'Past. Aprov.', 'Vendas (Meta)'],
            x=[a_ideal, vi_ideal, p_ideal, pa_ideal, v_meta],
            textinfo="value",
            marker={"color": ["#022654", "#04428f", "#1e60b3", "#cb0935", "#9e0828"]},
            connector={"fillcolor": "rgba(4, 66, 143, 0.15)"}
        ))
        fig_ideal.update_layout(margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, font=dict(family="Inter", color=COR_TEXTO_LABEL))
        st.plotly_chart(fig_ideal, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br><hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Funil de Marketing Digital")
    
    col_mkt_espaco, col_mkt_grafico, col_mkt_espaco2 = st.columns([1, 2, 1])
    with col_mkt_grafico:
        fig_mkt = go.Figure(go.Funnel(
            y=['Leads Digitais', 'Oport. Digitais', 'Vendas Dig. (40% DV RJ)'],
            x=[ld_ideal, od_ideal, vd_ideal],
            textinfo="value",
            marker={"color": ["#022654", "#1e60b3", "#cb0935"]},
            connector={"fillcolor": "rgba(4, 66, 143, 0.15)"}
        ))
        fig_mkt.update_layout(margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=300, font=dict(family="Inter", color=COR_TEXTO_LABEL))
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
    meta_vgv_proj = float(metas_mes_atual["Meta_VGV"].sum()) if not metas_mes_atual.empty else float(total_meta_vgv)

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
            proj = projetar_vendas_mes_atual(base_proj, col_contrato_gerado, meta_vgv_proj)
            if proj:
                render_projecao_vendas(proj)
            else:
                st.info("Dados insuficientes para treinar a projeção de vendas (janela de 52 semanas).")
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
        df_grafico_eficiencia["Data_Contrato_DT"] = pd.to_datetime(df_grafico_eficiencia[col_contrato_gerado], dayfirst=True, errors="coerce")
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
                font=dict(family="Inter", color=COR_TEXTO_LABEL),
                legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5),
                hovermode="x unified"
            )
            
            fig_linha.update_yaxes(title_text="Quantidade (Vendas)", secondary_y=False, showgrid=False)
            fig_linha.update_yaxes(title_text="VGV Real (R$)", secondary_y=True, showgrid=True, gridcolor="rgba(226, 232, 240, 0.5)")
            
            st.plotly_chart(fig_linha, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Não há dados de vendas no período acumulado de eficiência para exibir.")
    else:
        st.warning("A coluna de Contrato Gerado em não foi encontrada. Impossível renderizar a linha do tempo.")

    st.markdown(
        f'<div class="footer" style="text-align:center;padding:1rem 0;color:{COR_TEXTO_MUTED};font-size:0.82rem;">'
        f"Direcional Engenharia · Vendas — Acompanhamento de metas</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
