# -*- coding: utf-8 -*-
"""
Acompanhamento de vendas — metas vs realizado (Direcional).
Planilha: BD Vendas Completa + Metas.
Dependências: streamlit, pandas, plotly, gspread, google-auth

Secrets: `.streamlit/secrets.toml` com seções `[connections.gsheets]` (JSON da service account:
type, project_id, private_key_id, private_key em bloco multilinha TOML, client_email, URIs, etc.)
e opcional `spreadsheet_id`. Seção `[email]` para SMTP (smtp_server, smtp_port, sender_email,
sender_password) — lida pelo app mas não usada na leitura da planilha.
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
from plotly.subplots import make_subplots
import streamlit as st

# -----------------------------------------------------------------------------
# Identificação da planilha e Arquivos Visuais
# -----------------------------------------------------------------------------
SPREADSHEET_ID = "1wpuNQvksot9CLhGgQRe7JlyDeRISEh_sc3-6VRDyQYk"

WS_VENDAS = "BD Vendas Completa"
WS_METAS = "Metas"

# Novas abas adicionadas conforme solicitação
WS_METAS_IMOB = "Metas Coordenadores IMOB"
WS_METAS_COMERCIAIS = "Metas Coordenadores Comerciais"
WS_METAS_GC = "Metas Grandes Contas"
WS_DICIONARIO = "Dicionário Coordenadores"

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
    """String para `url(...)` no CSS: data-URL ou URL https (fallback)."""
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
    fid = (os.environ.get("DIRECIONAL_LOGO_FILE_ID") or "").strip()
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
            background: rgba(255,255,255,0.6);
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
    if not info: raise ValueError("Credenciais [connections.gsheets] ausentes ou incompletas.")
    return ler_aba_gsheets(info, spreadsheet_id, worksheet)


def parse_valor_br(val: Any) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace("R$", "").replace(" ", "").replace("\xa0", "")
    if not s or s.lower() == "nan": return 0.0
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s: return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0


def extrair_mes(val: Any) -> Optional[int]:
    if pd.isna(val): return None
    v = str(val).strip().lower()
    if not v: return None
    v_date = v.split()[0]
    if '-' in v_date:
        p = v_date.split('-')
        if len(p) == 3 and len(p[0]) == 4:
            try:
                m = int(p[1])
                if 1 <= m <= 12: return m
            except: pass
    for m_str, m_num in MESES_TEXTO_MAP.items():
        if m_str in v_date: return m_num
    try:
        m = int(float(v_date))
        if 1 <= m <= 12: return m
    except ValueError: pass
    if '/' in v_date:
        p = v_date.split('/')
        if len(p) == 3: 
            try:
                m = int(p[1])
                if 1 <= m <= 12: return m
            except: pass
        elif len(p) == 2: 
            try:
                m = int(p[0])
                if 1 <= m <= 12: return m
            except: pass
    return None


def extrair_ano(val: Any) -> Optional[int]:
    if pd.isna(val): return None
    v = str(val).strip()
    if not v: return None
    v_date = v.split()[0]
    if '-' in v_date:
        p = v_date.split('-')
        if len(p) == 3 and len(p[0]) == 4:
            try:
                ano = int(p[0])
                if ano > 2000: return ano
            except: pass
    try:
        ano = int(float(v_date))
        if ano > 2000: return ano
    except ValueError: pass
    if '/' in v_date:
        p = v_date.split('/')
        try:
            ano = int(p[-1])
            if ano > 2000: return ano
        except: pass
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


# -----------------------------------------------------------------------------
# Lógicas Auxiliares para Novas Abas
# -----------------------------------------------------------------------------
def extrair_lista_coords(coordenadores_str: str) -> List[str]:
    """Extrai nomes de {nome1, nome2} ou strings separadas por vírgula."""
    s = str(coordenadores_str).strip()
    if s.startswith("{") and s.endswith("}"):
        s = s[1:-1]
    return [c.strip() for c in s.split(",") if c.strip()]


def mapear_vendas_por_coordenador(df_vendas: pd.DataFrame, df_dic: pd.DataFrame, empreendimento: str = None) -> pd.DataFrame:
    """Filtra vendas cruzando Dicionário (Proprietário -> Coordenador)."""
    # Monta dict de Proprietário -> Coordenador
    dic_map = {}
    for _, row in df_dic.iterrows():
        coord = str(row.iloc[0]).strip()
        prop = str(row.iloc[1]).strip()
        if coord and prop:
            dic_map[prop.lower()] = coord
            
    vendas_map = df_vendas.copy()
    # Identifica coluna de proprietário (já feito no main via col_proprietario, mas garantindo)
    col_prop = achar_coluna(vendas_map, ["Proprietário da oportunidade", "Proprietario da oportunidade", "Nome da conta", "Proprietario", "Corretor"])
    
    if col_prop:
        vendas_map["_Coordenador_Mapeado"] = vendas_map[col_prop].astype(str).str.strip().str.lower().map(dic_map)
    else:
        vendas_map["_Coordenador_Mapeado"] = None
        
    if empreendimento:
        col_emp = achar_coluna(vendas_map, ["Empreendimento", "Obra", "Nome do Empreendimento"])
        if col_emp:
            vendas_map = vendas_map[vendas_map[col_emp].astype(str).str.strip().str.lower() == empreendimento.strip().lower()]
            
    return vendas_map


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
        df_vendas = ler_planilha_aba_df(sid, WS_VENDAS, cred_fp)
        df_metas_raw = ler_planilha_aba_df(sid, WS_METAS, cred_fp)
        
        # Tentativa de leitura das novas abas (evita quebrar se não existirem ainda)
        try: df_metas_imob = ler_planilha_aba_df(sid, WS_METAS_IMOB, cred_fp)
        except: df_metas_imob = pd.DataFrame()
        
        try: df_metas_comerciais = ler_planilha_aba_df(sid, WS_METAS_COMERCIAIS, cred_fp)
        except: df_metas_comerciais = pd.DataFrame()
        
        try: df_metas_gc = ler_planilha_aba_df(sid, WS_METAS_GC, cred_fp)
        except: df_metas_gc = pd.DataFrame()
        
        try: df_dic = ler_planilha_aba_df(sid, WS_DICIONARIO, cred_fp)
        except: df_dic = pd.DataFrame()
        
    except Exception as e:
        st.error(f"Erro ao ler as planilhas: {str(e)}")
        return

    df_vendas = normalizar_colunas(df_vendas)
    df_metas_melted = melt_metas(df_metas_raw)

    # -------------------------------------------------------------------------
    # Filtros Básicos de Metas (Visão Geral)
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
            new_row["Meta_Qtd"] = float(row["Meta_Qtd"]) / n
            new_row["Meta_VGV"] = float(row["Meta_VGV"]) / n
            new_row["_peso_coord"] = 1.0 / n
            rows_metas.append(new_row)
            
    df_metas = pd.DataFrame(rows_metas)

    # -------------------------------------------------------------------------
    # Colunas Vendas
    # -------------------------------------------------------------------------
    col_ano = achar_coluna(df_vendas, ["Ano da Venda", "Ano Venda", "Ano"])
    col_mes = achar_coluna(df_vendas, ["Mês da Venda - Looker", "Mês da Venda", "Mês Venda", "Mes Venda", "Mês", "Mes"])
    col_regiao = achar_coluna(df_vendas, ["Região", "Regiao"])
    col_canal = achar_coluna(df_vendas, ["Canal"])
    col_valor = achar_coluna(df_vendas, ["Valor Real de Venda", "Valor Real", "Valor"])
    col_emp = achar_coluna(df_vendas, ["Empreendimento", "Obra", "Nome do Empreendimento"])
    col_venda_comercial = achar_coluna(df_vendas, ["Venda Comercial?", "Venda Comercial"])
    col_venda_facilitada = achar_coluna(df_vendas, ["Venda facilitada", "Venda Facilitada", "Venda Facilitada?"])
    col_proprietario = achar_coluna(df_vendas, ["Proprietário da oportunidade", "Proprietario da oportunidade", "Nome da conta", "Proprietario", "Corretor"])
    col_ranking = achar_coluna(df_vendas, ["Ranking"])
    col_data_venda = achar_coluna(df_vendas, ["Data da venda", "Data Venda", "Data de venda", "Data"])

    if not col_ano and not col_mes:
        st.error("Não foi possível encontrar as colunas de Ano e Mês na aba de vendas.")
        return

    if col_emp and col_emp != "Empreendimento":
        df_vendas.rename(columns={col_emp: "Empreendimento"}, inplace=True)
        col_emp = "Empreendimento"
    if col_regiao and col_regiao != "Região":
        df_vendas.rename(columns={col_regiao: "Região"}, inplace=True)
        col_regiao = "Região"

    if col_emp:
        df_vendas = df_vendas[~df_vendas[col_emp].astype(str).str.strip().str.lower().isin(["total", "geral", ""])]

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

    df_vendas["_mes_raw"] = df_vendas[col_mes].apply(extrair_mes) if col_mes else None
    def aplicar_fallback_mes(row: pd.Series) -> Optional[int]:
        m = row["_mes_raw"]
        if pd.notna(m) and 1 <= m <= 12: return int(m)
        if col_data_venda and pd.notna(row[col_data_venda]):
            m2 = extrair_mes(row[col_data_venda])
            if m2 and 1 <= m2 <= 12: return int(m2)
        return None
    df_vendas["_mes"] = df_vendas.apply(aplicar_fallback_mes, axis=1)

    df_vendas["_ano_raw"] = df_vendas[col_ano].apply(extrair_ano) if col_ano else None
    def aplicar_fallback_ano(row: pd.Series) -> Optional[int]:
        ano = row["_ano_raw"]
        if pd.notna(ano) and ano > 2000:
            return int(ano)
        if col_mes and pd.notna(row[col_mes]):
            a = extrair_ano(row[col_mes])
            if a and a > 2000: return int(a)
        if col_data_venda and pd.notna(row[col_data_venda]):
            a = extrair_ano(row[col_data_venda])
            if a and a > 2000: return int(a)
        return None
    df_vendas["_ano"] = df_vendas.apply(aplicar_fallback_ano, axis=1)

    df_vendas["_vgv"] = df_vendas[col_valor].map(parse_valor_br) if col_valor else 0.0

    if col_canal:
        def agrupar_canal(c: Any) -> str:
            c_str = str(c).strip().upper()
            prefixo = c_str.split('-')[0].strip()
            # Canais que compõem a DV RJ / RIV (Internos)
            if prefixo in ['RJ', 'RJG'] or c_str in ['RJ', 'RJG']:
                return 'IMOB'
            return 'DV RJ'
        df_vendas['Canal_Agrupado'] = df_vendas[col_canal].apply(agrupar_canal)
    else:
        df_vendas['Canal_Agrupado'] = 'DV RJ'

    # -------------------------------------------------------------------------
    # Multiplicação e Distribuição das Vendas de Acordo com Coordenador (Peso)
    # -------------------------------------------------------------------------
    map_emp_regiao = df_metas[["Empreendimento", "Regiao_Coord", "_peso_coord"]].drop_duplicates()
    df_vendas = df_vendas.merge(map_emp_regiao, on="Empreendimento", how="left")
    df_vendas["_peso_coord"] = df_vendas["_peso_coord"].fillna(1.0)
    df_vendas["Regiao_Coord"] = df_vendas["Regiao_Coord"].fillna(df_vendas.get("Região", "Não Informado"))
    df_vendas["_qtd_venda"] = 1.0 * df_vendas["_peso_coord"]
    df_vendas["_vgv_venda"] = df_vendas["_vgv"] * df_vendas["_peso_coord"]

    # -------------------------------------------------------------------------
    # LINHA ÚNICA DE FILTROS GLOBAIS
    # -------------------------------------------------------------------------
    anos_disponiveis = sorted(int(x) for x in df_vendas["_ano"].dropna().unique().tolist() if pd.notna(x) and x > 2000)
    meses_no_ano = list(range(1, 13))
    mes_atual = datetime.now().month
    mes_padrao = mes_atual if mes_atual in meses_no_ano else 1
    regioes_disponiveis = sorted(set(str(x).strip() for x in df_metas["Regiao_Coord"].dropna().unique() if str(x).strip()))
    
    emps_comuns = []
    if coluna_existe(df_vendas, "Empreendimento") and coluna_existe(df_metas, "Empreendimento"):
        emps_vendas = set(str(x).strip() for x in df_vendas["Empreendimento"].dropna().unique() if str(x).strip())
        emps_metas = set(str(x).strip() for x in df_metas["Empreendimento"].dropna().unique() if str(x).strip())
        emps_comuns = sorted(list(emps_vendas & emps_metas))

    st.markdown("<div style='margin-bottom:1rem; text-align: center;'><strong>Filtros</strong></div>", unsafe_allow_html=True)
    
    col_filtros = st.columns(5)
    with col_filtros[0]:
        canais_sel = st.multiselect("Canal da Meta", ["RIO", "DIR", "PARC", "RJ"], default=["RIO"])
    with col_filtros[1]:
        anos_sel = st.multiselect("Ano", anos_disponiveis, default=[anos_disponiveis[-1]] if anos_disponiveis else [])
    with col_filtros[2]:
        meses_sel = st.multiselect("Mês", meses_no_ano, default=[mes_padrao])
    with col_filtros[3]:
        regioes_sel = st.multiselect("Região", regioes_disponiveis, default=[])
    with col_filtros[4]:
        emps_sel = st.multiselect("Empreendimento", emps_comuns, default=[])

    # String formatada de Data (MM/YYYY) baseada nos filtros (usado nas novas abas)
    meses_str_format = [f"{m:02d}" for m in meses_sel]
    datas_filtradas = [f"{m}/{a}" for m in meses_str_format for a in anos_sel] if anos_sel and meses_sel else []

    # -------------------------------------------------------------------------
    # Aplicação de Filtros (Geral)
    # -------------------------------------------------------------------------
    vendas_f = df_vendas.copy()
    metas_f = df_metas.copy()

    if anos_sel: vendas_f = vendas_f[vendas_f["_ano"].isin(anos_sel)]
    if meses_sel:
        vendas_f = vendas_f[vendas_f["_mes"].isin(meses_sel)]
        metas_f = metas_f[metas_f["Mes_Num"].isin(meses_sel)]
    if regioes_sel:
        metas_f = metas_f[metas_f["Regiao_Coord"].isin(regioes_sel)]
        vendas_f = vendas_f[vendas_f["Regiao_Coord"].isin(regioes_sel)]
    if emps_sel:
        metas_f = metas_f[metas_f["Empreendimento"].isin(emps_sel)]
        vendas_f = vendas_f[vendas_f["Empreendimento"].isin(emps_sel)]

    fator_meta = 0.0
    mask_vendas = pd.Series(False, index=vendas_f.index)

    # Lógica de canais para cálculo de fatores e meta proporcional
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

    fator_meta = min(1.0, fator_meta)
    vendas_f = vendas_f[mask_vendas]

    # Meta absoluta utilizada no cálculo do funil
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

    # -------------------------------------------------------------------------
    # NOVA PLATAFORMA DE PREMIAÇÕES (Abas)
    # -------------------------------------------------------------------------
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1.5rem 0;'/>", unsafe_allow_html=True)
    st.markdown("<h2>Painel de Premiações e Metas</h2>", unsafe_allow_html=True)
    
    tabs = st.tabs(["Visão Geral (Original)", "Coordenadores IMOB", "Coordenadores Comerciais", "Grandes Contas"])

    # =========================================================================
    # TAB 0: Visão Geral (Mantendo a funcionalidade original integralmente)
    # =========================================================================
    with tabs[0]:
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
                cols_reg = st.columns(min(3, len(regioes_m)) or 1)
                for i, regiao in enumerate(regioes_m):
                    with cols_reg[i % len(cols_reg)]:
                        m_reg = metas_f[metas_f["Regiao_Coord"] == regiao]
                        v_reg = vendas_f[vendas_f["Regiao_Coord"] == regiao]
                        criar_medidor(regiao, float(v_reg["_qtd_venda"].sum()), m_reg["Meta_Qtd"].sum(), v_reg["_vgv_venda"].sum(), m_reg["Meta_VGV"].sum(), float(v_reg["_qtd_venda"].sum()))

        st.subheader("Tabela Resumo: Por Região")
        if "Regiao_Coord" in metas_f.columns:
            vg_reg = vendas_f.groupby("Regiao_Coord", as_index=False).agg(real_qtd=("_qtd_venda", "sum"), real_vgv=("_vgv_venda", "sum")).rename(columns={"Regiao_Coord": "Região"})
            mg_reg = metas_f.groupby("Regiao_Coord", as_index=False).agg(meta_qtd=("Meta_Qtd", "sum"), meta_vgv=("Meta_VGV", "sum")).rename(columns={"Regiao_Coord": "Região"})
            tab_reg = vg_reg.merge(mg_reg, on="Região", how="outer").fillna(0)
            tab_reg["% Qtd"] = tab_reg.apply(lambda r: (r["real_qtd"] / r["meta_qtd"] * 100.0) if r["meta_qtd"] > 0 else 0.0, axis=1)
            tab_reg["% VGV"] = tab_reg.apply(lambda r: (r["real_vgv"] / r["meta_vgv"] * 100.0) if r["meta_vgv"] > 0 else 0.0, axis=1)
            st.dataframe(tab_reg.sort_values("meta_qtd", ascending=False), use_container_width=True, hide_index=True)

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

        st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
        dia_atual = datetime.now().day
        st.subheader(f"Comparativo de Vendas (Dia 01 ao Dia {dia_atual:02d} do Mês)")
        
        if col_data_venda:
            vendas_f["Data_Formatada"] = pd.to_datetime(vendas_f[col_data_venda], dayfirst=True, errors="coerce")
            df_parcial = vendas_f[vendas_f["Data_Formatada"].dt.day <= dia_atual].copy()
            
            if not df_parcial.empty:
                df_comp = df_parcial.groupby(["_ano", "_mes"], as_index=False).agg(
                    QTD=("_qtd_venda", "sum"),
                    VGV=("_vgv_venda", "sum")
                ).sort_values(["_ano", "_mes"])
                
                df_comp["Periodo"] = df_comp["_mes"].astype(str).str.zfill(2) + "/" + df_comp["_ano"].astype(str)
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
                st.info(f"Não há dados de vendas no período do dia 01 ao dia {dia_atual:02d} para os filtros selecionados.")
        else:
            st.warning("A coluna de Data da Venda não foi encontrada. Impossível filtrar os dias.")

    # =========================================================================
    # TAB 1: Coordenadores IMOB
    # =========================================================================
    with tabs[1]:
        st.subheader("Premiação - Coordenadores IMOB")
        if df_metas_imob.empty or df_dic.empty:
            st.warning("Dados de Metas IMOB ou Dicionário não encontrados/preenchidos na planilha.")
        else:
            df_imob = normalizar_colunas(df_metas_imob)
            if datas_filtradas:
                df_imob = df_imob[df_imob.get("DATA", pd.Series()).isin(datas_filtradas)]
            
            # Agrupar por região
            if not df_imob.empty and "REGIÃO" in df_imob.columns:
                regioes = sorted(df_imob["REGIÃO"].dropna().unique())
                for reg in regioes:
                    st.markdown(f"### Região: {reg}")
                    df_reg = df_imob[df_imob["REGIÃO"] == reg]
                    
                    dados_tabela_imob = []
                    for _, row in df_reg.iterrows():
                        coords_str = str(row.get("COORDENADORES", ""))
                        lista_coords = extrair_lista_coords(coords_str)
                        emp = str(row.get("EMPREENDIMENTO", ""))
                        m_dir = parse_valor_br(row.get("META DIRECIONAL", 0))
                        m_imob = parse_valor_br(row.get("META IMOB", 0))
                        m_imob2 = parse_valor_br(row.get("META IMOB 2", 0))
                        
                        for c in lista_coords:
                            vendas_coord = mapear_vendas_por_coordenador(df_vendas, df_dic, emp)
                            # Puxar realizado deste coordenador
                            realizado = vendas_coord[vendas_coord["_Coordenador_Mapeado"] == c.lower()]["_qtd_venda"].sum() if "_Coordenador_Mapeado" in vendas_coord.columns else 0
                            
                            resultado = "NÃO BATEU NENHUMA META"
                            if realizado >= m_imob2 and m_imob2 > 0:
                                resultado = "BATEU META IMOB 2"
                            elif realizado >= m_imob and m_imob > 0:
                                resultado = "BATEU META IMOB"
                            elif realizado >= m_dir and m_dir > 0:
                                resultado = "BATEU META DIRECIONAL"
                                
                            dados_tabela_imob.append({
                                "Coordenador": c,
                                "Empreendimento": emp,
                                "META DIRECIONAL": m_dir,
                                "META IMOB": m_imob,
                                "META IMOB 2": m_imob2,
                                "Realizado": realizado,
                                "Resultado": resultado
                            })
                            
                    if dados_tabela_imob:
                        df_res_imob = pd.DataFrame(dados_tabela_imob)
                        st.dataframe(df_res_imob, use_container_width=True, hide_index=True)
                    else:
                        st.info(f"Sem dados correspondentes para a região {reg}.")
            else:
                st.info("Nenhum dado encontrado para os meses/anos filtrados em Coordenadores IMOB.")

    # =========================================================================
    # TAB 2: Coordenadores Comerciais
    # =========================================================================
    with tabs[2]:
        st.subheader("Premiação - Coordenadores Comerciais")
        if df_metas_comerciais.empty or df_dic.empty:
            st.warning("Dados de Metas Comerciais ou Dicionário não encontrados/preenchidos na planilha.")
        else:
            df_com = normalizar_colunas(df_metas_comerciais)
            if datas_filtradas:
                df_com = df_com[df_com.get("DATA", pd.Series()).isin(datas_filtradas)]
                
            if not df_com.empty:
                # Expande a base por Coordenador para tratar agrupamentos
                expanded_rows = []
                for _, row in df_com.iterrows():
                    lista_coords = extrair_lista_coords(str(row.get("COORDENADORES", "")))
                    emp = str(row.get("EMPREENDIMENTO", ""))
                    m_desafio = parse_valor_br(row.get("META DESAFIO VENDAS", 0))
                    m_bp = parse_valor_br(row.get("META BP", 0))
                    m_bp70 = parse_valor_br(row.get("META BP 70%", 0))
                    
                    for c in lista_coords:
                        vendas_coord = mapear_vendas_por_coordenador(df_vendas, df_dic, emp)
                        realizado = vendas_coord[vendas_coord["_Coordenador_Mapeado"] == c.lower()]["_qtd_venda"].sum() if "_Coordenador_Mapeado" in vendas_coord.columns else 0
                        
                        resultado = "NÃO BATEU NENHUMA META"
                        if realizado >= m_desafio and m_desafio > 0:
                            resultado = "BATEU META DESAFIO"
                        elif realizado >= m_bp and m_bp > 0:
                            resultado = "BATEU META BP"
                        elif realizado >= m_bp70 and m_bp70 > 0:
                            resultado = "BATEU META BP 70%"
                            
                        expanded_rows.append({
                            "Coordenador": c,
                            "Empreendimento": emp,
                            "META DESAFIO VENDAS": m_desafio,
                            "META BP": m_bp,
                            "META BP 70%": m_bp70,
                            "Realizado": realizado,
                            "Resultado": resultado
                        })
                
                df_exp = pd.DataFrame(expanded_rows)
                
                if not df_exp.empty:
                    # 1. Tabela Separada por Coordenador e Empreendimento
                    st.markdown("### Visão por Coordenador")
                    coords_unicos = sorted(df_exp["Coordenador"].unique())
                    for c_unico in coords_unicos:
                        st.markdown(f"**Coordenador: {c_unico}**")
                        df_c = df_exp[df_exp["Coordenador"] == c_unico].drop(columns=["Coordenador"])
                        st.dataframe(df_c, use_container_width=True, hide_index=True)
                        
                    # 2. Tabela para empreendimentos separados
                    st.markdown("<br><hr style='border:none;border-top:1px dashed #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
                    st.markdown("### Visão Consolidada por Empreendimentos (Separados)")
                    st.dataframe(df_exp, use_container_width=True, hide_index=True)
                else:
                    st.info("Sem dados resultantes após o mapeamento.")
            else:
                st.info("Nenhum dado encontrado para os meses/anos filtrados em Coordenadores Comerciais.")

    # =========================================================================
    # TAB 3: Grandes Contas
    # =========================================================================
    with tabs[3]:
        st.subheader("Premiação - Grandes Contas")
        if df_metas_gc.empty or df_dic.empty:
            st.warning("Dados de Metas Grandes Contas ou Dicionário não encontrados/preenchidos na planilha.")
        else:
            df_gc = normalizar_colunas(df_metas_gc)
            if datas_filtradas:
                df_gc = df_gc[df_gc.get("DATA", pd.Series()).isin(datas_filtradas)]
                
            if not df_gc.empty:
                dados_gc = []
                for _, row in df_gc.iterrows():
                    coords_str = str(row.get("COORDENADORES", ""))
                    lista_coords = extrair_lista_coords(coords_str)
                    
                    m1 = parse_valor_br(row.get("META 1", 0))
                    m2 = parse_valor_br(row.get("META 2", 0))
                    produtos_foco = extrair_lista_coords(str(row.get("PRODUTOS FOCO", "")))
                    
                    for c in lista_coords:
                        # Para Grandes Contas, olhamos o realizado total do mes filtrado e o foco
                        vendas_gerais_coord = mapear_vendas_por_coordenador(vendas_f, df_dic, None)
                        if "_Coordenador_Mapeado" in vendas_gerais_coord.columns:
                            vendas_c = vendas_gerais_coord[vendas_gerais_coord["_Coordenador_Mapeado"] == c.lower()]
                            realizado_total = vendas_c["_qtd_venda"].sum()
                            
                            # Filtra produtos foco
                            col_emp_gc = achar_coluna(vendas_c, ["Empreendimento", "Obra", "Nome do Empreendimento"])
                            if col_emp_gc and produtos_foco:
                                mask_foco = vendas_c[col_emp_gc].astype(str).str.strip().str.lower().isin([p.lower() for p in produtos_foco])
                                realizado_foco = vendas_c[mask_foco]["_qtd_venda"].sum()
                            else:
                                realizado_foco = 0
                        else:
                            realizado_total = 0
                            realizado_foco = 0
                            
                        # A regra diz: "veja se bateu a META 1 ou a META 2"
                        resultado = "NÃO BATEU NENHUMA META"
                        if realizado_total >= m2 and m2 > 0:
                            resultado = "BATEU META 2"
                        elif realizado_total >= m1 and m1 > 0:
                            resultado = "BATEU META 1"
                            
                        dados_gc.append({
                            "Coordenador": c,
                            "META 1": m1,
                            "META 2": m2,
                            "Realizado Total": realizado_total,
                            "Realizado Produtos Foco": realizado_foco,
                            "Resultado": resultado
                        })
                        
                if dados_gc:
                    df_res_gc = pd.DataFrame(dados_gc)
                    st.dataframe(df_res_gc, use_container_width=True, hide_index=True)
                else:
                    st.info("Sem dados resultantes após o mapeamento.")
            else:
                st.info("Nenhum dado encontrado para os meses/anos filtrados em Grandes Contas.")

    st.markdown(
        f'<div class="footer" style="text-align:center;padding:1rem 0;color:{COR_TEXTO_MUTED};font-size:0.82rem;">'
        f"Direcional Engenharia · Vendas — Acompanhamento de metas</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
