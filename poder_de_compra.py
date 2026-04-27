# -*- coding: utf-8 -*-
"""
Análise de Gaps de Vendas — Dinheiro na Mesa (Direcional).
Planilha: BD COMPLETA (Gaps).
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
# Identificação da folha de cálculo e Ficheiros Visuais
# -----------------------------------------------------------------------------
SPREADSHEET_ID = "1PZS41l6oPxQp3qq-2_zcw7RbMK5W3XCIryE41hH6dmI"
WS_VENDAS = "BD COMPLETA"

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"
URL_LOGO_DIRECIONAL_EMAIL = "https://logodownload.org/wp-content/uploads/2021/04/direcional-engenharia-logo.png"

# Paleta alinhada à Ficha de Credenciamento / Vendas RJ
COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_FUNDO_CARD = "rgba(255, 255, 255, 0.78)"
COR_BORDA = "#eef2f6"
COR_TEXTO_MUTED = "#64748b"
COR_TEXTO_LABEL = "#1e293b"
COR_INPUT_BG = "#f0f2f6"


def _hex_rgb_triplet(hex_color: str) -> str:
    """Converte #RRGGBB em 'r, g, b' para uso em rgba(...)."""
    x = (hex_color or "").strip().lstrip("#")
    if len(x) != 6:
        return "0, 0, 0"
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}"


RGB_AZUL_CSS = _hex_rgb_triplet(COR_AZUL_ESC)
RGB_VERMELHO_CSS = _hex_rgb_triplet(COR_VERMELHO)


# -----------------------------------------------------------------------------
# Funções de Design
# -----------------------------------------------------------------------------
def _resolver_png_raiz(nome: str) -> Path | None:
    for base in (_DIR_APP, _DIR_APP.parent):
        p = base / nome
        if p.is_file():
            return p
    return None

def _resolver_imagem_fundo_local(nome: str) -> Path | None:
    for base in (_DIR_APP, _DIR_APP.parent):
        for ext in (".jpg", ".jpeg", ".png", ".PNG"):
            stem = Path(nome).stem
            p = base / f"{stem}{ext}"
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
    return "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?auto=format&fit=crop&w=1920&q=80"

def _logo_arquivo_local() -> str | None:
    p_topo = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
    if p_topo:
        return str(p_topo)
    for name in ("logo_direcional.png", "logo_direcional.jpg", "logo.png"):
        p = _DIR_APP / "assets" / name
        if p.is_file():
            return str(p)
    return None

def _logo_url_secrets() -> str | None:
    try:
        if hasattr(st, "secrets") and isinstance(st.secrets.get("branding"), dict):
            return (st.secrets["branding"].get("LOGO_URL") or "").strip() or None
    except Exception:
        pass
    return None

def _logo_url_drive_por_id_arquivo() -> str | None:
    fid = (os.environ.get("DIRECIONAL_LOGO_FILE_ID") or "").strip()
    if len(fid) < 10: return None
    return f"https://drive.google.com/uc?export=view&id={fid}"

def _exibir_logo_topo() -> None:
    path = _logo_arquivo_local()
    url = _logo_url_secrets() or _logo_url_drive_por_id_arquivo()
    try:
        if path:
            ext = Path(path).suffix.lower().lstrip(".")
            mime = "image/jpeg" if ext == "png" else "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            st.markdown(f'<div class="ficha-logo-wrap"><img src="data:{mime};base64,{b64}" alt="Direcional" /></div>', unsafe_allow_html=True)
            return
        if url:
            st.markdown(f'<div class="ficha-logo-wrap"><img src="{html.escape(url)}" alt="Direcional" /></div>', unsafe_allow_html=True)
    except Exception:
        pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Análise Realizado X Projetado</p>'
        f'<p class="ficha-sub">Gaps de Vendas — <strong>BD COMPLETA</strong>.</p>'
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
        [data-testid="stSidebar"] {{ display: none !important; }}
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
        h1, h2, h3, h4 {{ font-family: 'Montserrat', sans-serif !important; color: {COR_AZUL_ESC} !important; font-weight: 800 !important; text-align: center !important; }}
        .ficha-logo-wrap {{ text-align: center; padding: 0.1rem 0 0.45rem 0; }}
        .ficha-logo-wrap img {{ max-height: 72px; width: auto; max-width: min(280px, 85vw); object-fit: contain; }}
        .ficha-hero-stack {{ width: 100%; margin-bottom: 0.35rem; }}
        .ficha-hero {{ text-align: center; padding: 0.5rem 0 0 0; max-width: 640px; margin: 0 auto; animation: fichaFadeIn 0.85s both; }}
        .ficha-hero .ficha-title {{ font-family: 'Montserrat', sans-serif; font-size: clamp(1.35rem, 3.5vw, 1.75rem); font-weight: 900; color: {COR_AZUL_ESC}; margin: 0; }}
        .ficha-hero .ficha-sub {{ color: #475569; font-size: 0.95rem; margin: 0.45rem 0 0 0; }}
        .ficha-hero-bar-wrap {{ width: 100%; margin: clamp(0.85rem, 2.4vw, 1.2rem) 0; }}
        .ficha-hero-bar {{ height: 4px; border-radius: 999px; background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC}); background-size: 200% 100%; animation: fichaShimmer 4s infinite alternate; }}
        .vel-kpi-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 1.25rem; }}
        .vel-kpi {{
            flex: 1 1 20%; background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.9) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9); border-radius: 14px; padding: 14px 16px; text-align: center;
            box-shadow: 0 2px 8px rgba({RGB_AZUL_CSS}, 0.06); transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        .vel-kpi:hover {{ transform: translateY(-4px); box-shadow: 0 10px 20px -5px rgba({RGB_AZUL_CSS}, 0.15); }}
        .vel-kpi .lbl {{ font-size: 0.72rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: {COR_AZUL_ESC}; opacity: 0.85; }}
        .vel-kpi .val {{ font-family: 'Montserrat', sans-serif; font-size: 1.35rem; font-weight: 800; color: {COR_AZUL_ESC}; margin-top: 6px; }}
        .vel-kpi .val--red {{ color: {COR_VERMELHO} !important; }}
        div[data-baseweb="input"] {{ border-radius: 10px !important; border: 1px solid #e2e8f0 !important; background-color: {COR_INPUT_BG} !important; }}
        div[data-baseweb="input"]:focus-within {{ border-color: rgba({RGB_AZUL_CSS}, 0.35) !important; box-shadow: 0 0 0 3px rgba({RGB_AZUL_CSS}, 0.08) !important; }}
        /* Estilização Extra para as Abas */
        div[data-baseweb="tab-list"] {{ gap: 10px; }}
        div[data-baseweb="tab"] {{ background: rgba(255,255,255,0.6); border-radius: 8px 8px 0 0; padding: 10px 20px; font-weight: 600; color: {COR_AZUL_ESC}; }}
        div[data-baseweb="tab"][aria-selected="true"] {{ background: rgba(255,255,255,0.95); border-bottom: 3px solid {COR_VERMELHO}; }}
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
    if not s: return s
    if "\\n" in s and "\n" not in s: s = s.replace("\\n", "\n")
    return s

def montar_service_account_info(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not raw: return None
    chaves = ("type", "project_id", "private_key_id", "private_key", "client_email", "client_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url", "client_x509_cert_url")
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
    if not str(out.get("type") or "").strip(): out["type"] = "service_account"
    if "token_uri" not in out: out["token_uri"] = "https://oauth2.googleapis.com/token"
    if "auth_uri" not in out: out["auth_uri"] = "https://accounts.google.com/o/oauth2/auth"
    return out

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

def ler_aba_gsheets(service_account_info: Dict[str, Any], spreadsheet_id: str, worksheet: str) -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id.strip())
    nome = worksheet.strip()
    try:
        ws = sh.worksheet(nome)
    except gspread.WorksheetNotFound:
        for w in sh.worksheets():
            if w.title.strip() == nome:
                ws = w
                break
        else:
            for w in sh.worksheets():
                if w.title.strip().lower() == nome.lower():
                    ws = w
                    break
            else:
                raise gspread.WorksheetNotFound(f"Aba {nome!r} não encontrada.") from None
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

def achar_coluna(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    for a in aliases:
        for c in df.columns:
            if a.lower() == str(c).strip().lower(): return c
    for a in aliases:
        for c in df.columns:
            if a.lower() in str(c).strip().lower(): return c
    return None

def fmt_br_milhoes(v: float) -> str:
    if v == 0: return "R$ 0,00"
    if v >= 1e6: return f"R$ {v / 1e6:.2f} mi"
    if v >= 1e3: return f"R$ {v / 1e3:.1f} mil"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_br_porcentagem(v: float) -> str:
    return f"{v:.1f}%".replace(".", ",")

# -----------------------------------------------------------------------------
# Aplicação Principal
# -----------------------------------------------------------------------------
def main() -> None:
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(page_title="Análise Realizado X Projetado | Direcional", page_icon=str(fav) if fav else None, layout="wide")
    aplicar_estilo()
    _cabecalho_pagina()

    raw_gs = _secrets_connections_gsheets()
    info = montar_service_account_info(raw_gs)
    if not info:
        st.error("Credenciais Google em **[connections.gsheets]** incompletas. Preencha pelo menos **private_key** e **client_email**.")
        return

    sid = SPREADSHEET_ID
    cred_fp = _fingerprint_credenciais(info)

    with st.spinner("A carregar dados da folha de cálculo..."):
        try:
            df = ler_planilha_aba_df(sid, WS_VENDAS, cred_fp)
        except Exception as e:
            st.error(f"Erro ao ler a folha de cálculo: {str(e)}")
            return

    # Normalizar cabeçalhos
    df.columns = [str(c).strip() for c in df.columns]

    # Mapeamento de Colunas
    c_data = achar_coluna(df, ["CONTRATO GERADO EM", "Data do Contrato", "Contrato gerado"])
    c_emp = achar_coluna(df, ["Empreendimento"])
    c_reg_imob = achar_coluna(df, ["Regional ou Imob", "Regional"])
    c_imobiliaria = achar_coluna(df, ["Imobiliária", "Imobiliaria"])
    c_canal = achar_coluna(df, ["Canal"])
    c_rank = achar_coluna(df, ["RANKING", "Ranking"])
    
    c_v_real = achar_coluna(df, ["VALOR REAL DE VENDA"])
    c_v_dir = achar_coluna(df, ["VALOR DIRECIONAL DE VENDA"])
    c_v_emc = achar_coluna(df, ["VALOR EMCASH DE VENDA"])

    # Colunas de Vendas Facilitadas/Comerciais
    c_v_comercial = achar_coluna(df, ["VENDA COMERCIAL", "Venda Comercial"])
    c_v_facilitada = achar_coluna(df, ["VENDA FACILITADA", "Venda Facilitada"])

    # Colunas de Limpeza antigas
    c_ps_dir = achar_coluna(df, ["PS DIRECIONAL"])
    c_ps_emc = achar_coluna(df, ["PS EMCASH"])
    
    # Novas colunas solicitadas e de limpeza
    c_renda = achar_coluna(df, ["RENDA APURADA", "Renda Apurada"])
    c_val_fin = achar_coluna(df, ["Valor do Financiamento"])
    c_ato = achar_coluna(df, ["Ato Total", "Total Ato Pago"])
    c_pro_soluto = achar_coluna(df, ["Pro Soluto", "Pro soluto"])
    c_renda_procx = achar_coluna(df, ["Renda PROCX", "Renda procx"])
    c_finan = achar_coluna(df, ["FINANCIAMENTO MÁXIMO", "Financiamento Máximo"])
    c_subsi = achar_coluna(df, ["SUBSÍDIO DISPONÍVEL", "Subsídio Disponível", "Subsidio Disponivel"])
    c_avalia = achar_coluna(df, ["Avaliação", "Avaliacao", "Nome da Avaliação de crédito"])
    
    c_fin_real = achar_coluna(df, ["FINANCIAMENTO REAL", "Financiamento Real"])
    c_sub_real = achar_coluna(df, ["SUBSÍDIO REAL", "Subsídio Real", "Subsidio Real"])
    c_val_fin_sub_real = achar_coluna(df, ["VALOR C/ FIN E SUB REAIS", "Valor c/ Fin e Sub Reais"])

    if not all([c_v_real, c_v_dir, c_v_emc]):
        st.error("Colunas essenciais de valores (Real, Direcional, Emcash) não encontradas.")
        return

    if not c_data:
        st.error("Coluna de Data (CONTRATO GERADO EM) não encontrada para gerar a linha do tempo.")
        return

    # -------------------------------------------------------------------------
    # Filtro Obrigatório: Venda Comercial == 1
    # -------------------------------------------------------------------------
    if c_v_comercial and c_v_comercial in df.columns:
        df = df[pd.to_numeric(df[c_v_comercial], errors='coerce') == 1]

    # -------------------------------------------------------------------------
    # Limpeza de Dados: Ignorar linhas com colunas vazias, 0 ou #N/A
    # -------------------------------------------------------------------------
    def _is_invalid_value(val, is_numeric=False, allow_zero=False):
        if pd.isna(val): return True
        s = str(val).strip().upper()
        if s in ("", "#N/A", "NAN", "NA", "NONE", "NULL"): return True
        if not allow_zero:
            if is_numeric:
                if parse_valor_br(val) == 0.0: return True
            else:
                if s in ("0", "0.0", "0,0"): return True
        return False

    # Filtro rigoroso (Vazio, N/A ou Zero) para numéricas (EXCETO Subsídio)
    cols_limpeza_numericas = [c for c in [c_renda, c_val_fin, c_ato, c_pro_soluto, c_renda_procx, c_finan, c_fin_real, c_val_fin_sub_real] if c and c in df.columns]
    for col in cols_limpeza_numericas:
        df = df[~df[col].apply(lambda x: _is_invalid_value(x, is_numeric=True, allow_zero=False))]

    # Subsídio (Vazio, N/A, mas PERMITE Zero)
    cols_limpeza_subsidio = [c for c in [c_subsi, c_sub_real] if c and c in df.columns]
    for col in cols_limpeza_subsidio:
        df = df[~df[col].apply(lambda x: _is_invalid_value(x, is_numeric=True, allow_zero=True))]

    # Filtro rigoroso (Vazio, N/A ou string "0") para textos (como Avaliação)
    cols_limpeza_texto = [c for c in [c_avalia] if c and c in df.columns]
    for col in cols_limpeza_texto:
        df = df[~df[col].apply(lambda x: _is_invalid_value(x, is_numeric=False, allow_zero=False))]

    # Limpeza padrão original (apenas nulos) para outras colunas se necessário
    cols_limpeza_padrao = [c for c in [c_ps_dir, c_ps_emc] if c and c in df.columns]
    for col in cols_limpeza_padrao:
        df = df[df[col].notna() & (df[col].astype(str).str.strip() != "")]

    # Tratamento de Dados e Datas
    df["Data_Formatada"] = pd.to_datetime(df[c_data], dayfirst=True, errors="coerce")
    df["_ano"] = df["Data_Formatada"].dt.year
    df["_mes"] = df["Data_Formatada"].dt.month
    
    # Tratamento Financeiro
    if c_v_real in df.columns: df["VGV_Real"] = df[c_v_real].apply(parse_valor_br)
    if c_v_dir in df.columns: df["VGV_Dir"] = df[c_v_dir].apply(parse_valor_br)
    if c_v_emc in df.columns: df["VGV_Emc"] = df[c_v_emc].apply(parse_valor_br)
    
    # Tratamento Financeiro (Novos Campos)
    if c_fin_real and c_fin_real in df.columns: df["Financiamento_Real"] = df[c_fin_real].apply(parse_valor_br)
    else: df["Financiamento_Real"] = 0.0

    if c_finan and c_finan in df.columns: df["Financiamento_Max"] = df[c_finan].apply(parse_valor_br)
    else: df["Financiamento_Max"] = 0.0

    if c_sub_real and c_sub_real in df.columns: df["Subsidio_Real"] = df[c_sub_real].apply(parse_valor_br)
    else: df["Subsidio_Real"] = 0.0

    if c_subsi and c_subsi in df.columns: df["Subsidio_Disp"] = df[c_subsi].apply(parse_valor_br)
    else: df["Subsidio_Disp"] = 0.0

    if c_val_fin_sub_real and c_val_fin_sub_real in df.columns: df["Valor_FinSub_Real"] = df[c_val_fin_sub_real].apply(parse_valor_br)
    else: df["Valor_FinSub_Real"] = 0.0

    # Cálculo dos Gaps Nominais Originais
    if "VGV_Dir" in df.columns and "VGV_Real" in df.columns: df["Gap_Dir"] = df["VGV_Dir"] - df["VGV_Real"]
    if "VGV_Emc" in df.columns and "VGV_Real" in df.columns: df["Gap_Emc"] = df["VGV_Emc"] - df["VGV_Real"]

    # Cálculo dos Novos Gaps (c/ Fin e Sub Reais)
    if "VGV_Dir" in df.columns and "Valor_FinSub_Real" in df.columns: df["Gap_FinSub_Dir"] = df["VGV_Dir"] - df["Valor_FinSub_Real"]
    if "VGV_Emc" in df.columns and "Valor_FinSub_Real" in df.columns: df["Gap_FinSub_Emc"] = df["VGV_Emc"] - df["Valor_FinSub_Real"]

    # Diferenças de Crédito
    df["Dif_Financiamento"] = df["Financiamento_Max"] - df["Financiamento_Real"]
    df["Dif_Subsidio"] = df["Subsidio_Disp"] - df["Subsidio_Real"]

    # Limpeza de campos categóricos nulos
    for col in [c_emp, c_reg_imob, c_imobiliaria, c_canal, c_rank]:
        if col and col in df.columns:
            df[col] = df[col].fillna("Não Informado").astype(str).str.strip()

    # -------------------------------------------------------------------------
    # Filtros da UI
    # -------------------------------------------------------------------------
    anos_disp = sorted([int(x) for x in df["_ano"].dropna().unique() if x > 2000])
    meses_disp = list(range(1, 13))
    mes_atual = datetime.now().month
    mes_padrao = mes_atual if mes_atual in meses_disp else 1

    st.markdown("<div style='margin-bottom:1rem; text-align: center;'><strong>Filtros de Análise</strong></div>", unsafe_allow_html=True)
    
    row1_c1, row1_c2, row1_c3, row1_c4 = st.columns(4)
    with row1_c1: sel_ano = st.multiselect("Ano", anos_disp, default=[anos_disp[-1]] if anos_disp else [])
    with row1_c2: sel_mes = st.multiselect("Mês", meses_disp, default=[mes_padrao])
    with row1_c3: sel_canal = st.multiselect("Canal", sorted(df[c_canal].unique()) if c_canal and c_canal in df.columns else [])
    with row1_c4: sel_reg_imob_filt = st.multiselect("Regional ou Imob", sorted(df[c_reg_imob].unique()) if c_reg_imob and c_reg_imob in df.columns else [])

    row2_c1, row2_c2, row2_c3 = st.columns(3)
    with row2_c1: sel_emp = st.multiselect("Empreendimento", sorted(df[c_emp].unique()) if c_emp and c_emp in df.columns else [])
    with row2_c2: sel_imobiliaria = st.multiselect("Imobiliária", sorted(df[c_imobiliaria].unique()) if c_imobiliaria and c_imobiliaria in df.columns else [])
    with row2_c3: sel_rank = st.multiselect("Ranking", sorted(df[c_rank].unique()) if c_rank and c_rank in df.columns else [])

    # Novos botões de filtros opcionais para Gaps
    row3_c1, row3_c2 = st.columns(2)
    with row3_c1:
        exibir_negativos = st.checkbox("Exibir Gaps < 0 (Atos Expressivos)", value=False)
    with row3_c2:
        exibir_acima_100k = st.checkbox("Exibir Gaps > R$ 100k", value=False)

    df_f = df.copy()
    if sel_ano: df_f = df_f[df_f["_ano"].isin(sel_ano)]
    if sel_mes: df_f = df_f[df_f["_mes"].isin(sel_mes)]
    if c_canal and sel_canal and c_canal in df_f.columns: df_f = df_f[df_f[c_canal].isin(sel_canal)]
    if c_reg_imob and sel_reg_imob_filt and c_reg_imob in df_f.columns: df_f = df_f[df_f[c_reg_imob].isin(sel_reg_imob_filt)]
    if c_emp and sel_emp and c_emp in df_f.columns: df_f = df_f[df_f[c_emp].isin(sel_emp)]
    if c_imobiliaria and sel_imobiliaria and c_imobiliaria in df_f.columns: df_f = df_f[df_f[c_imobiliaria].isin(sel_imobiliaria)]
    if c_rank and sel_rank and c_rank in df_f.columns: df_f = df_f[df_f[c_rank].isin(sel_rank)]

    # -------------------------------------------------------------------------
    # Aplicando Regras de Gaps (Aplica aos Gaps Originais)
    # -------------------------------------------------------------------------
    if not exibir_negativos and "Gap_Dir" in df_f.columns and "Gap_Emc" in df_f.columns:
        df_f = df_f[(df_f["Gap_Dir"] >= 0) & (df_f["Gap_Emc"] >= 0)]
    
    if not exibir_acima_100k and "Gap_Dir" in df_f.columns and "Gap_Emc" in df_f.columns:
        df_f = df_f[(df_f["Gap_Dir"] <= 100000) & (df_f["Gap_Emc"] <= 100000)]

    # -------------------------------------------------------------------------
    # Componentes de Renderização
    # -------------------------------------------------------------------------
    def render_kpi_block(df_target: pd.DataFrame, title: str, col_vgv="VGV_Real", col_dir="Gap_Dir", col_emc="Gap_Emc", label_vgv="VGV Real Total"):
        st.subheader(title)
        if df_target.empty:
            st.info("Não há dados para os filtros selecionados.")
            return

        vendas_qtd = len(df_target)
        vgv_tot = df_target[col_vgv].sum() if col_vgv in df_target.columns else 0.0
        
        gap_dir_tot = df_target[col_dir].sum() if col_dir in df_target.columns else 0.0
        gap_dir_avg = df_target[col_dir].mean() if col_dir in df_target.columns else 0.0
        gap_dir_med = df_target[col_dir].median() if col_dir in df_target.columns else 0.0
        gap_dir_p10 = df_target[col_dir].quantile(0.1) if col_dir in df_target.columns else 0.0
        pct_gap_dir = (gap_dir_tot / vgv_tot * 100.0) if vgv_tot > 0 else 0.0

        gap_emc_tot = df_target[col_emc].sum() if col_emc in df_target.columns else 0.0
        gap_emc_avg = df_target[col_emc].mean() if col_emc in df_target.columns else 0.0
        gap_emc_med = df_target[col_emc].median() if col_emc in df_target.columns else 0.0
        gap_emc_p10 = df_target[col_emc].quantile(0.1) if col_emc in df_target.columns else 0.0
        pct_gap_emc = (gap_emc_tot / vgv_tot * 100.0) if vgv_tot > 0 else 0.0

        st.markdown(
            f"""
            <div class="vel-kpi-row">
                <div class="vel-kpi"><div class="lbl">Vendas (QTD)</div><div class="val">{vendas_qtd}</div></div>
                <div class="vel-kpi"><div class="lbl">{label_vgv}</div><div class="val">{fmt_br_milhoes(vgv_tot)}</div></div>
                <div class="vel-kpi"><div class="lbl">Gap Direcional (Tot)</div><div class="val val--red">{fmt_br_milhoes(gap_dir_tot)}</div></div>
                <div class="vel-kpi"><div class="lbl">Gap Emcash (Tot)</div><div class="val val--red">{fmt_br_milhoes(gap_emc_tot)}</div></div>
            </div>
            <div class="vel-kpi-row">
                <div class="vel-kpi"><div class="lbl">Média Gap (Dir)</div><div class="val">{fmt_br_milhoes(gap_dir_avg)}</div></div>
                <div class="vel-kpi"><div class="lbl">Mediana Gap (Dir)</div><div class="val">{fmt_br_milhoes(gap_dir_med)}</div></div>
                <div class="vel-kpi"><div class="lbl">P10 Gap (Dir)</div><div class="val">{fmt_br_milhoes(gap_dir_p10)}</div></div>
                <div class="vel-kpi"><div class="lbl">Aumento Possível (Dir)</div><div class="val">{fmt_br_porcentagem(pct_gap_dir)}</div></div>
            </div>
            <div class="vel-kpi-row" style="margin-bottom: 2rem;">
                <div class="vel-kpi"><div class="lbl">Média Gap (Emc)</div><div class="val">{fmt_br_milhoes(gap_emc_avg)}</div></div>
                <div class="vel-kpi"><div class="lbl">Mediana Gap (Emc)</div><div class="val">{fmt_br_milhoes(gap_emc_med)}</div></div>
                <div class="vel-kpi"><div class="lbl">P10 Gap (Emc)</div><div class="val">{fmt_br_milhoes(gap_emc_p10)}</div></div>
                <div class="vel-kpi"><div class="lbl">Aumento Possível (Emc)</div><div class="val">{fmt_br_porcentagem(pct_gap_emc)}</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    def render_graficos_gaps(df_target, col_dir="Gap_Dir", col_emc="Gap_Emc"):
        st.subheader("Evolução Mensal")
        if not df_target.empty and col_dir in df_target.columns and col_emc in df_target.columns:
            df_chart = df_target.groupby(["_ano", "_mes"], as_index=False).agg(
                G_Dir=(col_dir, "sum"),
                G_Emc=(col_emc, "sum")
            ).sort_values(["_ano", "_mes"])
            
            df_chart["Periodo"] = df_chart["_mes"].astype(str).str.zfill(2) + "/" + df_chart["_ano"].astype(str)
            
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_chart["Periodo"], y=df_chart["G_Dir"], name="Gap Direcional (R$)", marker_color=COR_AZUL_ESC))
            fig.add_trace(go.Bar(x=df_chart["Periodo"], y=df_chart["G_Emc"], name="Gap Emcash (R$)", marker_color=COR_VERMELHO))
            
            fig.update_layout(barmode="group", bargap=0.4, margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter", color=COR_TEXTO_LABEL), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

            st.subheader("Distribuição da Frequência de Gaps")
            fig_hist = go.Figure()
            fig_hist.add_trace(go.Histogram(x=df_target[col_dir], name="Distribuição Direcional", marker_color=COR_AZUL_ESC, opacity=0.75, xbins=dict(size=5000)))
            fig_hist.add_trace(go.Histogram(x=df_target[col_emc], name="Distribuição Emcash", marker_color=COR_VERMELHO, opacity=0.75, xbins=dict(size=5000)))
            fig_hist.update_layout(barmode="overlay", xaxis_title="Valor do Gap (R$)", yaxis_title="Frequência (Vendas)", margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter", color=COR_TEXTO_LABEL), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig_hist, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Não há dados para exibir gráficos.")

    def gerar_tabelas_gaps(df_target, col_vgv="VGV_Real", col_dir="Gap_Dir", col_emc="Gap_Emc", vgv_label="VGV Real"):
        def gerar_tabela_gap_local(df_input, coluna_agrupamento, label_tabela):
            if not coluna_agrupamento or coluna_agrupamento not in df_input.columns or df_input.empty: return
            st.subheader(f"Gaps por {label_tabela}")
            
            tab = df_input.groupby(coluna_agrupamento, as_index=False).agg(
                QTD=(coluna_agrupamento, "count"),
                ValBase=(col_vgv, "sum"),
                Gap_Dir=(col_dir, "sum"),
                Gap_Emc=(col_emc, "sum")
            )
            
            tab["% Gap Dir"] = tab.apply(lambda r: (r["Gap_Dir"] / r["ValBase"] * 100) if r["ValBase"] > 0 else 0, axis=1)
            tab["% Gap Emc"] = tab.apply(lambda r: (r["Gap_Emc"] / r["ValBase"] * 100) if r["ValBase"] > 0 else 0, axis=1)
            tab = tab.sort_values("Gap_Dir", ascending=False)
            
            show = tab.rename(columns={
                coluna_agrupamento: label_tabela,
                "ValBase": vgv_label,
                "Gap_Dir": "Gap Direcional (R$)",
                "Gap_Emc": "Gap Emcash (R$)"
            })
            
            show[vgv_label] = show[vgv_label].map(lambda x: fmt_br_milhoes(float(x)))
            show["Gap Direcional (R$)"] = show["Gap Direcional (R$)"].map(lambda x: fmt_br_milhoes(float(x)))
            show["Gap Emcash (R$)"] = show["Gap Emcash (R$)"].map(lambda x: fmt_br_milhoes(float(x)))
            show["% Gap Dir"] = show["% Gap Dir"].map(lambda x: f"{x:.1f}%")
            show["% Gap Emc"] = show["% Gap Emc"].map(lambda x: f"{x:.1f}%")
            
            st.dataframe(show, use_container_width=True, hide_index=True)

        if c_emp and c_emp in df_target.columns: gerar_tabela_gap_local(df_target, c_emp, "Empreendimento")
        if c_reg_imob and c_reg_imob in df_target.columns and c_canal and c_canal in df_target.columns:
            df_regionais = df_target[df_target[c_canal].isin(['DIR', 'RIV'])]
            if not df_regionais.empty: gerar_tabela_gap_local(df_regionais, c_reg_imob, "Regional (Canais DIR/RIV)")
            
            df_imobs = df_target[df_target[c_canal].isin(['RJ', 'RJG'])]
            if not df_imobs.empty: gerar_tabela_gap_local(df_imobs, c_reg_imob, "Imob (Canais RJ/RJG)")
                
        if c_imobiliaria and c_imobiliaria in df_target.columns: gerar_tabela_gap_local(df_target, c_imobiliaria, "Imobiliária")
        if c_rank and c_rank in df_target.columns: gerar_tabela_gap_local(df_target, c_rank, "Ranking")
        if c_canal and c_canal in df_target.columns: gerar_tabela_gap_local(df_target, c_canal, "Canal")

    # -------------------------------------------------------------------------
    # Renderização da Estrutura em Abas
    # -------------------------------------------------------------------------
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    
    tab_orig, tab_fin, tab_credito = st.tabs([
        "📊 Gaps (Base: VGV Real)", 
        "📊 Gaps (Base: Valor c/ Fin e Sub Reais)", 
        "💰 Análise de Crédito (Fin e Sub)"
    ])

    with tab_orig:
        render_kpi_block(df_f, "Estatísticas Consolidadas - VGV Real", col_vgv="VGV_Real", col_dir="Gap_Dir", col_emc="Gap_Emc", label_vgv="VGV Real Total")
        render_graficos_gaps(df_f, col_dir="Gap_Dir", col_emc="Gap_Emc")
        gerar_tabelas_gaps(df_f, col_vgv="VGV_Real", col_dir="Gap_Dir", col_emc="Gap_Emc", vgv_label="VGV Real")

    with tab_fin:
        render_kpi_block(df_f, "Estatísticas Consolidadas - Valor c/ Fin e Sub Reais", col_vgv="Valor_FinSub_Real", col_dir="Gap_FinSub_Dir", col_emc="Gap_FinSub_Emc", label_vgv="Valor c/ Fin e Sub Reais (Tot)")
        render_graficos_gaps(df_f, col_dir="Gap_FinSub_Dir", col_emc="Gap_FinSub_Emc")
        gerar_tabelas_gaps(df_f, col_vgv="Valor_FinSub_Real", col_dir="Gap_FinSub_Dir", col_emc="Gap_FinSub_Emc", vgv_label="Valor c/ Fin e Sub")

    with tab_credito:
        st.subheader("Indicadores de Crédito")
        if not df_f.empty:
            tot_fin_real = df_f["Financiamento_Real"].sum()
            tot_fin_max = df_f["Financiamento_Max"].sum()
            tot_dif_fin = df_f["Dif_Financiamento"].sum()
            
            tot_sub_real = df_f["Subsidio_Real"].sum()
            tot_sub_disp = df_f["Subsidio_Disp"].sum()
            tot_dif_sub = df_f["Dif_Subsidio"].sum()

            st.markdown(
                f"""
                <div class="vel-kpi-row">
                    <div class="vel-kpi"><div class="lbl">Financiamento Real (Tot)</div><div class="val">{fmt_br_milhoes(tot_fin_real)}</div></div>
                    <div class="vel-kpi"><div class="lbl">Financiamento Máximo (Tot)</div><div class="val">{fmt_br_milhoes(tot_fin_max)}</div></div>
                    <div class="vel-kpi"><div class="lbl">Dif. Financiamento (Dinheiro na Mesa)</div><div class="val val--red">{fmt_br_milhoes(tot_dif_fin)}</div></div>
                </div>
                <div class="vel-kpi-row" style="margin-bottom: 2rem;">
                    <div class="vel-kpi"><div class="lbl">Subsídio Real (Tot)</div><div class="val">{fmt_br_milhoes(tot_sub_real)}</div></div>
                    <div class="vel-kpi"><div class="lbl">Subsídio Disponível (Tot)</div><div class="val">{fmt_br_milhoes(tot_sub_disp)}</div></div>
                    <div class="vel-kpi"><div class="lbl">Dif. Subsídio (Dinheiro na Mesa)</div><div class="val val--red">{fmt_br_milhoes(tot_dif_sub)}</div></div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # Gráficos de Crédito
            st.subheader("Evolução Mensal - Crédito")
            df_cred_chart = df_f.groupby(["_ano", "_mes"], as_index=False).agg(
                Fin_Real=("Financiamento_Real", "sum"),
                Fin_Max=("Financiamento_Max", "sum"),
                Sub_Real=("Subsidio_Real", "sum"),
                Sub_Disp=("Subsidio_Disp", "sum")
            ).sort_values(["_ano", "_mes"])
            df_cred_chart["Periodo"] = df_cred_chart["_mes"].astype(str).str.zfill(2) + "/" + df_cred_chart["_ano"].astype(str)
            
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                fig_fin = go.Figure()
                fig_fin.add_trace(go.Bar(x=df_cred_chart["Periodo"], y=df_cred_chart["Fin_Real"], name="Financiamento Real", marker_color=COR_AZUL_ESC))
                fig_fin.add_trace(go.Bar(x=df_cred_chart["Periodo"], y=df_cred_chart["Fin_Max"], name="Financiamento Máximo", marker_color="#3b82f6")) # Azul claro
                fig_fin.update_layout(title="Financiamento Real x Máximo", barmode="group", margin=dict(l=20, r=20, t=40, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter", color=COR_TEXTO_LABEL), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                st.plotly_chart(fig_fin, use_container_width=True, config={"displayModeBar": False})

            with col_chart2:
                fig_sub = go.Figure()
                fig_sub.add_trace(go.Bar(x=df_cred_chart["Periodo"], y=df_cred_chart["Sub_Real"], name="Subsídio Real", marker_color=COR_VERMELHO))
                fig_sub.add_trace(go.Bar(x=df_cred_chart["Periodo"], y=df_cred_chart["Sub_Disp"], name="Subsídio Disponível", marker_color="#f87171")) # Vermelho claro
                fig_sub.update_layout(title="Subsídio Real x Disponível", barmode="group", margin=dict(l=20, r=20, t=40, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter", color=COR_TEXTO_LABEL), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                st.plotly_chart(fig_sub, use_container_width=True, config={"displayModeBar": False})

            # Tabela de Crédito por Empreendimento
            if c_emp and c_emp in df_f.columns:
                st.subheader("Desempenho de Crédito por Empreendimento")
                tab_cred = df_f.groupby(c_emp, as_index=False).agg(
                    QTD=(c_emp, "count"),
                    Fin_Real=("Financiamento_Real", "sum"),
                    Fin_Max=("Financiamento_Max", "sum"),
                    Dif_Fin=("Dif_Financiamento", "sum"),
                    Sub_Real=("Subsidio_Real", "sum"),
                    Sub_Disp=("Subsidio_Disp", "sum"),
                    Dif_Sub=("Dif_Subsidio", "sum")
                ).sort_values("Dif_Fin", ascending=False)
                
                show_cred = tab_cred.rename(columns={
                    c_emp: "Empreendimento",
                    "Fin_Real": "Fin. Real",
                    "Fin_Max": "Fin. Máximo",
                    "Dif_Fin": "Dif. Financiamento",
                    "Sub_Real": "Sub. Real",
                    "Sub_Disp": "Sub. Disponível",
                    "Dif_Sub": "Dif. Subsídio"
                })

                for c in ["Fin. Real", "Fin. Máximo", "Dif. Financiamento", "Sub. Real", "Sub. Disponível", "Dif. Subsídio"]:
                    show_cred[c] = show_cred[c].map(lambda x: fmt_br_milhoes(float(x)))
                
                st.dataframe(show_cred, use_container_width=True, hide_index=True)
        else:
            st.info("Não há dados para exibir a Análise de Crédito.")

    st.markdown(
        f'<div class="footer" style="text-align:center;padding:1rem 0;color:{COR_TEXTO_MUTED};font-size:0.82rem;">'
        f"Direcional Engenharia · Vendas — Análise de Gaps</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
