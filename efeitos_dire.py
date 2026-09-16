# -*- coding: utf-8 -*-
"""
Acompanhamento e Projeção Sazonal de Vendas & Funil — Direcional (RJ).
Design: Gaps Style (Transparência, Blur, Fundo de Cadastro, Inter/Montserrat).
"""
from __future__ import annotations

import base64
import calendar
import datetime
from datetime import date, datetime, timedelta
import html
import math
import os
from pathlib import Path
import re
import time
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# -----------------------------------------------------------------------------
# Configurações de Identificação e Padrões Visuais
# -----------------------------------------------------------------------------
SPREADSHEET_ID = "1wpuNQvksot9CLhGgQRe7JlyDeRISEh_sc3-6VRDyQYk"
SPREADSHEET_METAS_ID = "1cseWbys3GXd7Q70irMK_8Iw3xz5iQy0wmArgkSdd5BA"

SF_REPORT_AGENDAMENTOS_ID = "00OU600000AcFGPMA3"
SF_REPORT_PASTAS_ID = "00OU600000FEOoDMAX"
SF_REPORT_VENDAS_ID = "00O3Z000005ZsPmUAK"

FUNIL_ETAPAS = ("agendamentos", "visitas", "pastas", "pastas_aprovadas", "vendas")
FUNIL_LABELS = {
    "agendamentos": "Agendamentos",
    "visitas": "Visitas",
    "pastas": "Pastas",
    "pastas_aprovadas": "Pastas Aprovadas",
    "vendas": "Vendas",
}

COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_TEXTO_PRETO = "#000000"
COR_BORDA = "#eef2f6"

FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"
_DIR_APP = Path(__file__).resolve().parent if '__file__' in locals() else Path.cwd()

MESES_TEXTO_MAP = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}

DIAS_SEMANA_PT = {0: "segunda", 1: "terça", 2: "quarta", 3: "quinta", 4: "sexta", 5: "sábado", 6: "domingo"}
MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}
_TZ_BR = "America/Sao_Paulo"

def _hex_rgb_triplet(hex_color: str) -> str:
    x = (hex_color or "").strip().lstrip("#")
    if len(x) != 6: return "0, 0, 0"
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}"

RGB_AZUL_CSS = _hex_rgb_triplet(COR_AZUL_ESC)
RGB_VERMELHO_CSS = _hex_rgb_triplet(COR_VERMELHO)

def _resolver_imagem_fundo_local(nome: str) -> Path | None:
    for base in (_DIR_APP, _DIR_APP.parent):
        for ext in (".jpg", ".jpeg", ".JPG", ".JPEG", ".png", ".PNG"):
            stem = Path(nome).stem
            p = base / f"{stem}{ext}"
            if p.is_file(): return p
        p = base / nome
        if p.is_file(): return p
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
        except OSError: pass
    return "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?auto=format&fit=crop&w=1920&q=80"

def aplicar_estilo() -> None:
    bg_url = _css_url_fundo_cadastro()
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap');
        html, body, :root, [data-testid="stApp"] {{ color-scheme: light !important; }}
        html, body {{ font-family: 'Inter', sans-serif; color: {COR_TEXTO_PRETO}; background: transparent !important; }}
        
        /* Restaura o fundo original com gradiente e imagem */
        .stApp, [data-testid="stApp"] {{
            background: 
                linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%),
                url("{bg_url}") center / cover no-repeat !important;
            background-attachment: fixed !important;
        }}
        [data-testid="stHeader"] {{ background: transparent !important; }}
        [data-testid="stSidebar"] {{ display: none !important; }}
        .block-container {{
            max-width: 1700px !important; margin: 1rem auto !important;
            padding: 2rem !important; background: rgba(255, 255, 255, 0.85) !important;
            backdrop-filter: blur(18px) saturate(1.15); border-radius: 24px !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            box-shadow: 0 24px 48px -12px rgba({RGB_AZUL_CSS}, 0.18) !important;
        }}
        h1, h2, h3, h4 {{ font-family: 'Montserrat', sans-serif !important; color: {COR_AZUL_ESC} !important; font-weight: 800 !important; text-align: center !important; }}
        h5, h6 {{ font-family: 'Montserrat', sans-serif !important; color: {COR_TEXTO_PRETO} !important; font-weight: 700 !important; text-align: center !important; }}
        .vel-kpi-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 1.25rem; }}
        .vel-kpi {{
            flex: 1 1 18%; background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.9) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9); border-radius: 14px; padding: 14px 16px; text-align: center;
            box-shadow: 0 2px 8px rgba({RGB_AZUL_CSS}, 0.06);
        }}
        .vel-kpi .lbl {{ font-size: 0.72rem; font-weight: 700; text-transform: uppercase; color: {COR_TEXTO_PRETO}; opacity: 0.85; }}
        .vel-kpi .val {{ font-family: 'Montserrat', sans-serif; font-size: 1.35rem; font-weight: 800; color: {COR_AZUL_ESC} !important; margin-top: 6px; }}
        .vel-kpi .val--red {{ color: {COR_VERMELHO} !important; }}
        div.stDownloadButton > button {{
            background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}) !important;
            color: #ffffff !important; font-weight: 700 !important; border-radius: 10px !important; border: none !important;
        }}
        div.stDownloadButton > button:hover {{ opacity: 0.9 !important; color: #ffffff !important; }}
        </style>
        """,
        unsafe_allow_html=True,
    )

def _secrets_connections_gsheets() -> Dict[str, Any]:
    try:
        sec = st.secrets
        if hasattr(sec, "get") and sec.get("connections"):
            g = sec["connections"].get("gsheets")
            if g is not None: return dict(g)
    except Exception: pass
    return {}

def _normalizar_private_key_toml(pk: str) -> str:
    s = (pk or "").strip()
    if not s: return s
    if "\\n" in s and "\n" not in s: s = s.replace("\\n", "\n")
    return s

def montar_service_account_info(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not raw: return None
    chaves = ("type", "project_id", "private_key_id", "private_key", "client_email", "client_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url", "client_x509_cert_url")
    out = {}
    for k in chaves:
        v = raw.get(k)
        if v is not None:
            if isinstance(v, str): v = v.strip()
            if v != "": out[k] = v
    if "private_key" in out: out["private_key"] = _normalizar_private_key_toml(str(out["private_key"]))
    if "private_key" not in out or "client_email" not in out: return None
    if "type" not in out: out["type"] = "service_account"
    if "token_uri" not in out: out["token_uri"] = "https://oauth2.googleapis.com/token"
    if "auth_uri" not in out: out["auth_uri"] = "https://accounts.google.com/o/oauth2/auth"
    return out

def ler_aba_gsheets(service_account_info: Dict[str, Any], spreadsheet_id: str, worksheet: str) -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id.strip())
    try: ws = sh.worksheet(worksheet.strip())
    except Exception:
        for w in sh.worksheets():
            if w.title.strip().lower() == worksheet.strip().lower(): ws = w; break
        else: raise
    rows = ws.get_all_values()
    if not rows: return pd.DataFrame()
    return pd.DataFrame(rows[1:], columns=[str(c).strip() for c in rows[0]])

def parse_data_serie(serie: pd.Series) -> pd.Series:
    if serie is None: return pd.Series(dtype="datetime64[ns]")
    as_str = serie.map(lambda x: "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip())
    out = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    mask_iso = as_str.str.match(r"^\d{4}-\d{2}-\d{2}", na=False)
    if mask_iso.any():
        has_time = mask_iso & as_str.str.contains("T", na=False)
        date_only = mask_iso & ~has_time
        if has_time.any():
            ts = pd.to_datetime(serie.loc[has_time], errors="coerce", utc=True)
            out.loc[has_time] = ts.dt.tz_convert(_TZ_BR).dt.tz_localize(None)
        if date_only.any():
            out.loc[date_only] = pd.to_datetime(as_str.loc[date_only], format="%Y-%m-%d", errors="coerce")
    vazios = {"", "nan", "none", "nat", "null", "na", "n/a", "-"}
    mask_rest = out.isna() & ~as_str.str.lower().isin(vazios)
    if mask_rest.any():
        out.loc[mask_rest] = pd.to_datetime(serie.loc[mask_rest], dayfirst=True, errors="coerce")
    return out

def _aplicar_secrets_salesforce() -> None:
    try:
        if hasattr(st, "secrets") and "salesforce" in st.secrets:
            sec = st.secrets["salesforce"]
            if sec.get("USER"): os.environ["SALESFORCE_USER"] = str(sec["USER"]).strip()
            if sec.get("PASSWORD"): os.environ["SALESFORCE_PASSWORD"] = str(sec["PASSWORD"]).strip()
            if sec.get("TOKEN"): os.environ["SALESFORCE_TOKEN"] = str(sec["TOKEN"]).strip()
            dom = str(sec.get("DOMAIN") or sec.get("domain") or "").strip()
            if dom: os.environ["SALESFORCE_DOMAIN"] = dom
    except Exception: pass

def conectar_salesforce_app() -> Tuple[Any, Optional[str]]:
    _aplicar_secrets_salesforce()
    try:
        from simple_salesforce import Salesforce
    except ImportError: return None, "Pacote simple-salesforce não instalado."
    username = (os.environ.get("SALESFORCE_USER") or "").strip()
    password = (os.environ.get("SALESFORCE_PASSWORD") or "").strip()
    token = (os.environ.get("SALESFORCE_TOKEN") or "").strip()
    domain = (os.environ.get("SALESFORCE_DOMAIN") or "login").strip() or "login"
    if not username or not password: return None, "Credenciais Salesforce ausentes."
    try:
        kwargs = {"username": username, "password": password, "domain": domain}
        if token: kwargs["security_token"] = token
        return Salesforce(**kwargs), None
    except Exception as e: return None, f"{type(e).__name__}: {e}"

@st.cache_resource(ttl=3300, show_spinner=False)
def _cliente_salesforce_cache():
    sf, err = conectar_salesforce_app()
    if sf is None: raise RuntimeError(err or "Falha ao conectar no Salesforce.")
    return sf

@st.cache_data(ttl=3600, show_spinner=False)
def extrair_dados_sf_cached(ano_alvo: int, mes_alvo: int) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    sf = _cliente_salesforce_cache()
    data_inicio = datetime(ano_alvo - 3, mes_alvo, 1).strftime("%Y-%m-%d")
    
    soql_ag = (
        "SELECT Codigo_do_agendamento__c, CreatedDate, Data_da_Visita__c "
        "FROM Event "
        "WHERE Unidade_de_negocio__c = 'Direcional' "
        "AND Regional__c = 'RJ' "
        f"AND CreatedDate >= {data_inicio}T00:00:00Z"
    )
    try:
        res_ag = sf.query_all(soql_ag)
        df_ag = pd.DataFrame([{
            "Código do agendamento": r.get("Codigo_do_agendamento__c"),
            "Data de criação": r.get("CreatedDate"),
            "Data da visita": r.get("Data_da_Visita__c")
        } for r in (res_ag.get("records") or [])])
    except Exception: df_ag = pd.DataFrame()

    soql_pas = (
        "SELECT Name, CreatedDate, dataPrimeiroEnvioAnalise__c, dataAprovacaoSAFI__c "
        "FROM Avaliacao_credito__c "
        "WHERE Empreendimento__r.Regional__c = 'RJ' "
        f"AND CreatedDate >= {data_inicio}T00:00:00Z"
    )
    try:
        res_pas = sf.query_all(soql_pas)
        df_pas = pd.DataFrame([{
            "Nome da Avaliação de crédito": r.get("Name"),
            "Data de criação": r.get("CreatedDate"),
            "Data Primeiro Envio Análise": r.get("dataPrimeiroEnvioAnalise__c"),
            "Data Aprovação SAFI": r.get("dataAprovacaoSAFI__c")
        } for r in (res_pas.get("records") or [])])
    except Exception: df_pas = pd.DataFrame()

    soql_ven = (
        "SELECT Id, Name, Empreendimento__r.Name, Valor_Real_de_Venda__c, DirecionalVendas__c, "
        "ContratoGeradoEm__c, DataVenda__c, Imobiliaria__r.Name "
        "FROM Opportunity "
        "WHERE DirecionalVendas__c = true "
        "AND Empreendimento__r.Regional__c = 'RJ' "
        f"AND ContratoGeradoEm__c >= {data_inicio} "
        "AND Imobiliaria__r.Name LIKE 'DIR%'"
    )
    try:
        res_ven = sf.query_all(soql_ven)
        df_ven = pd.DataFrame([{
            "ID da Oportunidade": r.get("Id"),
            "Empreendimento": (r.get("Empreendimento__r") or {}).get("Name"),
            "Valor Real de Venda": r.get("Valor_Real_de_Venda__c"),
            "Venda Comercial?": r.get("DirecionalVendas__c"),
            "Contrato gerado em": r.get("ContratoGeradoEm__c"),
            "Data da venda": r.get("DataVenda__c"),
            "Imobiliária": (r.get("Imobiliaria__r") or {}).get("Name")
        } for r in (res_ven.get("records") or [])])
    except Exception: df_ven = pd.DataFrame()

    return df_ag, df_pas, df_ven

def processar_base_diaria(df_ag: pd.DataFrame, df_pas: pd.DataFrame, df_ven: pd.DataFrame, ano_alvo: int, mes_alvo: int) -> Tuple[pd.DataFrame, date, date]:
    inicio = date(ano_alvo - 3, mes_alvo, 1)
    fim_treino = date(ano_alvo, mes_alvo, 1) - timedelta(days=1)
    
    idx = pd.date_range(inicio, date(ano_alvo, mes_alvo, calendar.monthrange(ano_alvo, mes_alvo)[1]), freq="D")
    cal = pd.DataFrame({"data": [d.date() for d in idx]})
    cal["dia_mes"] = cal["data"].map(lambda d: d.day)
    cal["dia_semana"] = cal["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    cal["mes"] = cal["data"].map(lambda d: MESES_PT[d.month])
    cal["ano_num"] = cal["data"].map(lambda d: d.year)

    def preencher_etapa(df, col_data):
        if df is None or df.empty or col_data not in df.columns: return pd.Series(0.0, index=cal.index)
        dt = parse_data_serie(df[col_data]).dropna()
        if dt.empty: return pd.Series(0.0, index=cal.index)
        vc = dt.dt.normalize().value_counts()
        return cal["data"].map(lambda d: float(vc.get(pd.Timestamp(d), 0.0)))

    cal["agendamentos"] = preencher_etapa(df_ag, "Data de criação")
    cal["visitas"] = preencher_etapa(df_ag, "Data da visita")
    cal["pastas"] = preencher_etapa(df_pas, "Data Primeiro Envio Análise")
    cal["pastas_aprovadas"] = preencher_etapa(df_pas, "Data Aprovação SAFI")
    cal["vendas"] = preencher_etapa(df_ven, "Contrato gerado em")
    
    return cal, inicio, fim_treino

def _matriz_explicativas_com_tendencia(df: pd.DataFrame) -> Tuple[np.ndarray, List[str]]:
    n = len(df)
    n_cols = 31 + 7 + 12 + 1 + 1 
    X = np.zeros((n, n_cols), dtype=float)
    X[:, -1] = 1.0 

    dias_semana_idx = {nome: i for i, nome in DIAS_SEMANA_PT.items()}
    meses_idx = {nome: i for i, nome in MESES_PT.items()}
    data_min = pd.to_datetime(df["data"]).min()

    for i, row in enumerate(df.itertuples(index=False)):
        dia = int(row.dia_mes)
        if 1 <= dia <= 31: X[i, dia - 1] = 1.0
        ds = dias_semana_idx.get(str(row.dia_semana), None)
        if ds is not None: X[i, 31 + ds] = 1.0
        ms = meses_idx.get(str(row.mes), None)
        if ms is not None: X[i, 31 + 7 + (ms - 1)] = 1.0
        
        dt_atual = pd.to_datetime(row.data)
        anos_passados = (dt_atual - data_min).days / 365.25
        X[i, 31 + 7 + 12] = anos_passados

    names = [f"dia_{d}" for d in range(1, 32)] + list(DIAS_SEMANA_PT.values()) + list(MESES_PT.values()) + ["tendencia_anos", "intercepto"]
    return X, names

def treinar_regressao_com_estatisticas(treino: pd.DataFrame, alvo: str) -> Dict[str, Any]:
    X, feature_names = _matriz_explicativas_com_tendencia(treino)
    y = treino[alvo].astype(float).values
    beta, residuals, rank, s = np.linalg.lstsq(X, y, rcond=None)
    
    n = len(y)
    p = X.shape[1]
    y_hat = X @ beta
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    r2 = 1.0 - (ss_res / ss_tot) if ss_tot > 0 else 0.0
    mae = float(np.mean(np.abs(y - y_hat)))
    
    sigma2 = ss_res / max(1, n - p)
    try:
        cov_matrix = sigma2 * np.linalg.inv(X.T @ X)
        se_beta = np.sqrt(np.diagonal(cov_matrix))
    except Exception:
        se_beta = np.zeros_like(beta)
        
    t_stats = np.zeros_like(beta)
    p_values = np.ones_like(beta)
    for i in range(len(beta)):
        if se_beta[i] > 1e-9:
            t_stats[i] = beta[i] / se_beta[i]
            p_values[i] = float(2.0 * (1.0 - 0.5 * (1.0 + math.erf(abs(t_stats[i]) / math.sqrt(2.0)))))

    tabela_betas = pd.DataFrame({
        "Variável": feature_names,
        "Beta": beta,
        "Erro Padrão": se_beta,
        "Estatística t": t_stats,
        "Significância (p-value)": p_values
    })

    return {"beta": beta, "r2": r2, "mae": mae, "tabela_betas": tabela_betas}

def projetar_mes_alvo_diario(treino: pd.DataFrame, ano_alvo: int, mes_alvo: int, modelos_treinados: Dict[str, Any]) -> Dict[str, pd.DataFrame]:
    dias_mes = calendar.monthrange(ano_alvo, mes_alvo)[1]
    idx = pd.date_range(date(ano_alvo, mes_alvo, 1), date(ano_alvo, mes_alvo, dias_mes), freq="D")
    df_alvo = pd.DataFrame({"data": [d.date() for d in idx]})
    df_alvo["dia_mes"] = df_alvo["data"].map(lambda d: d.day)
    df_alvo["dia_semana"] = df_alvo["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    df_alvo["mes"] = df_alvo["data"].map(lambda d: MESES_PT[d.month])
    df_alvo["ano_num"] = ano_alvo

    X_alvo, _ = _matriz_explicativas_com_tendencia(df_alvo)
    
    resultados_proj = {}
    for etapa in FUNIL_ETAPAS:
        mod = modelos_treinados[etapa]
        beta = mod["beta"]
        y_proj = X_alvo @ beta
        y_proj = np.maximum(y_proj, 0.0) 
        
        margem = 0.5 
        df_res = pd.DataFrame({
            "dia": df_alvo["dia_mes"],
            "projetado": y_proj,
            "limite_inferior": np.maximum(0.0, y_proj - margem),
            "limite_superior": y_proj + margem
        })
        resultados_proj[etapa] = df_res
        
    return resultados_proj

def carregar_metas_sheets() -> pd.DataFrame:
    try:
        raw = _secrets_connections_gsheets()
        info = montar_service_account_info(raw)
        if not info: return pd.DataFrame()
        df = ler_aba_gsheets(info, SPREADSHEET_METAS_ID, "Meta Mensal")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

def calcular_conversoes_mensais(cal: pd.DataFrame) -> Dict[str, Any]:
    cal["ano_mes"] = pd.to_datetime(cal["data"]).dt.to_period("M")
    mensal = cal.groupby("ano_mes")[list(FUNIL_ETAPAS)].sum().reset_index()
    mensal["mes_num"] = mensal["ano_mes"].dt.month
    mensal["ano_num"] = mensal["ano_mes"].dt.year
    mensal["tempo_idx"] = (mensal["ano_num"] - mensal["ano_num"].min()) * 12 + mensal["mes_num"]

    pares = [
        ("agendamentos", "vendas", "Agendamentos → Vendas"),
        ("visitas", "vendas", "Visitas → Vendas"),
        ("pastas", "vendas", "Pastas → Vendas"),
        ("pastas_aprovadas", "vendas", "Pastas Aprovadas → Vendas"),
    ]

    conversoes_res = {}
    for origem, destino, label in pares:
        sub = mensal[mensal[origem] > 0].copy()
        sub["taxa"] = (sub[destino] / sub[origem]) * 100.0
        
        X_m = np.column_stack([np.ones(len(sub)), sub["mes_num"].values, sub["tempo_idx"].values])
        y_m = sub["taxa"].values
        beta_m, *_ = np.linalg.lstsq(X_m, y_m, rcond=None)
        
        sub["taxa_projetada"] = X_m @ beta_m
        
        conversoes_res[label] = {
            "df": sub,
            "beta": beta_m,
            "media": float(sub["taxa"].mean()),
            "mediana": float(sub["taxa"].median())
        }
    return conversoes_res

def main() -> None:
    st.set_page_config(page_title="Acompanhamento e Projeção Sazonal", layout="wide")
    aplicar_estilo()

    st.markdown('<div class="ficha-logo-wrap"><h1 style="color: #04428f; font-weight: 900;">Acompanhamento e Projeção Sazonal</h1></div>', unsafe_allow_html=True)

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        anos_disp = list(range(datetime.now().year - 1, datetime.now().year + 2))
        ano_alvo = st.selectbox("Ano Alvo", anos_disp, index=anos_disp.index(datetime.now().year))
    with col_f2:
        meses_disp = list(MESES_PT.keys())
        mes_alvo = st.selectbox("Mês Alvo", meses_disp, format_func=lambda x: MESES_PT[x].capitalize(), index=datetime.now().month - 1)

    prog_placeholder = st.empty()
    t_start = time.time()

    def atualizar_progresso(pct, msg):
        elapsed = time.time() - t_start
        grad = f"linear-gradient(90deg, {COR_AZUL_ESC} 0%, {COR_VERMELHO} 100%)"
        html_str = f"""
        <div style="margin: 1.5rem 0; padding: 1.25rem; background: rgba(255,255,255,0.9); border-radius: 12px; border: 1px solid #e2e8f0;">
            <div style="display: flex; justify-content: space-between; font-weight: 600; margin-bottom: 0.5rem; color: #1e293b;">
                <span>{msg}</span>
                <span style="font-family: monospace;">{pct}% | {elapsed:.1f}s</span>
            </div>
            <div style="width: 100%; background-color: #cbd5e1; border-radius: 999px; height: 10px; overflow: hidden;">
                <div style="width: {pct}%; background: {grad}; height: 100%;"></div>
            </div>
        </div>
        """
        prog_placeholder.markdown(html_str, unsafe_allow_html=True)

    atualizar_progresso(15, "Iniciando consultas no Salesforce...")
    try:
        df_ag, df_pas, df_ven = extrair_dados_sf_cached(ano_alvo, mes_alvo)
    except Exception as e:
        st.error(f"Erro ao conectar no Salesforce: {e}")
        return

    atualizar_progresso(60, "Processando matrizes e calculando regressões...")
    cal, inicio, fim_treino = processar_base_diaria(df_ag, df_pas, df_ven, ano_alvo, mes_alvo)

    modelos_treinados = {}
    for etapa in FUNIL_ETAPAS:
        modelos_treinados[etapa] = treinar_regressao_com_estatisticas(cal, etapa)

    proj_diaria = projetar_mes_alvo_diario(cal, ano_alvo, mes_alvo, modelos_treinados)
    conv_mensal = calcular_conversoes_mensais(cal)

    atualizar_progresso(90, "Lendo Metas do Google Sheets...")
    df_metas_sheets = carregar_metas_sheets()
    
    meta_vendas_mes = 0.0
    if not df_metas_sheets.empty:
        col_ano = [c for c in df_metas_sheets.columns if "ano" in c.lower()]
        col_mes = [c for c in df_metas_sheets.columns if "mês" in c.lower() or "mes" in c.lower()]
        col_meta = [c for c in df_metas_sheets.columns if "meta" in c.lower()]
        if col_ano and col_mes and col_meta:
            sub_m = df_metas_sheets[
                (pd.to_numeric(df_metas_sheets[col_ano[0]], errors="coerce") == ano_alvo) & 
                (pd.to_numeric(df_metas_sheets[col_mes[0]], errors="coerce") == mes_alvo)
            ]
            if not sub_m.empty:
                meta_vendas_mes = float(pd.to_numeric(sub_m[col_meta[0]], errors="coerce").sum())

    atualizar_progresso(100, "Concluído!")
    time.sleep(0.3)
    prog_placeholder.empty()

    # -------------------------------------------------------------------------
    # Renderização das Abas
    # -------------------------------------------------------------------------
    tab1, tab2, tab3, tab4 = st.tabs(["Projeção Diária", "Estatísticas & Betas", "Conversões Mensais", "Metas x Realizado"])

    with tab1:
        st.subheader(f"Projeção Diária — {MESES_PT[mes_alvo].capitalize()}/{ano_alvo}")
        
        hoje = date.today()
        eh_mes_atual = (hoje.year == ano_alvo and hoje.month == mes_alvo)
        dia_limite = hoje.day if eh_mes_atual else calendar.monthrange(ano_alvo, mes_alvo)[1]

        sub_real = cal[(pd.to_datetime(cal["data"]).dt.year == ano_alvo) & (pd.to_datetime(cal["data"]).dt.month == mes_alvo)]

        for etapa in FUNIL_ETAPAS:
            st.markdown(f"##### {FUNIL_LABELS[etapa]} — Projetado x Realizado (%)")
            df_p = proj_diaria[etapa]
            
            soma_real_mes = float(sub_real[sub_real["dia_mes"] <= dia_limite][etapa].sum()) if not sub_real.empty else 1.0
            if soma_real_mes <= 0: soma_real_mes = 1.0

            df_p["projetado_pct"] = (df_p["projetado"] / df_p["projetado"].sum()) * 100.0
            df_p["inf_pct"] = df_p["projetado_pct"] - 0.5
            df_p["sup_pct"] = df_p["projetado_pct"] + 0.5

            real_pct_lista = []
            for d in df_p["dia"]:
                if eh_mes_atual and d > dia_limite:
                    real_pct_lista.append(np.nan)
                else:
                    val_dia = float(sub_real[sub_real["dia_mes"] == d][etapa].sum()) if not sub_real.empty else 0.0
                    real_pct_lista.append((val_dia / soma_real_mes) * 100.0)
            df_p["realizado_pct"] = real_pct_lista

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=pd.concat([df_p["dia"], df_p["dia"][::-1]]),
                y=pd.concat([df_p["sup_pct"], df_p["inf_pct"][::-1]]),
                fill='tozerox', fillcolor='rgba(203, 9, 53, 0.12)',
                line=dict(color='rgba(255,255,255,0)'), name="Intervalo (±0,5%)", showlegend=True
            ))
            fig.add_trace(go.Scatter(
                x=df_p["dia"], y=df_p["projetado_pct"], mode="lines+markers", name="Projetado",
                line=dict(color=COR_VERMELHO, width=2, dash="dash"), marker=dict(size=6)
            ))
            fig.add_trace(go.Scatter(
                x=df_p["dia"], y=df_p["realizado_pct"], mode="lines+markers", name="Realizado",
                line=dict(color=COR_AZUL_ESC, width=3), marker=dict(size=7)
            ))

            fig.update_layout(
                margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                height=320, hovermode="x unified", legend=dict(orientation="h", y=1.15, x=0.5, xanchor="center")
            )
            fig.update_yaxes(title_text="Representatividade (%)", ticksuffix="%")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with tab2:
        st.subheader("Estatísticas e Coeficientes (Betas)")
        for etapa in FUNIL_ETAPAS:
            st.markdown(f"##### Indicador: {FUNIL_LABELS[etapa]}")
            m = modelos_treinados[etapa]
            c1, c2 = st.columns(2)
            with c1: st.metric("R² do Modelo", f"{m['r2']:.4f}")
            with c2: st.metric("Desvio Absoluto Médio (MAE)", f"{m['mae']:.2f}")
            st.dataframe(m["tabela_betas"], use_container_width=True, hide_index=True)

    with tab3:
        st.subheader("Projeção e Histórico de Conversões Mensais")
        for label, dados in conv_mensal.items():
            st.markdown(f"##### {label}")
            sub = dados["df"]
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=sub["ano_mes"].astype(str), y=sub["taxa"], mode="lines+markers", name="Realizado Histórico",
                line=dict(color=COR_AZUL_ESC, width=2), marker=dict(size=6)
            ))
            fig.add_trace(go.Scatter(
                x=sub["ano_mes"].astype(str), y=sub["taxa_projetada"], mode="lines+markers", name="Regressão Mensal (Tendência + Sazonalidade)",
                line=dict(color=COR_VERMELHO, width=2, dash="dash"), marker=dict(size=4)
            ))
            fig.update_layout(
                margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                height=300, hovermode="x unified", legend=dict(orientation="h", y=1.15, x=0.5, xanchor="center")
            )
            fig.update_yaxes(title_text="Taxa de Conversão (%)", ticksuffix="%")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
            
            col_m1, col_m2 = st.columns(2)
            with col_m1: st.metric("Média Histórica", f"{dados['media']:.2f}%")
            with col_m2: st.metric("Mediana Histórica", f"{dados['mediana']:.2f}%")

    with tab4:
        st.subheader(f"Acompanhamento de Metas de {MESES_PT[mes_alvo].capitalize()}/{ano_alvo}")
        st.caption("As metas de topo de funil foram calculadas com base na proporção histórica de conversão do respectivo mês em relação às vendas.")

        sub_mes_hist = cal[cal["mes"] == MESES_PT[mes_alvo]]
        total_vendas_mes_hist = sub_mes_hist["vendas"].sum() if not sub_mes_hist.empty else 1.0
        if total_vendas_mes_hist <= 0: total_vendas_mes_hist = 1.0

        metas_dict = {"vendas": meta_vendas_mes}
        for etapa in ("agendamentos", "visitas", "pastas", "pastas_aprovadas"):
            soma_etapa = sub_mes_hist[etapa].sum() if not sub_mes_hist.empty else 0.0
            razao = soma_etapa / total_vendas_mes_hist if total_vendas_mes_hist > 0 else 1.0
            metas_dict[etapa] = meta_vendas_mes * max(1.0, razao)

        metas_tabela = []
        for etapa in FUNIL_ETAPAS:
            m_est = metas_dict[etapa]
            r_acum = float(sub_real[sub_real["dia_mes"] <= dia_limite][etapa].sum()) if not sub_real.empty else 0.0
            ating = (r_acum / m_est * 100.0) if m_est > 0 else 0.0
            metas_tabela.append({
                "Indicador": FUNIL_LABELS[etapa],
                "Meta Mensal Estimada": f"{int(m_est):,}".replace(",", "."),
                "Realizado Acumulado": f"{int(r_acum):,}".replace(",", "."),
                "Atingimento (%)": f"{ating:.1f}%"
            })

        st.dataframe(pd.DataFrame(metas_tabela), use_container_width=True, hide_index=True)

        st.markdown("##### Meta x Realizado (Volume Acumulado Diário)")
        for etapa in FUNIL_ETAPAS:
            st.markdown(f"##### {FUNIL_LABELS[etapa]} — Meta x Realizado (Volume Acumulado)")
            df_p = proj_diaria[etapa].copy()
            meta_etapa_total = metas_dict[etapa]
            
            pesos_dia = df_p["projetado"] / df_p["projetado"].sum() if df_p["projetado"].sum() > 0 else (1/len(df_p))
            df_p["meta_diaria"] = meta_etapa_total * pesos_dia
            df_p["meta_acum"] = df_p["meta_diaria"].cumsum()

            real_acum_lista = []
            soma_parcial = 0.0
            for d in df_p["dia"]:
                if eh_mes_atual and d > dia_limite:
                    real_acum_lista.append(np.nan)
                else:
                    val_dia = float(sub_real[sub_real["dia_mes"] == d][etapa].sum()) if not sub_real.empty else 0.0
                    soma_parcial += val_dia
                    real_acum_lista.append(soma_parcial)
            df_p["realizado_acum"] = real_acum_lista

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=df_p["dia"], y=df_p["realizado_acum"], mode="lines+markers", name="Realizado (Acum.)",
                line=dict(color=COR_AZUL_ESC, width=3), marker=dict(size=7)
            ))
            fig.add_trace(go.Scatter(
                x=df_p["dia"], y=df_p["meta_acum"], mode="lines+markers", name="Meta Projetada (Acum.)",
                line=dict(color=COR_VERMELHO, width=2, dash="dash"), marker=dict(size=6)
            ))
            fig.update_layout(
                margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                height=300, hovermode="x unified", legend=dict(orientation="h", y=1.15, x=0.5, xanchor="center")
            )
            fig.update_yaxes(title_text="Volume Acumulado")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # -------------------------------------------------------------------------
    # Botão de Download
    # -------------------------------------------------------------------------
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:2rem 0;'/>", unsafe_allow_html=True)
    
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for etapa in FUNIL_ETAPAS:
            proj_diaria[etapa].to_excel(writer, sheet_name=FUNIL_LABELS[etapa], index=False)
        pd.DataFrame(metas_tabela).to_excel(writer, sheet_name="Resumo Metas", index=False)
    
    st.download_button(
        label="📥 Baixar Base de Projeções (Excel)",
        data=output.getvalue(),
        file_name=f"projecoes_sazonais_{mes_alvo:02d}_{ano_alvo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if __name__ == "__main__":
    main()
