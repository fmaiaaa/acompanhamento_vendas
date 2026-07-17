# -*- coding: utf-8 -*-
"""
Base compartilhada dos relatórios de funil por Regional / Gerente / Corretor.
Carrega Salesforce, monta eventos diários e aplica as mesmas regras de data
do velocímetro (pastas = 1º envio; aprovadas = SAFI; vendas comerciais).
"""
from __future__ import annotations

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


def aplicar_estilo_basico() -> None:
    st.markdown(
        f"""
        <style>
        .stApp {{ background: linear-gradient(160deg, #f4f7fb 0%, #eef2f7 40%, #f8fafc 100%); }}
        h1, h2, h3, h4, h5, h6, p, label, span, div {{ color: {COR_TEXTO_PRETO} !important; }}
        .bloco-pessoa {{
            background: {COR_FUNDO_CARD};
            border: 1px solid {COR_BORDA};
            border-radius: 12px;
            padding: 0.75rem 1rem;
            margin-bottom: 0.85rem;
            box-shadow: 0 4px 18px rgba(4,66,143,0.06);
        }}
        .bloco-pessoa .nome {{
            font-weight: 700;
            font-size: 1.05rem;
            color: {COR_AZUL_ESC} !important;
            margin-bottom: 0.4rem;
            padding-bottom: 0.35rem;
            border-bottom: 1px solid {COR_BORDA};
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


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
