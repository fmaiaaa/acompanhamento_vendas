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

import re
from typing import Any, Dict, List, Optional

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# -----------------------------------------------------------------------------
# Identificação da planilha (ID extraído da URL)
# -----------------------------------------------------------------------------
SPREADSHEET_ID = "1wpuNQvksot9CLhGgQRe7JlyDeRISEh_sc3-6VRDyQYk"
WS_VENDAS = "BD Vendas Completa"
WS_METAS = "Metas"

# Paleta alinhada à Ficha Credenciamento / Vendas RJ
COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_VERMELHO_ESCURO = "#9e0828"
COR_FUNDO_CARD = "rgba(255, 255, 255, 0.78)"
COR_BORDA = "#eef2f6"
COR_TEXTO_MUTED = "#64748b"
COR_TEXTO_LABEL = "#1e293b"

MESES_COL_MAP = {
    "jan./1": 1,
    "fev./1": 2,
    "mar./1": 3,
    "abr./1": 4,
    "mai./1": 5,
    "jun./1": 6,
    "jul./1": 7,
    "ago./1": 8,
    "set./1": 9,
    "out./1": 10,
    "nov./1": 11,
    "dez./1": 12,
}

MESES_TEXTO_MAP = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
}


def _secrets_connections_gsheets() -> Dict[str, Any]:
    """Lê `[connections.gsheets]` do secrets.toml → dict."""
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
    """Chave PEM colada no TOML com \\n literais vira quebras reais."""
    s = (pk or "").strip()
    if not s:
        return s
    if "\\n" in s and "\n" not in s:
        s = s.replace("\\n", "\n")
    return s


def montar_service_account_info(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Monta o dict esperado por `google.oauth2.service_account.Credentials.from_service_account_info`.
    Aceita `type` vazio nos secrets — assume service_account se houver private_key + client_email.
    """
    if not raw:
        return None
    chaves = (
        "type",
        "project_id",
        "private_key_id",
        "private_key",
        "client_email",
        "client_id",
        "auth_uri",
        "token_uri",
        "auth_provider_x509_cert_url",
        "client_x509_cert_url",
    )
    out: Dict[str, Any] = {}
    for k in chaves:
        v = raw.get(k)
        if v is None:
            continue
        if isinstance(v, str):
            v = v.strip()
        if v == "":
            continue
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
        if v:
            return v
    return SPREADSHEET_ID


def email_config_de_secrets() -> Dict[str, Any]:
    """Lê `[email]` — SMTP (reservado para alertas futuros; não usado no fluxo atual)."""
    try:
        e = st.secrets.get("email", {})
        if not isinstance(e, dict):
            return {}
        port_raw = e.get("smtp_port", 587)
        try:
            port = int(port_raw) if port_raw not in (None, "") else 587
        except (TypeError, ValueError):
            port = 587
        return {
            "smtp_server": str(e.get("smtp_server") or "").strip(),
            "smtp_port": port,
            "sender_email": str(e.get("sender_email") or "").strip(),
            "sender_password": str(e.get("sender_password") or "").strip(),
        }
    except Exception:
        return {}


def valores_para_dataframe(rows: List[List[str]]) -> pd.DataFrame:
    """Primeira linha = cabeçalho; alinha largura das linhas."""
    if not rows:
        return pd.DataFrame()
    header = [str(c).strip() for c in rows[0]]
    w = len(header)
    if w == 0:
        return pd.DataFrame()
    body = rows[1:]
    if not body:
        return pd.DataFrame(columns=header)
    norm: List[List[str]] = []
    for r in body:
        cells = [str(c) for c in r]
        if len(cells) < w:
            cells = cells + [""] * (w - len(cells))
        else:
            cells = cells[:w]
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
                if w.title.strip() == nome:
                    return w
            for w in sh.worksheets():
                if w.title.strip().lower() == nome.lower():
                    return w
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
    """
    Lê uma aba via gspread. `_cred_fp` deve ser o fingerprint da private_key nos secrets
    para invalidar o cache quando a chave mudar.
    """
    raw = _secrets_connections_gsheets()
    info = montar_service_account_info(raw)
    if not info:
        raise ValueError("Credenciais [connections.gsheets] ausentes ou incompletas.")
    return ler_aba_gsheets(info, spreadsheet_id, worksheet)


def _hex_rgb(hex_color: str) -> str:
    x = (hex_color or "").strip().lstrip("#")
    if len(x) != 6:
        return "4, 66, 143"
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}"


RGB_AZUL = _hex_rgb(COR_AZUL_ESC)
RGB_VERM = _hex_rgb(COR_VERMELHO)


def parse_valor_br(val: Any) -> float:
    """Converte valores como 215.625,50 ou 247072,61 para float."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace("R$", "").replace(" ", "")
    if not s or s.lower() == "nan":
        return 0.0
    # Emoji / texto colado no ranking (ex.: Ouro🥇) — ignora sufixo
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s:
        return 0.0
    if "," in s and "." in s:
        # BR: milhar com ponto, decimal com vírgula
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def extrair_mes(val: Any) -> Optional[int]:
    """Extrai mês numérico com alta precisão para diferentes formatações (ex: Maio/2026, 05/2026)."""
    if pd.isna(val): return None
    v = str(val).strip().lower()
    if not v: return None
    
    # 1. Busca por nome do mês dentro da string
    for m_str, m_num in MESES_TEXTO_MAP.items():
        if m_str in v:
            return m_num
            
    # 2. Busca por número solto (ex: 5 ou 05)
    try:
        m = int(float(v))
        if 1 <= m <= 12: return m
    except ValueError:
        pass
        
    # 3. Busca em formatos de data com barra (dd/mm/yyyy ou mm/yyyy)
    if '/' in v:
        p = v.split('/')
        if len(p) == 3:  # dd/mm/yyyy
            try:
                m = int(p[1])
                if 1 <= m <= 12: return m
            except: pass
        elif len(p) == 2:  # mm/yyyy
            try:
                m = int(p[0])
                if 1 <= m <= 12: return m
            except: pass
            
    return None


def extrair_ano(val: Any) -> Optional[int]:
    """Extrai ano numérico de strings complexas."""
    if pd.isna(val): return None
    v = str(val).strip()
    if not v: return None
    
    # 1. Ano direto como número
    try:
        ano = int(float(v))
        if ano > 2000: return ano
    except ValueError:
        pass
        
    # 2. Ano presente após a última barra
    if '/' in v:
        p = v.split('/')
        try:
            ano = int(p[-1])
            if ano > 2000: return ano
        except: pass
        
    return None


def achar_coluna(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    """Busca uma coluna no DataFrame por uma lista de aliases (busca exata ou parcial)."""
    for a in aliases:
        for c in df.columns:
            if a.lower() == str(c).strip().lower():
                return c
    for a in aliases:
        for c in df.columns:
            if a.lower() in str(c).strip().lower():
                return c
    return None


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def coluna_existe(df: pd.DataFrame, nome: str) -> bool:
    return nome in df.columns


def melt_metas(df_metas_raw: pd.DataFrame) -> pd.DataFrame:
    """Transforma Metas largas em longas (mês numérico + meta quantidade)."""
    df = normalizar_colunas(df_metas_raw)
    cols_mes = [c for c in df.columns if c in MESES_COL_MAP]
    if not cols_mes:
        # tenta casar variações com espaço
        alt = {}
        for c in df.columns:
            c0 = c.strip().lower().replace(" ", "")
            for k, v in MESES_COL_MAP.items():
                if k.replace("/", "").replace(".", "") in c0 or k in c:
                    alt[c] = v
                    break
        cols_mes = [c for c in df.columns if c in alt]
        if not cols_mes:
            return pd.DataFrame(columns=["Empreendimento", "Região", "Mes_Num", "Meta_Qtd"])

    id_vars = [c for c in df.columns if c.lower() in ["empreendimento", "região", "regiao", "obra"]]
    if not id_vars:
        return pd.DataFrame(columns=["Empreendimento", "Região", "Mes_Num", "Meta_Qtd"])

    out = df.melt(
        id_vars=id_vars,
        value_vars=cols_mes,
        var_name="Mes_Texto",
        value_name="Meta_Qtd",
    )
    out["Mes_Num"] = out["Mes_Texto"].map(lambda x: MESES_COL_MAP.get(x, MESES_COL_MAP.get(str(x).strip(), None)))
    out = out.dropna(subset=["Mes_Num"])
    out["Mes_Num"] = out["Mes_Num"].astype(int)
    out["Meta_Qtd"] = pd.to_numeric(out["Meta_Qtd"], errors="coerce").fillna(0)
    
    # Normaliza colunas ID para garantir consistência
    for c in out.columns:
        if c.lower() in ["empreendimento", "obra"]:
            out.rename(columns={c: "Empreendimento"}, inplace=True)
        elif c.lower() in ["região", "regiao"]:
            out.rename(columns={c: "Região"}, inplace=True)
            
    return out


def aplicar_estilo() -> None:
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800;900&family=Inter:wght@400;500;600;700&display=swap');
        html, body, [data-testid="stAppViewContainer"] {{
            font-family: 'Inter', sans-serif;
            color: {COR_TEXTO_LABEL};
        }}
        .stApp {{
            background: linear-gradient(135deg, rgba({RGB_AZUL}, 0.88) 0%, rgba(30, 58, 95, 0.55) 45%,
                rgba({RGB_VERM}, 0.18) 72%, rgba(15, 23, 42, 0.42) 100%) !important;
        }}
        [data-testid="stAppViewContainer"] {{ background: transparent !important; }}
        [data-testid="stHeader"] {{
            background: transparent !important;
            background-color: transparent !important;
        }}
        [data-testid="stSidebar"] {{
            background: rgba(255, 255, 255, 0.92) !important;
            border-right: 1px solid {COR_BORDA} !important;
        }}
        .block-container {{
            max-width: 1200px !important;
            padding: 1.5rem 2rem 2rem 2rem !important;
            background: {COR_FUNDO_CARD} !important;
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border-radius: 24px !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            box-shadow: 0 24px 48px -12px rgba({RGB_AZUL}, 0.2) !important;
        }}
        h1, h2, h3 {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_AZUL_ESC} !important;
            font-weight: 800 !important;
        }}
        .vel-hero-bar {{
            height: 4px;
            width: 100%;
            border-radius: 999px;
            margin: 0.75rem 0 1.25rem 0;
            background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC});
            background-size: 200% 100%;
        }}
        .vel-kpi-row {{
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            margin-bottom: 1.25rem;
        }}
        .vel-kpi {{
            flex: 1 1 160px;
            background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.9) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 14px;
            padding: 14px 16px;
            text-align: center;
            box-shadow: 0 2px 8px rgba({RGB_AZUL}, 0.06);
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
        </style>
        """,
        unsafe_allow_html=True,
    )


def fmt_br_milhoes(v: float) -> str:
    if v >= 1e6:
        return f"R$ {v / 1e6:.2f} mi"
    if v >= 1e3:
        return f"R$ {v / 1e3:.1f} mil"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def criar_medidor(titulo: str, realizado: float, meta: float, vgv: float, vendas_qtd: int) -> None:
    """Gauge estilo velocímetro com cores da marca."""
    meta_f = float(meta) if meta and meta > 0 else 0.0
    perc = min(150.0, (realizado / meta_f * 100.0)) if meta_f > 0 else 0.0
    # Eixo até 100% para leitura; valores >100% ainda aparecem no número
    axis_max = 100
    val_display = min(perc, axis_max) if perc <= axis_max else axis_max

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=val_display,
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
                "bar": {"color": COR_AZUL_ESC},
                "bgcolor": "white",
                "borderwidth": 2,
                "bordercolor": COR_BORDA,
                "steps": [
                    {"range": [0, 40], "color": "rgba(203, 9, 53, 0.25)"},
                    {"range": [40, 80], "color": "rgba(4, 66, 143, 0.12)"},
                    {"range": [80, 100], "color": "rgba(22, 163, 74, 0.28)"},
                ],
                "threshold": {
                    "line": {"color": COR_VERMELHO, "width": 3},
                    "thickness": 0.8,
                    "value": 100,
                },
            },
        )
    )
    if perc > 100:
        fig.update_layout(
            annotations=[
                dict(
                    text=f"Real: {perc:.1f}% da meta",
                    x=0.5,
                    y=0.15,
                    xref="paper",
                    yref="paper",
                    showarrow=False,
                    font=dict(size=12, color=COR_VERMELHO, family="Inter"),
                )
            ]
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
        <div style="text-align:center;font-size:0.9rem;color:{COR_TEXTO_LABEL};margin-top:-8px;">
            <strong>{int(vendas_qtd)}</strong> vendas · VGV <strong>{fmt_br_milhoes(float(vgv))}</strong> · 
            Meta <strong>{meta_f:g}</strong> un.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(
        page_title="Acompanhamento de Vendas | Direcional",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    aplicar_estilo()

    st.markdown(
        f"""
        <div style="text-align:center;padding:0.25rem 0 0 0;">
            <h1 style="margin:0;font-size:clamp(1.35rem,3vw,1.85rem);">Acompanhamento de metas de vendas</h1>
            <p style="color:{COR_TEXTO_MUTED};margin:0.5rem 0 0 0;font-size:0.95rem;">
                Realizado vs metas por período — BD Vendas Completa e aba Metas
            </p>
            <div class="vel-hero-bar"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    raw_gs = _secrets_connections_gsheets()
    info = montar_service_account_info(raw_gs)
    if not info:
        st.error(
            "Credenciais Google em **[connections.gsheets]** incompletas. "
            "Preencha pelo menos **private_key** e **client_email** (JSON da conta de serviço). "
            "O campo **type** pode ser `service_account` ou ficar vazio."
        )
        return

    sid = spreadsheet_id_de_secrets(raw_gs)
    cred_fp = _fingerprint_credenciais(info)

    try:
        df_vendas = ler_planilha_aba_df(sid, WS_VENDAS, cred_fp)
        df_metas_raw = ler_planilha_aba_df(sid, WS_METAS, cred_fp)
    except Exception as e:
        err = str(e)
        st.error(f"Erro ao ler a planilha: {err}")
        st.info(
            f"**ID:** `{sid}`  \n"
            f"**Abas:** `{WS_VENDAS}`, `{WS_METAS}`  \n"
            f"**Service account:** `{info.get('client_email', '')}`"
        )
        return

    df_vendas = normalizar_colunas(df_vendas)
    df_metas = melt_metas(df_metas_raw)

    # -------------------------------------------------------------------------
    # Filtra as linhas de TOTAL que vem das planilhas para não duplicar somas
    # -------------------------------------------------------------------------
    if "Empreendimento" in df_metas.columns:
        df_metas = df_metas[~df_metas["Empreendimento"].astype(str).str.strip().str.lower().isin(["total", "geral", ""])]
    if "Região" in df_metas.columns:
        df_metas = df_metas[~df_metas["Região"].astype(str).str.strip().str.lower().isin(["total", "geral"])]

    # -------------------------------------------------------------------------
    # Mapeamento Inteligente de Colunas
    # -------------------------------------------------------------------------
    col_ano = achar_coluna(df_vendas, ["Ano da Venda", "Ano Venda", "Ano"])
    col_mes = achar_coluna(df_vendas, ["Mês da Venda - Looker", "Mês da Venda", "Mês Venda", "Mes Venda", "Mês", "Mes"])
    col_regiao = achar_coluna(df_vendas, ["Região", "Regiao"])
    col_canal = achar_coluna(df_vendas, ["Canal"])
    col_valor = achar_coluna(df_vendas, ["Valor Real de Venda", "Valor Real", "Valor"])
    col_emp = achar_coluna(df_vendas, ["Empreendimento", "Obra", "Nome do Empreendimento"])
    col_venda_comercial = achar_coluna(df_vendas, ["Venda Comercial?", "Venda Comercial"])

    if not col_ano and not col_mes:
        st.error(
            "Não foi possível encontrar as colunas de Ano e Mês na aba de vendas. "
            f"Colunas disponíveis: {', '.join(df_vendas.columns[:25].tolist())}"
        )
        return

    # Padroniza as colunas de Vendas para baterem com as das Metas
    if col_emp and col_emp != "Empreendimento":
        df_vendas.rename(columns={col_emp: "Empreendimento"}, inplace=True)
        col_emp = "Empreendimento"
    if col_regiao and col_regiao != "Região":
        df_vendas.rename(columns={col_regiao: "Região"}, inplace=True)
        col_regiao = "Região"

    # Remove também eventuais linhas de "Total" na aba de vendas
    if col_emp:
        df_vendas = df_vendas[~df_vendas[col_emp].astype(str).str.strip().str.lower().isin(["total", "geral", ""])]

    # -------------------------------------------------------------------------
    # Filtro Obrigatório: Apenas Venda Comercial == 1
    # -------------------------------------------------------------------------
    if col_venda_comercial:
        # Usa pd.to_numeric com coerce para forçar tudo que não for número virar NaN
        # Depois filtra onde o valor numérico é igual a 1
        df_vendas = df_vendas[pd.to_numeric(df_vendas[col_venda_comercial], errors='coerce') == 1]
    else:
        st.warning("Coluna 'Venda Comercial?' não encontrada na base. O filtro de venda comercial não foi aplicado.")

    # Extração de Mês
    df_vendas["_mes"] = df_vendas[col_mes].apply(extrair_mes) if col_mes else None
    
    # Extração de Ano e Fallback: Se o Ano da Venda estiver zoado/zerado, tenta extrair de dentro da coluna Mês
    df_vendas["_ano"] = df_vendas[col_ano].apply(extrair_ano) if col_ano else None
    
    def aplicar_fallback_ano(row: pd.Series) -> Optional[int]:
        ano = row["_ano"]
        if pd.isna(ano) or ano < 2000:
            if col_mes:
                return extrair_ano(row[col_mes])
        return ano
        
    df_vendas["_ano"] = df_vendas.apply(aplicar_fallback_ano, axis=1)

    if col_valor:
        df_vendas["_vgv"] = df_vendas[col_valor].map(parse_valor_br)
    else:
        df_vendas["_vgv"] = 0.0
        st.warning("Coluna de Valor (VGV) não encontrada — o financeiro ficará zerado.")

    # Nova regra de criação do Canal Agrupado
    if col_canal:
        def agrupar_canal(c: Any) -> str:
            c_str = str(c).strip().upper()
            prefixo = c_str.split('-')[0].strip()
            # Se a coluna Canal for RJ ou RJG (ou começar com isso), o valor é IMOB.
            if prefixo in ['RJ', 'RJG'] or c_str in ['RJ', 'RJG']:
                return 'IMOB'
            return 'DV RJ'
        
        df_vendas['Canal_Agrupado'] = df_vendas[col_canal].apply(agrupar_canal)
    else:
        df_vendas['Canal_Agrupado'] = 'DV RJ'
        st.warning("Coluna **Canal** não encontrada na aba de vendas — assumindo tudo como DV RJ.")

    # -------------------------------------------------------------------------
    # Sidebar - Filtros
    # -------------------------------------------------------------------------
    st.sidebar.markdown("### Filtros")
    
    # Filtro principal do Canal
    filtro_canal = st.sidebar.selectbox(
        "Canal da Meta",
        ["RIO", "DIR", "PARC", "RJ"],
        index=0,
        help="RIO (100% da Meta, Todas as Vendas) | DIR (50% Meta, Vendas DV RJ) | PARC (25% Meta, Vendas RJG) | RJ (25% Meta, Vendas RJ)"
    )

    # Identifica apenas anos válidos (> 2000)
    anos_disponiveis: List[int] = sorted(
        int(x) for x in df_vendas["_ano"].dropna().unique().tolist() if pd.notna(x) and x > 2000
    )
    if not anos_disponiveis:
        st.warning("Nenhum ano numérico válido (> 2000) foi encontrado na base de vendas (ou todas as vendas comerciais = 1 foram filtradas para fora).")
        return

    ano_sel = st.sidebar.selectbox("Ano", anos_disponiveis, index=len(anos_disponiveis) - 1)

    # Identifica apenas meses válidos (1 a 12)
    meses_no_ano = sorted(
        int(x)
        for x in df_vendas.loc[df_vendas["_ano"] == ano_sel, "_mes"].dropna().unique().tolist()
        if pd.notna(x) and 1 <= int(x) <= 12
    )
    if not meses_no_ano:
        meses_no_ano = list(range(1, 13))

    idx_mes_padrao = max(0, len(meses_no_ano) - 1) if meses_no_ano else 0
    mes_sel = st.sidebar.selectbox("Mês", meses_no_ano, index=idx_mes_padrao)

    # Preparar listas para os novos filtros de Região e Empreendimento
    regioes_disponiveis = []
    if coluna_existe(df_metas, "Região"):
        regioes_disponiveis = sorted(set(str(x).strip() for x in df_metas["Região"].dropna().unique() if str(x).strip()))

    emps_comuns = []
    if coluna_existe(df_vendas, "Empreendimento") and coluna_existe(df_metas, "Empreendimento"):
        emps_vendas = set(str(x).strip() for x in df_vendas["Empreendimento"].dropna().unique() if str(x).strip())
        emps_metas = set(str(x).strip() for x in df_metas["Empreendimento"].dropna().unique() if str(x).strip())
        # Interseção: só empreendimentos que existem em ambas as abas
        emps_comuns = sorted(list(emps_vendas & emps_metas))

    filtro_regiao = st.sidebar.selectbox("Região (Metas)", ["Todas"] + regioes_disponiveis)
    filtro_emp = st.sidebar.selectbox("Empreendimento", ["Todos"] + emps_comuns)

    # -------------------------------------------------------------------------
    # Aplicação de Filtros
    # -------------------------------------------------------------------------
    vendas_f = df_vendas[(df_vendas["_ano"] == ano_sel) & (df_vendas["_mes"] == mes_sel)].copy()
    metas_f = df_metas[df_metas["Mes_Num"] == mes_sel].copy()

    # 1. Aplica lógica do canal escolhido na Venda
    fator_meta = 1.0
    if filtro_canal == "RIO":
        fator_meta = 1.0
        # vendas_f se mantém com todos os dados
    elif filtro_canal == "DIR":
        fator_meta = 0.50
        vendas_f = vendas_f[vendas_f["Canal_Agrupado"] == "DV RJ"]
    elif filtro_canal == "PARC":
        fator_meta = 0.25
        if col_canal:
            vendas_f = vendas_f[vendas_f[col_canal].astype(str).str.upper().str.strip().apply(
                lambda x: x.split('-')[0].strip() == 'RJG' or x == 'RJG'
            )]
    elif filtro_canal == "RJ":
        fator_meta = 0.25
        if col_canal:
            vendas_f = vendas_f[vendas_f[col_canal].astype(str).str.upper().str.strip().apply(
                lambda x: x.split('-')[0].strip() == 'RJ' or x == 'RJ'
            )]

    # 2. Aplica filtro de Região (baseado na aba de Metas)
    if filtro_regiao != "Todas" and coluna_existe(metas_f, "Região"):
        metas_f = metas_f[metas_f["Região"].astype(str).str.strip() == filtro_regiao]
        # Filtra a aba de vendas usando os empreendimentos vinculados a essa região nas metas
        emps_da_regiao = metas_f["Empreendimento"].astype(str).str.strip().unique()
        if coluna_existe(vendas_f, "Empreendimento"):
            vendas_f = vendas_f[vendas_f["Empreendimento"].astype(str).str.strip().isin(emps_da_regiao)]

    # 3. Aplica filtro de Empreendimento
    if filtro_emp != "Todos":
        if coluna_existe(vendas_f, "Empreendimento"):
            vendas_f = vendas_f[vendas_f["Empreendimento"].astype(str).str.strip() == filtro_emp]
        if coluna_existe(metas_f, "Empreendimento"):
            metas_f = metas_f[metas_f["Empreendimento"].astype(str).str.strip() == filtro_emp]

    # Ajusta as metas proporcionalmente ao canal escolhido (50%, 25%, etc)
    metas_f["Meta_Qtd"] = metas_f["Meta_Qtd"] * fator_meta

    total_realizado = int(vendas_f.shape[0])
    total_meta = float(metas_f["Meta_Qtd"].sum()) if not metas_f.empty else 0.0
    total_vgv = float(vendas_f["_vgv"].sum())

    pct_ating = (total_realizado / total_meta * 100.0) if total_meta > 0 else 0.0

    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Período</div><div class="val">{mes_sel:02d}/{ano_sel}</div></div>
            <div class="vel-kpi"><div class="lbl">Canal Filtro</div><div class="val">{filtro_canal}</div></div>
            <div class="vel-kpi"><div class="lbl">Vendas realizadas</div><div class="val">{total_realizado}</div></div>
            <div class="vel-kpi"><div class="lbl">Meta Ajustada (un.)</div><div class="val">{total_meta:g}</div></div>
            <div class="vel-kpi"><div class="lbl">VGV no mês</div><div class="val val--red">{fmt_br_milhoes(total_vgv)}</div></div>
            <div class="vel-kpi"><div class="lbl">% da meta</div><div class="val">{pct_ating:.1f}%</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Visão geral")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        criar_medidor("Geral — quantidade vs meta", float(total_realizado), total_meta, total_vgv, total_realizado)

    st.subheader("Por região")
    if "Região" in metas_f.columns and "Empreendimento" in metas_f.columns and "Empreendimento" in vendas_f.columns:
        regioes_m = sorted(str(x).strip() for x in metas_f["Região"].dropna().unique() if str(x).strip())

        if not regioes_m:
            st.info("Não há região definida nas metas deste mês.")
        else:
            cols = st.columns(min(3, len(regioes_m)) or 1)
            for i, regiao in enumerate(regioes_m):
                with cols[i % len(cols)]:
                    # 1. Filtra metas para esta região e soma a meta
                    metas_regiao = metas_f[metas_f["Região"].astype(str).str.strip() == regiao]
                    m_reg = float(metas_regiao["Meta_Qtd"].sum())

                    # 2. Identifica os empreendimentos vinculados a esta região na aba de metas
                    emps_da_regiao = metas_regiao["Empreendimento"].astype(str).str.strip().unique()

                    # 3. Conta as vendas cruzando apenas os empreendimentos correspondentes
                    v_reg = vendas_f[vendas_f["Empreendimento"].astype(str).str.strip().isin(emps_da_regiao)]

                    criar_medidor(
                        regiao,
                        float(v_reg.shape[0]),
                        m_reg,
                        float(v_reg["_vgv"].sum()),
                        int(v_reg.shape[0]),
                    )
    else:
        st.warning("As colunas **Região** e/ou **Empreendimento** não foram encontradas para cruzar as regiões.")

    st.caption(
        "Nota: O agrupamento por região utiliza automaticamente o mapeamento de **Empreendimentos** "
        "definido na aba de **Metas** para buscar as vendas correspondentes na base completa."
    )

    if coluna_existe(df_vendas, "Empreendimento") and coluna_existe(metas_f, "Empreendimento"):
        st.subheader("Por empreendimento")
        emp_v = vendas_f.assign(
            _emp=vendas_f["Empreendimento"].astype(str).str.strip()
        )
        vg = (
            emp_v.groupby("_emp", as_index=False)
            .agg(realizado=("Empreendimento", "count"), vgv=("_vgv", "sum"))
            .rename(columns={"_emp": "Empreendimento"})
        )
        emp_m = metas_f.assign(
            _emp=metas_f["Empreendimento"].astype(str).str.strip()
        )
        mg = (
            emp_m.groupby("_emp", as_index=False)["Meta_Qtd"]
            .sum()
            .rename(columns={"_emp": "Empreendimento", "Meta_Qtd": "Meta (un.)"})
        )
        tab = vg.merge(mg, on="Empreendimento", how="outer").fillna(0)
        tab["Meta (un.)"] = tab["Meta (un.)"].astype(float)
        tab["realizado"] = tab["realizado"].astype(float)
        tab["vgv"] = tab["vgv"].astype(float)
        tab["% meta"] = tab.apply(
            lambda r: (r["realizado"] / r["Meta (un.)"] * 100.0) if r["Meta (un.)"] > 0 else 0.0,
            axis=1,
        )
        tab = tab.sort_values(["Meta (un.)", "realizado"], ascending=False)
        show = tab.rename(columns={"vgv": "VGV (R$)", "realizado": "Realizado (un.)"})
        show["VGV (R$)"] = show["VGV (R$)"].map(lambda x: fmt_br_milhoes(float(x)))
        show["% meta"] = show["% meta"].map(lambda x: f"{x:.1f}%")
        st.dataframe(show, use_container_width=True, hide_index=True)
    else:
        st.info(
            "Para ver o cruzamento por empreendimento, as duas abas precisam da coluna **Empreendimento**."
        )

    with st.expander("Prévia dos dados (mês filtrado)"):
        st.dataframe(vendas_f.head(50), use_container_width=True)

    st.markdown(
        f'<div class="footer" style="text-align:center;padding:1rem 0;color:{COR_TEXTO_MUTED};font-size:0.82rem;">'
        f"Direcional Engenharia · Vendas — Acompanhamento de metas</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
