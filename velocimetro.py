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
    """Lê uma aba via gspread. `_cred_fp` deve ser o fingerprint da private_key."""
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
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s:
        return 0.0
    if "," in s and "." in s:
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
    """Extrai mês numérico com alta precisão."""
    if pd.isna(val): return None
    v = str(val).strip().lower()
    if not v: return None
    for m_str, m_num in MESES_TEXTO_MAP.items():
        if m_str in v:
            return m_num
    try:
        m = int(float(v))
        if 1 <= m <= 12: return m
    except ValueError:
        pass
    if '/' in v:
        p = v.split('/')
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
    """Extrai ano numérico de strings complexas."""
    if pd.isna(val): return None
    v = str(val).strip()
    if not v: return None
    try:
        ano = int(float(v))
        if ano > 2000: return ano
    except ValueError:
        pass
    if '/' in v:
        p = v.split('/')
        try:
            ano = int(p[-1])
            if ano > 2000: return ano
        except: pass
    return None


def achar_coluna(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    """Busca uma coluna no DataFrame por uma lista de aliases."""
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
    """Transforma Metas QTD e VGV largas em longas (mês numérico + meta_qtd + meta_vgv)."""
    df = normalizar_colunas(df_metas_raw)
    
    # Identifica as colunas padrão
    id_vars = [c for c in df.columns if c.lower() in ["empreendimento", "região", "regiao", "obra"]]
    
    # Identifica colunas QTD1 a QTD12
    cols_qtd = [c for c in df.columns if re.match(r'^qtd(1[0-2]|[1-9])$', c.lower().strip())]
    # Identifica colunas VGV1 a VGV12
    cols_vgv = [c for c in df.columns if re.match(r'^vgv(1[0-2]|[1-9])$', c.lower().strip())]

    if not id_vars or not cols_qtd:
        return pd.DataFrame(columns=["Empreendimento", "Região", "Mes_Num", "Meta_Qtd", "Meta_VGV"])

    # Melt de QTD
    df_qtd = df.melt(id_vars=id_vars, value_vars=cols_qtd, var_name="Mes_Str", value_name="Meta_Qtd")
    df_qtd["Mes_Num"] = df_qtd["Mes_Str"].str.upper().str.replace("QTD", "").astype(int)
    df_qtd.drop(columns=["Mes_Str"], inplace=True)
    df_qtd["Meta_Qtd"] = pd.to_numeric(df_qtd["Meta_Qtd"], errors="coerce").fillna(0)

    # Melt de VGV (se existir)
    if cols_vgv:
        df_vgv = df.melt(id_vars=id_vars, value_vars=cols_vgv, var_name="Mes_Str", value_name="Meta_VGV")
        df_vgv["Mes_Num"] = df_vgv["Mes_Str"].str.upper().str.replace("VGV", "").astype(int)
        df_vgv.drop(columns=["Mes_Str"], inplace=True)
        # Usa parse_valor_br caso o usuário tenha formatado como R$ 2.000.000,00
        df_vgv["Meta_VGV"] = df_vgv["Meta_VGV"].apply(parse_valor_br)
        
        # Junta QTD e VGV
        out = pd.merge(df_qtd, df_vgv, on=id_vars + ["Mes_Num"], how="outer").fillna(0)
    else:
        out = df_qtd.copy()
        out["Meta_VGV"] = 0.0

    # Normaliza colunas ID
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
        }}
        .block-container {{
            max-width: 1300px !important;
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
            flex: 1 1 16%;
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


def criar_medidor(titulo: str, realizado: float, meta: float, vgv: float, meta_vgv: float, vendas_qtd: int) -> None:
    """Gauge estilo velocímetro com cores da marca."""
    meta_f = float(meta) if meta and meta > 0 else 0.0
    perc = min(150.0, (realizado / meta_f * 100.0)) if meta_f > 0 else 0.0
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
        <div style="text-align:center;font-size:0.85rem;color:{COR_TEXTO_LABEL};margin-top:-8px;line-height:1.4;">
            <strong>Qtd:</strong> {int(vendas_qtd)} / {meta_f:g} <br/>
            <strong>VGV:</strong> {fmt_br_milhoes(float(vgv))} / {fmt_br_milhoes(float(meta_vgv))}
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(
        page_title="Acompanhamento de Vendas | Direcional",
        layout="wide",
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
            "Preencha pelo menos **private_key** e **client_email**."
        )
        return

    sid = spreadsheet_id_de_secrets(raw_gs)
    cred_fp = _fingerprint_credenciais(info)

    try:
        df_vendas = ler_planilha_aba_df(sid, WS_VENDAS, cred_fp)
        df_metas_raw = ler_planilha_aba_df(sid, WS_METAS, cred_fp)
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {str(e)}")
        return

    df_vendas = normalizar_colunas(df_vendas)
    df_metas = melt_metas(df_metas_raw)

    # Ignora os empreendimentos que nao possuem regiao na coluna metas
    if "Região" in df_metas.columns:
        df_metas = df_metas[df_metas["Região"].astype(str).str.strip().replace('nan', '') != ""]
        df_metas = df_metas[~df_metas["Região"].astype(str).str.strip().str.lower().isin(["total", "geral"])]
    if "Empreendimento" in df_metas.columns:
        df_metas = df_metas[~df_metas["Empreendimento"].astype(str).str.strip().str.lower().isin(["total", "geral", ""])]

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
    col_proprietario = achar_coluna(df_vendas, ["Proprietário da oportunidade", "Proprietario da oportunidade", "Nome da conta", "Proprietario", "Corretor"])
    col_ranking = achar_coluna(df_vendas, ["Ranking"])

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

    # Filtro Obrigatório: Apenas Venda Comercial == 1
    if col_venda_comercial:
        df_vendas = df_vendas[pd.to_numeric(df_vendas[col_venda_comercial], errors='coerce') == 1]
    else:
        st.warning("Coluna 'Venda Comercial?' não encontrada na base.")

    df_vendas["_mes"] = df_vendas[col_mes].apply(extrair_mes) if col_mes else None
    df_vendas["_ano"] = df_vendas[col_ano].apply(extrair_ano) if col_ano else None
    
    def aplicar_fallback_ano(row: pd.Series) -> Optional[int]:
        ano = row["_ano"]
        if pd.isna(ano) or ano < 2000:
            if col_mes:
                return extrair_ano(row[col_mes])
        return ano
        
    df_vendas["_ano"] = df_vendas.apply(aplicar_fallback_ano, axis=1)
    df_vendas["_vgv"] = df_vendas[col_valor].map(parse_valor_br) if col_valor else 0.0

    if col_canal:
        def agrupar_canal(c: Any) -> str:
            c_str = str(c).strip().upper()
            prefixo = c_str.split('-')[0].strip()
            if prefixo in ['RJ', 'RJG'] or c_str in ['RJ', 'RJG']:
                return 'IMOB'
            return 'DV RJ'
        df_vendas['Canal_Agrupado'] = df_vendas[col_canal].apply(agrupar_canal)
    else:
        df_vendas['Canal_Agrupado'] = 'DV RJ'

    # Mapear a Região oficial da Venda com base na tabela de Metas
    map_emp_regiao = df_metas.drop_duplicates("Empreendimento").set_index("Empreendimento")["Região"].to_dict()
    df_vendas["Regiao_Meta"] = df_vendas["Empreendimento"].map(map_emp_regiao)

    # -------------------------------------------------------------------------
    # LINHA ÚNICA DE FILTROS (Múltipla Seleção)
    # -------------------------------------------------------------------------
    anos_disponiveis = sorted(int(x) for x in df_vendas["_ano"].dropna().unique().tolist() if pd.notna(x) and x > 2000)
    meses_no_ano = sorted(int(x) for x in df_vendas["_mes"].dropna().unique().tolist() if pd.notna(x) and 1 <= int(x) <= 12)
    regioes_disponiveis = sorted(set(str(x).strip() for x in df_metas["Região"].dropna().unique() if str(x).strip()))
    
    emps_comuns = []
    if coluna_existe(df_vendas, "Empreendimento") and coluna_existe(df_metas, "Empreendimento"):
        emps_vendas = set(str(x).strip() for x in df_vendas["Empreendimento"].dropna().unique() if str(x).strip())
        emps_metas = set(str(x).strip() for x in df_metas["Empreendimento"].dropna().unique() if str(x).strip())
        emps_comuns = sorted(list(emps_vendas & emps_metas))

    st.markdown("<div style='margin-bottom:1rem;'><strong>Filtros</strong></div>", unsafe_allow_html=True)
    
    col_filtros = st.columns(5)
    with col_filtros[0]:
        canais_sel = st.multiselect("Canal da Meta", ["RIO", "DIR", "PARC", "RJ"], default=["RIO"])
    with col_filtros[1]:
        anos_sel = st.multiselect("Ano", anos_disponiveis, default=[anos_disponiveis[-1]] if anos_disponiveis else [])
    with col_filtros[2]:
        mes_padrao = max(0, len(meses_no_ano) - 1) if meses_no_ano else []
        meses_sel = st.multiselect("Mês", meses_no_ano, default=[meses_no_ano[mes_padrao]] if meses_no_ano else [])
    with col_filtros[3]:
        regioes_sel = st.multiselect("Região", regioes_disponiveis, default=[])
    with col_filtros[4]:
        emps_sel = st.multiselect("Empreendimento", emps_comuns, default=[])

    # -------------------------------------------------------------------------
    # Aplicação de Filtros (Listas)
    # -------------------------------------------------------------------------
    vendas_f = df_vendas.copy()
    metas_f = df_metas.copy()

    if anos_sel:
        vendas_f = vendas_f[vendas_f["_ano"].isin(anos_sel)]
    if meses_sel:
        vendas_f = vendas_f[vendas_f["_mes"].isin(meses_sel)]
        metas_f = metas_f[metas_f["Mes_Num"].isin(meses_sel)]

    if regioes_sel:
        metas_f = metas_f[metas_f["Região"].isin(regioes_sel)]
        vendas_f = vendas_f[vendas_f["Regiao_Meta"].isin(regioes_sel)]

    if emps_sel:
        metas_f = metas_f[metas_f["Empreendimento"].isin(emps_sel)]
        vendas_f = vendas_f[vendas_f["Empreendimento"].isin(emps_sel)]

    # Lógica de Canal
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

    fator_meta = min(1.0, fator_meta)
    vendas_f = vendas_f[mask_vendas]

    metas_f["Meta_Qtd"] = metas_f["Meta_Qtd"] * fator_meta
    metas_f["Meta_VGV"] = metas_f["Meta_VGV"] * fator_meta

    total_realizado_qtd = int(vendas_f.shape[0])
    total_meta_qtd = float(metas_f["Meta_Qtd"].sum()) if not metas_f.empty else 0.0
    total_vgv_realizado = float(vendas_f["_vgv"].sum())
    total_meta_vgv = float(metas_f["Meta_VGV"].sum()) if not metas_f.empty else 0.0

    pct_qtd = (total_realizado_qtd / total_meta_qtd * 100.0) if total_meta_qtd > 0 else 0.0
    pct_vgv = (total_vgv_realizado / total_meta_vgv * 100.0) if total_meta_vgv > 0 else 0.0

    st.markdown(
        f"""
        <div class="vel-kpi-row" style="margin-top: 1rem;">
            <div class="vel-kpi"><div class="lbl">Qtd Meta</div><div class="val">{total_meta_qtd:g}</div></div>
            <div class="vel-kpi"><div class="lbl">Qtd Realizado</div><div class="val">{total_realizado_qtd}</div></div>
            <div class="vel-kpi"><div class="lbl">% Qtd</div><div class="val">{pct_qtd:.1f}%</div></div>
            <div class="vel-kpi"><div class="lbl">VGV Meta</div><div class="val">{fmt_br_milhoes(total_meta_vgv)}</div></div>
            <div class="vel-kpi"><div class="lbl">VGV Realizado</div><div class="val val--red">{fmt_br_milhoes(total_vgv_realizado)}</div></div>
            <div class="vel-kpi"><div class="lbl">% VGV</div><div class="val">{pct_vgv:.1f}%</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Visão geral")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        criar_medidor("Geral — quantidade vs meta", float(total_realizado_qtd), total_meta_qtd, total_vgv_realizado, total_meta_vgv, total_realizado_qtd)

    st.subheader("Por região")
    if "Região" in metas_f.columns and "Empreendimento" in metas_f.columns:
        regioes_m = sorted(str(x).strip() for x in metas_f["Região"].dropna().unique() if str(x).strip())

        if not regioes_m:
            st.info("Não há região definida nas metas do período filtrado.")
        else:
            cols = st.columns(min(3, len(regioes_m)) or 1)
            for i, regiao in enumerate(regioes_m):
                with cols[i % len(cols)]:
                    metas_regiao = metas_f[metas_f["Região"].astype(str).str.strip() == regiao]
                    m_reg_qtd = float(metas_regiao["Meta_Qtd"].sum())
                    m_reg_vgv = float(metas_regiao["Meta_VGV"].sum())

                    emps_da_regiao = metas_regiao["Empreendimento"].astype(str).str.strip().unique()
                    v_reg = vendas_f[vendas_f["Empreendimento"].astype(str).str.strip().isin(emps_da_regiao)]

                    criar_medidor(
                        regiao,
                        float(v_reg.shape[0]),
                        m_reg_qtd,
                        float(v_reg["_vgv"].sum()),
                        m_reg_vgv,
                        int(v_reg.shape[0]),
                    )
    else:
        st.warning("As colunas **Região** e/ou **Empreendimento** não foram encontradas para cruzar as regiões.")

    # -------------------------------------------------------------------------
    # TABELA POR REGIÃO
    # -------------------------------------------------------------------------
    st.subheader("Tabela Resumo: Por Região")
    if "Região" in metas_f.columns:
        vg_reg = (
            vendas_f.groupby("Regiao_Meta", as_index=False)
            .agg(real_qtd=("Empreendimento", "count"), real_vgv=("_vgv", "sum"))
            .rename(columns={"Regiao_Meta": "Região"})
        )
        mg_reg = (
            metas_f.groupby("Região", as_index=False)
            .agg(meta_qtd=("Meta_Qtd", "sum"), meta_vgv=("Meta_VGV", "sum"))
        )
        tab_reg = vg_reg.merge(mg_reg, on="Região", how="outer").fillna(0)
        
        tab_reg["% Qtd"] = tab_reg.apply(lambda r: (r["real_qtd"] / r["meta_qtd"] * 100.0) if r["meta_qtd"] > 0 else 0.0, axis=1)
        tab_reg["% VGV"] = tab_reg.apply(lambda r: (r["real_vgv"] / r["meta_vgv"] * 100.0) if r["meta_vgv"] > 0 else 0.0, axis=1)
        
        tab_reg = tab_reg.sort_values(["meta_qtd", "real_qtd"], ascending=False)
        show_reg = tab_reg.rename(columns={
            "meta_qtd": "Meta Qtd (un.)", "real_qtd": "Realizado Qtd (un.)",
            "meta_vgv": "Meta VGV (R$)", "real_vgv": "Realizado VGV (R$)"
        })
        
        show_reg["Realizado VGV (R$)"] = show_reg["Realizado VGV (R$)"].map(lambda x: fmt_br_milhoes(float(x)))
        show_reg["Meta VGV (R$)"] = show_reg["Meta VGV (R$)"].map(lambda x: fmt_br_milhoes(float(x)))
        show_reg["% Qtd"] = show_reg["% Qtd"].map(lambda x: f"{x:.1f}%")
        show_reg["% VGV"] = show_reg["% VGV"].map(lambda x: f"{x:.1f}%")
        
        st.dataframe(show_reg, use_container_width=True, hide_index=True)

    # -------------------------------------------------------------------------
    # TABELA POR EMPREENDIMENTO
    # -------------------------------------------------------------------------
    st.subheader("Tabela Resumo: Por Empreendimento")
    if coluna_existe(df_vendas, "Empreendimento") and coluna_existe(metas_f, "Empreendimento"):
        emp_v = vendas_f.assign(_emp=vendas_f["Empreendimento"].astype(str).str.strip())
        vg_emp = (
            emp_v.groupby("_emp", as_index=False)
            .agg(real_qtd=("Empreendimento", "count"), real_vgv=("_vgv", "sum"))
            .rename(columns={"_emp": "Empreendimento"})
        )
        emp_m = metas_f.assign(_emp=metas_f["Empreendimento"].astype(str).str.strip())
        mg_emp = (
            emp_m.groupby("_emp", as_index=False)
            .agg(meta_qtd=("Meta_Qtd", "sum"), meta_vgv=("Meta_VGV", "sum"))
            .rename(columns={"_emp": "Empreendimento"})
        )
        tab_emp = vg_emp.merge(mg_emp, on="Empreendimento", how="outer").fillna(0)
        
        tab_emp["% Qtd"] = tab_emp.apply(lambda r: (r["real_qtd"] / r["meta_qtd"] * 100.0) if r["meta_qtd"] > 0 else 0.0, axis=1)
        tab_emp["% VGV"] = tab_emp.apply(lambda r: (r["real_vgv"] / r["meta_vgv"] * 100.0) if r["meta_vgv"] > 0 else 0.0, axis=1)
        
        tab_emp = tab_emp.sort_values(["meta_qtd", "real_qtd"], ascending=False)
        show_emp = tab_emp.rename(columns={
            "meta_qtd": "Meta Qtd (un.)", "real_qtd": "Realizado Qtd (un.)",
            "meta_vgv": "Meta VGV (R$)", "real_vgv": "Realizado VGV (R$)"
        })
        
        show_emp["Realizado VGV (R$)"] = show_emp["Realizado VGV (R$)"].map(lambda x: fmt_br_milhoes(float(x)))
        show_emp["Meta VGV (R$)"] = show_emp["Meta VGV (R$)"].map(lambda x: fmt_br_milhoes(float(x)))
        show_emp["% Qtd"] = show_emp["% Qtd"].map(lambda x: f"{x:.1f}%")
        show_emp["% VGV"] = show_emp["% VGV"].map(lambda x: f"{x:.1f}%")
        
        st.dataframe(show_emp, use_container_width=True, hide_index=True)
    else:
        st.info("Para ver o cruzamento por empreendimento, as duas abas precisam da coluna **Empreendimento**.")

    # -------------------------------------------------------------------------
    # TABELAS DE VENDAS (IMOB E REGIONAL)
    # -------------------------------------------------------------------------
    if col_proprietario:
        c_imob, c_reg = st.columns(2)
        
        with c_imob:
            st.subheader("Vendas por IMOB")
            df_imob = vendas_f[vendas_f["Canal_Agrupado"] == "IMOB"].copy()
            if not df_imob.empty:
                df_imob[col_proprietario] = df_imob[col_proprietario].fillna("Não Informado").astype(str)
                tab_imob = df_imob.groupby(col_proprietario, as_index=False).agg(
                    QTD=(col_proprietario, "count"),
                    VGV=("_vgv", "sum")
                ).sort_values("VGV", ascending=False)
                tab_imob["VGV"] = tab_imob["VGV"].apply(lambda x: fmt_br_milhoes(float(x)))
                tab_imob.rename(columns={col_proprietario: "Proprietário"}, inplace=True)
                st.dataframe(tab_imob, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhuma venda IMOB encontrada com os filtros atuais.")
                
        with c_reg:
            st.subheader("Vendas por Regional")
            df_regional = vendas_f[vendas_f["Canal_Agrupado"] == "DV RJ"].copy()
            if not df_regional.empty:
                df_regional[col_proprietario] = df_regional[col_proprietario].fillna("Não Informado").astype(str)
                tab_reg = df_regional.groupby(col_proprietario, as_index=False).agg(
                    QTD=(col_proprietario, "count"),
                    VGV=("_vgv", "sum")
                ).sort_values("VGV", ascending=False)
                tab_reg["VGV"] = tab_reg["VGV"].apply(lambda x: fmt_br_milhoes(float(x)))
                tab_reg.rename(columns={col_proprietario: "Proprietário"}, inplace=True)
                st.dataframe(tab_reg, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhuma venda Regional encontrada com os filtros atuais.")
    else:
        st.warning("Coluna 'Proprietário da oportunidade' não encontrada na base.")

    # -------------------------------------------------------------------------
    # TABELA DE VENDAS POR RANKING
    # -------------------------------------------------------------------------
    st.subheader("Vendas por Ranking")
    if col_ranking:
        if not vendas_f.empty:
            df_rank = vendas_f.copy()
            df_rank[col_ranking] = df_rank[col_ranking].fillna("Não Informado").astype(str)
            tab_rank = df_rank.groupby(col_ranking, as_index=False).agg(
                QTD=(col_ranking, "count"),
                VGV=("_vgv", "sum")
            ).sort_values("VGV", ascending=False)
            tab_rank["VGV"] = tab_rank["VGV"].apply(lambda x: fmt_br_milhoes(float(x)))
            tab_rank.rename(columns={col_ranking: "Ranking"}, inplace=True)
            st.dataframe(tab_rank, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma venda encontrada com os filtros atuais.")
    else:
        st.warning("Coluna 'Ranking' não encontrada na base.")

    st.markdown(
        f'<div class="footer" style="text-align:center;padding:1rem 0;color:{COR_TEXTO_MUTED};font-size:0.82rem;">'
        f"Direcional Engenharia · Vendas — Acompanhamento de metas</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
