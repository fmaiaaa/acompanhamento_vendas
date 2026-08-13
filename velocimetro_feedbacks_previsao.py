# Feedbacks comerciais + Previsão de vendas (formulários Google Sheets)
from __future__ import annotations

from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

SPREADSHEET_FEEDBACK_ID = "1wyGj3K4j7_v9YFad6A8JuRnV17B6JoYN7_FDHOfmYgk"
WS_FEEDBACK = "Respostas ao formulário 1"
SPREADSHEET_PREVISAO_ID = "1lBliB3AjR5vJyRy9SoDi6DQOA9x5LC5wYfNNo5cz0bE"
WS_PREVISAO = "Respostas ao formulário 1"

TEMAS_COMENTARIO: Dict[str, List[str]] = {
    "Agilidade / rapidez": [
        "rápid", "rapido", "rapidez", "ágil", "agil", "minutos", "imediata", "eficaz", "eficiente",
    ],
    "Elogio coordenador / comercial": [
        "nota 10", "parabéns", "parabens", "excelente", "melhor que", "proativo", "melhor que nós",
    ],
    "Suporte e parceria": ["suporte", "parceria", "ajudou", "solicito", "presente", "comprometido"],
    "Demora / espera": ["demor", "espera", "longa", "24 hs", "24h", "horas após"],
    "Sistema / Bora / cotação": ["bora", "cotação", "cotacao", "instabil", "sistema", "gerar", "editar"],
    "EmCash": ["emcash", "em cash"],
    "Contrato / processo": ["contrato", "minuta", "assin", "comunicada", "boas vindas"],
    "Unidade / estoque": ["mirror", "derrubada", "unidade", "estoque", "sugerida"],
    "Repasse / comissão / nota": ["repasse", "comissão", "comissao", "nota fiscal", "gerar nota"],
    "QR Code": ["qr code", "qr"],
    "Reclamação": ["insatisf", "chatead", "absurdo", "horrível", "horrivel", "reclama"],
    "Sugestão de melhoria": ["sugest", "poderia", "deixar porcesso", "deixar processo"],
}

ALIASES_COORD_FEEDBACK = [
    "Selecione o Coordenador Comercial responsável pelo processo da venda",
    "Coordenador Comercial",
    "Coordenador",
]
ALIASES_OBS_FEEDBACK = [
    "Deseja faazer alguma observação... Pode ser um elogio",
    "Deseja fazer alguma observação",
    "observação",
    "Observação",
]
ALIASES_NOTA_ATEND = [
    "De 1 a 10, avalie a Qualidade do Atendimento, sendo 1 caso esteja Muito Insatisfeito e 10 caso esteja Muito Satisfeito",
    "Qualidade do Atendimento",
]
ALIASES_NOTA_CCA = [
    "De 1 a 10, avalie de forma geral o CCA, sendo 1 caso esteja Muito Insatisfeito e 10 caso esteja Muito Satisfeito",
    "Nota CCA",
]
ALIASES_PROMOTOR_COM = ["Promotor Comercial", "Promotor Comercial "]
ALIASES_PROMOTOR_CCA = ["Promotor CCA", "Promotor CCA "]

ALIASES_PREV_SABADO = [
    "Data de referência (indicar sempre o sábado da semana)",
    "Data de referência",
]
ALIASES_PREV_VENDAS = ["Vendas Previsão", "Vendas Previsao", "QTD Vendas Normais Previstas", "QTD Vendas Facilitadas Previstas"]
ALIASES_PREV_VGV = ["VGV Previsão", "VGV Previsao"]
ALIASES_PREV_CANAL = ["Canal"]
ALIASES_PREV_EMP = ['Empreendimento (Previsto para ter venda)', "Empreendimento"]
ALIASES_PREV_IMOB = ["Imobiliária", "Imobiliaria"]
ALIASES_PREV_REG = ["Regional ou IMOB", "Regional"]
ALIASES_PREV_TRIM = ["Trimestre"]
ALIASES_PREV_SEM = ["Semestre"]
ALIASES_PREV_ANO = ["Ano"]


def _v():
    import sys
    return sys.modules.get("velocimetro") or sys.modules.get("__main__")


def _achar(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    v = _v()
    if v and hasattr(v, "achar_coluna"):
        return v.achar_coluna(df, aliases)
    lows = {str(c).strip().lower(): c for c in df.columns}
    for a in aliases:
        al = a.strip().lower()
        for k, orig in lows.items():
            if k == al or al in k:
                return orig
    return None


def _parse_data(val: Any) -> Optional[date]:
    v = _v()
    if v:
        s = v.parse_data_serie(pd.Series([val]))
        if len(s) and pd.notna(s.iloc[0]):
            d = s.iloc[0]
            return d.date() if hasattr(d, "date") else d
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S"):
        try:
            return pd.to_datetime(s, format=fmt).date()
        except Exception:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None


def _parse_num(val: Any) -> float:
    v = _v()
    if v and hasattr(v, "parse_valor_br"):
        try:
            return float(v.parse_valor_br(val))
        except Exception:
            pass
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    s = str(val).strip().replace("R$", "").replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def classificar_temas_comentario(texto: str) -> List[str]:
    if not texto or str(texto).strip().lower() in ("nan", "none", "", "não", "nao", "."):
        return []
    t = str(texto).lower()
    temas = []
    for tema, kws in TEMAS_COMENTARIO.items():
        if any(k in t for k in kws):
            temas.append(tema)
    if not temas and len(t.strip()) > 8:
        temas.append("Outros / geral")
    return temas


def _nps_de_coluna(serie: pd.Series) -> Optional[float]:
    s = serie.astype(str).str.strip().str.lower()
    s = s[s.isin(["promotor", "detrator", "neutro"])]
    if s.empty:
        return None
    n = len(s)
    prom = (s == "promotor").sum()
    det = (s == "detrator").sum()
    return (prom / n - det / n) * 100.0


@st.cache_data(ttl=600, show_spinner="Carregando feedbacks comerciais…")
def carregar_feedbacks_comerciais(cred_fp: str) -> pd.DataFrame:
    v = _v()
    df = v.ler_planilha_aba_df(SPREADSHEET_FEEDBACK_ID, WS_FEEDBACK, cred_fp)
    return v.normalizar_colunas(df) if hasattr(v, "normalizar_colunas") else df


@st.cache_data(ttl=600, show_spinner="Carregando previsões de vendas…")
def carregar_previsao_vendas(cred_fp: str) -> pd.DataFrame:
    v = _v()
    df = v.ler_planilha_aba_df(SPREADSHEET_PREVISAO_ID, WS_PREVISAO, cred_fp)
    return v.normalizar_colunas(df) if hasattr(v, "normalizar_colunas") else df


def sabado_referencia_de_data(dt: date) -> Optional[date]:
    wd = dt.weekday()
    if wd == 4:
        return dt + timedelta(days=1)
    if wd == 5:
        return dt
    if wd == 6:
        return dt - timedelta(days=1)
    return None


def grupo_canal_previsao(canal: Any) -> str:
    s = str(canal or "").strip().upper()
    if "IMOB" in s:
        return "IMOB"
    if "DV" in s or "DIR" in s or "RIV" in s:
        return "DV"
    return "OUTROS"


def grupo_canal_venda(imob: Any) -> str:
    v = _v()
    if v and hasattr(v, "canal_de_imobiliaria"):
        p = v.canal_de_imobiliaria(imob)
    else:
        p = str(imob or "")[:3].upper()
    if p in ("RJ", "RJG"):
        return "IMOB"
    if p in ("DIR", "RIV"):
        return "DV"
    return "OUTROS"


def preparar_previsao_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    col_sab = _achar(out, ALIASES_PREV_SABADO)
    col_v = _achar(out, ALIASES_PREV_VENDAS)
    col_vgv = _achar(out, ALIASES_PREV_VGV)
    col_canal = _achar(out, ALIASES_PREV_CANAL)
    col_emp = _achar(out, ALIASES_PREV_EMP)
    col_imob = _achar(out, ALIASES_PREV_IMOB)
    col_reg = _achar(out, ALIASES_PREV_REG)
    col_tri = _achar(out, ALIASES_PREV_TRIM)
    col_sem = _achar(out, ALIASES_PREV_SEM)
    col_ano = _achar(out, ALIASES_PREV_ANO)

    out["_sabado"] = out[col_sab].map(_parse_data) if col_sab else pd.NaT
    if col_v:
        out["_prev_qtd"] = out[col_v].map(_parse_num)
    else:
        q_cols = [c for c in out.columns if "previst" in str(c).lower() and "qtd" in str(c).lower()]
        out["_prev_qtd"] = out[q_cols].apply(lambda r: sum(_parse_num(x) for x in r), axis=1) if q_cols else 0.0
    out["_prev_vgv"] = out[col_vgv].map(_parse_num) if col_vgv else 0.0
    out["_grupo_canal"] = out[col_canal].map(grupo_canal_previsao) if col_canal else "OUTROS"
    out["_emp"] = out[col_emp].astype(str).str.strip() if col_emp else ""
    out["_imob"] = out[col_imob].astype(str).str.strip() if col_imob else ""
    out["_regional"] = out[col_reg].astype(str).str.strip() if col_reg else ""
    out["_trimestre"] = out[col_tri].astype(str).str.strip() if col_tri else ""
    out["_semestre"] = out[col_sem].astype(str).str.strip() if col_sem else ""
    out["_ano"] = out[col_ano].astype(str).str.strip() if col_ano else ""
    return out


def vendas_reais_por_fim_de_semana(df_vendas: pd.DataFrame, col_contrato: str) -> pd.DataFrame:
    if df_vendas is None or df_vendas.empty or not col_contrato or col_contrato not in df_vendas.columns:
        return pd.DataFrame(columns=["_sabado", "_grupo_canal", "_real_qtd", "_real_vgv"])
    v = _v()
    ven = df_vendas.copy()
    ven["_dt"] = v.parse_data_serie(ven[col_contrato]) if v else pd.to_datetime(ven[col_contrato], errors="coerce")
    ven = ven.dropna(subset=["_dt"])
    ven["_date"] = ven["_dt"].dt.date
    ven["_sabado"] = ven["_date"].map(sabado_referencia_de_data)
    ven = ven.dropna(subset=["_sabado"])
    col_imob = "Imobiliária" if "Imobiliária" in ven.columns else None
    ven["_grupo_canal"] = ven[col_imob].map(grupo_canal_venda) if col_imob else "OUTROS"
    col_vgv = v.achar_coluna(ven, ["Valor Real de Venda", "Valor Real", "_vgv"]) if v else None
    ven["_vgv_n"] = ven[col_vgv].map(_parse_num) if col_vgv else (ven["_vgv"].map(_parse_num) if "_vgv" in ven.columns else 0.0)
    return ven.groupby(["_sabado", "_grupo_canal"], as_index=False).agg(
        _real_qtd=("_date", "count"), _real_vgv=("_vgv_n", "sum"),
    )


def cruzar_previsao_realizado(df_prev: pd.DataFrame, df_vendas: pd.DataFrame, col_contrato: str) -> pd.DataFrame:
    prev = preparar_previsao_df(df_prev)
    if prev.empty:
        return pd.DataFrame()
    prev = prev.dropna(subset=["_sabado"])
    agg_prev = prev.groupby(["_sabado", "_grupo_canal"], as_index=False).agg(
        prev_qtd=("_prev_qtd", "sum"), prev_vgv=("_prev_vgv", "sum"),
        n_respostas=("_emp", "count"),
        n_imobs=("_imob", lambda s: len({x for x in s if str(x).strip() and str(x).lower() not in ("nan", "")})),
        n_regionais=("_regional", lambda s: len({x for x in s if "regional" in str(x).lower()})),
    )
    real = vendas_reais_por_fim_de_semana(df_vendas, col_contrato)
    if real.empty:
        merged = agg_prev.copy()
        merged["real_qtd"] = 0.0
        merged["real_vgv"] = 0.0
    else:
        merged = agg_prev.merge(real, on=["_sabado", "_grupo_canal"], how="left")
        merged["real_qtd"] = merged["_real_qtd"].fillna(0.0)
        merged["real_vgv"] = merged["_real_vgv"].fillna(0.0)
        merged = merged.drop(columns=[c for c in ("_real_qtd", "_real_vgv") if c in merged.columns])
    merged["erro_qtd"] = merged["real_qtd"] - merged["prev_qtd"]
    return merged[merged["n_respostas"] > 0].sort_values("_sabado")


def metricas_erro_previsao(df: pd.DataFrame, canal: str) -> Dict[str, Any]:
    sub = df[df["_grupo_canal"] == canal]
    if sub.empty:
        return {}
    err = sub["erro_qtd"].astype(float)
    out = {"n": len(sub), "erro_medio": float(err.mean()), "desvio_padrao": float(err.std()) if len(sub) > 1 else 0.0, "erro_maximo": float(err.abs().max())}
    for tol in (1, 2, 3, 4, 5):
        out[f"acertos_pm_{tol}"] = int((err.abs() <= tol).sum())
    return out


def render_grafico_previsto_realizado_linha(df: pd.DataFrame) -> None:
    if df.empty:
        st.info("Sem dados cruzados de previsão × realizado.")
        return
    fig = go.Figure()
    for canal, cor in (("IMOB", "#2563eb"), ("DV", "#dc2626")):
        sub = df[df["_grupo_canal"] == canal].sort_values("_sabado")
        if sub.empty:
            continue
        fig.add_trace(go.Scatter(x=sub["_sabado"], y=sub["prev_qtd"], name=f"Previsto {canal}", mode="lines+markers", line=dict(color=cor, dash="dash")))
        fig.add_trace(go.Scatter(x=sub["_sabado"], y=sub["real_qtd"], name=f"Realizado {canal}", mode="lines+markers", line=dict(color=cor)))
    fig.update_layout(height=420, xaxis_title="Sábado de referência", yaxis_title="Vendas (qtd)")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_scatter_previsto_realizado(df: pd.DataFrame) -> None:
    if df.empty:
        return
    fig = go.Figure()
    cores = {"IMOB": "#2563eb", "DV": "#dc2626"}
    max_val = max(df["prev_qtd"].max(), df["real_qtd"].max(), 1.0)
    fig.add_trace(go.Scatter(x=[0, max_val * 1.1], y=[0, max_val * 1.1], mode="lines", name="Perfeição (y=x)", line=dict(color="#94a3b8", dash="dot")))
    for canal in ("IMOB", "DV"):
        sub = df[df["_grupo_canal"] == canal]
        if sub.empty:
            continue
        fig.add_trace(go.Scatter(x=sub["prev_qtd"], y=sub["real_qtd"], mode="markers", name=canal, marker=dict(size=10, color=cores[canal]), text=sub["_sabado"].astype(str)))
    fig.update_layout(height=440, xaxis_title="Vendas previstas", yaxis_title="Vendas realizadas (SF)", yaxis=dict(scaleanchor="x", scaleratio=1))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_tabela_erros(df: pd.DataFrame) -> None:
    rows = []
    for canal in ("IMOB", "DV"):
        m = metricas_erro_previsao(df, canal)
        if m:
            rows.append({"Canal": canal, "Semanas": m["n"], "Erro médio": round(m["erro_medio"], 2), "Desvio padrão": round(m["desvio_padrao"], 2), "Erro máximo (abs)": round(m["erro_maximo"], 2), "±1": m.get("acertos_pm_1", 0), "±2": m.get("acertos_pm_2", 0), "±3": m.get("acertos_pm_3", 0), "±4": m.get("acertos_pm_4", 0), "±5": m.get("acertos_pm_5", 0)})
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def _stats_respondentes(serie: pd.Series) -> Dict[str, float]:
    vals = pd.to_numeric(serie, errors="coerce").dropna()
    vals = vals[vals > 0]
    if vals.empty:
        return {"media": 0.0, "max": 0.0, "min": 0.0}
    return {"media": float(vals.mean()), "max": float(vals.max()), "min": float(vals.min())}


def render_resumo_respondentes(df: pd.DataFrame) -> None:
    if df.empty:
        return
    c1, c2 = st.columns(2)
    imob, reg = _stats_respondentes(df["n_imobs"]), _stats_respondentes(df["n_regionais"])
    with c1:
        st.markdown("**Imobiliárias respondentes / fim de semana**")
        st.metric("Média", f"{imob['media']:.1f}")
        st.caption(f"Mín (sem zero): {imob['min']:.0f} · Máx: {imob['max']:.0f}")
    with c2:
        st.markdown("**Regionais respondentes / fim de semana**")
        st.metric("Média", f"{reg['media']:.1f}")
        st.caption(f"Mín (sem zero): {reg['min']:.0f} · Máx: {reg['max']:.0f}")


def agregar_prev_real_periodo(df_prev: pd.DataFrame, df_vendas: pd.DataFrame, col_contrato: str, chave: str) -> pd.DataFrame:
    prev = preparar_previsao_df(df_prev)
    if prev.empty:
        return pd.DataFrame()
    if chave == "mes_ano":
        prev["_periodo"] = prev["_sabado"].apply(lambda d: f"{d.month:02d}/{d.year}" if d is not None and pd.notna(d) else "")
    else:
        prev["_periodo"] = prev[chave].astype(str).str.strip() + " · " + prev["_ano"].astype(str).str.strip()
    prev = prev[prev["_periodo"].astype(str).str.strip() != ""]
    agg = prev.groupby(["_periodo", "_grupo_canal"], as_index=False).agg(prev_qtd=("_prev_qtd", "sum"), prev_vgv=("_prev_vgv", "sum"))
    real = vendas_reais_por_fim_de_semana(df_vendas, col_contrato)
    if real.empty:
        agg["real_qtd"] = 0.0
        agg["real_vgv"] = 0.0
        return agg
    if chave == "mes_ano":
        real["_periodo"] = real["_sabado"].apply(lambda d: f"{d.month:02d}/{d.year}" if d else "")
    else:
        real["_periodo"] = real["_sabado"].map(prev.drop_duplicates("_sabado").set_index("_sabado")["_periodo"])
    agg_r = real.groupby(["_periodo", "_grupo_canal"], as_index=False).agg(real_qtd=("_real_qtd", "sum"), real_vgv=("_real_vgv", "sum"))
    return agg.merge(agg_r, on=["_periodo", "_grupo_canal"], how="outer").fillna(0).sort_values("_periodo")


def render_barras_prev_real(df: pd.DataFrame, titulo: str, metrica: str = "qtd") -> None:
    if df.empty:
        st.info(f"Sem dados para {titulo}.")
        return
    fig = go.Figure()
    prev_col, real_col = ("prev_qtd", "real_qtd") if metrica == "qtd" else ("prev_vgv", "real_vgv")
    for canal, cor in (("IMOB", "#2563eb"), ("DV", "#dc2626")):
        sub = df[df["_grupo_canal"] == canal]
        if sub.empty:
            continue
        fig.add_trace(go.Bar(x=sub["_periodo"], y=sub[prev_col], name=f"Prev {canal}", marker_color=cor, opacity=0.45))
        fig.add_trace(go.Bar(x=sub["_periodo"], y=sub[real_col], name=f"Real {canal}", marker_color=cor))
    fig.update_layout(barmode="group", height=400, title=titulo)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_aba_feedbacks_comerciais(df: pd.DataFrame) -> None:
    st.subheader("Feedbacks Comerciais")
    if df is None or df.empty:
        st.warning("Não foi possível carregar a planilha de feedbacks.")
        return
    col_coord, col_obs = _achar(df, ALIASES_COORD_FEEDBACK), _achar(df, ALIASES_OBS_FEEDBACK)
    col_nota, col_nota_cca = _achar(df, ALIASES_NOTA_ATEND), _achar(df, ALIASES_NOTA_CCA)
    col_prom_com, col_prom_cca = _achar(df, ALIASES_PROMOTOR_COM), _achar(df, ALIASES_PROMOTOR_CCA)
    st.markdown("##### NPS e nota média por coordenador")
    if col_coord:
        rows = []
        for c in sorted(df[col_coord].dropna().astype(str).str.strip().unique()):
            sub = df[df[col_coord].astype(str).str.strip() == c]
            row = {"Coordenador": c, "Respostas": len(sub)}
            if col_nota:
                row["Nota média atendimento"] = round(pd.to_numeric(sub[col_nota], errors="coerce").mean(), 2)
            if col_nota_cca:
                row["Nota média CCA"] = round(pd.to_numeric(sub[col_nota_cca], errors="coerce").mean(), 2)
            if col_prom_com:
                nps = _nps_de_coluna(sub[col_prom_com])
                row["NPS Comercial"] = round(nps, 1) if nps is not None else None
            if col_prom_cca:
                nps = _nps_de_coluna(sub[col_prom_cca])
                row["NPS CCA"] = round(nps, 1) if nps is not None else None
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.markdown("##### Temas dos comentários (classificação automática)")
    st.caption("Cada observação pode ter mais de um tema.")
    if not col_obs:
        return
    contagem, exemplos = {}, {}
    for txt in df[col_obs].dropna().astype(str):
        if txt.strip().lower() in ("", "nan", "não", "nao", "."):
            continue
        for tema in classificar_temas_comentario(txt):
            contagem[tema] = contagem.get(tema, 0) + 1
            exemplos.setdefault(tema, [])
            if len(exemplos[tema]) < 3:
                exemplos[tema].append(txt[:120])
    if contagem:
        st.dataframe(pd.DataFrame([{"Tema": k, "Quantidade": v, "Exemplo": exemplos.get(k, [""])[0]} for k, v in sorted(contagem.items(), key=lambda x: -x[1])]), use_container_width=True, hide_index=True)


def render_aba_previsao_vendas(df_prev: pd.DataFrame, df_vendas: pd.DataFrame, col_contrato: str) -> None:
    st.subheader("Previsão de Vendas")
    if df_prev is None or df_prev.empty:
        st.warning("Planilha de previsão indisponível.")
        return
    prev = preparar_previsao_df(df_prev)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        emp_sel = st.multiselect("Empreendimento", sorted({e for e in prev["_emp"].unique() if str(e).strip() and str(e).lower() != "nan"}), default=[])
    with c2:
        canal_sel = st.selectbox("Canal", ["Todos", "IMOB", "DV"], index=0)
    datas = sorted(prev["_sabado"].dropna().unique())
    with c3:
        d_ini = st.date_input("Sábado inicial", value=min(datas) if datas else date.today())
    with c4:
        d_fim = st.date_input("Sábado final", value=max(datas) if datas else date.today())
    filtro = prev.copy()
    if emp_sel:
        filtro = filtro[filtro["_emp"].isin(emp_sel)]
    if canal_sel != "Todos":
        filtro = filtro[filtro["_grupo_canal"] == canal_sel]
    if datas:
        filtro = filtro[(filtro["_sabado"] >= d_ini) & (filtro["_sabado"] <= d_fim)]
    cruz = cruzar_previsao_realizado(filtro, df_vendas, col_contrato)
    if canal_sel != "Todos":
        cruz = cruz[cruz["_grupo_canal"] == canal_sel]
    st.markdown("##### Previsto × realizado por fim de semana")
    st.caption("Realizado SF = Contrato Gerado Em (sexta a domingo).")
    render_grafico_previsto_realizado_linha(cruz)
    st.markdown("##### Dispersão previsto × realizado")
    render_scatter_previsto_realizado(cruz)
    render_tabela_erros(cruz)
    render_resumo_respondentes(cruz)
    st.markdown("##### Agregado")
    t1, t2, t3 = st.tabs(["Trimestre · Ano", "Mês · Ano", "Semestre · Ano"])
    with t1:
        agg = agregar_prev_real_periodo(filtro, df_vendas, col_contrato, "_trimestre")
        render_barras_prev_real(agg, "Vendas — trimestre", "qtd")
        render_barras_prev_real(agg, "VGV — trimestre", "vgv")
    with t2:
        agg = agregar_prev_real_periodo(filtro, df_vendas, col_contrato, "mes_ano")
        render_barras_prev_real(agg, "Vendas — mês/ano", "qtd")
        render_barras_prev_real(agg, "VGV — mês/ano", "vgv")
    with t3:
        agg = agregar_prev_real_periodo(filtro, df_vendas, col_contrato, "_semestre")
        render_barras_prev_real(agg, "Vendas — semestre", "qtd")
        render_barras_prev_real(agg, "VGV — semestre", "vgv")
