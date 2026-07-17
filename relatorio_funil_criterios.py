# -*- coding: utf-8 -*-
"""
Relatório 2 — Funil filtrado por critérios de quantidade
(Regional / Gerente de vendas / Corretor).

Retorna apenas pessoas que atingem os mínimos escolhidos
em um ou mais indicadores do funil, no período informado.

Hospedagem Streamlit Cloud (mesma pasta / secrets do velocímetro).
"""
from __future__ import annotations

from datetime import date
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

from funil_pessoas_comum import (
    COR_AZUL_ESC,
    COR_TEXTO_PRETO,
    DIM_LABELS,
    DIMENSOES,
    FUNIL_ETAPAS,
    FUNIL_LABELS,
    agregar_funil_por_dimensao,
    aplicar_estilo_basico,
    carregar_eventos_funil_pessoas,
    filtrar_periodo,
    fmt_num,
    limpar_nome,
    n_dias_periodo,
    semana_iso_atual,
)


def _aplicar_criterios(
    df: pd.DataFrame,
    criterios: Dict[str, Optional[float]],
) -> pd.DataFrame:
    """Mantém linhas que satisfazem todos os mínimos informados (> None)."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    out = df.copy()
    for etapa, minimo in criterios.items():
        if minimo is None:
            continue
        if etapa not in out.columns:
            out[etapa] = 0.0
        out = out[out[etapa].astype(float) >= float(minimo)]
    return out.reset_index(drop=True)


def _fmt_funil_row(row: pd.Series) -> Dict[str, str]:
    return {FUNIL_LABELS[e]: fmt_num(float(row.get(e, 0.0)), 0) for e in FUNIL_ETAPAS}


def main() -> None:
    st.set_page_config(
        page_title="Funil · Filtro por critérios | Direcional",
        layout="wide",
        page_icon="🎯",
    )
    aplicar_estilo_basico()
    st.title("Funil por critérios de quantidade")
    st.caption(
        "Filtre o período e defina mínimos em um ou mais indicadores. "
        "O relatório devolve somente Regionais, Gerentes ou Corretores "
        "que atingem todos os critérios escolhidos — com o funil completo."
    )

    hoje = date.today()
    sem_ini, sem_fim = semana_iso_atual(hoje)

    with st.sidebar:
        st.header("Período de análise")
        ini = st.date_input("Início", value=sem_ini, key="crit_ini")
        fim = st.date_input("Fim", value=sem_fim, key="crit_fim")
        if st.button("Usar semana atual (seg–dom)"):
            st.session_state["crit_ini"] = sem_ini
            st.session_state["crit_fim"] = sem_fim
            st.rerun()

        st.header("Dimensão")
        dim = st.radio(
            "Analisar por",
            options=list(DIMENSOES),
            format_func=lambda x: DIM_LABELS[x],
            index=0,
        )

        st.header("Critérios (mínimos)")
        st.caption("Deixe em 0 para não filtrar aquele indicador.")
        criterios: Dict[str, Optional[float]] = {}
        ativos: List[str] = []
        for etapa in FUNIL_ETAPAS:
            v = st.number_input(
                FUNIL_LABELS[etapa],
                min_value=0,
                value=0,
                step=1,
                key=f"min_{etapa}",
            )
            if v > 0:
                criterios[etapa] = float(v)
                ativos.append(f"{FUNIL_LABELS[etapa]} ≥ {int(v)}")
            else:
                criterios[etapa] = None

    if fim < ini:
        st.error("O fim do período deve ser ≥ início.")
        return

    if not ativos:
        st.info(
            "Defina pelo menos um critério > 0 na barra lateral "
            "(ex.: Vendas ≥ 3, Pastas ≥ 10)."
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
        f"**Período:** {ini.strftime('%d/%m/%Y')} → {fim.strftime('%d/%m/%Y')} "
        f"({n_dias_periodo(ini, fim)} dias) · "
        f"**Dimensão:** {DIM_LABELS[dim]}"
    )
    if ativos:
        st.markdown("**Critérios ativos:** " + " · ".join(ativos))

    ev = filtrar_periodo(eventos, ini, fim)
    agg = agregar_funil_por_dimensao(ev, dim)
    filtrado = _aplicar_criterios(agg, criterios) if ativos else agg.iloc[0:0]

    if not ativos:
        return

    if filtrado.empty:
        st.warning("Nenhuma pessoa atende aos critérios no período.")
        return

    # Ordena pelo primeiro critério ativo (maior valor), depois vendas
    ordem = [e for e in FUNIL_ETAPAS if criterios.get(e) is not None]
    if "vendas" not in ordem:
        ordem.append("vendas")
    filtrado = filtrado.sort_values(ordem, ascending=[False] * len(ordem))

    st.success(f"{len(filtrado)} resultado(s) encontrados.")

    # Tabela consolidada
    cols_show = [dim] + list(FUNIL_ETAPAS)
    df_show = filtrado[cols_show].copy()
    df_show = df_show.rename(columns={dim: DIM_LABELS[dim], **FUNIL_LABELS})
    for c in FUNIL_LABELS.values():
        if c in df_show.columns:
            df_show[c] = df_show[c].map(lambda x: fmt_num(float(x), 0))
    st.dataframe(df_show, use_container_width=True, hide_index=True)

    st.markdown("##### Detalhamento")
    for _, row in filtrado.iterrows():
        nome = limpar_nome(row[dim])
        if not nome:
            continue
        partes = [f"{FUNIL_LABELS[e]}: {fmt_num(float(row.get(e, 0.0)), 0)}" for e in FUNIL_ETAPAS]
        st.markdown(
            f'<div class="bloco-pessoa">'
            f'<div class="nome">{nome}</div>'
            f'<div style="color:{COR_TEXTO_PRETO};font-size:0.95rem;">'
            + " &nbsp;·&nbsp; ".join(partes)
            + "</div></div>",
            unsafe_allow_html=True,
        )

    csv = filtrado.rename(columns={dim: DIM_LABELS[dim], **FUNIL_LABELS}).to_csv(
        index=False, sep=";", decimal=","
    )
    st.download_button(
        "Baixar CSV",
        data=csv.encode("utf-8-sig"),
        file_name=f"funil_criterios_{dim}_{ini}_{fim}.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()
