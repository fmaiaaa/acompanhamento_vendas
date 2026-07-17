# -*- coding: utf-8 -*-
"""
Relatório 1 — Funil: média histórica × período selecionado
por Regional, Gerente de vendas e Corretor.

Hospedagem Streamlit Cloud (mesma pasta / secrets do velocímetro).
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Dict, Optional, Tuple

import pandas as pd
import streamlit as st

from funil_pessoas_comum import (
    COR_AZUL_ESC,
    COR_TEXTO_PRETO,
    COR_VERMELHO,
    DIM_LABELS,
    DIMENSOES,
    FUNIL_ETAPAS,
    FUNIL_LABELS,
    agregar_funil_por_dimensao,
    aplicar_estilo_basico,
    carregar_eventos_funil_pessoas,
    domingo_da_semana,
    escalar_media_para_periodo,
    filtrar_periodo,
    fmt_num,
    fmt_pct,
    limpar_nome,
    n_dias_periodo,
    segunda_da_semana,
    semana_iso_atual,
)


def _default_base_periodo(semana_ini: date, semana_fim: date) -> Tuple[date, date]:
    """Base padrão: 12 semanas anteriores à semana selecionada (sem incluí-la)."""
    fim_base = semana_ini - timedelta(days=1)
    ini_base = segunda_da_semana(fim_base - timedelta(days=7 * 11))
    if ini_base > fim_base:
        ini_base = fim_base - timedelta(days=83)
    return ini_base, fim_base


def _montar_tabela_pessoa(
    nome: str,
    media: Dict[str, float],
    realizado: Dict[str, float],
) -> pd.DataFrame:
    rows = []
    for rotulo, fonte in (
        ("Média (equivalente ao período)", media),
        ("Realizado do período", realizado),
    ):
        row = {"Linha": rotulo}
        for e in FUNIL_ETAPAS:
            row[FUNIL_LABELS[e]] = float(fonte.get(e, 0.0))
        rows.append(row)

    row_pct = {"Linha": "% Realizado / Média"}
    for e in FUNIL_ETAPAS:
        m = float(media.get(e, 0.0))
        r = float(realizado.get(e, 0.0))
        row_pct[FUNIL_LABELS[e]] = (100.0 * r / m) if m > 1e-9 else None
    rows.append(row_pct)
    return pd.DataFrame(rows)


def _estilo_tabela(df: pd.DataFrame):
    df_fmt = df.copy()
    for i, row in df.iterrows():
        linha = str(row.get("Linha", ""))
        for c in df.columns:
            if c == "Linha":
                continue
            val = row[c]
            if val is None or (isinstance(val, float) and pd.isna(val)):
                df_fmt.at[i, c] = "—"
            elif linha.startswith("%"):
                df_fmt.at[i, c] = fmt_pct(float(val))
            else:
                df_fmt.at[i, c] = fmt_num(float(val), 1)

    def highlight_pct(row):
        if not str(row.get("Linha", "")).startswith("%"):
            return [""] * len(df.columns)
        cores = []
        orig = df.loc[row.name]
        for c in df.columns:
            if c == "Linha":
                cores.append(
                    f"background-color: #f8fafc; font-weight: 600; color: {COR_TEXTO_PRETO};"
                )
                continue
            v = orig.get(c)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                cores.append("")
            elif float(v) >= 100:
                cores.append("background-color: #ecfdf5; color: #065f46; font-weight: 600;")
            elif float(v) >= 80:
                cores.append("background-color: #fffbeb; color: #92400e; font-weight: 600;")
            else:
                cores.append(f"background-color: #fef2f2; color: {COR_VERMELHO}; font-weight: 600;")
        return cores

    return df_fmt.style.apply(highlight_pct, axis=1)


def _render_aba_dimensao(
    eventos: pd.DataFrame,
    dimensao: str,
    ini_periodo: date,
    fim_periodo: date,
    ini_base: date,
    fim_base: date,
) -> None:
    label_dim = DIM_LABELS.get(dimensao, dimensao)
    dias_periodo = n_dias_periodo(ini_periodo, fim_periodo)
    dias_base = n_dias_periodo(ini_base, fim_base)

    ev_periodo = filtrar_periodo(eventos, ini_periodo, fim_periodo)
    ev_base = filtrar_periodo(eventos, ini_base, fim_base)

    real = agregar_funil_por_dimensao(ev_periodo, dimensao)
    base = agregar_funil_por_dimensao(ev_base, dimensao)
    media = escalar_media_para_periodo(base, dias_base, dias_periodo, dimensao)

    # Somente quem tem ≥1 indicador positivo no período escolhido
    if real.empty:
        st.info(f"Nenhum {label_dim.lower()} com indicadores no período selecionado.")
        return

    mask_pos = real[list(FUNIL_ETAPAS)].sum(axis=1) > 0
    real = real.loc[mask_pos].copy()
    if real.empty:
        st.info(f"Nenhum {label_dim.lower()} com indicador positivo no período.")
        return

    media_idx = media.set_index(dimensao) if not media.empty else pd.DataFrame()
    real = real.sort_values("vendas", ascending=False)

    st.caption(
        f"{len(real)} {label_dim.lower()}(is) com pelo menos 1 indicador no período · "
        f"média = (total da base ÷ {dias_base} dias) × {dias_periodo} dias do período."
    )

    for _, row in real.iterrows():
        nome = limpar_nome(row[dimensao])
        if not nome:
            continue
        realizado = {e: float(row.get(e, 0.0)) for e in FUNIL_ETAPAS}
        if nome in media_idx.index:
            media_row = media_idx.loc[nome]
            media_vals = {e: float(media_row.get(e, 0.0)) for e in FUNIL_ETAPAS}
        else:
            media_vals = {e: 0.0 for e in FUNIL_ETAPAS}

        df_t = _montar_tabela_pessoa(nome, media_vals, realizado)
        # Exibe % já como número; formatação trata a linha
        st.markdown(
            f'<div class="bloco-pessoa"><div class="nome">{nome}</div></div>',
            unsafe_allow_html=True,
        )
        # Converte linha de % para exibição amigável numa cópia formatada
        df_show = df_t.copy()
        # Mantém numérico; o styler formata
        st.dataframe(_estilo_tabela(df_show), use_container_width=True, hide_index=True)


def main() -> None:
    st.set_page_config(
        page_title="Funil · Média × Período | Direcional",
        layout="wide",
        page_icon="📊",
    )
    aplicar_estilo_basico()
    st.title("Funil por pessoa — média × período")
    st.caption(
        "Compara o funil Agendamentos → Visitas → Pastas → Pastas aprovadas → Vendas "
        "da média histórica (escalada pelos dias) com o período escolhido. "
        "Abas: Regional, Gerente de vendas e Corretor."
    )

    hoje = date.today()
    sem_ini_padrao, sem_fim_padrao = semana_iso_atual(hoje)

    with st.sidebar:
        st.header("Período de análise")
        st.caption("Padrão: semana atual (segunda → domingo).")
        ini_periodo = st.date_input("Início do período", value=sem_ini_padrao, key="ini_per")
        fim_periodo = st.date_input("Fim do período", value=sem_fim_padrao, key="fim_per")
        if st.button("Usar semana atual (seg–dom)"):
            st.session_state["ini_per"] = sem_ini_padrao
            st.session_state["fim_per"] = sem_fim_padrao
            st.rerun()
        if st.button("Ajustar para seg–dom da data início"):
            d0 = segunda_da_semana(ini_periodo)
            st.session_state["ini_per"] = d0
            st.session_state["fim_per"] = domingo_da_semana(d0)
            st.rerun()

        st.header("Base da média")
        st.caption(
            "Período usado para calcular a média diária "
            "(padrão: 12 semanas anteriores, sem o período selecionado)."
        )
        ini_b_pad, fim_b_pad = _default_base_periodo(ini_periodo, fim_periodo)
        ini_base = st.date_input("Início da base", value=ini_b_pad, key="ini_base")
        fim_base = st.date_input("Fim da base", value=fim_b_pad, key="fim_base")

    if fim_periodo < ini_periodo:
        st.error("O fim do período deve ser ≥ início.")
        return
    if fim_base < ini_base:
        st.error("O fim da base deve ser ≥ início.")
        return

    # Evita vazamento: remove interseção período ∩ base
    if not (fim_base < ini_periodo or ini_base > fim_periodo):
        st.warning(
            "A base da média intersecta o período selecionado. "
            "A média ficará enviesada — ajuste os filtros se quiser excluir o período."
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
        f"**Período:** {ini_periodo.strftime('%d/%m/%Y')} → {fim_periodo.strftime('%d/%m/%Y')} "
        f"({n_dias_periodo(ini_periodo, fim_periodo)} dias) · "
        f"**Base da média:** {ini_base.strftime('%d/%m/%Y')} → {fim_base.strftime('%d/%m/%Y')} "
        f"({n_dias_periodo(ini_base, fim_base)} dias)"
    )

    tabs = st.tabs([DIM_LABELS[d] for d in DIMENSOES])
    for tab, dim in zip(tabs, DIMENSOES):
        with tab:
            _render_aba_dimensao(
                eventos, dim, ini_periodo, fim_periodo, ini_base, fim_base
            )


if __name__ == "__main__":
    main()
