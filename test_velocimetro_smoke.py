# -*- coding: utf-8 -*-
"""Smoke tests: coleta/transformação de todas as abas (sem SF/Sheets live)."""
from __future__ import annotations

import sys
import traceback
import types
from datetime import date, timedelta
from typing import Any, Callable, List
from unittest.mock import MagicMock

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Mock Streamlit antes de importar velocimetro
# ---------------------------------------------------------------------------
_st = types.ModuleType("streamlit")


def _identity_decorator(*_a, **_k):
    def _wrap(fn):
        return fn
    return _wrap


_st.cache_data = _identity_decorator
_st.cache_resource = _identity_decorator
_st.dialog = _identity_decorator
_st.set_page_config = lambda **_k: None
_st.markdown = lambda *_a, **_k: None
_st.caption = lambda *_a, **_k: None
_st.subheader = lambda *_a, **_k: None
_st.header = lambda *_a, **_k: None
_st.title = lambda *_a, **_k: None
_st.error = lambda *_a, **_k: None
_st.warning = lambda *_a, **_k: None
_st.info = lambda *_a, **_k: None
_st.success = lambda *_a, **_k: None
_st.divider = lambda *_a, **_k: None
_st.button = lambda *_a, **_k: False
_st.plotly_chart = lambda *_a, **_k: None
_st.dataframe = lambda *_a, **_k: None
_st.metric = lambda *_a, **_k: None
_st.columns = lambda n: [MagicMock() for _ in range(n if isinstance(n, int) else len(n))]
_st.tabs = lambda labels: [MagicMock() for _ in labels]
_st.selectbox = lambda _l, opts, **_k: opts[0] if opts else None
_st.multiselect = lambda _l, opts, default=None, **_k: default if default is not None else (opts or [])
_st.date_input = lambda _l, value=None, **_k: value or date.today()
_st.number_input = lambda _l, value=0.0, **_k: value
_st.slider = lambda _l, min_value=0, max_value=100, value=50, **_k: value
_st.text_input = lambda _l, value="", **_k: value
_st.checkbox = lambda _l, value=False, **_k: value
_st.expander = lambda *_a, **_k: MagicMock(__enter__=lambda s: s, __exit__=lambda *a: None)
_st.session_state = MagicMock()
_st.secrets = {"connections": {"gsheets": {}}}


class _TextColumnMock:
    def __init__(self, *_a, **_k):
        pass


_st.column_config = MagicMock()
_st.column_config.TextColumn = _TextColumnMock


class _ColMock:
    def __init__(self, *args, **kwargs):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def metric(self, *_a, **_k):
        pass

    def markdown(self, *_a, **_k):
        pass


_st.columns = lambda spec: [_ColMock() for _ in (range(spec) if isinstance(spec, int) else spec)]

sys.modules["streamlit"] = _st

import velocimetro as v  # noqa: E402
import velocimetro_feedbacks_previsao as vfp  # noqa: E402

FAILURES: List[str] = []


def run(name: str, fn: Callable[[], Any]) -> None:
    try:
        fn()
        print(f"  OK  {name}")
    except Exception as exc:
        FAILURES.append(name)
        print(f"  FAIL {name}: {exc}")
        traceback.print_exc(limit=3)


def _df_vendas(n: int = 20) -> pd.DataFrame:
    hoje = date.today()
    rows = []
    emps = ["Emp A", "Emp B", "Emp C"]
    coords = ["Dutra", "Leo", "Luciano"]
    imobs = ["RJ-001", "DIR-002", "RJG-003"]
    for i in range(n):
        d = hoje - timedelta(days=i % 14)
        rows.append({
            "Empreendimento": emps[i % 3],
            "Região": "RJ",
            "Coordenador": coords[i % 3],
            "Imobiliária": imobs[i % 3],
            "Canal": "RIO" if i % 2 == 0 else "DIR",
            "Contrato Gerado Em": d.isoformat(),
            "Valor Real de Venda": 250000 + i * 1000,
            "Venda Comercial?": "Sim",
            "Venda facilitada": "0",
            "_qtd_venda": 1.0,
            "_vgv_venda": 250000.0 + i * 1000,
            "_vgv": 250000.0 + i * 1000,
            "_peso_coord": 1.0,
            "Canal_Agrupado": "IMOB" if i % 2 == 0 else "DV RJ",
            "Gerente regional": "GR1",
        })
    return pd.DataFrame(rows)


def _df_metas_raw() -> pd.DataFrame:
    return pd.DataFrame({
        "Empreendimento": ["Emp A", "Emp B"],
        "Região": ["RJ", "RJ"],
        "Coordenador": ["Dutra", "Leo"],
        "Qtd 8": [10, 20],
        "VGV 8": ["1.000.000", "2.000.000"],
    })


def _df_metas_coord() -> pd.DataFrame:
    return pd.DataFrame({
        "Empreendimento": ["Emp A", "Emp B", "Emp C"],
        "Coordenador": ["Dutra", "Leo", "Luciano"],
        "Mes_Num": [8, 8, 8],
        "Ano_Num": [2026, 2026, 2026],
        "Meta Vendas Desafio": [43.0, 40.0, 70.0],
        "Meta Vendas BP": [35.0, 32.0, 55.0],
        "Meta Agendamentos Desafio": [100.0, 90.0, 110.0],
        "Meta Visitas Desafio": [50.0, 45.0, 55.0],
    })


def _df_estoque() -> pd.DataFrame:
    return pd.DataFrame({
        "Empreendimento": ["Emp A", "Emp A", "Emp B"],
        "Status da Unidade": ["Disponível", "Mirror", "Disponível"],
        "Valor Final com Kit": [300000, 320000, 280000],
        "Identificador": ["A101", "A102", "B201"],
        "Tipologia": ["2Q", "2Q", "3Q"],
    })


def _df_ag() -> pd.DataFrame:
    hoje = date.today()
    return pd.DataFrame({
        "Empreendimento": ["Emp A", "Emp B"],
        "CreatedDate": [(hoje - timedelta(days=2)).isoformat()] * 2,
        "Data da visita": [hoje.isoformat(), (hoje - timedelta(days=1)).isoformat()],
        "Código do agendamento": ["AG1", "AG2"],
    })


def _df_pastas() -> pd.DataFrame:
    hoje = date.today()
    return pd.DataFrame({
        "Empreendimento": ["Emp A", "Emp B"],
        "Data Primeiro Envio Análise": [hoje.isoformat()] * 2,
        "Data Aprovação SAFI": [hoje.isoformat(), pd.NA],
        "Nome da Avaliação": ["AV1", "AV2"],
        "Renda": [5000, 6000],
        "Total Sinal": [10000, 12000],
    })


def _filtros_glob():
    return v.FiltrosGlobais(
        data_ini=date(2026, 8, 1),
        data_fim=date(2026, 8, 13),
        mes_meta=8,
        ano_meta=2026,
        tipo_meta_col="Desafio",
        tipo_indicador="vendas",
        canal_sel="Todos",
        canal_meta="RIO",
        coords_sel=["Dutra", "Leo"],
        emps_sel=[],
        imobs_sel=[],
        status_estoque_sel=["Disponível", "Mirror"],
    )


def _filtros_dashboard():
    fg = _filtros_glob()
    return v.filtros_glob_to_dashboard(fg)


# ---------------------------------------------------------------------------
# Testes por área
# ---------------------------------------------------------------------------

def test_metas_transforms():
    raw = _df_metas_raw()
    melted = v.melt_metas(raw)
    assert not melted.empty
    adapted = v.adaptar_metas_melt_para_coord(melted, 2026)
    assert "Meta Vendas Desafio" in adapted.columns
    mapa = v.mapa_emp_coordenador(_df_metas_coord(), 8, 2026)
    assert len(mapa) >= 2
    qtd = v.soma_meta_coord(_df_metas_coord(), 8, 2026, "vendas", "Desafio")
    assert qtd > 0


def test_filtros_periodo():
    df = _df_ag()
    ini, fim = date(2026, 8, 1), date(2026, 8, 13)
    r1 = v._filtrar_df_periodo(df, v.ALIASES_DATA_CRIACAO, ini, fim)
    r2 = v._filtrar_df_periodo(df, "CreatedDate", ini, fim)
    assert isinstance(r1, pd.DataFrame)
    assert isinstance(r2, pd.DataFrame)
    empty = v._coalesce_df(pd.DataFrame())
    assert empty.empty
    d = {"x": pd.DataFrame()}
    assert v._coalesce_dict_df(d, "x").empty
    assert v._coalesce_dict_df(d, "missing").empty


def test_funil_empreendimento():
    df_ag, df_pas, df_ven = _df_ag(), _df_pastas(), _df_vendas(5)
    t = v.totais_funil_empreendimento(
        df_ag, df_pas, df_ven, "Emp A",
        date(2026, 8, 1), date(2026, 8, 13),
    )
    assert isinstance(t, dict)
    emps = v._empreendimentos_rj_direcional(
        v.melt_metas(_df_metas_raw()), df_ven, df_ag, df_pas, pd.DataFrame(),
    )
    assert isinstance(emps, list)


def test_pareto_abc():
    df = _df_vendas(30)
    prep = v._prep_vendas_canal(df)
    assert not prep.empty
    # _plot_pareto_abc chama st.plotly_chart (mocked)
    v._plot_pareto_abc(prep, "Empreendimento", "Teste", "test_pareto")
    v._plot_pareto_abc(prep, "Imobiliária", "Teste", "test_pareto_imob", top_n=10)
    v._plot_pareto_abc(pd.DataFrame(), "Empreendimento", "Vazio", "test_empty")


def test_dashboard_transforms():
    df_v = _df_vendas()
    df_est = _df_estoque()
    df_ag, df_pas = _df_ag(), _df_pastas()
    filtros = _filtros_dashboard()
    mapa = v.mapa_emp_coordenador(_df_metas_coord(), 8, 2026)
    col_data = "Contrato Gerado Em"
    df_f = v._aplicar_filtros_base(df_v, filtros, mapa, col_data, usar_periodo=True)
    assert isinstance(df_f, pd.DataFrame)
    sem, tot, pct = v.calcular_pastas_sem_visita(df_ag, df_pas, None)
    assert tot >= 0
    tab_ps = v.montar_tabela_ps_sinais_vgv(df_v, col_data, filtros, mapa)
    assert isinstance(tab_ps, pd.DataFrame)
    est_map = v.resumo_estoque_empreendimentos(df_est)
    sr_map = v.calcular_sinal_sobre_renda_por_emp(
        pd.DataFrame(), df_pas, filtros.data_ini, filtros.data_fim,
    )
    tab_sre = v.montar_tabela_sinal_renda_estoque_emp(
        list(mapa.keys()), est_map, {}, sr_map,
    )
    assert isinstance(tab_sre, pd.DataFrame)
    kpi, enr = v.agregar_estoque(df_est)
    assert "unidades" in kpi


def test_painel_v2_transforms():
    filtros = v.filtros_glob_to_v2(_filtros_glob())
    df_v = _df_vendas()
    col_c = "Contrato Gerado Em"
    vendas_f = v.filtrar_vendas_painel_v2(df_v, filtros, col_c, "Canal")
    assert isinstance(vendas_f, pd.DataFrame)
    real_q, real_v = v.realizado_vendas_periodo(
        df_v, col_c, filtros.data_ini, filtros.data_fim,
    )
    assert real_q >= 0
    tab = v.montar_tabela_analitica(
        ["Emp A", "Emp B"],
        pd.DataFrame(), pd.DataFrame(),
        df_v, _df_ag(), _df_pastas(),
        _df_metas_coord(), filtros, col_c, None,
        df_cotacoes=pd.DataFrame(),
        df_pastas_aprov=pd.DataFrame(),
        df_tabela_comp=pd.DataFrame(),
        estoque_map={}, total_unidades_por_emp={},
    )
    assert isinstance(tab, pd.DataFrame)


def test_poder_compra():
    df_pas = pd.DataFrame({
        "Empreendimento": ["Emp A", "Emp A"],
        "FGTS": [10000, 15000],
        "Valor de Subsidio": [20000, 25000],
        "Financiamento": [150000, 160000],
        "Renda": [5000, 6000],
        "Data Aprovação SAFI": [date.today().isoformat()] * 2,
    })
    kpi, df_est_enr = v.agregar_estoque(_df_estoque())
    resumo = v.calcular_resumo_ineficiencia_emp(
        df_pas, df_est_enr, _df_vendas(3), pd.DataFrame(), "Emp A",
    )
    assert resumo.pastas_aprovadas >= 0
    stats = v.estatisticas_preco_estoque(df_est_enr, "Emp A")
    assert "mediano" in stats
    stats_raw = v.estatisticas_preco_estoque(_df_estoque(), "Emp A")
    assert stats_raw["n"] >= 0


def test_projecao_funil():
    hoje = date.today()
    etapas = {
        e: {
            "mtd": 10.0, "projetado_reg": 50.0, "projetado_med": 48.0,
            "projetado_mtd_reg": 20.0, "r2": 0.5, "r2_medias": 0.3,
            "diaria": pd.DataFrame({"dia": [1, 2], "realizado": [1, 2]}),
        }
        for e in v.FUNIL_ETAPAS
    }
    proj = {
        "hoje": hoje,
        "ultimo_dia": 31,
        "real_mtd": {"vendas": 53, "agendamentos": 631},
        "proj_mes": {"vendas": 180.0},
        "pred_reg_mes": np.ones(31) * 3.0,
        "modelo_schema": "test",
        "inicio_treino": "2023-07-01",
        "fim_treino": "2026-06-30",
        "etapas": etapas,
        "r2s": {e: 0.5 for e in v.FUNIL_ETAPAS},
        "r2s_medias": {e: 0.3 for e in v.FUNIL_ETAPAS},
        "metas_diarias": {},
        "conversoes": {},
    }
    v.render_projecao_funil(proj)


def test_feedbacks_previsao():
    vfp.classificar_temas_comentario("Atendimento rápido e eficiente, parabéns ao coordenador")
    assert vfp.sabado_referencia_de_data(date(2026, 8, 8)) == date(2026, 8, 8)  # sábado
    df_prev = pd.DataFrame({
        "Data de referência (indicar sempre o sábado da semana)": ["08/08/2026"],
        "Vendas Previsão": ["10"],
        "VGV Previsão": ["2500000"],
        "Canal": ["IMOB"],
        "Empreendimento (Previsto para ter venda)": ["Emp A"],
    })
    prep = vfp.preparar_previsao_df(df_prev)
    assert not prep.empty
    real = vfp.vendas_reais_por_fim_de_semana(_df_vendas(), "Contrato Gerado Em")
    assert isinstance(real, pd.DataFrame)
    merged = vfp.cruzar_previsao_realizado(df_prev, _df_vendas(), "Contrato Gerado Em")
    assert isinstance(merged, pd.DataFrame)
    df_fb = pd.DataFrame({
        "Selecione o Coordenador Comercial responsável pelo processo da venda": ["Dutra"],
        "De 1 a 10, avalie a Qualidade do Atendimento, sendo 1 caso esteja Muito Insatisfeito e 10 caso esteja Muito Satisfeito": ["9"],
        "Promotor Comercial": ["promotor"],
        "Deseja faazer alguma observação... Pode ser um elogio": ["Atendimento ágil"],
    })
    # render sem crash
    vfp.render_aba_feedbacks_comerciais(df_fb)
    vfp.render_aba_previsao_vendas(df_prev, _df_vendas(), "Contrato Gerado Em")


def test_fallback_metas():
    melted = v.melt_metas(_df_metas_raw())
    df, aviso = v.carregar_metas_coordenadores_com_fallback("fake_cred", melted, 2026)
    assert not df.empty
    assert aviso is not None  # planilha real não carregou


def test_graficos_dashboard():
    df = v._prep_vendas_canal(_df_vendas())
    col = "Contrato Gerado Em"
    v.render_grafico_vendas_mes_canal(df, col)
    v.render_grafico_vendas_emp_canal(df)
    v.render_grafico_share_estoque_produto(
        _df_estoque(), ["Disponível", "Mirror"], _filtros_dashboard(),
        v.mapa_emp_coordenador(_df_metas_coord(), 8, 2026),
    )


def test_comparativos_mtd():
    v.render_comparativos_mtd_funil(_df_vendas(), "Contrato Gerado Em")


def test_render_dashboard_secoes():
    df_v = _df_vendas()
    df_est = _df_estoque()
    filtros = _filtros_dashboard()
    mapa = v.mapa_emp_coordenador(_df_metas_coord(), 8, 2026)
    col = "Contrato Gerado Em"
    v.render_secao_analises_avancadas(
        df_v, df_est, _df_ag(), _df_pastas(),
        filtros, mapa, col, ["Disponível", "Mirror"],
        df_cotacoes=pd.DataFrame(),
    )
    melted = v.melt_metas(_df_metas_raw())
    v.render_dashboard_comercial(
        df_v, df_est, pd.DataFrame(), "cred",
        col_data_venda=col, filtros_glob=_filtros_glob(),
        df_metas_fallback=melted,
    )


def test_preparar_vendas_painel():
    raw = _df_vendas(5)
    raw["Venda Comercial?"] = "Sim"
    raw["Contrato Gerado Em"] = date.today().isoformat()
    raw["Empreendimento"] = "Emp A"
    metas = v.preparar_metas_painel(_df_metas_raw())
    if metas.empty or "_peso_coord" not in metas.columns:
        metas = pd.DataFrame({
            "Empreendimento": ["Emp A"],
            "Região": ["RJ"],
            "Coordenador": ["Dutra"],
            "Mes_Num": [8],
            "Meta_Qtd": [10.0],
            "Meta_VGV": [1_000_000.0],
            "Regiao_Coord": ["RJ - Dutra"],
            "_peso_coord": [1.0],
        })
    prep = v.preparar_vendas_painel(raw, metas)
    assert "_vgv_venda" in prep["df_vendas_painel"].columns
    assert prep.get("col_contrato_gerado")


def test_cache_manifest():
    import velocimetro_cache as vc

    vc.gravar_manifest_local({"atualizado_em": "2099-01-01 12:00:00"})
    m = vc.ler_manifest_local()
    assert vc.manifest_valido(m)
    vc.limpar_cache_local()


def test_metricas_vendas_cache_sheets():
    """Colunas numéricas como string (Google Sheets) não devem quebrar KPIs."""
    df = _df_vendas()
    df["_qtd_venda"] = df["_qtd_venda"].astype(str)
    df["_vgv_venda"] = df["_vgv_venda"].astype(str)
    df["_peso_coord"] = df["_peso_coord"].astype(str)
    norm = v.assegurar_metricas_vendas(df)
    assert pd.api.types.is_numeric_dtype(norm["_qtd_venda"])
    col_c = "Contrato Gerado Em"
    qtd, vgv = v.realizado_vendas_periodo(
        norm, col_c, date(2026, 8, 1), date(2026, 8, 21),
    )
    assert qtd >= 0
    assert vgv >= 0


def test_metas_fallback_vgv_e_mes():
    legacy = pd.DataFrame({
        "Empreendimento": ["Emp A", "Emp B"],
        "Coordenador": ["Dutra", "Leo"],
        "Mes_Num": [8, 8],
        "Meta_Qtd": [10.0, 20.0],
        "Meta_VGV": [1_000_000.0, 2_000_000.0],
    })
    adapt = v.adaptar_metas_melt_para_coord(legacy, 2026)
    assert "Meta VGV Desafio (Caixa Único)" in adapt.columns
    total = v.soma_meta_vgv_coord(adapt, 8, 2026, "Desafio")
    assert total == 3_000_000.0
    assert v._parse_mes_num("Agosto") == 8
    assert v._parse_mes_num("8") == 8
    m = v._filtrar_metas_mes_ano(adapt, 8, 2026)
    assert len(m) == 2
    # Coordenadores com qtd mas sem VGV → fallback legado
    coord_sem_vgv = pd.DataFrame({
        "Empreendimento": ["Emp A"],
        "Coordenador": ["Dutra"],
        "Mes_Num": [8],
        "Ano_Num": [2026],
        "Meta Vendas Desafio": [5.0],
        "Meta VGV Desafio (Caixa Único)": [0.0],
    })
    assert not v._metas_coord_tem_vgv_mes(coord_sem_vgv, 8, 2026)
    assert v._metas_coord_tem_dados_mes(coord_sem_vgv, 8, 2026)


def test_tabela_analitica_styler():
    df = pd.DataFrame({
        "Empreendimento": ["Emp A", "Emp B"],
        "Desenquadramento_Pct": [30.0, 75.0],
        "Meta_Mes": [10.0, 20.0],
        "Meta_Dia": [10.0, 20.0],
    })
    prep = v.preparar_df_tabela_exibicao(df)
    assert prep.columns.is_unique
    col = v.ROTULOS_COLUNAS_TABELA["Desenquadramento_Pct"]
    styler = v._styler_desenquadramento(prep, col)
    assert styler is not None or col not in prep.columns
    v.render_tabela_analitica(df)


def main():
    print("=" * 60)
    print("Smoke tests — velocimetro (transformações)")
    print("=" * 60)
    tests = [
        ("Metas (melt/adaptar/mapa)", test_metas_transforms),
        ("Filtros período / coalesce", test_filtros_periodo),
        ("Funil por empreendimento", test_funil_empreendimento),
        ("Pareto ABC", test_pareto_abc),
        ("Dashboard transforms", test_dashboard_transforms),
        ("Painel v2 transforms", test_painel_v2_transforms),
        ("Poder de compra", test_poder_compra),
        ("Projeção funil (render)", test_projecao_funil),
        ("Feedbacks + Previsão", test_feedbacks_previsao),
        ("Fallback metas", test_fallback_metas),
        ("Gráficos dashboard", test_graficos_dashboard),
        ("Dashboard seções (render)", test_render_dashboard_secoes),
        ("Comparativos MTD", test_comparativos_mtd),
        ("Preparar vendas painel", test_preparar_vendas_painel),
        ("Cache manifest", test_cache_manifest),
        ("Métricas vendas cache Sheets", test_metricas_vendas_cache_sheets),
        ("Metas fallback VGV e mês", test_metas_fallback_vgv_e_mes),
        ("Tabela analítica styler", test_tabela_analitica_styler),
    ]
    for name, fn in tests:
        run(name, fn)
    print("=" * 60)
    if FAILURES:
        print(f"FALHOU: {len(FAILURES)} teste(s): {', '.join(FAILURES)}")
        sys.exit(1)
    print("TODOS OK —", len(tests), "suites")
    sys.exit(0)


if __name__ == "__main__":
    main()
