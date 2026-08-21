# -*- coding: utf-8 -*-
"""Verifica acesso às planilhas e datasets usados por cada aba do velocímetro."""
from __future__ import annotations

import sys
import types
from datetime import date
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

_DIR = Path(__file__).resolve().parent
_ROOT = _DIR.parent
sys.path.insert(0, str(_DIR))
sys.path.insert(0, str(_ROOT / "sf-sync-direcional"))

# Mock Streamlit antes de importar velocimetro
_st = types.ModuleType("streamlit")


def _identity_decorator(*_a, **_k):
    def _wrap(fn):
        return fn
    return _wrap


_st.dialog = _identity_decorator
_st.cache_data = lambda **_k: (lambda fn: fn)
_st.cache_resource = lambda **_k: (lambda fn: fn)
_st.set_page_config = lambda **_k: None
_st.secrets = {"connections": {"gsheets": {}}}


class _TextColumnMock:
    def __init__(self, *_a, **_k):
        pass


_st.column_config = types.SimpleNamespace(TextColumn=_TextColumnMock)
sys.modules["streamlit"] = _st

import velocimetro as v  # noqa: E402
import velocimetro_cache as vc  # noqa: E402
import velocimetro_feedbacks_previsao as vfp  # noqa: E402

from sf_credentials import aplicar_secrets_toml_local, credenciais_google_sheets  # noqa: E402

Resultado = Tuple[str, str, str]  # aba/grupo, recurso, status


def _carregar_credenciais() -> Optional[Dict[str, Any]]:
    aplicar_secrets_toml_local()
    raw = credenciais_google_sheets()
    if not raw:
        return None
    info = v.montar_service_account_info(raw)
    if info:
        return info
    # JSON aninhado → campos planos
    return v.montar_service_account_info(
        {
            "type": raw.get("type"),
            "project_id": raw.get("project_id"),
            "private_key_id": raw.get("private_key_id"),
            "private_key": raw.get("private_key"),
            "client_email": raw.get("client_email"),
            "client_id": raw.get("client_id"),
            "auth_uri": raw.get("auth_uri"),
            "token_uri": raw.get("token_uri"),
        }
    )


def _testar_aba(
    sid: str,
    ws: str,
    info: Dict[str, Any],
    cred_fp: str,
    aliases: Optional[tuple] = None,
) -> Tuple[bool, str]:
    try:
        df = v.ler_planilha_aba_df(sid, ws, cred_fp, aliases=aliases)
        n = len(df) if df is not None else 0
        cols = len(df.columns) if df is not None and not df.empty else 0
        return True, f"{n:,} linhas · {cols} colunas"
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"[:200]


def _testar_cache(dataset: str, info: Dict[str, Any]) -> Tuple[bool, str]:
    try:
        df = vc.ler_dataset(dataset, info, prefer_local=False)
        if df is None or df.empty:
            df_local = vc.ler_dataset_local(dataset)
            if df_local is not None and not df_local.empty:
                return True, f"{len(df_local):,} linhas (pickle local; Sheets vazio)"
            return False, "vazio no Sheets e sem pickle local"
        return True, f"{len(df):,} linhas · {len(df.columns)} colunas (Cache Ivan)"
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"[:200]


def main() -> int:
    print("=" * 70)
    print("Verificação de acesso — Acompanhamento de Vendas")
    print("=" * 70)

    info = _carregar_credenciais()
    if not info:
        print("ERRO: credenciais Google não encontradas.")
        print("  Esperado: salesforce/.streamlit/secrets.toml [google_sheets] ou service_account.json")
        return 1

    email = info.get("client_email", "?")
    print(f"Conta de serviço: {email}\n")

    _st.secrets = {
        "connections": {
            "gsheets": {
                **info,
                "spreadsheet_id": v.SPREADSHEET_CONSOLIDADA_ID,
            }
        }
    }
    cred_fp = v._fingerprint_credenciais(info)
    resultados: List[Resultado] = []

    # --- Planilha consolidada ---
    cons = v.SPREADSHEET_CONSOLIDADA_ID
    planilhas_consolidada = [
        ("Metas & Projeção / KPI", "Metas", v.WS_METAS, None),
        ("Metas & Projeção", "Metas Coordenadores", v.WS_METAS_COORD, None),
        ("Metas & Projeção / Dashboard", "Canal (Cópia de Canal)", v.WS_CANAL, v.WS_CANAL_ALIASES),
        ("Dashboard / fallback", "Base Estoque", v.WS_ESTOQUE, None),
        ("fallback vendas", "BD Vendas Completa", v.WS_VENDAS, None),
    ]
    print("[Planilha consolidada]", cons)
    for grupo, nome, ws, aliases in planilhas_consolidada:
        ok, msg = _testar_aba(cons, ws, info, cred_fp, aliases)
        resultados.append((grupo, f"{nome} ({ws})", "OK · " + msg if ok else "FALHA · " + msg))
        print(f"  {'OK' if ok else 'FALHA':5}  {nome}: {msg}")

    # --- Cache Ivan ---
    cache_id = vc.SPREADSHEET_CACHE_ID
    print(f"\n[Cache Ivan]", cache_id)
    manifest = vc.ler_manifest(info)
    ts = manifest.get("atualizado_em", "—")
    print(f"  Manifest: atualizado_em = {ts}")
    for key in (
        "vendas_painel", "vendas_raw", "estoque", "funil_ag", "funil_pastas",
        "cotacoes", "pc_pastas", "pc_tabela",
    ):
        ok, msg = _testar_cache(key, info)
        rotulo = vc.DATASETS.get(key, key)
        resultados.append(("Cache", rotulo, "OK · " + msg if ok else "FALHA · " + msg))
        print(f"  {'OK' if ok else 'FALHA':5}  {rotulo}: {msg}")

    # --- Formulários ---
    print("\n[Formulários Google]")
    forms = [
        ("Feedbacks Comerciais", vfp.SPREADSHEET_FEEDBACK_ID, vfp.WS_FEEDBACK),
        ("Previsão de Vendas", vfp.SPREADSHEET_PREVISAO_ID, vfp.WS_PREVISAO),
    ]
    for grupo, sid, ws in forms:
        ok, msg = _testar_aba(sid, ws, info, cred_fp)
        resultados.append((grupo, ws, "OK · " + msg if ok else "FALHA · " + msg))
        print(f"  {'OK' if ok else 'FALHA':5}  {grupo}: {msg}")

    # --- Lógica por aba (transformações) ---
    print("\n[Simulação de abas — transformações]")
    erros_aba: List[str] = []

    try:
        df_metas_raw = v.ler_planilha_aba_df(cons, v.WS_METAS, cred_fp)
        df_metas = v.preparar_metas_painel(df_metas_raw)
        df_mc, aviso = v.carregar_metas_coordenadores_com_fallback(
            cred_fp, df_metas, 2026, 8,
        )
        vgv = v.soma_meta_vgv_coord(df_mc, 8, 2026, "Desafio")
        qtd = v.soma_meta_coord(df_mc, 8, 2026, "vendas", "Desafio")
        assert vgv > 0, f"meta VGV agosto = {vgv}"
        assert 0 < qtd < 5_000, f"meta qtd absurda = {qtd}"
        print(f"  OK    Metas & Projeção: VGV={vgv:,.0f} qtd={qtd:,.0f} aviso={aviso!r}")
    except Exception as exc:
        erros_aba.append(f"Metas & Projeção: {exc}")
        print(f"  FALHA Metas & Projeção: {exc}")

    try:
        df_vp = vc.ler_dataset("vendas_painel", info, prefer_local=False)
        if df_vp is None or df_vp.empty:
            raise RuntimeError("vendas_painel vazio")
        df_vp = v.assegurar_metricas_vendas(df_vp)
        col_c = v.achar_coluna(df_vp, v.ALIASES_CONTRATO_GERADO) or "Contrato Gerado Em"
        q, vgv_r = v.realizado_vendas_periodo(
            df_vp, col_c, date(2026, 8, 1), date(2026, 8, 21),
        )
        assert q > 0, "sem vendas no período"
        print(f"  OK    Vendas painel: qtd={q:.0f} vgv={vgv_r:,.0f}")
    except Exception as exc:
        erros_aba.append(f"Vendas: {exc}")
        print(f"  FALHA Vendas: {exc}")

    try:
        df_est = vc.ler_dataset("estoque", info, prefer_local=False)
        kpi, _ = v.agregar_estoque(df_est if df_est is not None else __import__("pandas").DataFrame())
        assert kpi.get("unidades", 0) > 0
        print(f"  OK    Estoque KPI: {kpi.get('unidades')} unidades")
    except Exception as exc:
        erros_aba.append(f"Estoque: {exc}")
        print(f"  FALHA Estoque: {exc}")

    try:
        df_fb = vfp.carregar_feedbacks_comerciais(cred_fp)
        assert df_fb is not None and len(df_fb) > 0
        print(f"  OK    Feedbacks: {len(df_fb):,} linhas")
    except Exception as exc:
        erros_aba.append(f"Feedbacks: {exc}")
        print(f"  FALHA Feedbacks: {exc}")

    try:
        df_pr = vfp.carregar_previsao_vendas(cred_fp)
        assert df_pr is not None and len(df_pr) > 0
        prep = vfp.preparar_previsao_df(df_pr)
        print(f"  OK    Previsão: {len(df_pr):,} linhas · prep {len(prep):,}")
    except Exception as exc:
        erros_aba.append(f"Previsão: {exc}")
        print(f"  FALHA Previsão: {exc}")

    try:
        import pandas as pd
        tab = pd.DataFrame({
            "Empreendimento": ["Emp Teste"],
            "Desenquadramento_Pct": [75.0],
            "Meta_Mes": [10.0],
            "Meta_Dia": [1.0],
        })
        prep = v.preparar_df_tabela_exibicao(tab)
        assert prep.columns.is_unique
        sty = v._styler_desenquadramento(prep, v.ROTULOS_COLUNAS_TABELA["Desenquadramento_Pct"])
        print(f"  OK    Tabela analítica (styler): colunas únicas={prep.columns.is_unique}")
    except Exception as exc:
        erros_aba.append(f"Tabela analítica: {exc}")
        print(f"  FALHA Tabela analítica: {exc}")

    # --- Resumo ---
    print("\n" + "=" * 70)
    falhas = [r for r in resultados if r[2].startswith("FALHA")]
    falhas += [f"Aba: {e}" for e in erros_aba]
    if falhas:
        print(f"RESUMO: {len(falhas)} problema(s)")
        for f in falhas:
            print(f"  - {f}")
        return 1
    print("RESUMO: todos os acessos OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
