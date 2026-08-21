# -*- coding: utf-8 -*-
"""
Pré-computa bases do painel (pós-processamento para exibição) e grava cache
local (pickle) + abas «Cache · *» na planilha consolidada.

Uso:
  python velocimetro_precompute.py
  python velocimetro_precompute.py --sem-sheets   # só pickle local
"""
from __future__ import annotations

import argparse
import json
import sys
import time
import types
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd

_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(_DIR))

# Mock Streamlit antes de importar velocimetro
_st = types.ModuleType("streamlit")


def _noop(*_a, **_k):
    return None


def _identity_decorator(*_a, **_k):
    def _wrap(fn):
        return fn
    return _wrap


_st.cache_data = _identity_decorator
_st.cache_resource = _identity_decorator
_st.dialog = _identity_decorator
_st.secrets = {"connections": {"gsheets": {}}}
sys.modules["streamlit"] = _st

import velocimetro as v  # noqa: E402
import velocimetro_cache as vc  # noqa: E402


def _credenciais_google() -> Optional[Dict[str, Any]]:
    sf_dir = _DIR.parent / "salesforce"
    if str(sf_dir) not in sys.path:
        sys.path.insert(0, str(sf_dir))
    try:
        import sf_credentials as cred  # noqa: WPS433

        cred.aplicar_secrets_toml_local()
        c = cred.credenciais_google_sheets()
        if c:
            return c
    except Exception:
        pass
    raw = v._secrets_connections_gsheets()
    return v.montar_service_account_info(raw)


def _conectar_sf():
    sf_dir = _DIR.parent / "salesforce"
    if str(sf_dir) not in sys.path:
        sys.path.insert(0, str(sf_dir))
    try:
        import sf_credentials as cred  # noqa: WPS433

        cred.aplicar_secrets_toml_local()
        sf, err = cred.conectar_salesforce_diagnostico()
        if sf:
            return sf, None
        return None, err
    except Exception as exc:
        return v.conectar_salesforce_app()


def _sf_live(sf) -> None:
    """Injeta cliente SF no módulo velocimetro (substitui cache Streamlit)."""
    v._cliente_salesforce_cache = lambda: sf  # type: ignore[attr-defined]


def executar_precompute(gravar_sheets: bool = True) -> int:
    t0 = time.perf_counter()
    info = _credenciais_google()
    if not info:
        print("Credenciais Google ausentes.")
        return 1

    sf, err = _conectar_sf()
    if sf is None:
        print(f"Salesforce indisponível: {err}")
        return 1
    _sf_live(sf)

    sid_metas = v.SPREADSHEET_ID
    sid_cache = vc.SPREADSHEET_CACHE_ID

    print("Lendo metas…")
    df_metas_raw = v.ler_aba_gsheets(info, sid_metas, v.WS_METAS)
    df_metas = v.preparar_metas_painel(df_metas_raw)

    print("Extraindo vendas SF…")
    pacote_v = v._carregar_vendas_painel_sf_live()
    df_vendas_raw = pacote_v["vendas"]
    prep = v.preparar_vendas_painel(df_vendas_raw, df_metas)
    df_vendas_painel = prep["df_vendas_painel"]

    print("Extraindo estoque, funil, cotações, poder de compra…")
    df_estoque = v._carregar_estoque_painel_sf_live()
    funil = v._carregar_funil_painel_sf_live()
    funil_hist = v._carregar_funil_historico_painel_sf_live()
    df_cot = v._carregar_cotacoes_painel_sf_live()
    pacote_pc = v._carregar_pacote_poder_compra_sf_live()
    funil_emp = v._carregar_funil_empreendimento_sf_live()

    print("Contagem total de unidades por empreendimento…")
    total_unidades_map = v._sf_soql_contagem_unidades_total_por_emp(sf)

    manifest: Dict[str, Any] = {
        "atualizado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "col_contrato_gerado": prep.get("col_contrato_gerado") or "",
        "col_canal": prep.get("col_canal") or "",
        "col_data_venda": prep.get("col_data_venda") or "",
        "vendas_painel_linhas": len(df_vendas_painel),
        "vendas_raw_linhas": len(df_vendas_raw),
        "estoque_linhas": len(df_estoque),
        "funil_ag_linhas": len(funil.get("agendamentos", pd.DataFrame())),
        "funil_pastas_linhas": len(funil.get("pastas", pd.DataFrame())),
        "total_unidades_emp_json": json.dumps(total_unidades_map, ensure_ascii=False),
        "duracao_s": round(time.perf_counter() - t0, 1),
    }

    print("Gravando cache local…")
    vc.gravar_dataset("vendas_painel", df_vendas_painel, info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("vendas_raw", df_vendas_raw, info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("estoque", df_estoque, info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_ag", funil.get("agendamentos", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_pastas", funil.get("pastas", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_hist_ag", funil_hist.get("agendamentos", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_hist_pastas", funil_hist.get("pastas", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("cotacoes", df_cot, info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("pc_pastas", pacote_pc.get("pastas_aprovadas", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("pc_tabela", pacote_pc.get("tabela_comprometimento", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_emp_ag", funil_emp.get("agendamentos", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_emp_pastas", funil_emp.get("pastas", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_emp_opps", funil_emp.get("oportunidades", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_emp_ven", funil_emp.get("vendas", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_dataset("funil_emp_est", funil_emp.get("estoque", pd.DataFrame()), info, sid_cache, gravar_sheets=False)
    vc.gravar_manifest_local(manifest)

    if gravar_sheets:
        print("Publicando abas Cache · * na planilha consolidada…")
        for key in vc.DATASETS:
            df = vc.ler_dataset_local(key)
            if df is not None:
                vc.gravar_dataset(key, df, info, sid_cache, gravar_sheets=True)
        vc.gravar_manifest_sheets(manifest, info, sid_cache)

    print(f"Concluído em {time.perf_counter() - t0:.1f}s · vendas painel={len(df_vendas_painel):,} linhas")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Pré-computa cache de exibição do velocímetro")
    parser.add_argument("--sem-sheets", action="store_true", help="Grava só pickle local")
    args = parser.parse_args()
    return executar_precompute(gravar_sheets=not args.sem_sheets)


if __name__ == "__main__":
    raise SystemExit(main())
