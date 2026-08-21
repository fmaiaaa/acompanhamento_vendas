# -*- coding: utf-8 -*-
"""Republica abas Cache · * no Google Sheets a partir dos pickles locais (com truncamento)."""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(_DIR.parent / "sf-sync-direcional"))

import velocimetro_cache as vc  # noqa: E402
from sf_credentials import aplicar_secrets_toml_local, credenciais_google_sheets  # noqa: E402


def main() -> int:
    aplicar_secrets_toml_local()
    info = credenciais_google_sheets()
    if not info:
        print("Credenciais Google ausentes.")
        return 1

    sid = vc.SPREADSHEET_CACHE_ID
    print(f"Republicando cache → planilha {sid}")

    for key in vc.DATASETS:
        df = vc.ler_dataset_local(key)
        if df is None:
            print(f"  skip {key}: sem pickle")
            continue
        print(f"  {key}: {len(df):,} linhas (pickle)")
        vc.gravar_dataset(key, df, info, sid, gravar_sheets=True)

    manifest = vc.ler_manifest_local()
    if manifest:
        manifest["atualizado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        manifest["sheets_truncado"] = "sim"
        vc.gravar_manifest_local(manifest)
        vc.gravar_manifest_sheets(manifest, info, sid)

    print("Concluído.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
