# -*- coding: utf-8 -*-
"""Cache local (pickle) + abas Google Sheets com dados prontos para exibição no painel."""
from __future__ import annotations

import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

_DIR = Path(__file__).resolve().parent
_CACHE_DIR = _DIR / "data" / "cache_velocimetro"

SPREADSHEET_CONSOLIDADA_ID = "1wpuNQvksot9CLhGgQRe7JlyDeRISEh_sc3-6VRDyQYk"
# Abas «Cache · *» ficam na planilha Ivan (BD Vendas atinge limite de 10M células)
SPREADSHEET_CACHE_ID = "1cExor3vbUZWEeu8iaX7M0oNpjY3b-MssJBwkVhGKt5c"

WS_MANIFEST = "Cache · Manifest"
WS_VENDAS_PAINEL = "Cache · Vendas Painel"
WS_VENDAS_RAW = "Cache · Vendas Raw"
WS_ESTOQUE = "Cache · Estoque"
WS_FUNIL_AG = "Cache · Funil Ag"
WS_FUNIL_PASTAS = "Cache · Funil Pastas"
WS_FUNIL_HIST_AG = "Cache · Funil Hist Ag"
WS_FUNIL_HIST_PASTAS = "Cache · Funil Hist Pastas"
WS_COTACOES = "Cache · Cotações"
WS_PC_PASTAS = "Cache · PC Pastas"
WS_PC_TABELA = "Cache · PC Tabela"
WS_FUNIL_EMP_AG = "Cache · Funil Emp Ag"
WS_FUNIL_EMP_PASTAS = "Cache · Funil Emp Pastas"
WS_FUNIL_EMP_OPPS = "Cache · Funil Emp Opps"
WS_FUNIL_EMP_VEN = "Cache · Funil Emp Vendas"
WS_FUNIL_EMP_EST = "Cache · Funil Emp Estoque"

DATASETS: Dict[str, str] = {
    "vendas_painel": WS_VENDAS_PAINEL,
    "vendas_raw": WS_VENDAS_RAW,
    "estoque": WS_ESTOQUE,
    "funil_ag": WS_FUNIL_AG,
    "funil_pastas": WS_FUNIL_PASTAS,
    "funil_hist_ag": WS_FUNIL_HIST_AG,
    "funil_hist_pastas": WS_FUNIL_HIST_PASTAS,
    "cotacoes": WS_COTACOES,
    "pc_pastas": WS_PC_PASTAS,
    "pc_tabela": WS_PC_TABELA,
    "funil_emp_ag": WS_FUNIL_EMP_AG,
    "funil_emp_pastas": WS_FUNIL_EMP_PASTAS,
    "funil_emp_opps": WS_FUNIL_EMP_OPPS,
    "funil_emp_ven": WS_FUNIL_EMP_VEN,
    "funil_emp_est": WS_FUNIL_EMP_EST,
}


def cache_dir() -> Path:
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return _CACHE_DIR


def _pickle_path(dataset: str) -> Path:
    return cache_dir() / f"{dataset}.pkl"


def _manifest_path() -> Path:
    return cache_dir() / "manifest.json"


def _df_para_matriz(df: pd.DataFrame) -> List[List[Any]]:
    if df is None or df.empty:
        return [["(vazio)"]]
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d %H:%M:%S")
    out = out.where(pd.notna(out), "")
    header = [str(c) for c in out.columns]
    rows = [[str(v) if v != "" else "" for v in row] for row in out.astype(str).values.tolist()]
    return [header] + rows


def gravar_aba_gsheets(
    service_account_info: Dict[str, Any],
    spreadsheet_id: str,
    worksheet: str,
    df: pd.DataFrame,
) -> None:
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id.strip())
    nome = worksheet.strip()
    try:
        ws = sh.worksheet(nome)
    except gspread.WorksheetNotFound:
        nrows = max(len(df) + 5, 100) if df is not None and not df.empty else 100
        ncols = max(len(df.columns) + 2, 10) if df is not None and not df.empty else 10
        ws = sh.add_worksheet(title=nome[:99], rows=nrows, cols=ncols)
    matriz = _df_para_matriz(df if df is not None else pd.DataFrame())
    ws.clear()
    if matriz:
        ws.update(values=matriz, range_name="A1", value_input_option="USER_ENTERED")


def ler_manifest_local() -> Dict[str, Any]:
    p = _manifest_path()
    if not p.is_file():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def gravar_manifest_local(manifest: Dict[str, Any]) -> None:
    _manifest_path().write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def ler_manifest_sheets(
    service_account_info: Dict[str, Any],
    spreadsheet_id: str = SPREADSHEET_CACHE_ID,
) -> Dict[str, Any]:
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id.strip())
    try:
        ws = sh.worksheet(WS_MANIFEST)
    except Exception:
        return {}
    vals = ws.get_all_values()
    if not vals:
        return {}
    if len(vals) == 1 and vals[0] and str(vals[0][0]).startswith("{"):
        try:
            return json.loads(vals[0][0])
        except Exception:
            pass
    out: Dict[str, Any] = {}
    for row in vals:
        if len(row) >= 2 and row[0]:
            out[str(row[0]).strip()] = row[1]
    return out


def gravar_manifest_sheets(
    manifest: Dict[str, Any],
    service_account_info: Dict[str, Any],
    spreadsheet_id: str = SPREADSHEET_CACHE_ID,
) -> None:
    rows = [[k, str(v)] for k, v in manifest.items()]
    df = pd.DataFrame(rows, columns=["chave", "valor"])
    gravar_aba_gsheets(service_account_info, spreadsheet_id, WS_MANIFEST, df)


def ler_manifest(
    service_account_info: Optional[Dict[str, Any]] = None,
    spreadsheet_id: str = SPREADSHEET_CACHE_ID,
) -> Dict[str, Any]:
    """Preferência: manifest remoto (Actions) → manifest local (pickle)."""
    local = ler_manifest_local()
    if service_account_info:
        remote = ler_manifest_sheets(service_account_info, spreadsheet_id)
        if remote.get("atualizado_em"):
            gravar_manifest_local(remote)
            return remote
    return local


def manifest_valido(manifest: Dict[str, Any], max_horas: float = 48.0) -> bool:
    ts = manifest.get("atualizado_em")
    if not ts:
        return False
    try:
        dt = datetime.strptime(str(ts), "%Y-%m-%d %H:%M:%S")
        return (datetime.now() - dt).total_seconds() < max_horas * 3600
    except Exception:
        return bool(ts)


def gravar_dataset(
    dataset: str,
    df: pd.DataFrame,
    service_account_info: Optional[Dict[str, Any]] = None,
    spreadsheet_id: str = SPREADSHEET_CACHE_ID,
    gravar_sheets: bool = True,
) -> None:
    df = df if df is not None else pd.DataFrame()
    df.to_pickle(_pickle_path(dataset))
    ws = DATASETS.get(dataset)
    if gravar_sheets and ws and service_account_info:
        gravar_aba_gsheets(service_account_info, spreadsheet_id, ws, df)


def ler_dataset_local(dataset: str) -> Optional[pd.DataFrame]:
    p = _pickle_path(dataset)
    if not p.is_file():
        return None
    try:
        df = pd.read_pickle(p)
        return df if isinstance(df, pd.DataFrame) else None
    except Exception:
        return None


def ler_dataset_sheets(
    dataset: str,
    service_account_info: Dict[str, Any],
    spreadsheet_id: str = SPREADSHEET_CACHE_ID,
) -> Optional[pd.DataFrame]:
    ws = DATASETS.get(dataset)
    if not ws:
        return None
    # Import tardio para evitar dependência circular com velocimetro
    if "velocimetro" in sys.modules:
        v = sys.modules["velocimetro"]
        return v.ler_aba_gsheets(service_account_info, spreadsheet_id, ws)
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id.strip())
    try:
        ws_obj = sh.worksheet(ws)
    except Exception:
        return None
    vals = ws_obj.get_all_values()
    if not vals or len(vals) < 2:
        return None
    header = [str(h) for h in vals[0]]
    body = vals[1:]
    w = len(header)
    norm: List[List[str]] = []
    for r in body:
        cells = [str(c) for c in r]
        if len(cells) < w:
            cells = cells + [""] * (w - len(cells))
        else:
            cells = cells[:w]
        norm.append(cells)
    df = pd.DataFrame(norm, columns=header)
    if len(df) == 1 and "(vazio)" in str(df.iloc[0, 0]):
        return pd.DataFrame()
    return df


def ler_dataset(
    dataset: str,
    service_account_info: Optional[Dict[str, Any]] = None,
    spreadsheet_id: str = SPREADSHEET_CACHE_ID,
    prefer_local: bool = True,
) -> Optional[pd.DataFrame]:
    if prefer_local:
        local = ler_dataset_local(dataset)
        if local is not None and not local.empty:
            return local
    if service_account_info:
        remote = ler_dataset_sheets(dataset, service_account_info, spreadsheet_id)
        if remote is not None:
            if not remote.empty:
                remote.to_pickle(_pickle_path(dataset))
            return remote
    if not prefer_local:
        return ler_dataset_local(dataset)
    return None


def limpar_cache_local() -> None:
    for p in cache_dir().glob("*.pkl"):
        try:
            p.unlink()
        except Exception:
            pass
    mp = _manifest_path()
    if mp.is_file():
        try:
            mp.unlink()
        except Exception:
            pass
