"""Excel input/output helpers."""
import hashlib
import re
from pathlib import Path

import pandas as pd

CNPJ_COLUMN_CANDIDATES = ["cnpj", "documento", "doc", "cpf/cnpj", "cpf_cnpj"]
CEP_COLUMN_CANDIDATES = ["cep"]
NUMERO_COLUMN_CANDIDATES = ["numero", "número", "num", "nº", "n°", "numero_endereco"]


def _find_column(df, candidates):
    for cand in candidates:
        for c in df.columns:
            if c.lower() == cand:
                return c
    return None


def _digits(value) -> str:
    return re.sub(r"\D", "", str(value)) if value is not None else ""


def load_search_list_from_excel(path: str):
    """Read an .xlsx file and return (records, columns_used).

    Each record is {"cnpj": str, "cep": str, "numero": str}. The site search now
    uses CEP + Número (not Documento/CNPJ) -- CNPJ is kept purely so the output
    spreadsheet can still be cross-referenced back to the original record.

    Requires CEP and Número columns to be present; CNPJ is optional but expected.
    """
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    cnpj_col = _find_column(df, CNPJ_COLUMN_CANDIDATES)
    cep_col = _find_column(df, CEP_COLUMN_CANDIDATES)
    numero_col = _find_column(df, NUMERO_COLUMN_CANDIDATES)

    missing = [name for name, col in [("CEP", cep_col), ("Número", numero_col)] if col is None]
    if missing:
        raise ValueError(
            f"Não encontrei a(s) coluna(s) {', '.join(missing)} na planilha. "
            f"Colunas encontradas: {', '.join(df.columns)}"
        )

    records = []
    for _, row in df.iterrows():
        cep = _digits(row.get(cep_col))
        numero = _digits(row.get(numero_col))
        if not cep and not numero:
            continue
        cnpj = _digits(row.get(cnpj_col)) if cnpj_col else ""
        if not cnpj:
            # Fall back to the CEP+Número pair itself as the identifying key when
            # no CNPJ column exists, so results/checkpoint still have something unique.
            cnpj = f"{cep}-{numero}"
        records.append({"cnpj": cnpj, "cep": cep, "numero": numero})

    columns_used = {"cnpj": cnpj_col, "cep": cep_col, "numero": numero_col}
    return records, columns_used


def file_signature(path: str) -> str:
    """Cheap signature to detect 'same input list' across runs for checkpoint matching."""
    p = Path(path)
    stat = p.stat()
    raw = f"{p.name}:{stat.st_size}:{stat.st_mtime}"
    return hashlib.md5(raw.encode("utf-8")).hexdigest()


def write_output_excel(results: dict, out_path: str):
    """Write four sheets: 'Elegiveis' (kept = no data found), 'Log completo' (all
    statuses), 'Erros' (CNPJs that still errored out after the retry pass) and
    'CEP_ou_Numero_invalido' (rows where the source spreadsheet's CEP/Número was
    missing or rejected by the site -- not a real search outcome, worth fixing
    and re-running separately)."""
    rows = [{"CNPJ": cnpj, "status": status} for cnpj, status in results.items()]
    df_all = pd.DataFrame(rows)
    df_keep = df_all[df_all["status"] == "mantido"][["CNPJ"]]
    df_errors = df_all[df_all["status"] == "erro"][["CNPJ"]]
    df_invalid = df_all[df_all["status"] == "cep_ou_numero_invalido"][["CNPJ"]]

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_keep.to_excel(writer, sheet_name="Elegiveis", index=False)
        df_all.to_excel(writer, sheet_name="Log completo", index=False)
        df_errors.to_excel(writer, sheet_name="Erros", index=False)
        df_invalid.to_excel(writer, sheet_name="CEP_ou_Numero_invalido", index=False)


def write_errors_excel(results: dict, out_path: str) -> bool:
    """Write a standalone spreadsheet with just the CNPJs that errored out.
    Returns False (and writes nothing) if there are no errors."""
    error_cnpjs = [cnpj for cnpj, status in results.items() if status == "erro"]
    if not error_cnpjs:
        return False
    pd.DataFrame({"CNPJ": error_cnpjs}).to_excel(out_path, sheet_name="Erros", index=False)
    return True
