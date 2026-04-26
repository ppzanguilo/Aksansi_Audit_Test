# =============================================================================
# CAC AUDIT WORKFLOW — Google Colab
# =============================================================================
# Paste each "# ── CELL N" block into a separate Colab cell, or run the whole
# file top-to-bottom via:  !python cac_audit_workflow.py
#
# Required input files (Excel .xlsx):
#   1. Lista de Contas do PGC
#   2. Lista de Contas Sistema CAC
#   3. Integração de Balancete CAC (N e N-1)
#   4. Diário do Razão (N e N-1)
#
# Generated outputs (single Excel workbook with one sheet per result):
#   • CAC_Standard_List          • Balancete_SubContas
#   • Agrupamento_Contas         • Balanco
#   • Demonstracao_Resultados    • Notas
#   • Movimentos_Conciliados     • Movimentos_Nao_Conciliados
# =============================================================================


# ── CELL 1: Install dependencies ─────────────────────────────────────────────
# !pip install -q openpyxl xlsxwriter


# ── CELL 2: Imports ───────────────────────────────────────────────────────────
import io
import os
import pandas as pd
import numpy as np

try:
    from google.colab import files as colab_files
    IN_COLAB = True
except ImportError:
    IN_COLAB = False


# ── CELL 3: File-upload helper ────────────────────────────────────────────────
def upload_file(label: str) -> io.BytesIO:
    """
    In Colab: prompts the user to upload a file and returns its bytes.
    Outside Colab: prompts for a local file path.
    """
    if IN_COLAB:
        print(f"\nUpload: {label}")
        uploaded = colab_files.upload()
        name = list(uploaded.keys())[0]
        return io.BytesIO(uploaded[name])
    else:
        path = input(f"\nLocal path for [{label}]: ").strip()
        return open(path, "rb")


def read_excel(buf, label: str) -> pd.DataFrame:
    df = pd.read_excel(buf, dtype=str)
    df.columns = df.columns.str.strip()
    df.fillna("", inplace=True)
    print(f"  ✓ {label}: {len(df)} rows, columns: {list(df.columns)}")
    return df


# ── CELL 4: Step 1 — Read Lista de Contas do PGC ─────────────────────────────
# Expected columns (flexible — mapped by position keywords):
#   Classe | Desc_Classe | Conta | Desc_Conta | Sub_Conta | Desc_Sub_Conta
#   Tipo_Conta (BS/PL) | Natureza (D/C) | Tipo_Movimento (+/-) |
#   Ativo | Passivo | Gastos | Rendimentos

_PGC_COL_MAP = {
    "Classe": "Classe",
    "Desc Classe": "Desc_Classe",
    "Descrição Classe": "Desc_Classe",
    "Contas": "Conta",
    "Conta": "Conta",
    "Desc Conta": "Desc_Conta",
    "Descrição Conta": "Desc_Conta",
    "Sub Conta": "Sub_Conta",
    "SubConta": "Sub_Conta",
    "Desc Sub Conta": "Desc_Sub_Conta",
    "Descrição Sub Conta": "Desc_Sub_Conta",
    "Tipo de Conta (BS/PL)": "Tipo_Conta",
    "Tipo Conta": "Tipo_Conta",
    "Natureza (D/C)": "Natureza",
    "Natureza": "Natureza",
    "Tipo de Movimento (+/-)": "Tipo_Movimento",
    "Tipo Movimento": "Tipo_Movimento",
    "Ativo": "Ativo",
    "Passivo": "Passivo",
    "Gastos": "Gastos",
    "Rendimentos": "Rendimentos",
}

_PGC_REQUIRED = ["Conta", "Sub_Conta", "Natureza", "Tipo_Conta"]


def read_pgc(buf) -> pd.DataFrame:
    df = read_excel(buf, "Lista de Contas PGC")
    df.rename(columns=_PGC_COL_MAP, inplace=True)
    missing = [c for c in _PGC_REQUIRED if c not in df.columns]
    if missing:
        print(f"  ⚠ PGC — columns not found (check header names): {missing}")
    return df


# ── CELL 5: Step 2 — Read Lista de Contas Sistema CAC ────────────────────────
# Expected columns: Conta | Desc_Conta | Sub_Conta | Desc_Sub_Conta

_CAC_COL_MAP = {
    "Conta": "Conta",
    "Desc Conta": "Desc_Conta",
    "Descrição Conta": "Desc_Conta",
    "Sub Conta": "Sub_Conta",
    "SubConta": "Sub_Conta",
    "Desc Sub Conta": "Desc_Sub_Conta",
    "Descrição Sub Conta": "Desc_Sub_Conta",
}


def read_cac_accounts(buf) -> pd.DataFrame:
    df = read_excel(buf, "Lista de Contas Sistema CAC")
    df.rename(columns=_CAC_COL_MAP, inplace=True)
    return df


# ── CELL 6: Step 3 — Map CAC accounts to PGC → CAC Standard List ─────────────
# Output columns:
#   Classe | Desc_Classe | Conta | Desc_Conta | Sub_Conta | Desc_Sub_Conta
#   Tipo_Conta | Natureza | Tipo_Movimento | Ativo | Passivo | Gastos | Rendimentos

_PGC_CLASSIFICATION_COLS = [
    "Classe", "Desc_Classe", "Tipo_Conta", "Natureza", "Tipo_Movimento",
    "Ativo", "Passivo", "Gastos", "Rendimentos",
]


def map_cac_to_pgc(df_pgc: pd.DataFrame, df_cac: pd.DataFrame) -> pd.DataFrame:
    """
    Left-join CAC accounts onto PGC on (Conta, Sub_Conta) to inherit PGC
    classification columns. CAC descriptions take precedence.
    """
    pgc_class_cols = [c for c in _PGC_CLASSIFICATION_COLS if c in df_pgc.columns]
    pgc_merge = df_pgc[["Conta", "Sub_Conta"] + pgc_class_cols].drop_duplicates(
        subset=["Conta", "Sub_Conta"]
    )

    df_standard = pd.merge(df_cac, pgc_merge, on=["Conta", "Sub_Conta"], how="left")

    unmapped = df_standard[
        df_standard["Tipo_Conta"].isna() | (df_standard["Tipo_Conta"] == "")
    ]
    if not unmapped.empty:
        print(
            f"  ⚠ {len(unmapped)} CAC accounts not matched in PGC — "
            "verify Conta/Sub_Conta codes."
        )
        print(unmapped[["Conta", "Sub_Conta", "Desc_Sub_Conta"]].to_string(index=False))

    print(f"  ✓ CAC Standard List: {len(df_standard)} accounts mapped")
    return df_standard


# ── CELL 7: Step 4 — Read Balancete (integration N and N-1) ──────────────────
# Expected columns (one or two periods):
#   Conta | Sub_Conta | Desc_Conta | Desc_Sub_Conta |
#   Debito_N | Credito_N | Saldo_N | Debito_N1 | Credito_N1 | Saldo_N1
#   (or Debito | Credito | Saldo for single-period files)

_BAL_COL_MAP = {
    "Conta": "Conta",
    "Sub Conta": "Sub_Conta",
    "Sub_Conta": "Sub_Conta",
    "SubConta": "Sub_Conta",
    "Desc Conta": "Desc_Conta",
    "Descrição Conta": "Desc_Conta",
    "Desc Sub Conta": "Desc_Sub_Conta",
    "Descrição Sub Conta": "Desc_Sub_Conta",
    # Period N
    "Débito N": "Debito_N",
    "Debito N": "Debito_N",
    "Crédito N": "Credito_N",
    "Credito N": "Credito_N",
    "Saldo N": "Saldo_N",
    # Period N-1
    "Débito N-1": "Debito_N1",
    "Debito N-1": "Debito_N1",
    "Crédito N-1": "Credito_N1",
    "Credito N-1": "Credito_N1",
    "Saldo N-1": "Saldo_N1",
    # Generic (single period)
    "Débito": "Debito_N",
    "Debito": "Debito_N",
    "Crédito": "Credito_N",
    "Credito": "Credito_N",
    "Saldo": "Saldo_N",
}

_BAL_NUMERIC = ["Debito_N", "Credito_N", "Saldo_N", "Debito_N1", "Credito_N1", "Saldo_N1"]


def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False).str.replace(" ", "", regex=False),
        errors="coerce",
    ).fillna(0.0)


def read_balancete(buf) -> pd.DataFrame:
    df = read_excel(buf, "Balancete CAC")
    df.rename(columns=_BAL_COL_MAP, inplace=True)
    for col in _BAL_NUMERIC:
        if col in df.columns:
            df[col] = _to_numeric(df[col])
    # If Saldo_N is missing, derive it
    if "Saldo_N" not in df.columns and "Debito_N" in df.columns and "Credito_N" in df.columns:
        df["Saldo_N"] = df["Debito_N"] - df["Credito_N"]
    return df


# ── CELL 8: Step 5 — Aggregate Balancete to Sub-Conta level ──────────────────
# Output: Balancete da Entidade da CAC por Sub-Contas
#   Conta | Desc_Conta | Sub_Conta | Desc_Sub_Conta | Debito | Credito | Saldo

def aggregate_balancete(df_balancete: pd.DataFrame) -> pd.DataFrame:
    """
    Collapses any sub-sub-sub rows to (Conta, Sub_Conta) level by summing
    all numeric movement/balance columns.
    """
    group_keys = [c for c in ["Conta", "Sub_Conta", "Desc_Conta", "Desc_Sub_Conta"]
                  if c in df_balancete.columns]
    numeric_cols = [c for c in _BAL_NUMERIC if c in df_balancete.columns]

    df_agg = df_balancete.groupby(group_keys, as_index=False)[numeric_cols].sum()
    print(f"  ✓ Balancete aggregated: {len(df_agg)} sub-conta rows")
    return df_agg


# ── CELL 9: Step 6/7 — Read Diário do Razão (General Ledger) ─────────────────
# Expected columns:
#   Data | Conta | Sub_Conta | Desc_Sub_Conta | Historico |
#   Saldo_Inicial | Debito | Credito | Periodo (N or N-1)

_RAZAO_COL_MAP = {
    "Data": "Data",
    "Conta": "Conta",
    "Sub Conta": "Sub_Conta",
    "SubConta": "Sub_Conta",
    "Desc Sub Conta": "Desc_Sub_Conta",
    "Descrição Sub Conta": "Desc_Sub_Conta",
    "Histórico": "Historico",
    "Historico": "Historico",
    "Saldo Inicial": "Saldo_Inicial",
    "SI": "Saldo_Inicial",
    "Débito": "Debito",
    "Debito": "Debito",
    "Crédito": "Credito",
    "Credito": "Credito",
    "Período": "Periodo",
    "Periodo": "Periodo",
}

_RAZAO_NUMERIC = ["Saldo_Inicial", "Debito", "Credito"]


def read_diario_razao(buf) -> pd.DataFrame:
    df = read_excel(buf, "Diário do Razão")
    df.rename(columns=_RAZAO_COL_MAP, inplace=True)
    for col in _RAZAO_NUMERIC:
        if col in df.columns:
            df[col] = _to_numeric(df[col])
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    # Default period to N if column absent
    if "Periodo" not in df.columns:
        df["Periodo"] = "N"
    return df


# ── CELL 10: Step 8 — Calculate final balances (Si + A − D) ──────────────────
# For each (Conta, Sub_Conta, Periodo) in the GL:
#   Natureza D  →  Saldo_Final = Si + Debito  − Credito
#   Natureza C  →  Saldo_Final = Si + Credito − Debito

def calculate_final_balances(
    df_razao: pd.DataFrame, df_cac_standard: pd.DataFrame
) -> pd.DataFrame:
    """
    Identifies every sub-sub-sub-sub account in the GL, groups by
    (Conta, Sub_Conta, Periodo) and applies the Si + A − D formula.
    """
    natureza_map = (
        df_cac_standard.dropna(subset=["Sub_Conta"])
        .set_index("Sub_Conta")["Natureza"]
        .to_dict()
    )

    group_keys = [c for c in ["Conta", "Sub_Conta", "Periodo"] if c in df_razao.columns]

    agg = df_razao.groupby(group_keys, as_index=False).agg(
        Saldo_Inicial=("Saldo_Inicial", "first") if "Saldo_Inicial" in df_razao.columns
                      else ("Debito", lambda x: 0),
        Total_Debito=("Debito", "sum"),
        Total_Credito=("Credito", "sum"),
    )

    agg["Natureza"] = agg["Sub_Conta"].map(natureza_map).fillna("D")

    # Saldo Final = Si + D − C  (debit-nature)   /   Si + C − D  (credit-nature)
    agg["Saldo_Final"] = np.where(
        agg["Natureza"].str.upper() == "D",
        agg["Saldo_Inicial"] + agg["Total_Debito"] - agg["Total_Credito"],
        agg["Saldo_Inicial"] + agg["Total_Credito"] - agg["Total_Debito"],
    )

    print(f"  ✓ Final balances calculated for {len(agg)} GL account-period rows")
    return agg


# ── CELL 11: Step 9 — Reconciliation ─────────────────────────────────────────
# Verify each sub-sub-sub-sub account in the GL against the Balancete.
# Match on (Conta, Sub_Conta) for period N; check |Saldo_Final − Saldo_N| ≤ ε.
# Also verify the D/C nature is correct.

TOLERANCE = 0.01  # monetary tolerance for reconciliation


def reconcile(
    df_gl_balances: pd.DataFrame,
    df_balancete_agg: pd.DataFrame,
    df_cac_standard: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returns (df_conciliados, df_nao_conciliados).
    """
    natureza_map = (
        df_cac_standard.dropna(subset=["Sub_Conta"])
        .set_index("Sub_Conta")["Natureza"]
        .to_dict()
    )

    # Use period N GL rows
    df_n = df_gl_balances[df_gl_balances["Periodo"].astype(str).str.upper() == "N"].copy()

    bal_cols = ["Conta", "Sub_Conta", "Saldo_N"]
    bal_cols = [c for c in bal_cols if c in df_balancete_agg.columns]
    df_bal = df_balancete_agg[bal_cols].copy()

    merged = pd.merge(
        df_n,
        df_bal,
        on=["Conta", "Sub_Conta"],
        how="outer",
        indicator=True,
    )

    merged["Saldo_Final"] = merged.get("Saldo_Final", pd.Series(dtype=float)).fillna(0)
    merged["Saldo_N"] = merged.get("Saldo_N", pd.Series(dtype=float)).fillna(0)
    merged["Diferenca"] = (merged["Saldo_Final"] - merged["Saldo_N"]).abs()
    merged["Natureza_Esperada"] = merged["Sub_Conta"].map(natureza_map)

    is_conciliado = (merged["_merge"] == "both") & (merged["Diferenca"] <= TOLERANCE)

    df_conciliados = merged[is_conciliado].drop(columns=["_merge"])
    df_nao_conciliados = merged[~is_conciliado].drop(columns=["_merge"])

    # Flag reason for non-reconciliation
    reasons = []
    for _, row in df_nao_conciliados.iterrows():
        if row.get("_merge", "both") == "left_only":
            reasons.append("Conta não consta no Balancete")
        elif row.get("_merge", "both") == "right_only":
            reasons.append("Conta no Balancete sem movimentos no Razão")
        else:
            reasons.append(f"Diferença de {row['Diferenca']:.2f}")
    df_nao_conciliados = df_nao_conciliados.copy()
    df_nao_conciliados["Motivo"] = reasons

    print(
        f"  ✓ Conciliados: {len(df_conciliados)} | "
        f"Não Conciliados: {len(df_nao_conciliados)}"
    )
    return df_conciliados, df_nao_conciliados


# ── CELL 12: Generate financial statements ────────────────────────────────────

def build_balancete_report(
    df_bal_agg: pd.DataFrame, df_cac_standard: pd.DataFrame
) -> pd.DataFrame:
    """
    Balancete da Entidade da CAC por Sub-Contas
    Enrich aggregated balancete with PGC classification columns.
    """
    class_cols = [c for c in _PGC_CLASSIFICATION_COLS if c in df_cac_standard.columns]
    df_class = df_cac_standard[["Conta", "Sub_Conta"] + class_cols].drop_duplicates(
        subset=["Conta", "Sub_Conta"]
    )
    report = pd.merge(df_bal_agg, df_class, on=["Conta", "Sub_Conta"], how="left")
    return report


def build_agrupamento(df_report: pd.DataFrame) -> pd.DataFrame:
    """Agrupamento das sub-contas por contas (sum numeric cols at Conta level)."""
    numeric = [c for c in df_report.columns if c in _BAL_NUMERIC]
    group_keys = [c for c in ["Conta", "Desc_Conta", "Tipo_Conta"] if c in df_report.columns]
    return df_report.groupby(group_keys, as_index=False)[numeric].sum()


def build_balanco(df_report: pd.DataFrame) -> pd.DataFrame:
    """Balanço — BS accounts only, grouped by Conta."""
    df_bs = df_report[df_report.get("Tipo_Conta", pd.Series()).str.upper() == "BS"]
    numeric = [c for c in df_bs.columns if c in _BAL_NUMERIC]
    group_keys = [c for c in [
        "Conta", "Desc_Conta", "Sub_Conta", "Desc_Sub_Conta",
        "Tipo_Conta", "Natureza", "Tipo_Movimento"
    ] if c in df_bs.columns]
    return df_bs.groupby(group_keys, as_index=False)[numeric].sum()


def build_demonstracao_resultados(df_report: pd.DataFrame) -> pd.DataFrame:
    """Demonstração de Resultados — PL accounts only, grouped by Conta."""
    df_pl = df_report[df_report.get("Tipo_Conta", pd.Series()).str.upper() == "PL"]
    numeric = [c for c in df_pl.columns if c in _BAL_NUMERIC]
    group_keys = [c for c in ["Conta", "Desc_Conta", "Tipo_Conta", "Natureza"]
                  if c in df_pl.columns]
    return df_pl.groupby(group_keys, as_index=False)[numeric].sum()


def build_notas(df_report: pd.DataFrame) -> pd.DataFrame:
    """
    Notas da entidade da CAC por Sub Conta.
    Rubrica = Desc_Sub_Conta; Valor_Bruto = Saldo_N (period N balance).
    """
    saldo_cols = [c for c in df_report.columns if "Saldo" in c]
    keep = ["Sub_Conta", "Desc_Sub_Conta", "Tipo_Conta", "Natureza"] + saldo_cols
    keep = [c for c in keep if c in df_report.columns]

    df_notas = df_report[keep].copy()
    df_notas.rename(columns={"Desc_Sub_Conta": "Rubrica"}, inplace=True)
    if "Saldo_N" in df_notas.columns:
        df_notas.rename(columns={"Saldo_N": "Valor_Bruto"}, inplace=True)

    # Classify by account type
    tipo_map = {"BS": "Balanço", "PL": "Demonstração de Resultados"}
    if "Tipo_Conta" in df_notas.columns:
        df_notas["Classificacao"] = df_notas["Tipo_Conta"].str.upper().map(tipo_map).fillna("Outros")

    return df_notas.sort_values(["Tipo_Conta", "Sub_Conta"], ignore_index=True)


# ── CELL 13: Export to Excel ──────────────────────────────────────────────────

def export_to_excel(sheets: dict[str, pd.DataFrame], filename: str = "CAC_Audit_Output.xlsx"):
    """Write all result DataFrames to a multi-sheet Excel workbook."""
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "bg_color": "#4472C4", "font_color": "#FFFFFF"})

        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
            ws = writer.sheets[safe_name]
            for i, col in enumerate(df.columns):
                width = max(df[col].astype(str).map(len).max() if len(df) else 0, len(col)) + 2
                ws.set_column(i, i, min(width, 50))
                ws.write(0, i, col, header_fmt)

    print(f"  ✓ Saved: {filename}")
    if IN_COLAB:
        colab_files.download(filename)


# ── CELL 14: MAIN — orchestrate the full workflow ─────────────────────────────

def run_workflow():
    print("=" * 64)
    print("  CAC AUDIT WORKFLOW")
    print("=" * 64)

    # ── 1. Ingest inputs ────────────────────────────────────────────
    print("\n[1/4] Loading input files…")
    df_pgc = read_pgc(upload_file("Lista de Contas do PGC (Excel)"))
    df_cac_accounts = read_cac_accounts(upload_file("Lista de Contas Sistema CAC (Excel)"))
    df_balancete_raw = read_balancete(upload_file("Integração de Balancete CAC N e N-1 (Excel)"))
    df_razao = read_diario_razao(upload_file("Diário do Razão N e N-1 (Excel)"))

    # ── 2. Map CAC → PGC Standard ───────────────────────────────────
    print("\n[2/4] Building CAC Standard List…")
    df_cac_standard = map_cac_to_pgc(df_pgc, df_cac_accounts)

    # ── 3. Process Balancete ────────────────────────────────────────
    print("\n[3/4] Processing Balancete…")
    df_bal_agg = aggregate_balancete(df_balancete_raw)
    df_bal_report = build_balancete_report(df_bal_agg, df_cac_standard)

    # ── 4. GL balances + reconciliation ────────────────────────────
    print("\n[4/4] Calculating GL balances and reconciling…")
    df_gl_balances = calculate_final_balances(df_razao, df_cac_standard)
    df_conciliados, df_nao_conciliados = reconcile(df_gl_balances, df_bal_agg, df_cac_standard)

    # ── 5. Financial statements ─────────────────────────────────────
    print("\n[Reports] Generating financial statements…")
    df_agrup = build_agrupamento(df_bal_report)
    df_balanco = build_balanco(df_bal_report)
    df_dr = build_demonstracao_resultados(df_bal_report)
    df_notas = build_notas(df_bal_report)

    # ── 6. Export ───────────────────────────────────────────────────
    print("\n[Export] Writing output workbook…")
    export_to_excel(
        {
            "CAC_Standard_List": df_cac_standard,
            "Balancete_SubContas": df_bal_report,
            "Agrupamento_Contas": df_agrup,
            "Balanco": df_balanco,
            "Demonstracao_Resultados": df_dr,
            "Notas": df_notas,
            "Movimentos_Conciliados": df_conciliados,
            "Movimentos_Nao_Conciliados": df_nao_conciliados,
        }
    )

    print("\n" + "=" * 64)
    print("  WORKFLOW COMPLETE")
    print("=" * 64)

    return {
        "cac_standard": df_cac_standard,
        "balancete": df_bal_report,
        "agrupamento": df_agrup,
        "balanco": df_balanco,
        "demonstracao_resultados": df_dr,
        "notas": df_notas,
        "conciliados": df_conciliados,
        "nao_conciliados": df_nao_conciliados,
    }


# ── CELL 15: Entry point ──────────────────────────────────────────────────────
if __name__ == "__main__":
    results = run_workflow()
