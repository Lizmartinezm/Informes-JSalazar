from __future__ import annotations

from io import BytesIO

import pandas as pd

from utils.tlg_data_cleaning import load_tlg_trial_balance
from utils.tlg_financial_statements import prepare_tlg_detail


MONTH_NUMBERS = {
    "Enero": 1,
    "Febrero": 2,
    "Marzo": 3,
    "Abril": 4,
    "Mayo": 5,
    "Junio": 6,
    "Julio": 7,
    "Agosto": 8,
    "Septiembre": 9,
    "Octubre": 10,
    "Noviembre": 11,
    "Diciembre": 12,
}


def period_label(metadata: dict[str, str | None]) -> str:
    month = metadata.get("mes") or "Periodo"
    year = metadata.get("anio") or ""
    return f"{month} {year}".strip()


def build_monthly_balance(detail: pd.DataFrame, label: str) -> pd.DataFrame:
    balance = detail[detail["CLASE"].isin(["1", "2", "3"])].copy()
    balance[label] = balance["SALDO_FINAL"]
    return balance[["CODIGO_CUENTA", "NOMBRE_CUENTA", "CLASE", "CLASIFICACION", label]]


def build_monthly_income_statement(detail: pd.DataFrame, label: str) -> pd.DataFrame:
    pyg = detail[detail["CLASE"].isin(["4", "5", "6", "7"])].copy()
    pyg[label] = pyg["MOVIMIENTO_DEBITO"] - pyg["MOVIMIENTO_CREDITO"]
    pyg.loc[pyg["CLASE"] == "4", label] = (
        pyg.loc[pyg["CLASE"] == "4", "MOVIMIENTO_CREDITO"]
        - pyg.loc[pyg["CLASE"] == "4", "MOVIMIENTO_DEBITO"]
    )
    return pyg[["CODIGO_CUENTA", "NOMBRE_CUENTA", "CLASE", "CLASIFICACION", label]]


def _merge_month(base: pd.DataFrame, current: pd.DataFrame, label: str) -> pd.DataFrame:
    keys = ["CODIGO_CUENTA", "NOMBRE_CUENTA", "CLASE", "CLASIFICACION"]
    if base.empty:
        return current.sort_values("CODIGO_CUENTA").reset_index(drop=True)
    base = base.copy()
    current = current.copy()
    for key in keys:
        base[key] = base[key].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        current[key] = current[key].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    if label in base.columns:
        base = base.drop(columns=[label])
    merged = base.merge(current, on=keys, how="outer")
    value_cols = [col for col in merged.columns if col not in keys]
    merged[value_cols] = merged[value_cols].fillna(0)
    return merged.sort_values("CODIGO_CUENTA").reset_index(drop=True)


def _read_generated_previous(file) -> tuple[pd.DataFrame, pd.DataFrame] | None:
    file.seek(0)
    excel = pd.ExcelFile(file, engine="openpyxl")
    if "Balance Mensual" not in excel.sheet_names or "PYG Mensual" not in excel.sheet_names:
        return None
    balance = pd.read_excel(file, sheet_name="Balance Mensual", dtype={"CODIGO_CUENTA": str}, engine="openpyxl")
    file.seek(0)
    pyg = pd.read_excel(file, sheet_name="PYG Mensual", dtype={"CODIGO_CUENTA": str}, engine="openpyxl")
    return balance, pyg


def load_previous_accumulated(file) -> tuple[pd.DataFrame, pd.DataFrame, str]:
    generated = _read_generated_previous(file)
    if generated is not None:
        return generated[0], generated[1], "Version mensualizada anterior"

    file.seek(0)
    previous_df, previous_metadata = load_tlg_trial_balance(file)
    previous_detail = prepare_tlg_detail(previous_df)
    label = period_label(previous_metadata)
    return (
        build_monthly_balance(previous_detail, label),
        build_monthly_income_statement(previous_detail, label),
        f"Balance acumulado usado como base: {label}",
    )


def build_monthly_reports(
    monthly_file,
    previous_file=None,
) -> dict[str, object]:
    monthly_file.seek(0)
    raw_df, metadata = load_tlg_trial_balance(monthly_file)
    detail = prepare_tlg_detail(raw_df)
    label = period_label(metadata)
    current_balance = build_monthly_balance(detail, label)
    current_pyg = build_monthly_income_statement(detail, label)

    base_balance = pd.DataFrame()
    base_pyg = pd.DataFrame()
    previous_message = "No se cargo una version acumulada anterior."
    if previous_file is not None:
        base_balance, base_pyg, previous_message = load_previous_accumulated(previous_file)

    balance = _merge_month(base_balance, current_balance, label)
    pyg = _merge_month(base_pyg, current_pyg, label)

    income = current_pyg.loc[current_pyg["CLASE"] == "4", label].sum()
    expenses = current_pyg.loc[current_pyg["CLASE"].isin(["5", "6", "7"]), label].sum()
    result = income - expenses
    assets = current_balance.loc[current_balance["CLASE"] == "1", label].sum()
    liabilities_signed = current_balance.loc[current_balance["CLASE"] == "2", label].sum()
    equity_signed = current_balance.loc[current_balance["CLASE"] == "3", label].sum()
    liabilities = abs(liabilities_signed)
    equity = abs(equity_signed)
    balance_difference = detail["SALDO_FINAL"].sum()

    summary = pd.DataFrame(
        {
            "Indicador": [
                "Total activo",
                "Total pasivo",
                "Total patrimonio",
                "Ingresos del mes",
                "Costos y gastos del mes",
                "Resultado del mes",
                "Diferencia balance",
            ],
            label: [assets, liabilities, equity, income, expenses, result, balance_difference],
        }
    )

    return {
        "metadata": metadata,
        "label": label,
        "detail": detail,
        "balance": balance,
        "pyg": pyg,
        "summary": summary,
        "previous_message": previous_message,
    }


def export_monthly_reports(report: dict[str, object]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        report["summary"].to_excel(writer, sheet_name="Resumen Mensual", index=False)
        report["balance"].to_excel(writer, sheet_name="Balance Mensual", index=False)
        report["pyg"].to_excel(writer, sheet_name="PYG Mensual", index=False)
        report["detail"].to_excel(writer, sheet_name=f"Detalle {report['label']}"[:31], index=False)
        pd.DataFrame([report["metadata"]]).to_excel(writer, sheet_name="Datos Empresa", index=False)

        workbook = writer.book
        money_format = workbook.add_format({"num_format": "$#,##0.00"})
        header_format = workbook.add_format({"bold": True, "bg_color": "#0F766E", "font_color": "#FFFFFF"})
        for sheet_name, worksheet in writer.sheets.items():
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, 0, max(0, worksheet.dim_colmax))
            worksheet.set_row(0, 22, header_format)
            worksheet.set_column(0, 0, 18)
            worksheet.set_column(1, 1, 34)
            worksheet.set_column(2, 3, 18)
            worksheet.set_column(4, max(4, worksheet.dim_colmax), 18, money_format)
    return output.getvalue()
