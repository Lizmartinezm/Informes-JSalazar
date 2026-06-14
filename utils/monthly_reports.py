from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from utils.tlg_data_cleaning import MONTHS, load_tlg_trial_balance
from utils.tlg_financial_statements import prepare_tlg_detail


TEMPLATE_PATH = (
    Path(__file__).resolve().parents[1]
    / "Informes_mensualizados_template.xlsx"
)


def _period_from_metadata(metadata: dict[str, str | None]) -> tuple[int, int]:
    month_name = str(metadata.get("mes") or "").strip().lower()
    year_text = str(metadata.get("anio") or "").strip()
    month = MONTHS.get(month_name)
    if month is None or not year_text.isdigit():
        raise ValueError(
            "No fue posible identificar el mes y el año del balance. "
            "Verifica que el encabezado indique el periodo del informe."
        )
    return int(year_text), month


def _account4_values(detail: pd.DataFrame) -> tuple[dict[str, float], dict[str, float]]:
    data = detail.copy()
    data["CUENTA_4"] = (
        data["CODIGO_CUENTA"]
        .fillna("")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\D", "", regex=True)
        .str[:4]
    )
    data = data[data["CUENTA_4"].str.len() == 4].copy()

    balance = data.groupby("CUENTA_4")["SALDO_FINAL"].sum().to_dict()
    data["VALOR_PYG"] = data["MOVIMIENTO_DEBITO"] - data["MOVIMIENTO_CREDITO"]
    income = data["CLASE"] == "4"
    data.loc[income, "VALOR_PYG"] = (
        data.loc[income, "MOVIMIENTO_CREDITO"]
        - data.loc[income, "MOVIMIENTO_DEBITO"]
    )
    pyg = data.groupby("CUENTA_4")["VALOR_PYG"].sum().to_dict()
    return balance, pyg


def _find_bce_month_column(worksheet, year: int, month: int) -> int:
    for column in range(1, worksheet.max_column + 1):
        value = worksheet.cell(2, column).value
        if isinstance(value, (datetime, pd.Timestamp)):
            if value.year == year and value.month == month:
                return column
    raise ValueError(
        f"La plantilla no tiene una columna disponible para {month:02d}/{year} en la hoja BCE."
    )


def _find_pyg_month_column(worksheet, bce_column: int, year: int, month: int) -> int:
    bce_reference = f"BCE!{get_column_letter(bce_column)}2".upper()
    for column in range(1, worksheet.max_column + 1):
        value = worksheet.cell(3, column).value
        if isinstance(value, str) and bce_reference in value.replace("$", "").upper():
            return column

    known_year_starts = {2024: 5, 2025: 18, 2026: 31}
    if year in known_year_starts:
        return known_year_starts[year] + month - 1
    raise ValueError(
        f"La plantilla no tiene una columna disponible para {month:02d}/{year} en la hoja P Y G."
    )


def _write_mapped_accounts(worksheet, column: int, values: dict[str, float]) -> int:
    updated = 0
    for row in range(1, worksheet.max_row + 1):
        raw_code = worksheet.cell(row, 3).value
        if raw_code is None:
            continue
        code = str(raw_code).replace(".0", "").strip()
        if not code.isdigit() or len(code) != 4:
            continue
        worksheet.cell(row, column).value = float(values.get(code, 0.0))
        updated += 1
    return updated


def _load_base_workbook(previous_file: BinaryIO | None):
    if previous_file is not None:
        previous_file.seek(0)
        source = BytesIO(previous_file.read())
        source_name = "Informe mensualizado anterior"
    else:
        if not TEMPLATE_PATH.exists():
            raise FileNotFoundError(
                "No se encontró la plantilla base de Informes mensualizados."
            )
        source = TEMPLATE_PATH
        source_name = "Plantilla base"

    workbook = load_workbook(source, data_only=False, keep_links=True)
    missing = {"BCE", "P Y G"} - set(workbook.sheetnames)
    if missing:
        raise ValueError(
            "El archivo acumulado no corresponde a la plantilla de informes mensualizados. "
            "Debe contener las hojas BCE y P Y G."
        )
    return workbook, source_name


def build_monthly_reports(
    monthly_files: list[BinaryIO],
    previous_file: BinaryIO | None = None,
) -> dict[str, object]:
    if not monthly_files:
        raise ValueError("Debes cargar al menos un balance de prueba por tercero.")

    periods: list[dict[str, object]] = []
    seen_periods: set[tuple[int, int]] = set()
    companies: set[str] = set()

    for uploaded_file in monthly_files:
        uploaded_file.seek(0)
        raw_df, metadata = load_tlg_trial_balance(uploaded_file)
        year, month = _period_from_metadata(metadata)
        if (year, month) in seen_periods:
            raise ValueError(
                f"Se cargó más de un balance para {month:02d}/{year}. "
                "Deja solamente el archivo que deseas usar para ese mes."
            )
        seen_periods.add((year, month))
        company = str(metadata.get("empresa") or "").strip()
        if company:
            companies.add(company)
        detail = prepare_tlg_detail(raw_df)
        balance_values, pyg_values = _account4_values(detail)
        periods.append(
            {
                "year": year,
                "month": month,
                "metadata": metadata,
                "balance_values": balance_values,
                "pyg_values": pyg_values,
                "file_name": getattr(uploaded_file, "name", f"{month:02d}-{year}.xlsx"),
                "accounts": int(detail["CODIGO_CUENTA"].nunique()),
            }
        )

    if len(companies) > 1:
        raise ValueError("Los balances cargados parecen pertenecer a empresas diferentes.")

    periods.sort(key=lambda item: (item["year"], item["month"]))
    workbook, source_name = _load_base_workbook(previous_file)
    bce = workbook["BCE"]
    pyg = workbook["P Y G"]

    summary_rows: list[dict[str, object]] = []
    for period in periods:
        year = int(period["year"])
        month = int(period["month"])
        bce_column = _find_bce_month_column(bce, year, month)
        pyg_column = _find_pyg_month_column(pyg, bce_column, year, month)
        bce_count = _write_mapped_accounts(
            bce, bce_column, period["balance_values"]
        )
        pyg_count = _write_mapped_accounts(
            pyg, pyg_column, period["pyg_values"]
        )
        summary_rows.append(
            {
                "Archivo": period["file_name"],
                "Periodo": f"{month:02d}/{year}",
                "Cuentas leídas": period["accounts"],
                "Filas BCE actualizadas": bce_count,
                "Filas PYG actualizadas": pyg_count,
            }
        )

    if hasattr(workbook, "calculation"):
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"

    output = BytesIO()
    workbook.save(output)
    last_period = periods[-1]
    return {
        "output": output.getvalue(),
        "periods": pd.DataFrame(summary_rows),
        "source_name": source_name,
        "last_period": f"{int(last_period['month']):02d}/{int(last_period['year'])}",
    }


def export_monthly_reports(report: dict[str, object]) -> bytes:
    return bytes(report["output"])
