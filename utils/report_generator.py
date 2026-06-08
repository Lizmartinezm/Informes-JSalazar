from __future__ import annotations

from io import BytesIO

import pandas as pd


def format_cop(value: float) -> str:
    try:
        return "$" + f"{float(value):,.0f}".replace(",", ".")
    except Exception:
        return "$0"


def format_percent(value: float) -> str:
    try:
        return f"{float(value) * 100:.2f}%"
    except Exception:
        return "0.00%"


def build_excel_report(
    executive: pd.DataFrame,
    monthly: pd.DataFrame,
    payments: pd.DataFrame,
    sellers: pd.DataFrame,
    clients: pd.DataFrame,
    tips: pd.DataFrame,
    expenses: pd.DataFrame,
    operating: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheets = {
            "Resumen Ejecutivo": executive,
            "Informe Mensual": monthly,
            "Formas de Pago": payments,
            "Vendedores": sellers,
            "Clientes": clients,
            "Propinas": tips,
            "Gastos": expenses,
            "Resultado Operativo": operating,
        }
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
            worksheet = writer.sheets[name]
            for idx, column in enumerate(df.columns):
                width = min(max(len(str(column)) + 4, 14), 34)
                worksheet.set_column(idx, idx, width)
    return output.getvalue()


def build_pdf_report_placeholder() -> bytes | None:
    return None


def automatic_interpretation(
    sales_total: float,
    expenses_total: float,
    profit_total: float,
    margin: float,
    top_sales_month: str,
    top_expense_month: str,
    top_payment: str,
    top_payment_share: float,
    top_seller: str,
) -> str:
    recommendation = (
        f"Se recomienda revisar con detalle los gastos de {top_expense_month}, especialmente si crecieron por encima de las ventas."
        if top_expense_month != "Sin datos"
        else "Se recomienda mantener un registro de gastos más detallado para mejorar el seguimiento del resultado."
    )
    return (
        f"Durante el periodo analizado, el restaurante registró ventas acumuladas por {format_cop(sales_total)}, "
        f"gastos acumulados por {format_cop(expenses_total)} y una utilidad estimada de {format_cop(profit_total)}, "
        f"equivalente a un margen aproximado del {format_percent(margin)}. "
        f"El mes con mayor venta fue {top_sales_month} y el mes con mayor gasto fue {top_expense_month}. "
        f"La forma de pago más representativa fue {top_payment}, con una participación del {format_percent(top_payment_share)} "
        f"sobre el recaudo analizado. El vendedor con mayor participación fue {top_seller}. {recommendation}"
    )
