from __future__ import annotations

from io import BytesIO

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


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


def _pdf_table(df: pd.DataFrame, max_rows: int = 12) -> Table:
    shown = df.head(max_rows).fillna("").astype(str)
    data = [shown.columns.tolist()] + shown.values.tolist()
    table = Table(data, repeatRows=1, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F766E")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    return table


def _section(title: str, styles: dict) -> list:
    return [Spacer(1, 0.18 * inch), Paragraph(title, styles["SectionTitle"]), Spacer(1, 0.08 * inch)]


def build_pdf_report(
    executive: pd.DataFrame,
    monthly: pd.DataFrame,
    payments: pd.DataFrame,
    sellers: pd.DataFrame,
    clients: pd.DataFrame,
    tips: pd.DataFrame,
    expenses: pd.DataFrame,
    operating: pd.DataFrame,
    interpretation: str,
) -> bytes:
    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=42,
        bottomMargin=36,
        title="Informe Gerencial Restaurante Sazon",
    )

    base_styles = getSampleStyleSheet()
    styles = {
        "Title": ParagraphStyle(
            "ReportTitle",
            parent=base_styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=17,
            leading=21,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#0F172A"),
            spaceAfter=8,
        ),
        "Subtitle": ParagraphStyle(
            "ReportSubtitle",
            parent=base_styles["BodyText"],
            fontSize=10,
            leading=13,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#475569"),
            spaceAfter=14,
        ),
        "SectionTitle": ParagraphStyle(
            "SectionTitle",
            parent=base_styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=15,
            textColor=colors.HexColor("#0F766E"),
        ),
        "Body": ParagraphStyle(
            "ReportBody",
            parent=base_styles["BodyText"],
            fontSize=9,
            leading=13,
            textColor=colors.HexColor("#1F2937"),
        ),
    }

    story = [
        Paragraph("Informe Gerencial de Ventas y Gastos - Restaurante Sazon", styles["Title"]),
        Paragraph("Analisis mensual y acumulado del periodo cargado", styles["Subtitle"]),
    ]

    story += _section("Resumen ejecutivo", styles)
    story.append(_pdf_table(executive, max_rows=20))
    story += _section("Interpretacion automatica", styles)
    story.append(Paragraph(interpretation, styles["Body"]))
    story += _section("Informe mensual", styles)
    story.append(_pdf_table(monthly, max_rows=14))
    story += _section("Formas de pago", styles)
    story.append(_pdf_table(payments, max_rows=10))

    story.append(PageBreak())
    story += _section("Ventas por vendedor", styles)
    story.append(_pdf_table(sellers, max_rows=15))
    story += _section("Clientes principales", styles)
    story.append(_pdf_table(clients, max_rows=15))
    story += _section("Propinas por mes", styles)
    story.append(_pdf_table(tips, max_rows=12))
    story += _section("Gastos por mes", styles)
    story.append(_pdf_table(expenses, max_rows=12))
    story += _section("Resultado operativo", styles)
    story.append(_pdf_table(operating, max_rows=14))

    doc.build(story)
    output.seek(0)
    return output.getvalue()


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
