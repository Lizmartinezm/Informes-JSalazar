from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Flowable, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


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


PRIMARY = colors.HexColor("#0F766E")
INK = colors.HexColor("#0F172A")
MUTED = colors.HexColor("#475569")
LINE = colors.HexColor("#CBD5E1")
SOFT = colors.HexColor("#F8FAFC")
WARNING = colors.HexColor("#B45309")


def _as_number(value: object) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("$", "").replace("%", "").replace(".", "").replace(",", ".").strip()
    try:
        return float(text)
    except ValueError:
        return 0.0


def _format_df_money(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy()
    for column in columns:
        if column in out.columns:
            out[column] = out[column].map(format_cop)
    return out


def _format_df_percent(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy()
    for column in columns:
        if column in out.columns:
            out[column] = out[column].map(format_percent)
    return out


def _pdf_table(df: pd.DataFrame, max_rows: int = 12, widths: list[float] | None = None) -> Table:
    shown = df.head(max_rows).fillna("").astype(str)
    data = [shown.columns.tolist()] + shown.values.tolist()
    table = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
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


def _paragraph(text: str, styles: dict, style: str = "Body") -> Paragraph:
    return Paragraph(str(text).replace("\n", "<br/>"), styles[style])


class KPIGrid(Flowable):
    def __init__(self, items: list[tuple[str, str]], width: float = 700, columns: int = 4):
        super().__init__()
        self.items = items
        self.width = width
        self.columns = columns
        self.card_h = 58
        self.gap = 10
        self.height = ((len(items) + columns - 1) // columns) * (self.card_h + self.gap)

    def wrap(self, avail_width, avail_height):
        self.width = min(self.width, avail_width)
        return self.width, self.height

    def draw(self):
        card_w = (self.width - self.gap * (self.columns - 1)) / self.columns
        for index, (label, value) in enumerate(self.items):
            row, col = divmod(index, self.columns)
            x = col * (card_w + self.gap)
            y = self.height - (row + 1) * self.card_h - row * self.gap
            self.canv.setStrokeColor(LINE)
            self.canv.setFillColor(colors.white)
            self.canv.roundRect(x, y, card_w, self.card_h, 6, stroke=1, fill=1)
            self.canv.setFillColor(MUTED)
            self.canv.setFont("Helvetica", 7.5)
            self.canv.drawString(x + 8, y + self.card_h - 18, label[:34])
            self.canv.setFillColor(INK)
            self.canv.setFont("Helvetica-Bold", 12)
            self.canv.drawString(x + 8, y + 16, value[:24])


class BarChart(Flowable):
    def __init__(self, labels: list[str], values: list[float], width: float = 620, height: float = 145, color=PRIMARY):
        super().__init__()
        self.labels = labels
        self.values = values
        self.width = width
        self.height = height
        self.color = color

    def wrap(self, avail_width, avail_height):
        self.width = min(self.width, avail_width)
        return self.width, self.height

    def draw(self):
        if not self.labels or not self.values:
            return
        left = 58
        bottom = 24
        chart_w = self.width - left - 12
        chart_h = self.height - bottom - 10
        max_value = max(self.values) or 1
        bar_gap = 7
        bar_h = max(8, (chart_h - bar_gap * (len(self.values) - 1)) / len(self.values))
        self.canv.setFont("Helvetica", 7)
        for index, (label, value) in enumerate(zip(self.labels, self.values)):
            y = bottom + (len(self.values) - index - 1) * (bar_h + bar_gap)
            bar_w = chart_w * (value / max_value)
            self.canv.setFillColor(MUTED)
            self.canv.drawRightString(left - 6, y + 2, label[:12])
            self.canv.setFillColor(self.color)
            self.canv.roundRect(left, y, bar_w, bar_h, 3, stroke=0, fill=1)
            self.canv.setFillColor(INK)
            self.canv.drawString(left + bar_w + 4, y + 2, format_cop(value))


class GroupedBarChart(Flowable):
    def __init__(self, labels: list[str], first: list[float], second: list[float], first_label: str = "Ventas", second_label: str = "Gastos", width: float = 620, height: float = 150):
        super().__init__()
        self.labels = labels
        self.first = first
        self.second = second
        self.first_label = first_label
        self.second_label = second_label
        self.width = width
        self.height = height

    def wrap(self, avail_width, avail_height):
        self.width = min(self.width, avail_width)
        return self.width, self.height

    def draw(self):
        if not self.labels:
            return
        left = 40
        bottom = 28
        chart_w = self.width - left - 16
        chart_h = self.height - bottom - 20
        max_value = max(self.first + self.second) or 1
        group_w = chart_w / len(self.labels)
        bar_w = max(8, group_w * 0.28)
        self.canv.setFont("Helvetica", 7)
        for idx, label in enumerate(self.labels):
            x = left + idx * group_w + group_w * 0.16
            h1 = chart_h * (self.first[idx] / max_value)
            h2 = chart_h * (self.second[idx] / max_value)
            self.canv.setFillColor(PRIMARY)
            self.canv.rect(x, bottom, bar_w, h1, stroke=0, fill=1)
            self.canv.setFillColor(colors.HexColor("#D97706"))
            self.canv.rect(x + bar_w + 3, bottom, bar_w, h2, stroke=0, fill=1)
            self.canv.setFillColor(MUTED)
            self.canv.drawCentredString(x + bar_w, bottom - 12, label[:8])
        self.canv.setFillColor(PRIMARY)
        self.canv.rect(left, self.height - 12, 8, 8, stroke=0, fill=1)
        self.canv.setFillColor(INK)
        self.canv.drawString(left + 12, self.height - 12, self.first_label)
        self.canv.setFillColor(colors.HexColor("#D97706"))
        self.canv.rect(left + 80, self.height - 12, 8, 8, stroke=0, fill=1)
        self.canv.setFillColor(INK)
        self.canv.drawString(left + 92, self.height - 12, self.second_label)


class MultiBarChart(Flowable):
    def __init__(
        self,
        labels: list[str],
        series: list[tuple[str, list[float], colors.Color]],
        width: float = 690,
        height: float = 150,
    ):
        super().__init__()
        self.labels = labels
        self.series = series
        self.width = width
        self.height = height

    def wrap(self, avail_width, avail_height):
        self.width = min(self.width, avail_width)
        return self.width, self.height

    def draw(self):
        if not self.labels or not self.series:
            return
        left = 42
        bottom = 28
        chart_w = self.width - left - 12
        chart_h = self.height - bottom - 20
        all_values = [value for _, values, _ in self.series for value in values]
        max_value = max(all_values) if all_values else 1
        max_value = max_value or 1
        group_w = chart_w / len(self.labels)
        bar_w = max(5, group_w * 0.18)
        self.canv.setFont("Helvetica", 7)
        for idx, label in enumerate(self.labels):
            group_x = left + idx * group_w + group_w * 0.13
            for s_idx, (_, values, color) in enumerate(self.series):
                value = values[idx] if idx < len(values) else 0
                h = chart_h * (value / max_value)
                self.canv.setFillColor(color)
                self.canv.rect(group_x + s_idx * (bar_w + 2), bottom, bar_w, h, stroke=0, fill=1)
            self.canv.setFillColor(MUTED)
            self.canv.drawCentredString(group_x + bar_w, bottom - 12, label[:8])
        legend_x = left
        for name, _, color in self.series:
            self.canv.setFillColor(color)
            self.canv.rect(legend_x, self.height - 12, 8, 8, stroke=0, fill=1)
            self.canv.setFillColor(INK)
            self.canv.drawString(legend_x + 12, self.height - 12, name)
            legend_x += 120


def _get_indicator(executive: pd.DataFrame, indicator: str, default: str = "Sin datos") -> str:
    if executive.empty:
        return default
    mask = executive["Indicador"].astype(str).str.lower() == indicator.lower()
    if not mask.any():
        return default
    return str(executive.loc[mask, "Valor"].iloc[0])


def _money_col(df: pd.DataFrame, col: str) -> pd.Series:
    return df[col].map(_as_number) if col in df.columns else pd.Series(dtype=float)


def _standardize_pdf_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename = {}
    for column in df.columns:
        lowered = str(column).lower()
        if "participaci" in lowered and "%" in lowered:
            rename[column] = "Participacion %"
        elif "numero" in lowered and "factura" in lowered or "factura" in lowered and "nã" in lowered:
            rename[column] = "Numero de facturas"
        elif "mero de facturas" in lowered:
            rename[column] = "Numero de facturas"
        elif "compras" in lowered:
            rename[column] = "Numero de compras"
        elif "ltima fecha" in lowered or "ultima fecha" in lowered:
            rename[column] = "Ultima fecha de compra"
        elif "cortes" in lowered and "total" in lowered:
            rename[column] = "Total cortesías"
        elif "cortes" in lowered:
            rename[column] = "Cortesías"
    return df.rename(columns=rename)


def _month_comment(monthly: pd.DataFrame) -> str:
    if monthly.empty:
        return "No hay informacion mensual suficiente para evaluar la tendencia de ventas."
    top = monthly.sort_values("Ventas", ascending=False).iloc[0]
    low = monthly.sort_values("Ventas", ascending=True).iloc[0]
    last = monthly.iloc[-1]
    prior = monthly.iloc[-2] if len(monthly) > 1 else None
    direction = "se mantuvieron sin variacion frente al mes anterior"
    if prior is not None and prior["Ventas"]:
        change = (last["Ventas"] - prior["Ventas"]) / prior["Ventas"]
        direction = f"{'subieron' if change >= 0 else 'bajaron'} {format_percent(abs(change))} frente al mes anterior"
    gap = 0 if not top["Ventas"] else (last["Ventas"] - top["Ventas"]) / top["Ventas"]
    return (
        f"El mes con mayor venta fue {top['Mes']} con {format_cop(top['Ventas'])}, mientras que el mes con menor venta fue "
        f"{low['Mes']} con {format_cop(low['Ventas'])}. En el ultimo mes, las ventas {direction}. "
        f"Frente al mejor mes, el ultimo mes tuvo una variacion de {format_percent(gap)}."
    )


def _expense_warning(operating: pd.DataFrame) -> str:
    if operating.empty or "Gastos" not in operating.columns:
        return "No se encontraron gastos para contrastar contra las ventas."
    zero_months = operating.loc[operating["Gastos"].fillna(0) == 0, "Mes"].astype(str).tolist()
    months_with_expenses = operating.loc[operating["Gastos"].fillna(0) > 0, "Mes"].astype(str).tolist()
    if len(months_with_expenses) == 1:
        return (
            f"Los gastos cargados se concentran unicamente en el mes de {months_with_expenses[0]}. "
            "Se recomienda validar si los gastos de los demas meses estan pendientes por cargar."
        )
    if zero_months:
        return (
            "Advertencia: Se identificaron meses con gastos en $0. Esto puede indicar que el archivo de gastos no tiene "
            "informacion completa para esos meses. Por lo tanto, la utilidad estimada y el margen pueden estar sobreestimados."
        )
    return "Los gastos fueron cargados para todos los meses analizados, lo que permite una lectura mas completa del resultado estimado."


def _payment_comment(payments: pd.DataFrame) -> str:
    recaudo = payments[~payments["Forma de pago"].astype(str).str.lower().str.contains("cortesia|cortesía|descuento", regex=True)]
    if recaudo.empty or recaudo["Valor"].sum() == 0:
        return "No hay informacion suficiente para identificar una forma de pago predominante."
    top = recaudo.sort_values("Valor", ascending=False).iloc[0]
    comment = (
        f"La forma de pago predominante fue {top['Forma de pago']} con {format_cop(top['Valor'])}, "
        f"equivalente al {format_percent(top['Valor'] / recaudo['Valor'].sum() if recaudo['Valor'].sum() else 0)} del recaudo."
    )
    cash_value = recaudo.loc[recaudo["Forma de pago"].astype(str).str.lower().str.contains("efectivo"), "Valor"].sum()
    if recaudo["Valor"].sum() and cash_value / recaudo["Valor"].sum() > 0.5:
        comment += " Como el efectivo supera el 50%, se recomienda realizar arqueos diarios, conciliacion de caja y validacion contra ventas registradas."
    return comment


def _seller_comment(sellers: pd.DataFrame, sales_total: float) -> str:
    if sellers.empty:
        return "No hay informacion suficiente para evaluar vendedores."
    sellers = sellers.copy()
    sellers["Participacion sobre ventas"] = sellers["Total ventas"] / sales_total if sales_total else 0
    top_sales = sellers.sort_values("Total ventas", ascending=False).iloc[0]
    top_ticket = sellers.sort_values("Ticket promedio", ascending=False).iloc[0]
    return (
        f"El vendedor con mayor participacion fue {top_sales['Vendedor']} con ventas por {format_cop(top_sales['Total ventas'])}, "
        f"equivalente al {format_percent(top_sales['Participacion sobre ventas'])} del total. "
        f"Tambien se destaca {top_ticket['Vendedor']} por tener el mayor ticket promedio."
    )


def _seller_control_comment(sellers: pd.DataFrame) -> str:
    if sellers.empty:
        return "No hay informacion suficiente para evaluar vendedores."
    top_sales = sellers.sort_values("Total ventas", ascending=False).iloc[0]
    top_ticket = sellers.sort_values("Ticket promedio", ascending=False).iloc[0]
    if "Total cortesias" in sellers.columns:
        top_courtesy = sellers.sort_values("Total cortesias", ascending=False).iloc[0]
        courtesy_text = f" El mayor valor de cortesias lo registra {top_courtesy['Vendedor']} con {format_cop(top_courtesy['Total cortesias'])}."
    elif "Total cortesías" in sellers.columns:
        top_courtesy = sellers.sort_values("Total cortesías", ascending=False).iloc[0]
        courtesy_text = f" El mayor valor de cortesias lo registra {top_courtesy['Vendedor']} con {format_cop(top_courtesy['Total cortesías'])}."
    else:
        courtesy_text = " No hay cortesias por vendedor para comparar."
    return (
        f"El vendedor con mayor venta fue {top_sales['Vendedor']} con {format_cop(top_sales['Total ventas'])}."
        f"{courtesy_text} El mayor ticket promedio lo registra {top_ticket['Vendedor']}."
    )


def _client_sections(clients: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, str]:
    if clients.empty:
        return pd.DataFrame(), pd.DataFrame(), "No hay informacion suficiente para evaluar clientes."
    names = clients["Cliente"].astype(str).str.upper().str.strip()
    general_mask = names.str.contains("CLIENTE GENERAL|CONSUMIDOR FINAL|SIN CLIENTE", regex=True)
    general = clients.loc[general_mask].copy()
    identified = clients.loc[~general_mask].copy().head(10)
    comment = (
        "El alto volumen registrado como Cliente General limita el analisis de fidelizacion. Se recomienda mejorar la identificacion "
        "de clientes para medir recurrencia, frecuencia de compra y clientes de mayor valor."
        if not general.empty
        else "La base de clientes esta identificada con mayor detalle, lo que permite analizar recurrencia y clientes de mayor valor."
    )
    return general, identified, comment


def _tips_comment(tips: pd.DataFrame, sales_total: float) -> str:
    if tips.empty:
        return "No hay informacion suficiente para analizar propinas."
    top = tips.sort_values("Propinas", ascending=False).iloc[0]
    total = tips["Propinas"].sum()
    return f"Las propinas acumuladas fueron {format_cop(total)}, equivalentes al {format_percent(total / sales_total if sales_total else 0)} de las ventas. El mes con mayores propinas fue {top['Mes']}."


def _courtesy_comment(monthly: pd.DataFrame, gross_sales_total: float, courtesy_total: float) -> str:
    if monthly.empty or "Cortesías" not in monthly.columns:
        return "No se encontro informacion de cortesias o descuentos en el archivo de ventas."
    top = monthly.sort_values("Cortesías", ascending=False).iloc[0]
    pct = courtesy_total / gross_sales_total if gross_sales_total else 0
    comment = (
        f"Durante el periodo analizado se otorgaron cortesias o descuentos por {format_cop(courtesy_total)}, "
        f"equivalentes al {format_percent(pct)} de las ventas brutas. Este valor no representa dinero recibido, "
        "pero si ventas que el restaurante dejo de cobrar. Se recomienda hacer seguimiento mensual a las cortesias, "
        "identificar quien las autoriza y revisar si corresponden a estrategia comercial, atencion al cliente o ajustes operativos."
    )
    if pct > 0.03:
        comment += " Al superar el 3% de las ventas, conviene establecer topes, responsables de autorizacion y una revision semanal."
    comment += f" El mes con mayor valor de cortesias fue {top['Mes']} con {format_cop(top['Cortesías'])}."
    return comment


def _conclusions(monthly: pd.DataFrame, payments: pd.DataFrame, operating: pd.DataFrame, clients: pd.DataFrame, sales_total: float) -> list[str]:
    top_month = monthly.sort_values("Ventas", ascending=False).iloc[0]["Mes"] if not monthly.empty else "Sin datos"
    last_change = "sin informacion suficiente"
    if len(monthly) > 1 and monthly.iloc[-2]["Ventas"]:
        last_change = format_percent((monthly.iloc[-1]["Ventas"] - monthly.iloc[-2]["Ventas"]) / monthly.iloc[-2]["Ventas"])
    recaudo = payments[~payments["Forma de pago"].astype(str).str.lower().str.contains("cortesia|cortesía|descuento", regex=True)]
    payment = recaudo.sort_values("Valor", ascending=False).iloc[0]["Forma de pago"] if not recaudo.empty and recaudo["Valor"].sum() else "Sin datos"
    zero_expense = not operating.empty and (operating["Gastos"].fillna(0) == 0).any()
    has_general = not clients.empty and clients["Cliente"].astype(str).str.upper().str.contains("CLIENTE GENERAL").any()
    conclusions = [
        f"Las ventas acumuladas del periodo fueron {format_cop(sales_total)}, siendo {top_month} el mejor mes.",
        f"El ultimo mes presento una variacion de {last_change} frente al mes anterior.",
        f"La mayoria del recaudo se realizo por {payment}; conviene fortalecer los controles y conciliaciones de este medio de pago.",
    ]
    conclusions.append(
        "La utilidad estimada debe analizarse con cuidado porque hay meses sin gastos cargados."
        if zero_expense
        else "La utilidad estimada se calculo con los gastos disponibles en el archivo cargado."
    )
    if has_general:
        conclusions.append('Se recomienda mejorar la identificacion de clientes para no depender del registro generico "Cliente General".')
    else:
        conclusions.append("Se recomienda seguir fortaleciendo el registro de clientes para medir frecuencia y recurrencia.")
    return conclusions


def _manager_recommendations(
    monthly: pd.DataFrame,
    payments: pd.DataFrame,
    operating: pd.DataFrame,
    clients: pd.DataFrame,
    gross_sales_total: float,
    courtesy_total: float,
    expenses_total: float,
) -> list[str]:
    recommendations = []
    top_month = monthly.sort_values("Ventas brutas", ascending=False).iloc[0]["Mes"] if "Ventas brutas" in monthly.columns and not monthly.empty else "Sin datos"
    recommendations.append(f"Las ventas brutas del periodo fueron {format_cop(gross_sales_total)}; el mejor mes fue {top_month}.")
    if gross_sales_total and courtesy_total / gross_sales_total > 0.03:
        recommendations.append("Las cortesias superan el 3% de las ventas brutas; defina responsables, motivos autorizados y revision semanal.")
    else:
        recommendations.append("Mantenga seguimiento mensual de cortesias para evitar que se vuelvan una perdida silenciosa de ingreso potencial.")
    real_payments = payments[~payments["Forma de pago"].astype(str).str.lower().str.contains("cortesia|cortesía|descuento|cortes", regex=True)]
    cash = real_payments.loc[real_payments["Forma de pago"].astype(str).str.lower().str.contains("efectivo"), "Valor"].sum()
    if real_payments["Valor"].sum() and cash / real_payments["Valor"].sum() > 0.5:
        recommendations.append("El efectivo supera el 50% del recaudo; haga arqueos diarios y conciliacion contra ventas registradas.")
    else:
        recommendations.append("Concilie diariamente efectivo, transferencias y tarjetas contra el cierre de ventas.")
    if not operating.empty and (operating["Gastos"].fillna(0) == 0).any():
        recommendations.append("Hay meses sin gastos registrados; la utilidad estimada puede estar sobreestimada.")
    else:
        recommendations.append(f"El impacto total de cortesias y gastos fue {format_cop(courtesy_total + expenses_total)}; revise si es consistente con la operacion.")
    has_general = not clients.empty and clients["Cliente"].astype(str).str.upper().str.contains("CLIENTE GENERAL").any()
    if has_general:
        recommendations.append("Reduzca el uso de Cliente General para medir recurrencia, frecuencia y clientes de mayor valor.")
    else:
        recommendations.append("Use la identificacion de clientes para activar acciones de fidelizacion y seguimiento de recurrencia.")
    return recommendations[:5]


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
    monthly = monthly.copy()
    payments = payments.copy()
    sellers = sellers.copy()
    clients = clients.copy()
    tips = tips.copy()
    expenses = expenses.copy()
    operating = operating.copy()
    monthly = _standardize_pdf_columns(monthly)
    payments = _standardize_pdf_columns(payments)
    sellers = _standardize_pdf_columns(sellers)
    clients = _standardize_pdf_columns(clients)
    tips = _standardize_pdf_columns(tips)
    expenses = _standardize_pdf_columns(expenses)
    operating = _standardize_pdf_columns(operating)

    for col in ["Ventas", "Valor neto recibido", "Propinas", "Gastos", "Utilidad estimada", "Ticket promedio"]:
        if col in monthly.columns:
            monthly[col] = _money_col(monthly, col)
        if col in operating.columns:
            operating[col] = _money_col(operating, col)
    for col in ["Valor", "Participacion %"]:
        if col in payments.columns:
            payments[col] = payments[col].map(_as_number)
    for col in ["Total ventas", "Total valor neto", "Total propinas", "Ticket promedio"]:
        if col in sellers.columns:
            sellers[col] = _money_col(sellers, col)
    for col in ["Total vendido", "Ticket promedio"]:
        if col in clients.columns:
            clients[col] = _money_col(clients, col)
    if "Propinas" in tips.columns:
        tips["Propinas"] = _money_col(tips, "Propinas")
    if "Gastos" in expenses.columns:
        expenses["Gastos"] = _money_col(expenses, "Gastos")
    if "Margen estimado" in operating.columns:
        operating["Margen estimado"] = operating["Margen estimado"].map(_as_number)
    if "Margen estimado" in monthly.columns:
        monthly["Margen estimado"] = monthly["Margen estimado"].map(_as_number)

    sales_total = monthly["Ventas netas"].sum() if "Ventas netas" in monthly.columns else monthly["Ventas"].sum() if "Ventas" in monthly.columns else 0
    gross_sales_total = monthly["Ventas brutas"].sum() if "Ventas brutas" in monthly.columns else sales_total
    courtesy_total = monthly["Cortesías"].sum() if "Cortesías" in monthly.columns else 0
    real_collected_total = monthly["Recaudo real"].sum() if "Recaudo real" in monthly.columns else payments.loc[~payments["Forma de pago"].astype(str).str.lower().str.contains("cortesia|cortesía|descuento", regex=True), "Valor"].sum()
    expenses_total = operating["Gastos"].sum() if "Gastos" in operating.columns else 0
    profit_total = sales_total - expenses_total
    invoice_count = int(monthly["Numero de facturas"].map(_as_number).sum()) if "Numero de facturas" in monthly.columns else 0
    avg_ticket = sales_total / invoice_count if invoice_count else 0
    margin_total = profit_total / sales_total if sales_total else 0
    last_month_sales = monthly.iloc[-1]["Ventas"] if not monthly.empty else 0
    recaudo_payments = payments[~payments["Forma de pago"].astype(str).str.lower().str.contains("cortesia|cortesía|descuento", regex=True)]
    top_payment = recaudo_payments.sort_values("Valor", ascending=False).iloc[0] if not recaudo_payments.empty and recaudo_payments["Valor"].sum() else None
    main_payment = str(top_payment["Forma de pago"]) if top_payment is not None else "Sin datos"
    period = f"{monthly.iloc[0]['Mes']} - {monthly.iloc[-1]['Mes']}" if not monthly.empty else "Sin datos"

    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(letter),
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
            alignment=TA_LEFT,
            textColor=colors.HexColor("#1F2937"),
        ),
        "Alert": ParagraphStyle(
            "ReportAlert",
            parent=base_styles["BodyText"],
            fontSize=9,
            leading=13,
            textColor=WARNING,
            backColor=colors.HexColor("#FEF3C7"),
            borderPadding=7,
            spaceBefore=4,
            spaceAfter=4,
        ),
    }

    story = [
        Paragraph("Informe Gerencial de Ventas y Gastos - Restaurante Sazon", styles["Title"]),
        Paragraph("Analisis mensual y acumulado del periodo cargado", styles["Subtitle"]),
        Spacer(1, 0.2 * inch),
        Paragraph(f"Periodo analizado: {period}", styles["Subtitle"]),
        Paragraph(f"Fecha de generacion: {datetime.now().strftime('%Y-%m-%d')}", styles["Subtitle"]),
        PageBreak(),
    ]

    kpis = [
        ("Ventas brutas", format_cop(gross_sales_total)),
        ("Cortesias otorgadas", format_cop(courtesy_total)),
        ("Ventas netas", format_cop(sales_total)),
        ("Recaudo real", format_cop(real_collected_total)),
        ("% cortesias", format_percent(courtesy_total / gross_sales_total if gross_sales_total else 0)),
        ("Gastos acumulados", format_cop(expenses_total)),
        ("Utilidad estimada", format_cop(profit_total)),
        ("Facturas", f"{invoice_count:,.0f}".replace(",", ".")),
        ("Pago principal", main_payment),
    ]
    story += _section("Resumen ejecutivo", styles)
    story.append(KPIGrid(kpis))
    top_month = monthly.sort_values("Ventas", ascending=False).iloc[0] if not monthly.empty else {"Mes": "Sin datos", "Ventas": 0}
    low_month = monthly.sort_values("Ventas", ascending=True).iloc[0] if not monthly.empty else {"Mes": "Sin datos", "Ventas": 0}
    payment_share = top_payment["Participacion %"] if top_payment is not None and "Participacion %" in top_payment else 0
    executive_text = (
        f"El restaurante registro ventas brutas por {format_cop(gross_sales_total)} y ventas netas por {format_cop(sales_total)} durante el periodo analizado. "
        f"El mes con mayor venta fue {top_month['Mes']} con {format_cop(top_month['Ventas'])}, mientras que el mes con menor venta fue "
        f"{low_month['Mes']} con {format_cop(low_month['Ventas'])}. La forma de pago mas utilizada fue {main_payment}, "
        f"representando el {format_percent(payment_share)} del recaudo real. La utilidad estimada fue de {format_cop(profit_total)}; "
        "sin embargo, este resultado depende de que los gastos cargados esten completos para todos los meses."
    )
    story += _section("Lectura ejecutiva", styles)
    story.append(_paragraph(executive_text, styles))
    story.append(PageBreak())

    sales_table = monthly[["Mes", "Ventas", "Ticket promedio"]].copy()
    if "Numero de facturas" in monthly.columns:
        sales_table.insert(2, "Facturas", monthly["Numero de facturas"].map(lambda value: f"{_as_number(value):,.0f}".replace(",", ".")))
    sales_table["Variacion mensual %"] = monthly["Ventas"].pct_change().fillna(0)
    sales_table = _format_df_percent(_format_df_money(sales_table, ["Ventas", "Ticket promedio"]), ["Variacion mensual %"])
    story += _section("Comportamiento mensual de ventas", styles)
    story.append(_pdf_table(sales_table, max_rows=12))
    story.append(Spacer(1, 0.08 * inch))
    story.append(BarChart(monthly["Mes"].astype(str).tolist(), monthly["Ventas"].tolist(), width=680, height=155))
    story.append(_paragraph(_month_comment(monthly), styles))
    story.append(PageBreak())

    result_table = operating[["Mes", "Ventas", "Gastos", "Utilidad estimada", "Margen estimado"]].copy()
    result_table = _format_df_percent(_format_df_money(result_table, ["Ventas", "Gastos", "Utilidad estimada"]), ["Margen estimado"])
    story += _section("Resultado estimado y gastos", styles)
    story.append(_pdf_table(result_table, max_rows=12))
    story.append(Spacer(1, 0.08 * inch))
    story.append(GroupedBarChart(operating["Mes"].astype(str).tolist(), operating["Ventas"].tolist(), operating["Gastos"].tolist(), width=680, height=155))
    expense_note = _expense_warning(operating)
    story.append(_paragraph(expense_note, styles, "Alert" if "Advertencia" in expense_note or "unicamente" in expense_note else "Body"))
    story.append(PageBreak())

    payment_table = _format_df_percent(_format_df_money(payments[["Forma de pago", "Valor", "Participacion %"]], ["Valor"]), ["Participacion %"])
    story += _section("Recaudo y ventas no cobradas", styles)
    story.append(_pdf_table(payment_table, max_rows=8))
    story.append(Spacer(1, 0.08 * inch))
    story.append(BarChart(payments["Forma de pago"].astype(str).tolist(), payments["Valor"].tolist(), width=680, height=145, color=colors.HexColor("#2563EB")))
    story.append(_paragraph(_payment_comment(payments), styles))
    story.append(_paragraph("Las cortesias corresponden a consumos registrados como venta, pero no cobrados al cliente. Por esta razon no representan entrada de dinero, pero si deben controlarse porque reducen el ingreso potencial del restaurante.", styles, "Alert"))
    story.append(PageBreak())

    seller_pdf = sellers.copy()
    seller_pdf["Participacion sobre ventas"] = seller_pdf["Total ventas"] / sales_total if sales_total else 0
    seller_pdf = seller_pdf.sort_values("Total ventas", ascending=False).head(5)
    seller_pdf = seller_pdf[["Vendedor", "Total ventas", "Participacion sobre ventas", "Numero de facturas", "Ticket promedio", "Total propinas"]]
    seller_pdf = seller_pdf.rename(columns={"Numero de facturas": "Facturas", "Total propinas": "Propinas"})
    seller_pdf = _format_df_percent(_format_df_money(seller_pdf, ["Total ventas", "Ticket promedio", "Propinas"]), ["Participacion sobre ventas"])
    story += _section("Desempeno por vendedor", styles)
    story.append(_pdf_table(seller_pdf, max_rows=5))
    story.append(_paragraph(_seller_comment(sellers, sales_total), styles))
    story.append(PageBreak())

    general, identified, client_comment = _client_sections(clients)
    client_cols = ["Cliente", "Numero de compras", "Total vendido", "Ticket promedio", "Ultima fecha de compra"]
    client_cols = [col for col in client_cols if col in clients.columns]
    story += _section("Clientes", styles)
    if not general.empty:
        general_pdf = _format_df_money(general[client_cols].head(3), ["Total vendido", "Ticket promedio"])
        story.append(_paragraph("Cliente General", styles, "SectionTitle"))
        story.append(_pdf_table(general_pdf, max_rows=3))
        story.append(Spacer(1, 0.08 * inch))
    identified_pdf = _format_df_money(identified[client_cols].head(10), ["Total vendido", "Ticket promedio"]) if not identified.empty else pd.DataFrame({"Mensaje": ["No hay clientes identificados para mostrar."]})
    story.append(_paragraph("Clientes identificados", styles, "SectionTitle"))
    story.append(_pdf_table(identified_pdf, max_rows=10))
    story.append(_paragraph(client_comment, styles, "Alert" if not general.empty else "Body"))
    story.append(PageBreak())

    courtesy_month = monthly[["Mes", "Cortesías"]].copy() if "Cortesías" in monthly.columns else pd.DataFrame({"Mes": [], "Cortesías": []})
    courtesy_seller = pd.DataFrame()
    if "Total cortesías" in sellers.columns:
        courtesy_seller = sellers[["Vendedor", "Total cortesías"]].sort_values("Total cortesías", ascending=False).head(8).rename(columns={"Total cortesías": "Cortesías"})
    story += _section("Analisis de cortesias", styles)
    story.append(_paragraph(f"Total de cortesias acumuladas: {format_cop(courtesy_total)}", styles))
    story.append(_paragraph(f"Porcentaje de cortesias sobre ventas brutas: {format_percent(courtesy_total / gross_sales_total if gross_sales_total else 0)}", styles))
    story.append(_pdf_table(_format_df_money(courtesy_month, ["Cortesías"]), max_rows=12))
    if not courtesy_seller.empty:
        story.append(Spacer(1, 0.08 * inch))
        story.append(_pdf_table(_format_df_money(courtesy_seller, ["Cortesías"]), max_rows=8))
    courtesy_style = "Alert" if (courtesy_total / gross_sales_total if gross_sales_total else 0) > 0.03 else "Body"
    story.append(_paragraph(_courtesy_comment(monthly, gross_sales_total, courtesy_total), styles, courtesy_style))
    story.append(PageBreak())

    tip_seller_col = "Vendedor" if "Vendedor" in sellers.columns else sellers.columns[0]
    tip_seller = sellers[[tip_seller_col, "Total propinas"]].sort_values("Total propinas", ascending=False).head(8).rename(columns={tip_seller_col: "Vendedor", "Total propinas": "Propinas"})
    story += _section("Propinas", styles)
    story.append(_paragraph(f"Propinas acumuladas: {format_cop(tips['Propinas'].sum() if 'Propinas' in tips.columns else 0)}", styles))
    story.append(_paragraph(f"Participacion de propinas sobre ventas: {format_percent((tips['Propinas'].sum() if 'Propinas' in tips.columns else 0) / sales_total if sales_total else 0)}", styles))
    story.append(_pdf_table(_format_df_money(tips[["Mes", "Propinas"]], ["Propinas"]), max_rows=12))
    story.append(Spacer(1, 0.08 * inch))
    story.append(_pdf_table(_format_df_money(tip_seller, ["Propinas"]), max_rows=8))
    story.append(_paragraph(_tips_comment(tips, sales_total), styles))
    story.append(PageBreak())

    story += _section("Conclusiones y recomendaciones", styles)
    conclusions = _conclusions(monthly, payments, operating, clients, sales_total)
    for idx, conclusion in enumerate(conclusions, start=1):
        story.append(_paragraph(f"{idx}. {conclusion}", styles))
        story.append(Spacer(1, 0.05 * inch))

    doc.build(story)
    output.seek(0)
    return output.getvalue()


def build_pdf_report_v2(
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
    monthly = _standardize_pdf_columns(monthly.copy())
    payments = _standardize_pdf_columns(payments.copy())
    sellers = _standardize_pdf_columns(sellers.copy())
    clients = _standardize_pdf_columns(clients.copy())
    operating = _standardize_pdf_columns(operating.copy())

    for frame, cols in [
        (monthly, ["Ventas brutas", "Cortesías", "CortesÃ­as", "Ventas netas", "Ventas", "Gastos", "Resultado estimado", "Utilidad estimada", "Recaudo real"]),
        (payments, ["Valor", "Participacion %"]),
        (sellers, ["Total ventas", "Total cortesías", "Total cortesÃ­as", "Total propinas", "Numero de facturas", "Ticket promedio"]),
        (clients, ["Numero de compras", "Total vendido", "Ticket promedio"]),
        (operating, ["Ventas", "Gastos", "Utilidad estimada", "Margen estimado"]),
    ]:
        for col in cols:
            if col in frame.columns:
                frame[col] = frame[col].map(_as_number)

    courtesy_col = "Cortesías" if "Cortesías" in monthly.columns else "CortesÃ­as" if "CortesÃ­as" in monthly.columns else None
    seller_courtesy_col = "Total cortesías" if "Total cortesías" in sellers.columns else "Total cortesÃ­as" if "Total cortesÃ­as" in sellers.columns else None
    gross_col = "Ventas brutas" if "Ventas brutas" in monthly.columns else "Ventas"
    net_col = "Ventas netas" if "Ventas netas" in monthly.columns else "Ventas"

    gross_sales_total = monthly[gross_col].sum() if gross_col in monthly.columns else 0
    courtesy_total = monthly[courtesy_col].sum() if courtesy_col else 0
    net_sales_total = monthly[net_col].sum() if net_col in monthly.columns else gross_sales_total - courtesy_total
    expenses_total = operating["Gastos"].sum() if "Gastos" in operating.columns else 0
    result_total = net_sales_total - expenses_total
    impact_total = courtesy_total + expenses_total
    courtesy_pct = courtesy_total / gross_sales_total if gross_sales_total else 0
    impact_pct = impact_total / gross_sales_total if gross_sales_total else 0
    period = f"{monthly.iloc[0]['Mes']} - {monthly.iloc[-1]['Mes']}" if not monthly.empty else "Sin datos"

    real_payments = payments[~payments["Forma de pago"].astype(str).str.lower().str.contains("cortesia|cortesía|descuento|cortes", regex=True)].copy()
    real_collected_total = real_payments["Valor"].sum() if "Valor" in real_payments.columns else 0
    main_payment = "Sin datos"
    if not real_payments.empty and real_collected_total:
        main_payment = str(real_payments.sort_values("Valor", ascending=False).iloc[0]["Forma de pago"])

    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(letter),
        rightMargin=34,
        leftMargin=34,
        topMargin=34,
        bottomMargin=30,
        title="Informe Gerencial Restaurante Sazon",
    )
    base_styles = getSampleStyleSheet()
    styles = {
        "Title": ParagraphStyle("ReportTitleV2", parent=base_styles["Title"], fontName="Helvetica-Bold", fontSize=17, leading=20, alignment=TA_CENTER, textColor=INK),
        "Subtitle": ParagraphStyle("ReportSubtitleV2", parent=base_styles["BodyText"], fontSize=9, leading=12, alignment=TA_CENTER, textColor=MUTED),
        "SectionTitle": ParagraphStyle("SectionTitleV2", parent=base_styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=14, textColor=PRIMARY),
        "Body": ParagraphStyle("ReportBodyV2", parent=base_styles["BodyText"], fontSize=8.5, leading=11.5, alignment=TA_LEFT, textColor=INK),
        "Alert": ParagraphStyle("ReportAlertV2", parent=base_styles["BodyText"], fontSize=8.5, leading=11.5, textColor=WARNING, backColor=colors.HexColor("#FEF3C7"), borderPadding=6),
    }

    story = [
        Paragraph("Informe Gerencial de Ventas y Gastos - Restaurante Sazon", styles["Title"]),
        Paragraph(f"Analisis mensual y acumulado del periodo cargado | Periodo: {period} | Generado: {datetime.now().strftime('%Y-%m-%d')}", styles["Subtitle"]),
    ]

    kpis = [
        ("Ventas brutas", format_cop(gross_sales_total)),
        ("Cortesias/descuentos", format_cop(courtesy_total)),
        ("Ventas netas", format_cop(net_sales_total)),
        ("Recaudo real", format_cop(real_collected_total)),
        ("Gastos registrados", format_cop(expenses_total)),
        ("Resultado estimado", format_cop(result_total)),
        ("% cortesias", format_percent(courtesy_pct)),
        ("Pago principal", main_payment),
    ]
    story += _section("Pagina 1: Resumen ejecutivo", styles)
    story.append(KPIGrid(kpis, columns=4))
    story.append(Spacer(1, 0.08 * inch))
    story.append(_paragraph(
        f"Estructura del resultado: Ventas brutas {format_cop(gross_sales_total)} (-) cortesias/descuentos {format_cop(courtesy_total)} "
        f"= ventas netas {format_cop(net_sales_total)} (-) gastos registrados {format_cop(expenses_total)} "
        f"= resultado operativo estimado {format_cop(result_total)}. El impacto total sobre ventas brutas fue {format_cop(impact_total)} "
        f"({format_percent(impact_pct)}). "
        + ("Alerta: las cortesias superan el 3% de las ventas brutas; conviene revisar autorizaciones, motivos y responsables." if courtesy_pct > 0.03 else "Las cortesias se mantienen por debajo del 3% de las ventas brutas."),
        styles,
        "Alert" if courtesy_pct > 0.03 else "Body",
    ))
    story.append(PageBreak())

    monthly_result = pd.DataFrame({
        "Mes": monthly["Mes"],
        "Ventas brutas": monthly[gross_col],
        "Cortesias": monthly[courtesy_col] if courtesy_col else 0,
        "Ventas netas": monthly[net_col],
        "Gastos": operating["Gastos"].values if "Gastos" in operating.columns else 0,
    })
    monthly_result["Resultado estimado"] = monthly_result["Ventas netas"] - monthly_result["Gastos"]
    monthly_result["% cortesias"] = monthly_result["Cortesias"] / monthly_result["Ventas brutas"].replace(0, pd.NA)
    monthly_result["Variacion ventas %"] = monthly_result["Ventas brutas"].pct_change().fillna(0)
    monthly_result = monthly_result.fillna(0)
    monthly_display = _format_df_percent(
        _format_df_money(monthly_result, ["Ventas brutas", "Cortesias", "Ventas netas", "Gastos", "Resultado estimado"]),
        ["% cortesias", "Variacion ventas %"],
    )
    story += _section("Pagina 2: Ventas, cortesias y resultado mensual", styles)
    story.append(_pdf_table(monthly_display, max_rows=8))
    story.append(Spacer(1, 0.08 * inch))
    story.append(MultiBarChart(monthly_result["Mes"].astype(str).tolist(), [
        ("Ventas brutas", monthly_result["Ventas brutas"].tolist(), PRIMARY),
        ("Ventas netas", monthly_result["Ventas netas"].tolist(), colors.HexColor("#2563EB")),
        ("Cortesias", monthly_result["Cortesias"].tolist(), colors.HexColor("#D97706")),
    ], height=140))
    story.append(PageBreak())

    recaudo_table = real_payments[["Forma de pago", "Valor"]].copy() if not real_payments.empty else pd.DataFrame(columns=["Forma de pago", "Valor"])
    recaudo_table["Participacion recaudo %"] = recaudo_table["Valor"] / real_collected_total if real_collected_total else 0
    recaudo_table = pd.concat([recaudo_table, pd.DataFrame({"Forma de pago": ["Total recaudo real"], "Valor": [real_collected_total], "Participacion recaudo %": [1 if real_collected_total else 0]})], ignore_index=True)
    unpaid_table = pd.DataFrame({"Concepto": ["Ventas no cobradas: cortesias/descuentos"], "Valor": [courtesy_total], "% ventas brutas": [courtesy_pct]})
    cash_share = real_payments.loc[real_payments["Forma de pago"].astype(str).str.lower().str.contains("efectivo"), "Valor"].sum() / real_collected_total if real_collected_total else 0
    story += _section("Pagina 3: Recaudo y control de caja", styles)
    story.append(_pdf_table(_format_df_percent(_format_df_money(recaudo_table, ["Valor"]), ["Participacion recaudo %"]), max_rows=6))
    story.append(Spacer(1, 0.08 * inch))
    story.append(_pdf_table(_format_df_percent(_format_df_money(unpaid_table, ["Valor"]), ["% ventas brutas"]), max_rows=2))
    story.append(_paragraph(
        "El efectivo supera el 50% del recaudo. Se recomienda arqueo diario, conciliacion de caja y validacion contra ventas registradas."
        if cash_share > 0.5 else "Concilie diariamente efectivo, transferencias, tarjeta y credito contra el cierre de ventas.",
        styles,
        "Alert" if cash_share > 0.5 else "Body",
    ))
    story.append(PageBreak())

    seller_pdf = sellers.copy()
    if seller_courtesy_col:
        seller_pdf["% cortesias"] = seller_pdf[seller_courtesy_col] / seller_pdf["Total ventas"].replace(0, pd.NA)
    else:
        seller_pdf["Cortesias"] = 0
        seller_courtesy_col = "Cortesias"
        seller_pdf["% cortesias"] = 0
    seller_pdf = seller_pdf.sort_values("Total ventas", ascending=False).head(5)
    seller_pdf = seller_pdf[["Vendedor", "Total ventas", seller_courtesy_col, "% cortesias", "Total propinas", "Numero de facturas", "Ticket promedio"]]
    seller_pdf = seller_pdf.rename(columns={"Total ventas": "Ventas", seller_courtesy_col: "Cortesias", "Total propinas": "Propinas", "Numero de facturas": "Facturas"})
    story += _section("Pagina 4: Vendedores", styles)
    story.append(_pdf_table(_format_df_percent(_format_df_money(seller_pdf, ["Ventas", "Cortesias", "Propinas", "Ticket promedio"]), ["% cortesias"]), max_rows=5))
    sellers_for_comment = sellers.rename(columns={seller_courtesy_col: "Total cortesias"}) if seller_courtesy_col in sellers.columns else sellers
    story.append(_paragraph(_seller_control_comment(sellers_for_comment), styles))
    story.append(PageBreak())

    general, identified, client_comment = _client_sections(clients)
    client_cols = [col for col in ["Cliente", "Numero de compras", "Total vendido", "Ticket promedio", "Ultima fecha de compra"] if col in clients.columns]
    story += _section("Pagina 5: Clientes y recomendaciones", styles)
    if not general.empty:
        story.append(_paragraph("Cliente General", styles, "SectionTitle"))
        story.append(_pdf_table(_format_df_money(general[client_cols].head(2), ["Total vendido", "Ticket promedio"]), max_rows=2))
    story.append(_paragraph("Top 5 clientes identificados", styles, "SectionTitle"))
    identified_pdf = _format_df_money(identified[client_cols].head(5), ["Total vendido", "Ticket promedio"]) if not identified.empty else pd.DataFrame({"Mensaje": ["No hay clientes identificados para mostrar."]})
    story.append(_pdf_table(identified_pdf, max_rows=5))
    story.append(_paragraph(client_comment, styles, "Alert" if not general.empty else "Body"))
    story += _section("Recomendaciones accionables", styles)
    for idx, rec in enumerate(_manager_recommendations(monthly_result, payments, operating, clients, gross_sales_total, courtesy_total, expenses_total), start=1):
        story.append(_paragraph(f"{idx}. {rec}", styles))

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
