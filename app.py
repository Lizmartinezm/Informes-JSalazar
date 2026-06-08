from __future__ import annotations

import pandas as pd
import streamlit as st

from utils.charts import (
    empty_chart,
    expenses_by_month,
    monthly_profit,
    monthly_sales_expenses,
    payment_distribution,
    seller_sales,
    tips_by_month,
)
from utils.data_cleaning import MONTHS_ES, load_expenses_data, load_sales_data, validate_excel_file
from utils.report_generator import (
    automatic_interpretation,
    build_excel_report,
    build_pdf_report,
    format_cop,
    format_percent,
)


st.set_page_config(page_title="Informe Gerencial - Restaurante Sazón", layout="wide")


st.markdown(
    """
    <style>
    .main .block-container {padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1320px;}
    .metric-card {
        border: 1px solid #E5E7EB; border-radius: 8px; padding: 14px 16px; background: #FFFFFF;
        min-height: 96px; box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }
    .metric-label {color: #475569; font-size: 0.82rem; margin-bottom: 6px;}
    .metric-value {color: #111827; font-size: 1.35rem; font-weight: 700; line-height: 1.2;}
    .section-title {margin-top: 1.5rem; padding-top: 0.25rem;}
    div[data-testid="stDataFrame"] {border: 1px solid #E5E7EB; border-radius: 8px;}
    </style>
    """,
    unsafe_allow_html=True,
)


def kpi_card(label: str, value: str) -> None:
    st.markdown(
        f"<div class='metric-card'><div class='metric-label'>{label}</div><div class='metric-value'>{value}</div></div>",
        unsafe_allow_html=True,
    )


def money_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy()
    for col in columns:
        if col in out.columns:
            out[col] = out[col].map(format_cop)
    return out


def percent_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy()
    for col in columns:
        if col in out.columns:
            out[col] = out[col].map(format_percent)
    return out


def build_dashboard_tables(sales: pd.DataFrame, expenses: pd.DataFrame, expense_details: pd.DataFrame):
    monthly_sales = (
        sales.groupby(["NUMERO_MES", "MES"], as_index=False)
        .agg(
            Ventas=("VALOR FACTURA", "sum"),
            **{
                "Valor neto recibido": ("VALOR NETO", "sum"),
                "Propinas": ("PROPINA", "sum"),
                "Número de facturas": ("NUMERO FACTURA", "nunique"),
            },
        )
        .rename(columns={"MES": "Mes"})
    )
    monthly = monthly_sales.merge(expenses.rename(columns={"MES": "Mes", "GASTOS": "Gastos"}), on=["NUMERO_MES", "Mes"], how="outer")
    monthly = monthly.fillna(0).sort_values("NUMERO_MES")
    monthly["Utilidad estimada"] = monthly["Ventas"] - monthly["Gastos"]
    monthly["Margen estimado"] = monthly["Utilidad estimada"] / monthly["Ventas"].replace(0, pd.NA)
    monthly["Margen estimado"] = monthly["Margen estimado"].fillna(0)
    monthly["Ticket promedio"] = monthly["Ventas"] / monthly["Número de facturas"].replace(0, pd.NA)
    monthly["Ticket promedio"] = monthly["Ticket promedio"].fillna(0)

    payment_values = {
        "Efectivo": sales["EFECTIVO"].sum(),
        "Tarjeta": sales["TARJETA"].sum(),
        "Transferencia": sales["TRANSFERENCIA"].sum(),
        "Crédito": sales["CREDITO"].sum(),
        "Cortesía": sales["CORTESIA"].sum(),
    }
    payments = pd.DataFrame({"Forma de pago": payment_values.keys(), "Valor": payment_values.values()})
    payment_total = payments["Valor"].sum()
    payments["Participación %"] = payments["Valor"] / payment_total if payment_total else 0
    payments = payments.sort_values("Valor", ascending=False)

    sellers = (
        sales.groupby("ATENDIO", as_index=False)
        .agg(
            **{
                "Total ventas": ("VALOR FACTURA", "sum"),
                "Total valor neto": ("VALOR NETO", "sum"),
                "Total propinas": ("PROPINA", "sum"),
                "Número de facturas": ("NUMERO FACTURA", "nunique"),
            }
        )
        .rename(columns={"ATENDIO": "Vendedor"})
    )
    sellers["Ticket promedio"] = sellers["Total ventas"] / sellers["Número de facturas"].replace(0, pd.NA)
    sellers["Ticket promedio"] = sellers["Ticket promedio"].fillna(0)
    sellers = sellers.sort_values("Total ventas", ascending=False)

    client_key = "CLIENTE" if sales["CLIENTE"].nunique() > 1 else "ID CLIENTE"
    clients = (
        sales.groupby(client_key, as_index=False)
        .agg(
            **{
                "Número de compras": ("NUMERO FACTURA", "nunique"),
                "Total vendido": ("VALOR FACTURA", "sum"),
                "Última fecha de compra": ("FECHA", "max"),
            }
        )
        .rename(columns={client_key: "Cliente"})
    )
    clients["Ticket promedio"] = clients["Total vendido"] / clients["Número de compras"].replace(0, pd.NA)
    clients["Ticket promedio"] = clients["Ticket promedio"].fillna(0)
    clients = clients.sort_values("Total vendido", ascending=False)

    tips_month = sales.groupby(["NUMERO_MES", "MES"], as_index=False)["PROPINA"].sum().rename(columns={"MES": "Mes", "PROPINA": "Propinas"})
    tips_seller = sales.groupby("ATENDIO", as_index=False)["PROPINA"].sum().rename(columns={"ATENDIO": "Vendedor", "PROPINA": "Propinas"})
    tips = tips_month.merge(pd.DataFrame({"_": []}), how="left") if False else tips_month

    provider_summary = pd.DataFrame()
    if not expense_details.empty and "PROVEEDOR_CONCEPTO" in expense_details.columns:
        provider_summary = (
            expense_details.groupby("PROVEEDOR_CONCEPTO", as_index=False)["GASTOS"]
            .sum()
            .rename(columns={"PROVEEDOR_CONCEPTO": "Proveedor o concepto", "GASTOS": "Gastos"})
            .sort_values("Gastos", ascending=False)
        )

    operating = monthly[["Mes", "Ventas", "Gastos", "Utilidad estimada", "Margen estimado"]].copy()
    return monthly, payments, sellers, clients, tips, tips_seller, provider_summary, operating


def apply_filters(sales: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filtros")
    min_date = sales["FECHA"].min().date()
    max_date = sales["FECHA"].max().date()
    date_range = st.sidebar.date_input("Rango de fechas", value=(min_date, max_date), min_value=min_date, max_value=max_date)

    selected_months = st.sidebar.multiselect("Mes", options=[MONTHS_ES[i] for i in sorted(sales["NUMERO_MES"].unique())])
    selected_sellers = st.sidebar.multiselect("Vendedor", options=sorted(sales["ATENDIO"].dropna().unique()))
    selected_payments = st.sidebar.multiselect("Forma de pago", options=sorted(sales["FORMA DE PAGO"].dropna().unique()))
    selected_clients = st.sidebar.multiselect("Cliente", options=sorted(sales["CLIENTE"].dropna().unique()))

    filtered = sales.copy()
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        filtered = filtered[(filtered["FECHA"] >= start) & (filtered["FECHA"] <= end)]
    if selected_months:
        filtered = filtered[filtered["MES"].isin(selected_months)]
    if selected_sellers:
        filtered = filtered[filtered["ATENDIO"].isin(selected_sellers)]
    if selected_payments:
        filtered = filtered[filtered["FORMA DE PAGO"].isin(selected_payments)]
    if selected_clients:
        filtered = filtered[filtered["CLIENTE"].isin(selected_clients)]
    return filtered


def filtered_expenses(expenses: pd.DataFrame, sales_filtered: pd.DataFrame) -> pd.DataFrame:
    months = sorted(sales_filtered["NUMERO_MES"].unique()) if not sales_filtered.empty else []
    if not months:
        return expenses.iloc[0:0].copy()
    return expenses[expenses["NUMERO_MES"].isin(months)].copy()


st.title("Informe Gerencial de Ventas y Gastos - Restaurante Sazón")
st.caption("Análisis mensual y acumulado del periodo cargado")

col1, col2 = st.columns(2)
with col1:
    sales_file = st.file_uploader("Subir archivo de ventas", type=["xlsx"])
with col2:
    expenses_file = st.file_uploader("Subir archivo de gastos", type=["xlsx"])

validation_errors = [err for err in [validate_excel_file(sales_file, "ventas"), validate_excel_file(expenses_file, "gastos")] if err]
if validation_errors:
    st.info("Por favor cargue el archivo de ventas y el archivo de gastos para generar el informe.")
    for error in validation_errors:
        if "Falta" not in error:
            st.warning(error)
    st.stop()

try:
    sales_df, sales_messages = load_sales_data(sales_file)
    expenses_df, expense_details_df, expense_messages = load_expenses_data(expenses_file)
except ValueError as exc:
    st.error(str(exc))
    st.stop()
except Exception:
    st.error("No fue posible generar el informe. Revise que los archivos tengan datos válidos y vuelva a intentarlo.")
    st.stop()

for message in sales_messages + expense_messages:
    st.sidebar.caption(message)

sales_filtered = apply_filters(sales_df)
if sales_filtered.empty:
    st.warning("No hay ventas para los filtros seleccionados.")
    st.stop()

expenses_filtered = filtered_expenses(expenses_df, sales_filtered)
monthly, payments, sellers, clients, tips, tips_seller, provider_summary, operating = build_dashboard_tables(
    sales_filtered, expenses_filtered, expense_details_df
)

sales_total = sales_filtered["VALOR FACTURA"].sum()
net_total = sales_filtered["VALOR NETO"].sum()
tips_total = sales_filtered["PROPINA"].sum()
expenses_total = expenses_filtered["GASTOS"].sum() if not expenses_filtered.empty else 0
profit_total = sales_total - expenses_total
margin_total = profit_total / sales_total if sales_total else 0
invoice_count = sales_filtered["NUMERO FACTURA"].nunique()
avg_ticket = sales_total / invoice_count if invoice_count else 0
top_sales_month = monthly.sort_values("Ventas", ascending=False)["Mes"].iloc[0] if not monthly.empty else "Sin datos"
top_expense_month = monthly.sort_values("Gastos", ascending=False)["Mes"].iloc[0] if not monthly.empty and monthly["Gastos"].sum() else "Sin datos"
top_payment = payments.iloc[0]["Forma de pago"] if not payments.empty and payments["Valor"].sum() else "Sin datos"
top_payment_share = payments.iloc[0]["Participación %"] if not payments.empty and payments["Valor"].sum() else 0
top_seller = sellers.iloc[0]["Vendedor"] if not sellers.empty else "Sin datos"

kpi_rows = [
    ("Ventas acumuladas", format_cop(sales_total)),
    ("Gastos acumulados", format_cop(expenses_total)),
    ("Utilidad estimada acumulada", format_cop(profit_total)),
    ("Margen estimado acumulado", format_percent(margin_total)),
    ("Valor neto recibido acumulado", format_cop(net_total)),
    ("Total propinas", format_cop(tips_total)),
    ("Número de facturas", f"{invoice_count:,.0f}".replace(",", ".")),
    ("Ticket promedio", format_cop(avg_ticket)),
    ("Mes con mayor venta", top_sales_month),
    ("Forma de pago más usada", top_payment),
    ("Vendedor con mayor venta", top_seller),
]

for i in range(0, len(kpi_rows), 4):
    cols = st.columns(4)
    for col, (label, value) in zip(cols, kpi_rows[i : i + 4]):
        with col:
            kpi_card(label, value)

st.markdown("<h3 class='section-title'>Informe mensual</h3>", unsafe_allow_html=True)
monthly_display = percent_columns(
    money_columns(monthly.drop(columns=["NUMERO_MES"], errors="ignore"), ["Ventas", "Valor neto recibido", "Propinas", "Gastos", "Utilidad estimada", "Ticket promedio"]),
    ["Margen estimado"],
)
st.dataframe(monthly_display, use_container_width=True, hide_index=True)

st.markdown("<h3 class='section-title'>Resultado operativo estimado</h3>", unsafe_allow_html=True)
chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    st.plotly_chart(monthly_sales_expenses(monthly), use_container_width=True)
with chart_col2:
    st.plotly_chart(monthly_profit(monthly), use_container_width=True)
operating_display = percent_columns(money_columns(operating, ["Ventas", "Gastos", "Utilidad estimada"]), ["Margen estimado"])
st.dataframe(operating_display, use_container_width=True, hide_index=True)

st.markdown("<h3 class='section-title'>Informe acumulado</h3>", unsafe_allow_html=True)
accumulated = pd.DataFrame(
    {
        "Indicador": [
            "Total ventas acumuladas",
            "Total gastos acumulados",
            "Utilidad estimada acumulada",
            "Margen estimado acumulado",
            "Total propinas acumuladas",
            "Total recaudo en efectivo",
            "Total recaudo en tarjeta",
            "Total recaudo en transferencia",
            "Total crédito",
            "Total cortesías",
        ],
        "Valor": [
            format_cop(sales_total),
            format_cop(expenses_total),
            format_cop(profit_total),
            format_percent(margin_total),
            format_cop(tips_total),
            format_cop(sales_filtered["EFECTIVO"].sum()),
            format_cop(sales_filtered["TARJETA"].sum()),
            format_cop(sales_filtered["TRANSFERENCIA"].sum()),
            format_cop(sales_filtered["CREDITO"].sum()),
            format_cop(sales_filtered["CORTESIA"].sum()),
        ],
    }
)
st.dataframe(accumulated, use_container_width=True, hide_index=True)
interpretation = automatic_interpretation(
    sales_total,
    expenses_total,
    profit_total,
    margin_total,
    top_sales_month,
    top_expense_month,
    top_payment,
    top_payment_share,
    top_seller,
)
st.success(interpretation)

st.markdown("<h3 class='section-title'>Análisis de formas de pago</h3>", unsafe_allow_html=True)
pay_col1, pay_col2 = st.columns([1, 1])
with pay_col1:
    st.plotly_chart(payment_distribution(payments) if payments["Valor"].sum() else empty_chart(), use_container_width=True)
with pay_col2:
    payments_display = percent_columns(money_columns(payments, ["Valor"]), ["Participación %"])
    st.dataframe(payments_display, use_container_width=True, hide_index=True)

st.markdown("<h3 class='section-title'>Ventas por vendedor</h3>", unsafe_allow_html=True)
seller_col1, seller_col2 = st.columns([1, 1])
with seller_col1:
    st.plotly_chart(seller_sales(sellers) if not sellers.empty else empty_chart(), use_container_width=True)
with seller_col2:
    sellers_display = money_columns(sellers, ["Total ventas", "Total valor neto", "Total propinas", "Ticket promedio"])
    st.dataframe(sellers_display, use_container_width=True, hide_index=True)

st.markdown("<h3 class='section-title'>Análisis de clientes</h3>", unsafe_allow_html=True)
clients_display = clients.head(10).copy()
clients_display["Última fecha de compra"] = pd.to_datetime(clients_display["Última fecha de compra"]).dt.strftime("%Y-%m-%d")
clients_display = money_columns(clients_display, ["Total vendido", "Ticket promedio"])
st.dataframe(clients_display, use_container_width=True, hide_index=True)

st.markdown("<h3 class='section-title'>Análisis de propinas</h3>", unsafe_allow_html=True)
tip_col1, tip_col2 = st.columns([1, 1])
with tip_col1:
    st.plotly_chart(tips_by_month(tips) if not tips.empty else empty_chart(), use_container_width=True)
with tip_col2:
    st.write(f"Propinas totales: **{format_cop(tips_total)}**")
    st.write(f"Propinas en tarjeta: **{format_cop(sales_filtered['PROPINA TARJETA'].sum())}**")
    st.write(f"Propinas en efectivo/resto: **{format_cop(sales_filtered['PROPINA RESTO'].sum())}**")
    st.dataframe(money_columns(tips_seller.sort_values("Propinas", ascending=False), ["Propinas"]), use_container_width=True, hide_index=True)

st.markdown("<h3 class='section-title'>Análisis de gastos</h3>", unsafe_allow_html=True)
g_col1, g_col2 = st.columns([1, 1])
with g_col1:
    st.plotly_chart(expenses_by_month(expenses_filtered) if not expenses_filtered.empty else empty_chart("No hay gastos para los meses filtrados"), use_container_width=True)
with g_col2:
    ratio = expenses_total / sales_total if sales_total else 0
    st.write(f"Gastos acumulados: **{format_cop(expenses_total)}**")
    st.write(f"Mes con mayor gasto: **{top_expense_month}**")
    st.write(f"Relación gastos / ventas: **{format_percent(ratio)}**")
    if not provider_summary.empty:
        st.dataframe(money_columns(provider_summary.head(10), ["Gastos"]), use_container_width=True, hide_index=True)

executive = pd.DataFrame(
    {
        "Indicador": [label for label, _ in kpi_rows],
        "Valor": [value for _, value in kpi_rows],
    }
)
excel_bytes = build_excel_report(
    executive,
    monthly.drop(columns=["NUMERO_MES"], errors="ignore"),
    payments,
    sellers,
    clients.head(50),
    tips,
    expenses_filtered,
    operating,
)
pdf_bytes = build_pdf_report(
    executive,
    monthly_display,
    payments_display,
    sellers_display,
    clients_display,
    money_columns(tips, ["Propinas"]),
    money_columns(
        expenses_filtered.rename(columns={"MES": "Mes", "GASTOS": "Gastos"}).drop(columns=["NUMERO_MES"], errors="ignore"),
        ["Gastos"],
    ),
    operating_display,
    interpretation,
)

st.markdown("<h3 class='section-title'>Descargas</h3>", unsafe_allow_html=True)
download_col1, download_col2 = st.columns(2)
with download_col1:
    st.download_button(
        "Descargar informe en Excel",
        data=excel_bytes,
        file_name="informe_restaurante_sazon.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with download_col2:
    st.download_button(
        "Descargar informe en PDF",
        data=pdf_bytes,
        file_name="informe_restaurante_sazon.pdf",
        mime="application/pdf",
    )
