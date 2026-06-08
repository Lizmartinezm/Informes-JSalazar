from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


COLOR_SEQUENCE = ["#0F766E", "#2563EB", "#D97706", "#7C3AED", "#DC2626", "#475569"]


def money_axis(fig):
    fig.update_layout(
        hovermode="x unified",
        margin=dict(l=10, r=10, t=45, b=10),
        legend_title_text="",
        font=dict(family="Arial", size=13),
    )
    fig.update_yaxes(tickprefix="$", separatethousands=True)
    return fig


def monthly_sales_expenses(monthly: pd.DataFrame):
    chart_data = monthly[["Mes", "Ventas", "Gastos"]].melt("Mes", var_name="Concepto", value_name="Valor")
    fig = px.bar(chart_data, x="Mes", y="Valor", color="Concepto", barmode="group", color_discrete_sequence=COLOR_SEQUENCE)
    return money_axis(fig)


def monthly_profit(monthly: pd.DataFrame):
    fig = px.line(monthly, x="Mes", y="Utilidad estimada", markers=True, color_discrete_sequence=[COLOR_SEQUENCE[0]])
    fig.update_traces(line=dict(width=3))
    return money_axis(fig)


def payment_distribution(payments: pd.DataFrame):
    fig = px.pie(payments, names="Forma de pago", values="Valor", hole=0.42, color_discrete_sequence=COLOR_SEQUENCE)
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10), legend_title_text="")
    return fig


def seller_sales(sellers: pd.DataFrame):
    data = sellers.head(12).sort_values("Total ventas")
    fig = px.bar(data, x="Total ventas", y="Vendedor", orientation="h", color_discrete_sequence=[COLOR_SEQUENCE[1]])
    return money_axis(fig)


def tips_by_month(tips: pd.DataFrame):
    fig = px.bar(tips, x="Mes", y="Propinas", color_discrete_sequence=[COLOR_SEQUENCE[2]])
    return money_axis(fig)


def expenses_by_month(expenses: pd.DataFrame):
    fig = px.bar(expenses, x="MES", y="GASTOS", color_discrete_sequence=[COLOR_SEQUENCE[4]])
    return money_axis(fig)


def empty_chart(message: str = "Sin datos disponibles"):
    fig = go.Figure()
    fig.add_annotation(text=message, showarrow=False, x=0.5, y=0.5, xref="paper", yref="paper")
    fig.update_layout(height=280, margin=dict(l=10, r=10, t=20, b=10))
    return fig
