from __future__ import annotations

from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from utils.monthly_reports import build_monthly_reports, export_monthly_reports
from utils.tlg_data_cleaning import load_tlg_trial_balance
from utils.ui_components import info_panel, process_steps, section_header


def _money(value: float) -> str:
    return "$" + f"{float(value):,.0f}".replace(",", ".")


def _monthly_metrics(metrics: pd.DataFrame) -> pd.DataFrame:
    return metrics[metrics["Tipo"] == "Mensual"].sort_values(["Año", "Mes"])


def _file_signature(file_obj) -> tuple[str, int] | None:
    if file_obj is None:
        return None
    return (
        str(getattr(file_obj, "name", "archivo.xlsx")),
        int(getattr(file_obj, "size", 0) or 0),
    )


def _report_signature(
    mode: str,
    start_year: int | None,
    initial_balance_file,
    previous_file,
    monthly_files,
) -> tuple[object, ...]:
    return (
        mode,
        start_year,
        _file_signature(initial_balance_file),
        _file_signature(previous_file),
        tuple(_file_signature(file_obj) for file_obj in (monthly_files or [])),
    )


def _style_chart(figure, *, currency_axis: bool = True):
    figure.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#344054", family="Arial"),
        legend_title_text="",
        margin=dict(l=20, r=20, t=55, b=20),
        hoverlabel=dict(bgcolor="#101828", font_color="#FFFFFF"),
    )
    figure.update_xaxes(showgrid=False, linecolor="#D0D5DD")
    figure.update_yaxes(
        showgrid=True,
        gridcolor="#EAECF0",
        linecolor="#D0D5DD",
        tickformat="$,.0f" if currency_axis else None,
    )
    return figure


def _management_findings(metrics: pd.DataFrame) -> list[str]:
    monthly = _monthly_metrics(metrics)
    findings: list[str] = []
    if monthly.empty:
        return ["No hay periodos mensuales suficientes para generar hallazgos."]
    current = monthly.iloc[-1]
    if len(monthly) > 1:
        previous = monthly.iloc[-2]
        previous_income = float(previous["Ventas"])
        variation = (
            (float(current["Ventas"]) - previous_income) / abs(previous_income)
            if previous_income
            else 0
        )
        direction = "aumentaron" if variation >= 0 else "disminuyeron"
        findings.append(
            f"Las ventas {direction} {abs(variation):.1%} frente al periodo anterior."
        )
    margin = (
        float(current["Resultado"]) / float(current["Ventas"])
        if float(current["Ventas"])
        else 0
    )
    findings.append(f"El margen neto estimado del último periodo fue {margin:.1%}.")
    debt = (
        float(current["Pasivos"]) / float(current["Activos"])
        if float(current["Activos"])
        else 0
    )
    findings.append(f"El nivel de endeudamiento estimado fue {debt:.1%} de los activos.")
    current_ratio = (
        float(current["Activos corrientes"]) / float(current["Pasivos corrientes"])
        if float(current["Pasivos corrientes"])
        else 0
    )
    findings.append(f"La razón corriente estimada fue {current_ratio:.2f} veces.")
    if debt > 0.7:
        findings.append("Se recomienda revisar obligaciones y capacidad de pago de corto plazo.")
    elif margin < 0:
        findings.append("Se recomienda revisar costos y gastos que están presionando el resultado.")
    elif current_ratio < 1:
        findings.append("Se recomienda priorizar caja y recuperación de cartera para cubrir obligaciones corrientes.")
    else:
        findings.append("Se recomienda mantener seguimiento mensual a margen, liquidez y endeudamiento.")
    return findings


def _series_value(row: pd.Series, *keys: str, default: object = "") -> object:
    for key in keys:
        if key in row.index:
            return row[key]
    return default


def _build_pdf(report: dict[str, object]) -> bytes:
    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=0.55 * inch,
        rightMargin=0.55 * inch,
        topMargin=0.55 * inch,
        bottomMargin=0.55 * inch,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", fontName="Helvetica", fontSize=8.2, leading=10.2))
    story = [
        Paragraph("Informe mensualizado", styles["Title"]),
        Spacer(1, 0.12 * inch),
        Paragraph("Resumen ejecutivo de los periodos procesados.", styles["BodyText"]),
        Spacer(1, 0.18 * inch),
    ]

    table_data = [["Tipo", "Periodo", "Archivo", "Cuentas leidas", "BCE", "PYG"]]
    for _, row in report["periods"].iterrows():
        table_data.append(
            [
                str(_series_value(row, "Tipo", default="Mensual")),
                str(_series_value(row, "Periodo", default="-")),
                str(_series_value(row, "Archivo", default="-")),
                str(int(_series_value(row, "Cuentas leídas", "Cuentas leidas", default=0))),
                str(int(_series_value(row, "Filas BCE actualizadas", default=0))),
                str(int(_series_value(row, "Filas PYG actualizadas", default=0))),
            ]
        )
    table = Table(
        table_data,
        repeatRows=1,
        colWidths=[0.85 * inch, 0.95 * inch, 2.65 * inch, 0.9 * inch, 0.7 * inch, 0.7 * inch],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B6B57")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LEADING", (0, 0), (-1, -1), 10),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D0D5DD")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (3, 1), (-1, -1), "CENTER"),
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 0.18 * inch))
    metrics = report.get("metrics", pd.DataFrame())
    monthly_metrics = _monthly_metrics(metrics) if not metrics.empty else pd.DataFrame()
    if not monthly_metrics.empty:
        latest = monthly_metrics.iloc[-1]
        sales = float(latest["Ventas"])
        gross_profit = float(latest["Utilidad bruta"])
        operating_profit = float(latest["Utilidad operacional"])
        result = float(latest["Resultado"])
        assets = float(latest["Activos"])
        liabilities = float(latest["Pasivos"])
        margin = result / sales if sales else 0
        debt = liabilities / assets if assets else 0
        current_ratio = (
            float(latest["Activos corrientes"]) / float(latest["Pasivos corrientes"])
            if float(latest["Pasivos corrientes"])
            else 0
        )
        summary_data = [
            ["Indicador", "Valor"],
            ["Ventas", _money(sales)],
            ["Utilidad bruta", _money(gross_profit)],
            ["Utilidad operacional", _money(operating_profit)],
            ["Resultado neto", _money(result)],
            ["Margen neto", f"{margin:.1%}"],
            ["Activos", _money(assets)],
            ["Pasivos", _money(liabilities)],
            ["Endeudamiento", f"{debt:.1%}"],
            ["Razón corriente", f"{current_ratio:.2f}"],
            ["Diferencia de cuadre", _money(float(latest["Diferencia de cuadre"]))],
        ]
        summary_table = Table(summary_data, colWidths=[2.2 * inch, 1.7 * inch])
        summary_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D0D5DD")),
                    ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ]
            )
        )
        story.append(Paragraph("Indicadores del último periodo", styles["Heading2"]))
        story.append(summary_table)
        story.append(Spacer(1, 0.16 * inch))
        story.append(Paragraph("Hallazgos y recomendaciones", styles["Heading2"]))
        for finding in _management_findings(metrics):
            story.append(Paragraph(f"- {finding}", styles["Small"]))
            story.append(Spacer(1, 0.04 * inch))
        receivables = report.get("third_party_receivables", pd.DataFrame()).head(5)
        payables = report.get("third_party_payables", pd.DataFrame()).head(5)
        if not receivables.empty or not payables.empty:
            story.append(Spacer(1, 0.1 * inch))
            story.append(Paragraph("Principales terceros", styles["Heading2"]))
        if not receivables.empty:
            third_party_data = [["Cartera", "Saldo"]]
            third_party_data.extend(
                [str(row["Tercero"])[:42], _money(float(row["Saldo"]))]
                for _, row in receivables.iterrows()
            )
            third_party_table = Table(third_party_data, colWidths=[3.7 * inch, 1.4 * inch])
            third_party_table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B6B57")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D0D5DD")),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
                        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                    ]
                )
            )
            story.append(third_party_table)
            story.append(Spacer(1, 0.1 * inch))
        if not payables.empty:
            third_party_data = [["Obligaciones", "Saldo"]]
            third_party_data.extend(
                [str(row["Tercero"])[:42], _money(float(row["Saldo"]))]
                for _, row in payables.iterrows()
            )
            third_party_table = Table(third_party_data, colWidths=[3.7 * inch, 1.4 * inch])
            third_party_table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#344054")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D0D5DD")),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
                        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                    ]
                )
            )
            story.append(third_party_table)
    story.append(
        Paragraph(
            "El Excel conserva las hojas oficiales BCE y P Y G, sus formatos y sus fórmulas internas, sin vínculos externos.",
            styles["Small"],
        )
    )
    document.build(story)
    return buffer.getvalue()


def _preview_file(file_obj, file_type: str) -> dict[str, str]:
    try:
        file_obj.seek(0)
        _, metadata = load_tlg_trial_balance(file_obj)
        period = f"{metadata.get('mes') or 'N/D'} {metadata.get('anio') or ''}".strip()
    except Exception:
        period = "No se pudo leer"
    finally:
        try:
            file_obj.seek(0)
        except Exception:
            pass
    return {
        "Tipo": file_type,
        "Archivo": getattr(file_obj, "name", "archivo.xlsx"),
        "Periodo detectado": period,
    }


def _render_generated_report(report: dict[str, object]) -> None:
    st.success(
        f"Informe actualizado hasta {report['last_period']}. Base utilizada: {report['source_name']}."
    )
    metrics = report["metrics"]
    monthly_metrics = _monthly_metrics(metrics)
    if monthly_metrics.empty:
        st.warning("No se encontraron periodos mensuales para construir el análisis.")
        return

    latest = monthly_metrics.iloc[-1]
    previous = monthly_metrics.iloc[-2] if len(monthly_metrics) > 1 else None
    sales = float(latest["Ventas"])
    gross_profit = float(latest["Utilidad bruta"])
    operating_profit = float(latest["Utilidad operacional"])
    net_result = float(latest["Resultado"])
    assets = float(latest["Activos"])
    liabilities = float(latest["Pasivos"])
    equity = float(latest["Patrimonio"])
    current_assets = float(latest["Activos corrientes"])
    current_liabilities = float(latest["Pasivos corrientes"])
    gross_margin = gross_profit / sales if sales else 0
    operating_margin = operating_profit / sales if sales else 0
    net_margin = net_result / sales if sales else 0
    debt = liabilities / assets if assets else 0
    current_ratio = current_assets / current_liabilities if current_liabilities else 0
    working_capital = current_assets - current_liabilities
    sales_variation = None
    if previous is not None and float(previous["Ventas"]):
        sales_variation = (
            sales - float(previous["Ventas"])
        ) / abs(float(previous["Ventas"]))

    financial_tab, management_tab, third_party_tab = st.tabs(
        ["Estados Financieros", "Dashboard Gerencial", "Revisión por tercero"]
    )

    with financial_tab:
        check_difference = float(latest["Diferencia de cuadre"])
        tolerance = max(1.0, abs(assets) * 0.000001)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Activo", _money(assets))
        c2.metric("Pasivo", _money(liabilities))
        c3.metric("Patrimonio", _money(equity))
        c4.metric("Diferencia de cuadre", _money(check_difference))
        if abs(check_difference) <= tolerance:
            st.success(
                "Validación contable aprobada: el balance cuadra dentro de la tolerancia."
            )
        else:
            st.error(
                "El balance no cuadra. Revisa el mapeo antes de entregar este informe al cliente."
            )

        st.subheader("Control de archivos procesados")
        st.dataframe(report["periods"], width="stretch", hide_index=True)
        account_additions = report.get("account_additions", pd.DataFrame())
        if not account_additions.empty:
            st.subheader("Cuentas incorporadas automáticamente")
            st.info(
                "Estas cuentas existían en los balances cargados, pero no en la plantilla original. "
                "La app las ubicó en el estado financiero correspondiente."
            )
            st.dataframe(account_additions, width="stretch", hide_index=True)

        balance_tab, pyg_tab = st.tabs(["Balance General", "Estado de Resultados"])
        with balance_tab:
            balance_preview = report["balance_preview"]
            if balance_preview.empty:
                st.warning("La plantilla no contiene filas de cuentas reconocibles para el Balance General.")
            else:
                st.dataframe(
                    balance_preview,
                    width="stretch",
                    hide_index=True,
                    column_config={
                        balance_preview.columns[-1]: st.column_config.NumberColumn(
                            format="$ %.0f"
                        )
                    },
                )
        with pyg_tab:
            pyg_preview = report["pyg_preview"]
            if pyg_preview.empty:
                st.warning("La plantilla no contiene filas de cuentas reconocibles para el Estado de Resultados.")
            else:
                st.dataframe(
                    pyg_preview,
                    width="stretch",
                    hide_index=True,
                    column_config={
                        pyg_preview.columns[-1]: st.column_config.NumberColumn(
                            format="$ %.0f"
                        )
                    },
                )

        st.download_button(
            "Descargar Estados Financieros",
            data=export_monthly_reports(report),
            file_name="Estados_Financieros_mensualizados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
            key="monthly_excel_download",
        )

    with management_tab:
        st.subheader(f"Resumen ejecutivo · {latest['Periodo']}")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric(
            "Ventas",
            _money(sales),
            f"{sales_variation:.1%}" if sales_variation is not None else None,
        )
        k2.metric("Utilidad bruta", _money(gross_profit), f"{gross_margin:.1%}")
        k3.metric(
            "Utilidad operacional",
            _money(operating_profit),
            f"{operating_margin:.1%}",
        )
        k4.metric("Resultado neto", _money(net_result), f"{net_margin:.1%}")

        k5, k6, k7, k8 = st.columns(4)
        k5.metric("EBITDA estimado", _money(float(latest["EBITDA"])))
        k6.metric("Razón corriente", f"{current_ratio:.2f}x")
        k7.metric("Endeudamiento", f"{debt:.1%}")
        k8.metric("Capital de trabajo", _money(working_capital))

        chart_col1, chart_col2 = st.columns(2, gap="large")
        with chart_col1:
            trend = monthly_metrics.melt(
                id_vars=["Periodo"],
                value_vars=["Ventas", "Utilidad operacional", "Resultado"],
                var_name="Indicador",
                value_name="Valor",
            )
            trend_chart = px.line(
                trend,
                x="Periodo",
                y="Valor",
                color="Indicador",
                markers=True,
                title="Evolución mensual de ventas y resultados",
                color_discrete_map={
                    "Ventas": "#0B6B57",
                    "Utilidad operacional": "#D7A53F",
                    "Resultado": "#344054",
                },
            )
            st.plotly_chart(
                _style_chart(trend_chart),
                width="stretch",
                key="monthly_results_chart",
            )
        with chart_col2:
            margins = monthly_metrics[["Periodo", "Ventas", "Utilidad bruta", "Utilidad operacional", "Resultado"]].copy()
            for column, source in [
                ("Margen bruto", "Utilidad bruta"),
                ("Margen operacional", "Utilidad operacional"),
                ("Margen neto", "Resultado"),
            ]:
                margins[column] = margins[source].div(
                    margins["Ventas"].replace(0, pd.NA)
                )
            margin_chart_data = margins.melt(
                id_vars=["Periodo"],
                value_vars=["Margen bruto", "Margen operacional", "Margen neto"],
                var_name="Indicador",
                value_name="Valor",
            )
            margin_chart = px.line(
                margin_chart_data,
                x="Periodo",
                y="Valor",
                color="Indicador",
                markers=True,
                title="Evolución de márgenes",
                color_discrete_map={
                    "Margen bruto": "#0B6B57",
                    "Margen operacional": "#D7A53F",
                    "Margen neto": "#344054",
                },
            )
            margin_chart.update_yaxes(tickformat=".1%")
            st.plotly_chart(
                _style_chart(margin_chart, currency_axis=False),
                width="stretch",
                key="monthly_margin_chart",
            )

        chart_col3, chart_col4 = st.columns(2, gap="large")
        with chart_col3:
            expense_composition = report.get("expense_composition", pd.DataFrame())
            if not expense_composition.empty:
                expense_chart = px.bar(
                    expense_composition.sort_values("Valor"),
                    x="Valor",
                    y="Concepto",
                    orientation="h",
                    title="Composición de costos y gastos",
                    color_discrete_sequence=["#D7A53F"],
                )
                st.plotly_chart(
                    _style_chart(expense_chart),
                    width="stretch",
                    key="monthly_expense_chart",
                )
            else:
                st.info("No hay información suficiente para analizar costos y gastos.")
        with chart_col4:
            composition = pd.DataFrame(
                {
                    "Componente": ["Activos", "Pasivos", "Patrimonio"],
                    "Valor": [assets, liabilities, equity],
                }
            )
            composition_chart = px.bar(
                composition,
                x="Componente",
                y="Valor",
                color="Componente",
                title="Estructura financiera",
                color_discrete_map={
                    "Activos": "#0B6B57",
                    "Pasivos": "#D7A53F",
                    "Patrimonio": "#344054",
                },
            )
            composition_chart.update_layout(showlegend=False)
            st.plotly_chart(
                _style_chart(composition_chart),
                width="stretch",
                key="monthly_structure_chart",
            )

        st.subheader("Indicadores mensuales")
        indicator_columns = [
            "Periodo",
            "Ventas",
            "Utilidad bruta",
            "Utilidad operacional",
            "Resultado",
            "Activos",
            "Pasivos",
            "Diferencia de cuadre",
        ]
        st.dataframe(
            monthly_metrics[indicator_columns],
            width="stretch",
            hide_index=True,
            column_config={
                column: st.column_config.NumberColumn(format="$ %.0f")
                for column in indicator_columns
                if column != "Periodo"
            },
        )

        st.subheader("Lectura ejecutiva")
        for finding in _management_findings(metrics):
            st.info(finding)

        try:
            pdf_bytes = _build_pdf(report)
        except Exception as exc:
            st.warning(f"No se pudo preparar el PDF ejecutivo: {exc}")
        else:
            st.download_button(
                "Descargar Informe Gerencial PDF",
                data=pdf_bytes,
                file_name="Informe_Gerencial_mensualizado.pdf",
                mime="application/pdf",
                width="stretch",
                key="monthly_pdf_download",
            )

    with third_party_tab:
        receivables = report.get("third_party_receivables", pd.DataFrame())
        payables = report.get("third_party_payables", pd.DataFrame())
        activity = report.get("third_party_activity", pd.DataFrame())
        total_receivables = float(receivables["Saldo"].sum()) if not receivables.empty else 0
        total_payables = float(payables["Saldo"].sum()) if not payables.empty else 0
        total_activity = float(activity["Movimiento"].sum()) if not activity.empty else 0

        t1, t2, t3 = st.columns(3)
        t1.metric("Cartera por tercero", _money(total_receivables))
        t2.metric("Obligaciones por tercero", _money(total_payables))
        t3.metric("Movimiento analizado", _money(total_activity))

        receivable_col, payable_col = st.columns(2, gap="large")
        with receivable_col:
            st.subheader("Principales saldos por cobrar")
            if receivables.empty:
                st.info("No se detectaron saldos de cartera por tercero.")
            else:
                top_receivables = receivables.head(10).sort_values("Saldo")
                chart = px.bar(
                    top_receivables,
                    x="Saldo",
                    y="Tercero",
                    orientation="h",
                    color_discrete_sequence=["#0B6B57"],
                )
                st.plotly_chart(
                    _style_chart(chart),
                    width="stretch",
                    key="monthly_receivables_chart",
                )
                st.dataframe(
                    receivables.head(20),
                    width="stretch",
                    hide_index=True,
                    column_config={"Saldo": st.column_config.NumberColumn(format="$ %.0f")},
                )
        with payable_col:
            st.subheader("Principales obligaciones")
            if payables.empty:
                st.info("No se detectaron obligaciones por tercero.")
            else:
                top_payables = payables.head(10).sort_values("Saldo")
                chart = px.bar(
                    top_payables,
                    x="Saldo",
                    y="Tercero",
                    orientation="h",
                    color_discrete_sequence=["#D7A53F"],
                )
                st.plotly_chart(
                    _style_chart(chart),
                    width="stretch",
                    key="monthly_payables_chart",
                )
                st.dataframe(
                    payables.head(20),
                    width="stretch",
                    hide_index=True,
                    column_config={"Saldo": st.column_config.NumberColumn(format="$ %.0f")},
                )

        st.subheader("Terceros con mayor movimiento del periodo")
        if activity.empty:
            st.info("No se detectó movimiento transaccional por tercero.")
        else:
            st.dataframe(
                activity.head(30),
                width="stretch",
                hide_index=True,
                column_config={
                    "Débitos": st.column_config.NumberColumn(format="$ %.0f"),
                    "Créditos": st.column_config.NumberColumn(format="$ %.0f"),
                    "Movimiento": st.column_config.NumberColumn(format="$ %.0f"),
                },
            )


def render_monthly_reports() -> None:
    section_header(
        "Informes mensualizados",
        "Crea o actualiza el informe mensualizado con una base inicial por año o con una versión acumulada anterior.",
        eyebrow="Actualización acumulativa",
        badge="Balance + PYG",
    )
    process_steps(
        [
            "Elige el tipo de inicio",
            "Carga el saldo inicial o la base anterior",
            "Sube uno o varios balances y descarga",
        ]
    )

    mode = st.segmented_control(
        "Tipo de inicio",
        ["Primera vez", "Actualizar base existente"],
        default="Primera vez",
        key="monthly_start_mode",
        label_visibility="collapsed",
    )

    left, right = st.columns([1, 1], gap="large")
    initial_balance_file = None
    previous_file = None
    start_year = None

    with left:
        st.subheader("1. Base del informe")
        if mode == "Primera vez":
            start_year = st.selectbox(
                "Año con el que inicia el informe",
                [2024, 2025, 2026],
                index=1,
                key="monthly_start_year",
            )
            initial_balance_file = st.file_uploader(
                "Archivo de saldo inicial",
                type=["xlsx"],
                key="monthly_opening_balance_uploader",
                help="Carga el balance que servirá como punto de arranque del año base.",
            )
            info_panel(
                "Arranque por año",
                "Este modo inicia el informe desde un año base. El archivo inicial no se muestra como mes, sino como punto de partida del periodo.",
            )
        else:
            previous_file = st.file_uploader(
                "Informe mensualizado anterior",
                type=["xlsx"],
                key="monthly_previous_uploader",
                help="Debe contener la hoja de resumen final generada anteriormente.",
            )
            info_panel(
                "Base acumulada",
                "Este modo toma un informe ya iniciado y solo le agrega nuevos periodos.",
            )

    with right:
        st.subheader("2. Balances a procesar")
        monthly_files = st.file_uploader(
            "Subir uno o varios balances de prueba por tercero",
            type=["xlsx"],
            accept_multiple_files=True,
            key="monthly_trial_balances_uploader",
            help="Selecciona todos los meses que deseas crear o actualizar.",
        )
        if monthly_files:
            info_panel(
                "Carga recibida",
                f"Se detectaron {len(monthly_files)} archivo(s) listos para procesar.",
            )

    preview_rows = []
    if initial_balance_file is not None:
        preview_rows.append(_preview_file(initial_balance_file, "Saldo inicial"))
    if monthly_files:
        preview_rows.extend(_preview_file(uploaded, "Balance mensual") for uploaded in monthly_files)

    if preview_rows:
        st.subheader("Vista previa")
        st.dataframe(pd.DataFrame(preview_rows), width="stretch", hide_index=True)
        k1, k2, k3 = st.columns(3)
        k1.metric("Archivos listos", len(preview_rows))
        k2.metric("Modo", "Primera vez" if mode == "Primera vez" else "Actualizar")
        k3.metric("Año base", start_year if start_year else "N/A")

    if not monthly_files:
        st.info("Carga al menos un balance que deseas crear o actualizar.")
        return
    if mode == "Primera vez" and initial_balance_file is None:
        st.warning("En la primera vez debes cargar el archivo de saldo inicial.")
        return
    if mode == "Actualizar base existente" and previous_file is None:
        st.warning("Indicaste que tienes una base existente. Cárgala para conservar los meses acumulados.")
        return

    current_signature = _report_signature(
        mode,
        start_year,
        initial_balance_file,
        previous_file,
        monthly_files,
    )
    if st.session_state.get("monthly_report_signature") != current_signature:
        st.session_state.pop("monthly_generated_report", None)
        st.session_state["monthly_report_signature"] = current_signature

    st.divider()
    st.subheader("3. Generación, análisis y descarga")
    generate = st.button(
        "Procesar balances y generar informe",
        type="primary",
        width="stretch",
        key="monthly_generate_button",
    )
    if generate:
        try:
            with st.spinner("Conciliando cuentas y construyendo el informe gerencial..."):
                report = build_monthly_reports(
                    monthly_files,
                    previous_file=previous_file if mode == "Actualizar base existente" else None,
                    initial_balance_file=initial_balance_file if mode == "Primera vez" else None,
                    start_year=start_year,
                )
                st.session_state["monthly_generated_report"] = report
        except Exception as exc:
            st.session_state.pop("monthly_generated_report", None)
            st.error(f"No fue posible generar el informe mensualizado: {exc}")

    report = st.session_state.get("monthly_generated_report")
    if report is not None:
        _render_generated_report(report)
