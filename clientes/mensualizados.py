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


def _management_findings(metrics: pd.DataFrame) -> list[str]:
    monthly = metrics[metrics["Tipo"] == "Mensual"].sort_values(["Año", "Mes"])
    findings: list[str] = []
    if monthly.empty:
        return ["No hay periodos mensuales suficientes para generar hallazgos."]
    current = monthly.iloc[-1]
    if len(monthly) > 1:
        previous = monthly.iloc[-2]
        previous_income = float(previous["Ingresos"])
        variation = (
            (float(current["Ingresos"]) - previous_income) / abs(previous_income)
            if previous_income
            else 0
        )
        direction = "aumentaron" if variation >= 0 else "disminuyeron"
        findings.append(
            f"Los ingresos {direction} {abs(variation):.1%} frente al periodo anterior."
        )
    margin = (
        float(current["Resultado"]) / float(current["Ingresos"])
        if float(current["Ingresos"])
        else 0
    )
    findings.append(f"El margen neto estimado del último periodo fue {margin:.1%}.")
    debt = (
        float(current["Pasivos"]) / float(current["Activos"])
        if float(current["Activos"])
        else 0
    )
    findings.append(f"El nivel de endeudamiento estimado fue {debt:.1%} de los activos.")
    if debt > 0.7:
        findings.append("Se recomienda revisar obligaciones y capacidad de pago de corto plazo.")
    elif margin < 0:
        findings.append("Se recomienda revisar costos y gastos que están presionando el resultado.")
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
    monthly_metrics = metrics[metrics["Tipo"] == "Mensual"].sort_values(
        ["Año", "Mes"]
    ) if not metrics.empty else pd.DataFrame()
    if not monthly_metrics.empty:
        latest = monthly_metrics.iloc[-1]
        income = float(latest["Ingresos"])
        result = float(latest["Resultado"])
        assets = float(latest["Activos"])
        liabilities = float(latest["Pasivos"])
        margin = result / income if income else 0
        debt = liabilities / assets if assets else 0
        summary_data = [
            ["Indicador", "Valor"],
            ["Ingresos", _money(income)],
            ["Resultado neto", _money(result)],
            ["Margen neto", f"{margin:.1%}"],
            ["Activos", _money(assets)],
            ["Pasivos", _money(liabilities)],
            ["Endeudamiento", f"{debt:.1%}"],
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
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
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

    st.divider()
    st.subheader("3. Generación y descarga")
    if not st.button(
        "Procesar balances y generar informe",
        type="primary",
        use_container_width=True,
        key="monthly_generate_button",
    ):
        return

    try:
        with st.spinner("Actualizando la plantilla mes a mes..."):
            report = build_monthly_reports(
                monthly_files,
                previous_file=previous_file if mode == "Actualizar base existente" else None,
                initial_balance_file=initial_balance_file if mode == "Primera vez" else None,
                start_year=start_year,
            )
    except Exception as exc:
        st.error(f"No fue posible generar el informe mensualizado: {exc}")
        return

    st.success(
        f"Informe actualizado hasta {report['last_period']}. Base utilizada: {report['source_name']}."
    )
    financial_tab, management_tab = st.tabs(
        ["Estados Financieros", "Informe Gerencial"]
    )

    with financial_tab:
        st.subheader("Control de archivos procesados")
        st.dataframe(report["periods"], use_container_width=True, hide_index=True)

        balance_tab, pyg_tab = st.tabs(["Balance General", "Estado de Resultados"])
        with balance_tab:
            st.dataframe(
                report["balance_preview"],
                use_container_width=True,
                hide_index=True,
                column_config={
                    report["balance_preview"].columns[-1]: st.column_config.NumberColumn(
                        format="$ %.0f"
                    )
                },
            )
        with pyg_tab:
            st.dataframe(
                report["pyg_preview"],
                use_container_width=True,
                hide_index=True,
                column_config={
                    report["pyg_preview"].columns[-1]: st.column_config.NumberColumn(
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
            use_container_width=True,
        )

    with management_tab:
        metrics = report["metrics"]
        monthly_metrics = metrics[metrics["Tipo"] == "Mensual"].sort_values(
            ["Año", "Mes"]
        )
        latest = monthly_metrics.iloc[-1]
        income = float(latest["Ingresos"])
        result = float(latest["Resultado"])
        assets = float(latest["Activos"])
        liabilities = float(latest["Pasivos"])
        margin = result / income if income else 0
        debt = liabilities / assets if assets else 0
        working_capital = assets - liabilities

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Ingresos", _money(income))
        k2.metric("Resultado neto", _money(result))
        k3.metric("Margen neto", f"{margin:.1%}")
        k4.metric("Endeudamiento", f"{debt:.1%}")
        k5, k6, k7 = st.columns(3)
        k5.metric("Activos", _money(assets))
        k6.metric("Pasivos", _money(liabilities))
        k7.metric("Capital de trabajo", _money(working_capital))

        chart_col1, chart_col2 = st.columns(2)
        with chart_col1:
            trend = monthly_metrics.melt(
                id_vars=["Periodo"],
                value_vars=["Ingresos", "Resultado"],
                var_name="Indicador",
                value_name="Valor",
            )
            trend_chart = px.line(
                trend,
                x="Periodo",
                y="Valor",
                color="Indicador",
                markers=True,
                title="Evolución de ingresos y resultado",
                color_discrete_map={"Ingresos": "#0B6B57", "Resultado": "#C68A2D"},
            )
            trend_chart.update_layout(
                legend_title_text="",
                yaxis_title="Valor",
                xaxis_title="",
                margin=dict(l=20, r=20, t=55, b=20),
            )
            st.plotly_chart(trend_chart, use_container_width=True)
        with chart_col2:
            composition = pd.DataFrame(
                {
                    "Componente": ["Activos", "Pasivos", "Patrimonio"],
                    "Valor": [
                        float(latest["Activos"]),
                        float(latest["Pasivos"]),
                        float(latest["Patrimonio"]),
                    ],
                }
            )
            composition_chart = px.bar(
                composition,
                x="Componente",
                y="Valor",
                color="Componente",
                title="Estructura financiera",
                color_discrete_sequence=["#0B6B57", "#C68A2D", "#344054"],
            )
            composition_chart.update_layout(
                showlegend=False,
                xaxis_title="",
                yaxis_title="Valor",
                margin=dict(l=20, r=20, t=55, b=20),
            )
            st.plotly_chart(composition_chart, use_container_width=True)

        st.subheader("Hallazgos y recomendaciones")
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
                use_container_width=True,
            )
