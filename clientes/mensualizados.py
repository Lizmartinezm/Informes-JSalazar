from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from utils.monthly_reports import build_monthly_reports, export_monthly_reports
from utils.ui_components import info_panel, process_steps, section_header


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

    table_data = [["Periodo", "Archivo", "Cuentas leídas", "BCE", "PYG"]]
    for _, row in report["periods"].iterrows():
        table_data.append(
            [
                str(row["Periodo"]),
                str(row["Archivo"]),
                str(int(row["Cuentas leídas"])),
                str(int(row["Filas BCE actualizadas"])),
                str(int(row["Filas PYG actualizadas"])),
            ]
        )
    table = Table(table_data, repeatRows=1, colWidths=[0.85 * inch, 3.15 * inch, 0.9 * inch, 0.7 * inch, 0.7 * inch])
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
                ("ALIGN", (2, 1), (-1, -1), "CENTER"),
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 0.18 * inch))
    story.append(Paragraph("Las hojas BCE y P Y G se ocultan en la descarga para entregar un archivo final sin plantillas de ejemplo ni vínculos externos.", styles["Small"]))
    document.build(story)
    return buffer.getvalue()


def render_monthly_reports() -> None:
    section_header(
        "Informes mensualizados",
        "Actualiza Balance y PYG sobre la plantilla oficial, conservando los periodos anteriores y el formato de presentación.",
        eyebrow="Actualización acumulativa",
        badge="Balance + PYG",
    )
    process_steps(
        [
            "Selecciona la base acumulada",
            "Carga uno o varios balances",
            "Valida y descarga el informe",
        ]
    )

    left, right = st.columns([0.9, 1.1], gap="large")
    with left:
        st.subheader("1. Base del informe")
        has_previous = st.radio(
            "¿Tienes una versión acumulada anterior?",
            ["Sí", "No"],
            horizontal=True,
            key="monthly_has_previous",
        )

        previous_file = None
        if has_previous == "Sí":
            previous_file = st.file_uploader(
                "Informe mensualizado anterior",
                type=["xlsx"],
                key="monthly_previous_uploader",
                help="Debe contener las hojas BCE y P Y G.",
            )
            info_panel(
                "Actualización controlada",
                "Los meses existentes se conservan. Solo se reemplazan los periodos incluidos en los nuevos balances.",
            )
        else:
            info_panel(
                "Plantilla oficial incluida",
                "La aplicación iniciará desde el formato de Informes mensualizados suministrado como referencia.",
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
            st.caption(f"{len(monthly_files)} archivo(s) listo(s) para procesar")

    if not monthly_files:
        st.info("Carga al menos un balance del mes que deseas crear o actualizar.")
        return
    if has_previous == "Sí" and previous_file is None:
        st.warning(
            "Indicaste que tienes una versión anterior. Cárgala para conservar los meses acumulados."
        )
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
            report = build_monthly_reports(monthly_files, previous_file)
    except Exception as exc:
        st.error(f"No fue posible generar el informe mensualizado: {exc}")
        return

    st.success(
        f"Informe actualizado hasta {report['last_period']}. Base utilizada: {report['source_name']}."
    )
    st.subheader("Vista previa ejecutiva")
    st.dataframe(report["periods"], use_container_width=True, hide_index=True)
    st.caption("La descarga final oculta las hojas técnicas y entrega solo el resultado terminado.")

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(
            "Descargar Excel final",
            data=export_monthly_reports(report),
            file_name="Informes_mensualizados_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with dl_col2:
        st.download_button(
            "Descargar PDF ejecutivo",
            data=_build_pdf(report),
            file_name="Informes_mensualizados_resumen.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
