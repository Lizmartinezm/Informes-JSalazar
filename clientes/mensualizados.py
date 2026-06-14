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
from utils.tlg_data_cleaning import load_tlg_trial_balance
from utils.ui_components import info_panel, process_steps, section_header


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
    story.append(
        Paragraph(
            "Las hojas BCE y P Y G se ocultan en la descarga para entregar un archivo final sin plantillas de ejemplo ni vínculos externos.",
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
                [2024, 2025, 2026, 2027],
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
    st.subheader("Vista previa ejecutiva")
    st.dataframe(report["periods"], use_container_width=True, hide_index=True)
    st.caption("La descarga final oculta las hojas técnicas y entrega solo el resultado terminado.")

    st.download_button(
        "Descargar Excel final",
        data=export_monthly_reports(report),
        file_name="Informes_mensualizados_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    try:
        pdf_bytes = _build_pdf(report)
    except Exception as exc:
        st.warning(f"No se pudo preparar el PDF ejecutivo: {exc}")
    else:
        st.download_button(
            "Descargar PDF ejecutivo",
            data=pdf_bytes,
            file_name="Informes_mensualizados_resumen.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
