from __future__ import annotations

import streamlit as st

from utils.monthly_reports import build_monthly_reports, export_monthly_reports
from utils.ui_components import section_header


def render_monthly_reports() -> None:
    section_header(
        "Informes mensualizados",
        "Actualiza el Balance y el PYG en el formato mensualizado definido. "
        "Puedes cargar uno o varios balances de prueba por tercero al mismo tiempo.",
    )

    has_previous = st.radio(
        "¿Tienes una versión acumulada anterior?",
        ["Sí", "No"],
        horizontal=True,
        key="monthly_has_previous",
    )

    previous_file = None
    if has_previous == "Sí":
        previous_file = st.file_uploader(
            "Subir informe mensualizado acumulado anterior",
            type=["xlsx"],
            key="monthly_previous_uploader",
            help="Debe ser el archivo con las hojas BCE y P Y G generado o actualizado anteriormente.",
        )
        st.caption(
            "Se conservarán los meses existentes y solo se reemplazarán los periodos de los balances nuevos."
        )
    else:
        st.info(
            "El informe comenzará desde la plantilla base que suministraste, conservando su formato."
        )

    monthly_files = st.file_uploader(
        "Subir uno o varios balances de prueba por tercero",
        type=["xlsx"],
        accept_multiple_files=True,
        key="monthly_trial_balances_uploader",
        help="Puedes seleccionar varios archivos de meses diferentes en una sola carga.",
    )

    if not monthly_files:
        st.info("Carga al menos un balance del mes que deseas crear o actualizar.")
        return
    if has_previous == "Sí" and previous_file is None:
        st.warning(
            "Indicaste que tienes una versión anterior. Cárgala para conservar los meses acumulados."
        )
        return

    st.caption(f"Archivos seleccionados: {len(monthly_files)}")
    if not st.button(
        "Generar informe mensualizado",
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
    st.subheader("Periodos procesados")
    st.dataframe(report["periods"], use_container_width=True, hide_index=True)

    st.download_button(
        "Descargar Informes mensualizados",
        data=export_monthly_reports(report),
        file_name="Informes_mensualizados_actualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
