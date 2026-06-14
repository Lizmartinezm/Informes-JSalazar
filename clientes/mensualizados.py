from __future__ import annotations

import streamlit as st

from utils.monthly_reports import build_monthly_reports, export_monthly_reports
from utils.ui_components import info_panel, process_steps, section_header


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
    st.subheader("Control de periodos procesados")
    st.dataframe(report["periods"], use_container_width=True, hide_index=True)

    st.download_button(
        "Descargar Informes mensualizados",
        data=export_monthly_reports(report),
        file_name="Informes_mensualizados_actualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
