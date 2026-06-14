from __future__ import annotations

import pandas as pd
import streamlit as st

from utils.monthly_reports import build_monthly_reports, export_monthly_reports
from utils.ui_components import section_header


def _money(value: float) -> str:
    return "$" + f"{float(value):,.0f}".replace(",", ".")


def _format_table(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    fixed = {"CODIGO_CUENTA", "NOMBRE_CUENTA", "CLASE", "CLASIFICACION", "Indicador"}
    for col in out.columns:
        if col not in fixed:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).map(_money)
    return out


def render_monthly_reports() -> None:
    section_header(
        "Informes mensualizados",
        "Construye Balance y PYG mes a mes. Puedes cargar una version acumulada anterior para actualizar solamente "
        "el mes que necesitas, sin volver a subir todos los balances previos.",
    )

    monthly_file = st.file_uploader(
        "Subir balance de prueba por tercero del mes a actualizar",
        type=["xlsx"],
        key="monthly_trial_balance_uploader",
    )
    has_previous = st.radio(
        "¿Tienes una version acumulada anterior?",
        ["Si", "No"],
        horizontal=True,
        key="monthly_has_previous",
    )
    previous_file = None
    if has_previous == "Si":
        previous_file = st.file_uploader(
            "Subir version acumulada anterior o balance acumulado anterior",
            type=["xlsx"],
            key="monthly_previous_uploader",
        )
        st.caption(
            "Puedes cargar un Excel mensualizado generado anteriormente o un Balance de prueba por tercero acumulado."
        )

    if monthly_file is None:
        st.info("Carga el balance del mes que deseas actualizar.")
        return
    if has_previous == "Si" and previous_file is None:
        st.warning("Indicaste que tienes una version anterior. Cargala para conservar los periodos previos.")
        return

    try:
        report = build_monthly_reports(monthly_file, previous_file)
    except Exception as exc:
        st.error(f"No fue posible generar los informes mensualizados: {exc}")
        return

    metadata = report["metadata"]
    st.success(
        f"Periodo detectado: {metadata.get('periodo') or report['label']}. "
        f"{report['previous_message']}"
    )

    summary_tab, balance_tab, pyg_tab, download_tab = st.tabs(
        ["Resumen", "Balance mensualizado", "PYG mensualizado", "Descarga"]
    )

    with summary_tab:
        st.subheader(f"Resumen - {report['label']}")
        st.dataframe(_format_table(report["summary"]), use_container_width=True, hide_index=True)
        values = dict(zip(report["summary"]["Indicador"], report["summary"][report["label"]]))
        cols = st.columns(4)
        cols[0].metric("Activo", _money(values.get("Total activo", 0)))
        cols[1].metric("Pasivo", _money(values.get("Total pasivo", 0)))
        cols[2].metric("Patrimonio", _money(values.get("Total patrimonio", 0)))
        cols[3].metric("Resultado mes", _money(values.get("Resultado del mes", 0)))
        difference = values.get("Diferencia balance", 0)
        if abs(difference) > 1000:
            st.warning(f"Diferencia de balance detectada: {_money(difference)}.")
        else:
            st.success("La diferencia de balance esta dentro de un margen razonable.")

    with balance_tab:
        st.subheader("Balance mensualizado")
        st.dataframe(_format_table(report["balance"]), use_container_width=True, hide_index=True)

    with pyg_tab:
        st.subheader("Estado de Resultados mensualizado")
        st.dataframe(_format_table(report["pyg"]), use_container_width=True, hide_index=True)

    with download_tab:
        st.subheader("Descargar version actualizada")
        output = export_monthly_reports(report)
        st.download_button(
            "Descargar Balance y PYG mensualizados",
            data=output,
            file_name=f"Informes_mensualizados_{report['label'].replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
