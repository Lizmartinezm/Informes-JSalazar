from __future__ import annotations

import pandas as pd
import streamlit as st

from utils.tlg_data_cleaning import load_tlg_trial_balance, validate_tlg_company
from utils.tlg_excel_writer import build_tlg_management_pdf, build_tlg_summary_excel, update_tlg_financial_statements
from utils.tlg_financial_statements import build_tlg_financial_summary, build_tlg_management_text
from utils.ui_components import section_header


def _money(value: float) -> str:
    return "$" + f"{float(value):,.0f}".replace(",", ".")


def _metric_grid(metrics: dict[str, float]) -> None:
    rows = [
        ("Total activos", metrics.get("total_activo", 0)),
        ("Total pasivos", metrics.get("total_pasivo", 0)),
        ("Total patrimonio", metrics.get("total_patrimonio", 0)),
        ("Total ingresos", metrics.get("total_ingresos", 0)),
        ("Total costos", metrics.get("total_costos", 0)),
        ("Total gastos", metrics.get("total_gastos", 0)),
        ("Resultado periodo", metrics.get("resultado", 0)),
        ("Diferencia cuadre", metrics.get("diferencia_cuadre", 0)),
    ]
    for start in range(0, len(rows), 4):
        cols = st.columns(4)
        for col, (label, value) in zip(cols, rows[start : start + 4]):
            col.metric(label, _money(value))


def _format_money_df(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    for col in cols:
        if col in out.columns:
            out[col] = out[col].map(_money)
    return out


def render_tlg() -> None:
    section_header(
        "Estados Financieros TLG",
        "Carga el Balance de prueba por tercero de TLG para validar la informacion, clasificar cuentas y preparar "
        "Estados Financieros iniciales sin afectar el informe de Restaurante Sazon.",
    )

    tab_load, tab_validation, tab_balance, tab_income, tab_summary, tab_downloads = st.tabs(
        [
            "Carga de balance",
            "Validacion",
            "Estado de Situacion Financiera",
            "Estado de Resultados",
            "Resumen gerencial",
            "Descargas",
        ]
    )

    with tab_load:
        st.subheader("Carga de balance")
        balance_file = st.file_uploader(
            "Subir Balance de prueba por tercero",
            type=["xlsx"],
            key="tlg_balance_uploader",
        )
        template_file = st.file_uploader(
            "Subir plantilla Estados Financieros TLG 2026 (opcional)",
            type=["xlsx"],
            key="tlg_template_uploader",
        )
        confirmed = st.checkbox("Confirmo visualmente que el archivo cargado pertenece a TLG")

        if template_file is not None:
            st.session_state["tlg_template_file"] = template_file

        if balance_file:
            try:
                st.session_state["tlg_balance_file"] = balance_file
                df, metadata = load_tlg_trial_balance(balance_file)
                valid_company, validation_message = validate_tlg_company(metadata, confirmed)
                st.session_state["tlg_df"] = df
                st.session_state["tlg_metadata"] = metadata
                st.session_state["tlg_valid_company"] = valid_company
                st.session_state["tlg_validation_message"] = validation_message
                if valid_company:
                    summary = build_tlg_financial_summary(df)
                    st.session_state["tlg_summary"] = summary
                    st.success("Balance cargado y procesado correctamente.")
                else:
                    st.error(validation_message)
            except Exception as exc:
                st.error(f"No fue posible leer el balance de TLG: {exc}")
        else:
            st.info("Carga el balance de prueba por tercero para iniciar.")

    df = st.session_state.get("tlg_df")
    metadata = st.session_state.get("tlg_metadata", {})
    summary = st.session_state.get("tlg_summary")
    valid_company = st.session_state.get("tlg_valid_company", False)

    with tab_validation:
        st.subheader("Validacion del archivo")
        if df is None:
            st.info("Primero carga un balance en la pestaña Carga de balance.")
        else:
            st.write(st.session_state.get("tlg_validation_message", ""))
            validation_rows = {
                "Empresa detectada": metadata.get("empresa") or "No detectada",
                "NIT detectado": metadata.get("nit") or "No detectado",
                "Periodo detectado": metadata.get("periodo") or "No detectado",
                "Mes detectado": metadata.get("mes") or "No detectado",
                "Numero de cuentas leidas": df["CODIGO_CUENTA"].nunique(),
                "Numero de terceros leidos": df["NOMBRE_TERCERO"].nunique(),
                "Total saldo inicial": _money(df["SALDO_INICIAL"].sum()),
                "Total movimientos debito": _money(df["MOVIMIENTO_DEBITO"].sum()),
                "Total movimientos credito": _money(df["MOVIMIENTO_CREDITO"].sum()),
                "Total saldo final": _money(df["SALDO_FINAL"].sum()),
            }
            st.dataframe(pd.DataFrame(validation_rows.items(), columns=["Validacion", "Resultado"]), use_container_width=True, hide_index=True)
            if summary:
                diff = summary["metrics"].get("diferencia_cuadre", 0)
                if abs(diff) > 1000:
                    st.warning(f"Diferencia de cuadre detectada: {_money(diff)}. Revise cuentas de cierre, patrimonio y resultado.")
                else:
                    st.success("La diferencia de cuadre esta dentro de un margen razonable.")
            if not valid_company:
                st.error("El archivo cargado no parece corresponder a TLG. Por seguridad, no se actualizaran los Estados Financieros.")

    with tab_balance:
        st.subheader("Estado de Situacion Financiera")
        if not summary:
            st.info("Primero carga y valida el balance de TLG.")
        else:
            metrics = summary["metrics"]
            _metric_grid(metrics)
            balance = summary["balance"]
            display = _format_money_df(balance, ["SALDO_INICIAL", "MOVIMIENTO_DEBITO", "MOVIMIENTO_CREDITO", "SALDO_FINAL"])
            st.dataframe(display, use_container_width=True, hide_index=True)

    with tab_income:
        st.subheader("Estado de Resultados")
        if not summary:
            st.info("Primero carga y valida el balance de TLG.")
        else:
            metrics = summary["metrics"]
            cols = st.columns(4)
            cols[0].metric("Ingresos", _money(metrics.get("total_ingresos", 0)))
            cols[1].metric("Costos", _money(metrics.get("total_costos", 0)))
            cols[2].metric("Gastos", _money(metrics.get("total_gastos", 0)))
            cols[3].metric("Utilidad / perdida", _money(metrics.get("resultado", 0)))
            income = summary["income_statement"]
            display = _format_money_df(income, ["SALDO_INICIAL", "MOVIMIENTO_DEBITO", "MOVIMIENTO_CREDITO", "SALDO_FINAL"])
            st.dataframe(display, use_container_width=True, hide_index=True)

    with tab_summary:
        st.subheader("Resumen gerencial")
        if not summary:
            st.info("Primero carga y valida el balance de TLG.")
        else:
            text = build_tlg_management_text(summary["metrics"], metadata)
            st.success(text)
            if abs(summary["metrics"].get("diferencia_cuadre", 0)) > 1000:
                st.warning("Alerta contable: existe diferencia de cuadre. Antes de emitir estados financieros, valide cuentas de resultado, cierre y patrimonio.")
            if not valid_company:
                st.error("No se recomienda descargar estados financieros actualizados hasta confirmar que el archivo pertenece a TLG.")

    with tab_downloads:
        st.subheader("Descargas")
        if not summary:
            st.info("Primero carga y valida el balance de TLG.")
        elif not valid_company:
            st.error("El archivo cargado no parece corresponder a TLG. Por seguridad, no se actualizaran los Estados Financieros.")
        else:
            updated_excel = update_tlg_financial_statements(
                st.session_state.get("tlg_template_file"),
                summary["detail"],
                metadata,
            )
            summary_excel = build_tlg_summary_excel(summary, metadata)
            pdf = build_tlg_management_pdf(build_tlg_management_text(summary["metrics"], metadata), summary["metrics"], metadata)
            st.download_button(
                "Descargar Excel actualizado de Estados Financieros TLG",
                data=updated_excel,
                file_name="Estados_Financieros_TLG_actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "Descargar resumen financiero en Excel",
                data=summary_excel,
                file_name="Resumen_financiero_TLG.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "Descargar PDF resumen gerencial",
                data=pdf,
                file_name="Resumen_gerencial_TLG.pdf",
                mime="application/pdf",
            )
