from __future__ import annotations

import importlib
import sys

import streamlit as st

from clientes.tlg import render_tlg
from clientes.mensualizados import render_monthly_reports
from utils.ui_components import feature_grid, section_header


st.set_page_config(page_title="Panel Financiero y Tributario - Julio Salazar", layout="wide")


def render_home() -> None:
    section_header(
        "Panel Financiero y Tributario - Julio Salazar",
        "Bienvenido, Julio Salazar. Desde este panel podras consultar informes personalizados de clientes, "
        "generar Estados Financieros y preparar analisis tributarios.",
    )
    feature_grid(
        [
            (
                "Informes personalizados",
                "Reportes gerenciales por cliente, con logica independiente para cada negocio.",
            ),
            (
                "Estados Financieros",
                "Modulo futuro para clasificar balances y generar estados financieros completos.",
            ),
            (
                "Impuestos",
                "Modulo futuro para revisar IVA, retenciones, ICA, renta, exogena y conciliacion fiscal.",
            ),
        ]
    )


def render_custom_reports() -> None:
    section_header(
        "Informes personalizados de clientes",
        "Selecciona el cliente que quieres analizar. Cada informe vive en su propio modulo para evitar cruces de logica.",
    )
    client = st.radio(
        "Tipo de informe",
        ["Restaurante Sazon", "TLG", "Informes mensualizados"],
        horizontal=True,
    )
    st.divider()
    if client == "Restaurante Sazon":
        if "clientes.sazon" in sys.modules:
            importlib.reload(sys.modules["clientes.sazon"])
        else:
            importlib.import_module("clientes.sazon")
    elif client == "TLG":
        render_tlg()
    else:
        render_monthly_reports()


def render_financial_statements() -> None:
    section_header(
        "Estados Financieros",
        "Esta seccion estara destinada a la generacion automatica de Estados Financieros. Proximamente permitira "
        "cargar balances de prueba, clasificar cuentas contables y generar Estado de Situacion Financiera, "
        "Estado de Resultados, Flujo de Efectivo, Cambios en el Patrimonio y Notas.",
    )
    feature_grid(
        [
            ("Estado de Situacion Financiera", "Clasificacion de activos, pasivos y patrimonio."),
            ("Estado de Resultados", "Ingresos, costos, gastos y resultado del periodo."),
            ("Flujo de Efectivo", "Actividades de operacion, inversion y financiacion."),
            ("Estado de Cambios en el Patrimonio", "Movimientos patrimoniales y resultados acumulados."),
            ("Notas a los Estados Financieros", "Revelaciones y comentarios de soporte."),
            ("Analisis financiero", "Indicadores, variaciones y lectura gerencial."),
        ]
    )


def render_taxes() -> None:
    section_header(
        "Impuestos",
        "Esta seccion estara destinada a la automatizacion y revision de impuestos. Proximamente permitira analizar "
        "IVA, retencion en la fuente, ICA, renta, exogena y conciliacion fiscal.",
    )
    feature_grid(
        [
            ("IVA", "Revision de impuestos generados, descontables y saldos."),
            ("Retencion en la fuente", "Control de retenciones practicadas y asumidas."),
            ("ICA", "Analisis municipal y bases gravables."),
            ("Renta", "Depuracion fiscal y conciliaciones principales."),
            ("Informacion exogena", "Preparacion y cruces de informacion reportable."),
            ("Conciliacion fiscal", "Diferencias contables y fiscales."),
            ("Calendario tributario", "Fechas, obligaciones y alertas de cumplimiento."),
        ]
    )


st.sidebar.title("Julio Salazar")
section = st.sidebar.radio(
    "Menu principal",
    ["Inicio", "Informes personalizados", "Estados Financieros", "Impuestos"],
)

if section == "Inicio":
    render_home()
elif section == "Informes personalizados":
    render_custom_reports()
elif section == "Estados Financieros":
    render_financial_statements()
else:
    render_taxes()
