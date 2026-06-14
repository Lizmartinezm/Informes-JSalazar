from __future__ import annotations

import importlib
import sys

import streamlit as st

from clientes.mensualizados import render_monthly_reports
from clientes.tlg import render_tlg
from utils.ui_components import (
    feature_grid,
    inject_app_styles,
    section_header,
    sidebar_brand,
    top_bar,
)


st.set_page_config(
    page_title="Julio Salazar | Gestión financiera",
    layout="wide",
    initial_sidebar_state="expanded",
)
inject_app_styles()


def render_home() -> None:
    section_header(
        "Control financiero en un solo lugar",
        "Un entorno de trabajo para convertir archivos contables y operativos en informes claros, "
        "comparables y listos para la toma de decisiones.",
        eyebrow="Panel ejecutivo",
        badge="Información centralizada",
    )
    feature_grid(
        [
            (
                "Informes personalizados",
                "Reportes gerenciales adaptados a la operación y estructura de cada cliente.",
            ),
            (
                "Estados financieros",
                "Validación de balances, clasificación contable y preparación de estados financieros.",
            ),
            (
                "Informes mensualizados",
                "Actualización acumulativa de Balance y PYG a partir de uno o varios periodos.",
            ),
            (
                "Control tributario",
                "Espacio preparado para IVA, retenciones, ICA, renta y conciliación fiscal.",
            ),
            (
                "Trazabilidad",
                "Cada proceso conserva su fuente, periodo y versión para facilitar la revisión.",
            ),
            (
                "Descargas ejecutivas",
                "Resultados listos para compartir en Excel o PDF según el tipo de informe.",
            ),
        ]
    )


def render_custom_reports() -> None:
    section_header(
        "Informes personalizados",
        "Selecciona el proceso que deseas ejecutar. Cada módulo conserva su propia lógica, archivos y resultados.",
        eyebrow="Operación por cliente",
        badge="3 módulos disponibles",
    )
    client = st.segmented_control(
        "Tipo de informe",
        ["Restaurante Sazón", "TLG", "Informes mensualizados"],
        default="Restaurante Sazón",
        key="custom_report_selector",
        label_visibility="collapsed",
    )
    st.divider()
    if client == "Restaurante Sazón":
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
        "Estados financieros",
        "Área de preparación y análisis de información financiera basada en balances de prueba validados.",
        eyebrow="Estructura contable",
        badge="En evolución",
    )
    feature_grid(
        [
            ("Situación financiera", "Activos, pasivos, patrimonio y validaciones de cuadre."),
            ("Estado de resultados", "Ingresos, costos, gastos y resultado del periodo."),
            ("Flujo de efectivo", "Movimientos de operación, inversión y financiación."),
            ("Cambios en patrimonio", "Resultados, reservas y movimientos patrimoniales."),
            ("Notas financieras", "Revelaciones y soportes para la presentación de cifras."),
            ("Análisis financiero", "Variaciones, indicadores y lectura ejecutiva del periodo."),
        ]
    )


def render_taxes() -> None:
    section_header(
        "Gestión tributaria",
        "Espacio para automatizar revisiones, cruces y obligaciones tributarias con una visión de control.",
        eyebrow="Cumplimiento",
        badge="Próximamente",
    )
    feature_grid(
        [
            ("IVA", "Control de impuestos generados, descontables y saldos."),
            ("Retención en la fuente", "Revisión de retenciones practicadas y asumidas."),
            ("ICA", "Análisis territorial, actividades y bases gravables."),
            ("Renta", "Depuración fiscal y conciliaciones principales."),
            ("Información exógena", "Preparación y validación de información reportable."),
            ("Calendario tributario", "Fechas, obligaciones y alertas de cumplimiento."),
        ]
    )


sidebar_brand()
st.sidebar.caption("NAVEGACIÓN")
section = st.sidebar.radio(
    "Menú principal",
    ["Inicio", "Informes personalizados", "Estados financieros", "Impuestos"],
    label_visibility="collapsed",
)
st.sidebar.divider()
st.sidebar.caption("Versión operativa 2026")

top_bar(section)
if section == "Inicio":
    render_home()
elif section == "Informes personalizados":
    render_custom_reports()
elif section == "Estados financieros":
    render_financial_statements()
else:
    render_taxes()
