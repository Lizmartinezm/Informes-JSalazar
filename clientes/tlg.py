from __future__ import annotations

import streamlit as st

from utils.ui_components import feature_grid, section_header


def render_tlg() -> None:
    section_header(
        "Informe personalizado - TLG",
        "Este espacio queda separado del informe de Restaurante Sazon. Aqui se construira la logica especifica de TLG "
        "sin afectar calculos, tablas, PDF ni graficos existentes de Sazon.",
    )
    st.info("Modulo TLG en preparacion. En el siguiente paso se definiran archivos de entrada, indicadores y estructura del informe.")
    feature_grid(
        [
            ("Carga de informacion", "Pendiente definir archivos base y estructura de datos."),
            ("Indicadores TLG", "Pendiente definir KPIs gerenciales propios del cliente."),
            ("Reportes y descargas", "Pendiente definir Excel/PDF y conclusiones automaticas."),
        ]
    )
