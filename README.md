# Informe Gerencial - Restaurante Sazón

Aplicación web en Streamlit para generar un informe automático de ventas, gastos, utilidad estimada, formas de pago, propinas, vendedores y clientes del Restaurante Sazón.

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app.py
```

## Archivos que debe subir el usuario

La aplicación solicita únicamente dos archivos Excel con extensión `.xlsx`:

- Archivo de ventas, por ejemplo `Ventas ENERO MAYO 2026 APP.xlsx`.
- Archivo de gastos, por ejemplo `gastos RESTAURANTE SAZÓN APP.xlsx`.

El archivo de ventas puede tener hojas como `Resumen`, `Resumen (2)` o meses como `ENERO`. El archivo de gastos puede tener una hoja `Tendencias anuales` o hojas mensuales como `Ene.`, `Feb.`, `Mar.`, `Abr.`, `Mayo`.

## Indicadores generados

- Ventas acumuladas.
- Gastos acumulados.
- Utilidad estimada acumulada.
- Margen estimado acumulado.
- Valor neto recibido.
- Propinas.
- Número de facturas.
- Ticket promedio.
- Mes con mayor venta.
- Forma de pago más usada.
- Vendedor con mayor venta.

## Secciones del informe

- Informe mensual.
- Informe acumulado.
- Análisis de formas de pago.
- Ventas por vendedor.
- Análisis de clientes.
- Análisis de propinas.
- Análisis de gastos.
- Resultado operativo estimado.
- Conclusiones automáticas en lenguaje sencillo.

## Filtros

La barra lateral permite filtrar por:

- Rango de fechas.
- Mes.
- Vendedor.
- Forma de pago.
- Cliente.

Los filtros afectan las tablas, indicadores y gráficos del dashboard.

## Descarga del informe

El botón **Descargar informe en Excel** genera un archivo con estas hojas:

- Resumen Ejecutivo.
- Informe Mensual.
- Formas de Pago.
- Vendedores.
- Clientes.
- Propinas.
- Gastos.
- Resultado Operativo.

La descarga en PDF queda preparada como función pendiente para una siguiente versión.

## Manejo de errores

La aplicación muestra mensajes claros cuando detecta archivos vacíos, columnas faltantes, fechas inválidas, valores monetarios como texto, hojas con nombres diferentes, meses sin datos o un archivo de gastos sin información suficiente.
