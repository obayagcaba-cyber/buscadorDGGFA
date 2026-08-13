# Buscador DGGFA — Flota Automotor

Consulta por dominio y resumen de la flota automotor del Gobierno de la Ciudad
de Buenos Aires — DG. Gestión de la Flota Automotor.

Sitio estático, sin build ni dependencias: se publica tal cual.

| Archivo | Qué es |
|---|---|
| `index.html` | Página |
| `estilos.css` | Diseño según Manual de Marca GCBA 2026 |
| `app.js` | Buscador, ficha y cálculo del resumen |
| `datos.js` | Datos del período (**generado — no editar a mano**) |

`datos.js` se regenera cada mes desde la planilla de origen con el script
`generar_datos.py`, que vive fuera de este repositorio junto al archivo Excel.
La planilla no se publica.
