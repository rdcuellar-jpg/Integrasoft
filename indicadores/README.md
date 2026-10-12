# Dashboard de Excelencia Operativa — Indicadores

Aplicación web estática para revisar los indicadores de Historias de Usuario (HU) de la Dirección de Soluciones.
Usa como fuente la **Hoja1** del libro *Base para Indicadores.xlsx*, que se carga manualmente, y reproduce los
cálculos de las hojas *Datos Indicadores*, *Indicadores Seguimiento*, *Ficha de Indicadores* y *Dashboard*.

## Cómo usarla

1. Sirve la carpeta con cualquier servidor web estático, por ejemplo:
   ```bash
   cd indicadores
   python3 -m http.server 8080
   ```
   y abre <http://localhost:8080>. También se puede publicar tal cual en GitHub Pages, IIS, Nginx, etc.
   (Abrir `index.html` directamente con doble clic también funciona para cargar archivos, pero el botón
   *Datos de ejemplo* necesita servidor).
2. Pulsa **Cargar Excel / CSV** (o arrastra el archivo) y selecciona el `.xlsx` exportado. Se lee la hoja `Hoja1`;
   si no existe, se usa la primera hoja que tenga las columnas `ID` y `Estado`.
3. Los datos quedan guardados en el navegador: al volver a abrir la app se muestra la última carga.
   Para actualizar, vuelve a cargar el archivo.

Todo el procesamiento ocurre en el navegador: el archivo no se envía a ningún servidor.

## Formato esperado de la Hoja1

La fila de encabezados se detecta automáticamente (en la base actual es la fila 2). Se usan las columnas del bloque
*Paquetes de trabajo* (el bloque *Relaciones* repite encabezados y se ignora):

`ID`, `Asunto`, `Tipo`, `Cliente`, `Estado`, `En flujo`, `Trabajo total`, `Asignado a`, `Fecha de inicio`,
`Fecha de finalización` (dd-mm-aaaa o fecha de Excel), `Responsable`, `Tipo de Solicitud`, `Sprint`.

## Reglas de cálculo (igual que el Excel)

- Cada HU se cuenta **una sola vez** (la base repite filas por cada relación; se toma la primera aparición del ID).
- **HU consideradas** = HU únicas excluyendo los estados *Nuevo / Identificado* y *Asignación*.
- Fases: *Análisis* (Levantamiento Requerimiento, Realizar Mockups, Documentación, Aprobación cliente, Revisión (sol)),
  *Desarrollo* (En desarrollo, Revisión (IDS)), *Pruebas* (Testing), *Cerrado* (Entrega, CIERRE en Desarrollo o Solución).
  Si aparece un estado no listado, la app lo advierte en la parte superior.
- KPIs:
  | KPI | Cálculo | Meta |
  |---|---|---|
  | % Deuda técnica / backlog total | HU con Tipo de Solicitud = Deuda Tecnica / HU consideradas | ≤ 20% |
  | % Cumplimiento del Roadmap | HU con En flujo = Cerrado / HU consideradas | ≥ 90% |
  | % Automatización de pruebas | dato manual | ≥ 70% |
  | Precisión backlog refinement | dato manual | ≥ 85% |
  | % Cambios de alcance post-firma | dato manual | ≤ 10% |
  | % Retrabajo por ambigüedad | dato manual | ≤ 8% |
  | % HU en Análisis / Desarrollo / Pruebas | HU en la fase / HU consideradas | sin meta |
  | % HU Abiertas | HU con En flujo = Abierto / HU consideradas | sin meta |
- El **módulo** se asigna por palabra clave en el Asunto (primera coincidencia de arriba hacia abajo).

Las metas, los datos manuales de los KPIs sin información en la base y la tabla de palabras clave se editan en la
pestaña **Configuración** y se guardan en el navegador.

## Vistas

- **Dashboard**: resumen de la base, tarjetas de KPI con meta y estado, HU por En flujo, fase, estado, mes de
  finalización y Tipo de Solicitud.
- **Seguimiento HU abiertas**: HU abiertas por Cliente, Sprint, Módulo y Responsable, cruzadas por Tipo de Solicitud.
- **Detalle HU**: listado con búsqueda, filtros, orden por columna y exportación a CSV.
- **Ficha de indicadores**: definición, fórmula, frecuencia, meta, responsable y resultado actual.

Los filtros superiores (Cliente, Tipo de Solicitud, Sprint, Responsable, Asignado a, Módulo) se aplican a todas las vistas.

## Estructura

```
indicadores/
├── index.html
├── css/styles.css
├── js/indicadores.js        # motor de cálculo (sin dependencias, usable en Node)
├── js/app.js                # interfaz y gráficos
├── js/vendor/               # Chart.js 4.4.1 (MIT) y SheetJS 0.18.5 (Apache-2.0), locales para uso sin internet
├── data/hoja1-ejemplo.json  # copia de la Hoja1 actual para el botón "Datos de ejemplo"
└── tests/                   # pruebas que comparan contra los valores del Excel
```

## Pruebas

```bash
node --test indicadores/tests/*.test.js
```

Las pruebas verifican con la Hoja1 de ejemplo los mismos resultados del libro: 133 HU únicas, 111 consideradas,
deuda técnica 6,3 %, cumplimiento del roadmap 67,6 %, distribución por estado/fase/mes y las tablas de seguimiento.
