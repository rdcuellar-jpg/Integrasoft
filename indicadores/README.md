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

`ID`, `Asunto`, `Sprint`, `Tipo`, `Tipo de Solicitud`, `Cliente`, `Estado`, `Trabajo`, `Trabajo total`, `Asignado a`,
`Fecha de inicio`, `Fecha de finalización` (dd-mm-aaaa o fecha de Excel) y `Precisión HU`.
Si la hoja trae además `Responsable`, se habilita su filtro y su columna en el detalle.

## Reglas de cálculo (igual que las hojas Base Calculo y Datos Indicadores)

- Cada HU se cuenta **una sola vez** (la base repite filas por cada relación; se toma la primera aparición del ID).
- Cada **Estado** tiene una **Fase** (tabla editable en Configuración):
  *Análisis* (Levantamiento Requerimiento, especificaciones Funcionales, Aprobación cliente, Aprobacion PO, Backlog),
  *Desarrollo* (En desarrollo, Revisión (IDS)), *Pruebas* (Pruebas funcionales y Documentación),
  *Cerrado* (Entrega a Operaciones, Entrega, CIERRE en Desarrollo o Solución, Cierre, Garantía HU) y
  *Excluido* (Nuevo / Identificado, Asignación).
- **En flujo** se deriva de la fase: *Cerrado* si la fase es Cerrado, *Abierto* en otro caso.
- **HU consideradas** = HU únicas cuya fase no es Excluido. Un estado que no esté en la tabla queda fuera y la app lo advierte.
- KPIs:
  | KPI | Cálculo | Meta |
  |---|---|---|
  | % Deuda técnica / backlog total | HU con Tipo de Solicitud = Deuda Tecnica / HU consideradas | ≤ 20% |
  | % Cumplimiento del Backlog | HU cerradas / HU consideradas, ambas sin Deuda Tecnica | ≥ 90% |
  | Control de consumo de horas de desarrollo | Σ Trabajo total de las HU consideradas / horas contratadas (parámetro) | ≤ 100% |
  | Precisión backlog refinement | HU consideradas con sprint y sin marca en Precisión HU / HU consideradas con sprint | ≥ 85% |
  | % Cambios de alcance post-firma | HU con Precisión HU = "Cambio de Alcance" / HU consideradas | ≤ 10% |
  | % Retrabajo por ambigüedad | HU con Precisión HU = "Retrabajo*" / HU consideradas | ≤ 8% |
  | % HU en Análisis / Desarrollo / Pruebas | HU en la fase / HU consideradas | sin meta |
  | % HU Abiertas | HU con En flujo = Abierto / HU consideradas | sin meta |
- El gráfico mensual muestra el año elegido (por defecto el actual); las HU sin fecha o fuera del año se informan en la nota.
- El **módulo** se asigna por palabra clave en el Asunto (primera coincidencia de arriba hacia abajo).

En la pestaña **Configuración** se editan, y quedan guardados en el navegador: las metas, los parámetros
(horas contratadas, Tipo de Solicitud excluido del cumplimiento, marcas de Precisión HU con comodines `*` y `?`),
la tabla de fase por estado y la tabla de palabras clave.

## Vistas

- **Resumen**: titular con cuántos indicadores están en meta, tabla de KPIs con barra frente a la meta,
  distribución de HU por fase y estado, cierres por mes y HU por Tipo de Solicitud.
- **Seguimiento HU abiertas**: HU abiertas por Cliente, Sprint, Módulo y Asignado a, cruzadas por Tipo de Solicitud.
- **Detalle HU**: listado con búsqueda, filtros, orden por columna y exportación a CSV.
- **Ficha de indicadores**: definición, fórmula, frecuencia, meta, responsable y resultado actual.

Los filtros superiores (Cliente, Tipo de Solicitud, Sprint, Asignado a, Módulo, Precisión HU) se aplican a todas las vistas.

## Estructura

```
indicadores/
├── index.html
├── css/styles.css
├── js/indicadores.js        # motor de cálculo (sin dependencias, usable en Node)
├── js/app.js                # interfaz y gráficos
├── js/vendor/               # Chart.js 4.4.1 (MIT) y SheetJS 0.18.5 (Apache-2.0), locales para uso sin internet
├── fonts/                   # Schibsted Grotesk (SIL OFL 1.1), local para uso sin internet
├── data/hoja1-ejemplo.json  # copia de la Hoja1 actual para el botón "Datos de ejemplo"
└── tests/                   # pruebas que comparan contra los valores del Excel
```

## Pruebas

```bash
node --test indicadores/tests/*.test.js
```

Las pruebas verifican con la Hoja1 de ejemplo los mismos resultados del libro: 173 HU únicas, 151 consideradas,
deuda técnica 30,5 %, cumplimiento del backlog 74,3 %, 267 horas ejecutadas, precisión del refinement 95,2 %,
distribución por estado/fase/mes y las tablas de seguimiento.
