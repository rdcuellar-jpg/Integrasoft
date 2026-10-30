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
2. Pulsa **Cargar Excel** (o arrastra el archivo) y selecciona el Excel. Sirve directamente el `.xls` de paquetes de
   trabajo tal como se descarga de InteTrack (hoja *Paquetes de trabajo*) o el libro `Base para Indicadores.xlsx`.
   Se lee la hoja `Hoja1`; si no existe, se usa la primera hoja que tenga las columnas `ID` y `Estado`.
3. Al cargar un Excel la app pregunta si corresponde a **Indicadores de gestión** o a un **Sprint**. Los cortes de
   Sprint se usan además para el Burndown. El tipo se puede cambiar después en *Archivos cargados*.
4. Cada Excel cargado se guarda como un **informe** en el histórico del navegador, con una fecha de corte
   (la fecha AAAA-MM-DD del nombre del archivo, como en las descargas de InteTrack; si no tiene, la fecha de
   modificación del archivo; se puede cambiar). Si se vuelve a cargar el mismo archivo
   no se duplica. Al volver a abrir la app se muestra el último informe visto.

Todo el procesamiento ocurre en el navegador: el archivo no se envía a ningún servidor.

## Formato esperado de la Hoja1

La fila de encabezados se detecta automáticamente (en la base actual es la fila 2). Se usan las columnas del bloque
*Paquetes de trabajo* (el bloque *Relaciones* repite encabezados y se ignora):

`ID`, `Asunto`, `Sprint`, `Tipo`, `Tipo de Solicitud`, `Cliente`, `Estado`, `Trabajo`, `Trabajo total`, `Asignado a`,
`Fecha de inicio`, `Fecha de finalización` (dd-mm-aaaa o fecha de Excel), `Precisión HU`, `Módulo afectado`,
`Puntos de Historia` y `Tiempo invertido`. Si la hoja trae además `Responsable`, se habilita su filtro y su columna en el detalle.

Del bloque *Relaciones* (desde la columna `Tipo de relación`) se leen los paquetes de trabajo relacionados con cada HU
(`ID`, `Sprint`, `Tipo`, `Estado`, `Puntos de Historia`, `Fecha de finalización`…). Los de Tipo *Diseño y Desarrollo*
alimentan la pestaña **Sprint**.

## Reglas de cálculo (igual que las hojas Base Calculo y Datos Indicadores)

- Cada HU se cuenta **una sola vez** (la base repite filas por cada relación; se toma la primera aparición del ID).
- Cada **Estado** tiene una **Fase** (tabla editable en Configuración):
  *Análisis* (Levantamiento Requerimiento, especificaciones Funcionales, Aprobación cliente, Aprobacion PO, Backlog),
  *Desarrollo* (En desarrollo, Revisión (IDS)), *Pruebas* (Pruebas funcionales y Documentación),
  *Cerrado* (Entrega a Operaciones, Entrega, CIERRE en Desarrollo o Solución, Cierre, Garantía HU) y
  *Excluido* (Nuevo / Identificado, Asignación).
- Tipos de Solicitud: Mejora Funcional, Deuda Tecnica, Desarrollo Contratado, Garantía y Sin clasificar (vacío).
- **Estados equivalentes** (editables en Configuración): un estado del ERP que no está en la tabla de fases se puede
  contar como otro. Por defecto, cualquier estado que coincida con `Entrega*cliente*` (Entrega a Clientes,
  Entrega clientes…) se suma a **Entrega a Operaciones**. En el detalle de HU se ve el estado original del ERP.
- **En flujo** se deriva de la fase: *Cerrado* si la fase es Cerrado, *Abierto* en otro caso.
- **HU consideradas** = HU únicas cuya fase no es Excluido. Un estado que no esté en la tabla queda fuera y la app lo advierte.
- KPIs:
  | KPI | Cálculo | Meta |
  |---|---|---|
  | % Deuda técnica / backlog total | HU con Tipo de Solicitud = Deuda Tecnica / HU consideradas | ≤ 20% |
  | % Cumplimiento del Backlog | HU cerradas / HU consideradas, ambas sin Deuda Tecnica | ≥ 90% |
  | Control de consumo de horas de desarrollo | Σ Trabajo total de las HU consideradas / horas contratadas (parámetro) | ≤ 100% |
  | Precisión backlog refinement | HU consideradas sin marca en Precisión HU / HU consideradas | ≥ 85% |
  | % Cambios de alcance post-firma | HU con Precisión HU = "Cambio de Alcance" / HU consideradas | ≤ 10% |
  | % Retrabajo por ambigüedad | HU con Precisión HU = "Retrabajo*" / HU consideradas | ≤ 8% |
  | Efectividad (horas por punto de historia) | Σ Tiempo invertido / Σ Puntos de Historia de las HU consideradas con puntos | ≤ 8 h/pt |
  | % HU en Análisis / Desarrollo / Pruebas | HU en la fase / HU consideradas | sin meta |
  | % HU Abiertas | HU con En flujo = Abierto / HU consideradas | sin meta |
- El gráfico mensual muestra el año elegido (por defecto el actual); las HU sin fecha o fuera del año se informan en la nota.
- El **módulo** se toma de la columna `Módulo afectado` del Excel (sin distinguir mayúsculas ni tildes; si no
  existe, de la columna `Módulo`). Las HU sin módulo, o los archivos sin esa columna, quedan como
  *Sin módulo identificado*, y la app avisa cuando el archivo no trae la columna.

En la pestaña **Configuración** se editan, y quedan guardados en el navegador: las metas, los parámetros
(horas contratadas, Tipo de Solicitud excluido del cumplimiento, marcas de Precisión HU con comodines `*` y `?`),
la tabla de fase por estado, el avance por estado, el avance de Diseño y Desarrollo (tipo de paquete, acciones y estados
terminados) y los estados equivalentes.

## Vistas

Arriba se elige el **informe** (corte) que se quiere ver; todas las vistas muestran los indicadores de ese corte.

- **Indicadores Dirección de Soluciones** (vista principal): panel de KPI con el título y cuántos indicadores están
  en meta, los cortes evaluados, tarjetas de HU consideradas, abiertas, cerradas y horas, el estado de las metas con
  el avance promedio, y gráficos de HU por Tipo de Solicitud, por fase, abiertas y cerradas, abiertas por módulo y
  por cliente. Debajo: indicadores de la plantilla, distribución de HU por fase y estado, cierres por mes y HU por Tipo de Solicitud.
  Los indicadores se pueden ver de cuatro formas: *Tabla con meta* (barra frente a la meta), *Distancia a la meta*
  (puntos por encima o por debajo), *Por estado* (requieren atención, en meta, sin datos) y *Por cliente*
  (matriz de indicadores calculados con las HU de cada cliente).
- **Avance por estado**: avance de cada HU según su estado, con la tabla de esfuerzo del área:
  | Acción | % del esfuerzo | Avance de una HU en ese estado |
  |---|---|---|
  | Levantamiento de requerimientos | 10 % | 0 % |
  | Especificaciones funcionales | 15 % | 10 % |
  | Aprobación cliente | 5 % | 25 % |
  | Desarrollo | 45 % | 30 % |
  | Pruebas funcionales y documentación | 25 % | 75 % |
  | Entrega a Operaciones, Entrega cliente y cerradas (fase Cerrado) | | 100 % |
  | Estados no mencionados (Backlog, Aprobacion PO, Revisión (IDS)…) | | 0 % |

  Cada HU suma el esfuerzo de las acciones anteriores a su estado. Muestra el avance promedio, la distribución por
  tramos, el avance por estado (con su aporte al promedio), el avance promedio por cliente, sprint, módulo o
  asignado, y la lista de HU sin terminar. La tabla de acciones, sus % y estados, y la opción de sumar también la
  acción del estado actual se editan en Configuración. El Detalle de HU muestra el avance de cada HU y el
  Comparativo de informes compara el avance promedio entre cortes.
- **Sprint**: se calcula con los paquetes *Diseño y Desarrollo* relacionados con las HU (bloque Relaciones). Solo cuentan
  Desarrollo (45 %) y Pruebas funcionales (25 %), escalados a 100 %: un paquete *En desarrollo* lleva 0 %, uno en
  *Pruebas funcionales y Documentación* 64 % (45 de 70), los terminados (Despliegue, Entrega…, Cierre…, Garantía…) 100 % y
  los demás estados 0 %. Todo se edita en Configuración.
  - **Avance del sprint**: promedio del avance de los paquetes de cada sprint. El sprint es el de la **Historia de Usuario**
    a la que se relaciona el paquete (los paquetes de HU sin sprint se informan aparte y no entran), con paquetes
    terminados, en pruebas, sin avance y puntos terminados.
  - **Burndown**: puntos de historia pendientes (paquetes no terminados) del sprint elegido en cada Excel cargado como
    Sprint, ordenados por fecha de corte, frente a una línea ideal que llega a 0 en la fecha de finalización más tardía
    de los paquetes del sprint. Lista los **ID que pasaron de un sprint a otro** entre cortes (salieron o entraron) porque
    cambió el sprint de su HU.
    Si los paquetes no tienen puntos, se cuentan paquetes.
  - **Avance por ID**: gráfica de barras con el avance de cada HU del sprint elegido, como el promedio de sus paquetes
    Diseño y Desarrollo (el detalle de cada paquete aparece al pasar el cursor); se puede buscar y ver solo las HU sin terminar.
- **Comparativo de informes**: elige un informe base y otro para comparar, y muestra:
  - los 11 indicadores de cada corte, la variación en puntos (verde si mejora frente a la meta, rojo si empeora)
    y el cambio de estado;
  - el volumen de HU (únicas, consideradas, abiertas, cerradas, por fase y horas) con su diferencia;
  - la evolución de un indicador en todos los informes guardados, ordenados por fecha de corte, con la meta;
  - las HU que cambiaron entre los dos cortes: nuevas, cerradas, reabiertas, con cambio de estado y retiradas
    (las que ya no aparecen entre las consideradas del informe posterior).
  Los filtros activos y la configuración actual se aplican a los dos informes.
- **Archivos cargados**: lista de los Excel guardados en el navegador. Permite cambiar el nombre, la fecha de corte o el tipo (Gestión o Sprint),
  abrir un archivo, **eliminar** uno, varios (con las casillas y *Eliminar seleccionados*) o todos, y
  **exportar o importar el histórico** (archivo JSON) para llevarlo a otro equipo o respaldarlo. Eliminar pide
  confirmación y no modifica el Excel original. También se llega con *Administrar archivos* en el encabezado.
- **HU abiertas**: gráficas de HU abiertas por Cliente, Sprint, Módulo y Asignado a, cruzadas por Tipo de Solicitud; cada
  barra muestra la cantidad y su % sobre el total de HU abiertas.
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
├── js/historico.js          # histórico de informes en el navegador (IndexedDB)
├── js/vendor/               # Chart.js 4.4.1 (MIT) y SheetJS 0.18.5 (Apache-2.0), locales para uso sin internet
├── fonts/                   # Source Serif 4 (SIL OFL 1.1) para títulos, local para uso sin internet
├── data/hoja1-ejemplo.json  # copia de la Hoja1 actual para el botón "Datos de ejemplo"
└── tests/                   # pruebas que comparan contra los valores del Excel
```

## Pruebas

```bash
node --test indicadores/tests/*.test.js
```

Las pruebas verifican con la Hoja1 de ejemplo los mismos resultados del libro: 168 HU únicas, 122 consideradas,
deuda técnica 14,8 %, cumplimiento del backlog 76,0 %, 271 horas ejecutadas, precisión del refinement 94,3 %,
cambios de alcance 2,5 %, retrabajo 3,3 %, distribución por estado/fase/mes y las tablas de seguimiento. También
verifican Efectividad, Avance del sprint, Avance por ID y Burndown con una hoja en el formato de InteTrack.
