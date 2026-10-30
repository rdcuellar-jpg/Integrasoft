/*
 * Motor de cálculo de indicadores — replica las fórmulas del libro
 * "Base para Indicadores.xlsx" (hojas Base Calculo, Datos Indicadores,
 * Indicadores Seguimiento y Ficha de Indicadores) a partir de la fuente
 * manual de la Hoja1.
 *
 * Funciona en el navegador (window.Indicadores) y en Node (module.exports).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Indicadores = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var EXCLUIDO = 'Excluido';
  var SIN_MAPEAR = 'Sin mapear';

  // Estado -> Fase (Datos Indicadores!A13:B27, editable en Configuración).
  // La fase "Cerrado" define En flujo = Cerrado; el resto de fases, Abierto.
  var ESTADOS = [
    ['Levantamiento Requerimiento', 'Análisis'],
    ['especificaciones Funcionales', 'Análisis'],
    ['Aprobación cliente', 'Análisis'],
    ['Aprobacion PO', 'Análisis'],
    ['Backlog', 'Análisis'],
    ['En desarrollo', 'Desarrollo'],
    ['Revisión (IDS)', 'Desarrollo'],
    ['Pruebas funcionales y Documentación', 'Pruebas'],
    ['Entrega a Operaciones', 'Cerrado'],
    ['Entrega', 'Cerrado'],
    ['CIERRE en Desarrollo o Solución', 'Cerrado'],
    ['Cierre', 'Cerrado'],
    ['Garantía HU', 'Cerrado'],
    ['Nuevo / Identificado', EXCLUIDO],
    ['Asignación', EXCLUIDO]
  ];
  var FASES = ['Análisis', 'Desarrollo', 'Pruebas', 'Cerrado'];
  var TIPOS_SOLICITUD = ['Desarrollo Contratado', 'Mejora Funcional', 'Deuda Tecnica', 'Garantía', 'Sin clasificar'];
  var SIN_MODULO = 'Sin módulo identificado';
  var SIN_SPRINT = 'Sin sprint';
  var SIN_CLIENTE = 'Sin cliente';
  var SIN_CLASIFICAR = 'Sin clasificar';

  // Parámetros de Datos Indicadores (columna J y celdas de entrada)
  // Estados equivalentes (editable en Configuración): el estado del ERP que
  // coincide con el patrón se cuenta como el estado destino. Admite * y ?.
  var EQUIVALENCIAS = [
    ['Entrega*cliente*', 'Entrega a Operaciones']
  ];

  // Avance por estado: % del esfuerzo de cada acción y los estados del ERP que le
  // corresponden, en orden. Una HU en un estado suma las acciones anteriores a ese estado;
  // las HU en fase Cerrado (Entrega a Operaciones, Entrega cliente, CIERRE…) van al 100%
  // y los estados no mencionados suman 0%.
  var AVANCE = {
    acciones: [
      { accion: 'Levantamiento de requerimientos', esfuerzo: 0.10, estados: ['Levantamiento Requerimiento'] },
      { accion: 'Especificaciones funcionales', esfuerzo: 0.15, estados: ['especificaciones Funcionales'] },
      { accion: 'Aprobación cliente', esfuerzo: 0.05, estados: ['Aprobación cliente'] },
      { accion: 'Desarrollo', esfuerzo: 0.45, estados: ['En desarrollo'] },
      { accion: 'Pruebas funcionales y documentación', esfuerzo: 0.25, estados: ['Pruebas funcionales y Documentación'] }
    ],
    // false: la HU suma solo las acciones anteriores a su estado (la actual está en curso)
    incluirActual: false
  };

  // Avance de los paquetes Diseño y Desarrollo relacionados con las HU (Avance del sprint,
  // Avance por ID y Burndown). Solo cuentan Desarrollo y Pruebas funcionales, escalados a 100%:
  // un paquete suma las acciones anteriores a su estado y los estados terminados van al 100%.
  var AVANCE_DYD = {
    tipo: 'Diseño y Desarrollo',
    acciones: [
      { accion: 'Desarrollo', esfuerzo: 0.45, estados: ['En desarrollo'] },
      { accion: 'Pruebas funcionales', esfuerzo: 0.25, estados: ['Pruebas funcionales y Documentación'] }
    ],
    // Admiten comodines * y ?, sin distinguir mayúsculas ('Cierre*' incluye CIERRE en Desarrollo o Solución y Cierre DyD)
    terminados: ['Despliegue', 'Entrega*', 'Cierre*', 'Garantía*'],
    incluirActual: false
  };

  var PARAMETROS = {
    tipoExcluidoBacklog: 'Deuda Tecnica', // J55
    marcaAlcance: 'Cambio de Alcance',    // J58
    marcaRetrabajo: 'Retrabajo*',         // J59 (admite comodines * y ?)
    horasContratadas: null,               // C56
    anio: null                            // E65 (null = año actual)
  };
  // KPIs (Datos Indicadores!A54:I63 + Ficha de Indicadores)
  var KPIS = [
    { id: 'deuda', nombre: '% Deuda técnica / backlog total', meta: 0.20, tipoMeta: '<=', grupo: 'plantilla',
      definicion: 'Proporción de ítems del backlog clasificados como deuda técnica respecto al total de ítems del backlog.',
      formula: '(N° de ítems de deuda técnica / N° total de ítems del backlog) x 100', frecuencia: 'Mensual',
      responsable: 'Product Owner', criterio: 'HU con Tipo de Solicitud = Deuda Tecnica / total HU consideradas' },
    { id: 'backlog', nombre: '% Cumplimiento del Backlog', meta: 0.90, tipoMeta: '>=', grupo: 'plantilla',
      definicion: 'Porcentaje de HU del backlog cerradas (Entrega / CIERRE) respecto al total de HU consideradas, sin incluir las de deuda técnica.',
      formula: '(N° HU cerradas sin deuda técnica / N° HU consideradas sin deuda técnica) x 100', frecuencia: 'Mensual',
      responsable: 'Director de Soluciones / PO', criterio: 'HU cerradas / HU consideradas, sin incluir el Tipo de Solicitud excluido (parámetro)' },
    { id: 'horas', nombre: 'Control de consumo de horas de desarrollo', meta: 1, tipoMeta: '<=', grupo: 'plantilla', unidad: 'h',
      definicion: 'Cantidad de horas que se consumen de las horas contratadas, no debe ser superior al 100%.',
      formula: '(N° horas ejecutadas / N° horas contratadas) x 100', frecuencia: 'Mensual',
      responsable: 'Líder de QA', criterio: 'Horas ejecutadas (Trabajo total de las HU consideradas) vs horas contratadas (ingresar en Configuración)' },
    { id: 'refinement', nombre: 'Precisión backlog refinement', meta: 0.85, tipoMeta: '>=', grupo: 'plantilla',
      definicion: 'Historias en sprint que no tienen marca de retrabajo/cambio de alcance en la columna Precisión HU.',
      formula: '(N° de historias sin retrabajo / N° total de historias refinadas) x 100', frecuencia: 'Mensual',
      responsable: 'Product Owner', criterio: 'HU consideradas sin marca en Precisión HU / total HU consideradas (excluye solo Nuevo / Identificado y Asignación)' },
    { id: 'alcance', nombre: '% Cambios de alcance post-firma', meta: 0.10, tipoMeta: '<=', grupo: 'plantilla',
      definicion: 'Requerimientos modificados después de su aprobación formal.',
      formula: '(N° de requerimientos con cambio post-firma / N° total firmados) x 100', frecuencia: 'Mensual',
      responsable: 'Analista Funcional Líder', criterio: "HU con marca 'Cambio de Alcance' en Precisión HU / total HU consideradas" },
    { id: 'ambiguedad', nombre: '% Retrabajo por ambigüedad', meta: 0.08, tipoMeta: '<=', grupo: 'plantilla',
      definicion: 'Requerimientos retrabajados por especificaciones ambiguas o incompletas.',
      formula: '(N° retrabajados por ambigüedad / N° total de requerimientos) x 100', frecuencia: 'Mensual',
      responsable: 'Analista Funcional Líder', criterio: "HU con marca 'Retrabajo por ambigüedad' en Precisión HU / total HU consideradas" },
    { id: 'efectividad', nombre: 'Efectividad (horas por punto de historia)', meta: 8, tipoMeta: '<=', grupo: 'plantilla', unidad: 'h/pt',
      definicion: 'Uso eficiente del tiempo: horas invertidas por cada punto de historia. Menos horas por punto es más eficiente.',
      formula: 'Σ Tiempo invertido / Σ Puntos de Historia', frecuencia: 'Mensual',
      responsable: 'Líder de Desarrollo', criterio: 'Tiempo invertido / Puntos de Historia de las HU consideradas que tienen puntos' },
    { id: 'analisis', nombre: '% HU en Análisis', meta: null, tipoMeta: '', grupo: 'flujo',
      definicion: 'HU en Levantamiento Requerimiento, especificaciones Funcionales, Aprobación cliente, Aprobacion PO o Backlog.',
      formula: '(N° HU en Análisis / N° HU consideradas) x 100', frecuencia: 'Mensual',
      responsable: 'Analista Funcional Líder', criterio: 'Estados con fase Análisis' },
    { id: 'desarrollo', nombre: '% HU en Desarrollo', meta: null, tipoMeta: '', grupo: 'flujo',
      definicion: 'HU en estado En desarrollo o Revisión (IDS).',
      formula: '(N° HU en Desarrollo / N° HU consideradas) x 100', frecuencia: 'Mensual',
      responsable: 'Líder de Desarrollo', criterio: 'Estados con fase Desarrollo' },
    { id: 'pruebas', nombre: '% HU en Pruebas', meta: null, tipoMeta: '', grupo: 'flujo',
      definicion: 'HU en estado Pruebas funcionales y Documentación.',
      formula: '(N° HU en Pruebas / N° HU consideradas) x 100', frecuencia: 'Mensual',
      responsable: 'Líder de QA', criterio: 'Estados con fase Pruebas' },
    { id: 'abiertas', nombre: '% HU Abiertas', meta: null, tipoMeta: '', grupo: 'flujo',
      definicion: 'HU con En flujo = Abierto sobre el total considerado.',
      formula: '(N° HU abiertas / N° HU consideradas) x 100', frecuencia: 'Mensual',
      responsable: 'Product Owner', criterio: 'HU con En flujo = Abierto / total HU consideradas' }
  ];

  // Columnas de la Hoja1 (primer bloque "Paquetes de trabajo"; el bloque
  // "Relaciones" repite encabezados y se ignora).
  var COLUMNAS = {
    id: 'ID', asunto: 'Asunto', sprint: 'Sprint', tipo: 'Tipo', tipoSolicitud: 'Tipo de Solicitud',
    cliente: 'Cliente', estado: 'Estado', trabajo: 'Trabajo', trabajoTotal: 'Trabajo total',
    asignado: 'Asignado a', fechaInicio: 'Fecha de inicio', fechaFin: 'Fecha de finalización',
    precision: 'Precisión HU', responsable: 'Responsable',
    puntos: 'Puntos de Historia', tiempo: 'Tiempo invertido',
    // El módulo viene de la columna "Módulo afectado" del ERP (si no está, se acepta "Módulo")
    modulo: ['Módulo afectado', 'Módulo']
  };

  /** Texto sin tildes y en minúsculas, para comparar encabezados. */
  function sinTildes(s) {
    return texto(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ');
  }

  function texto(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function normalizar(s) {
    return texto(s).toLowerCase();
  }

  /** Coincidencia tipo COUNTIFS: sin distinguir mayúsculas y con comodines * y ?. */
  function coincideCriterio(valor, patron) {
    var p = texto(patron);
    if (!p) return texto(valor) === '';
    var re = p.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
    return new RegExp('^' + re + '$', 'i').test(texto(valor));
  }

  /** Convierte "dd-mm-aaaa", Date o serial de Excel a Date (o null). */
  function parseFecha(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number') {
      // Serial de Excel (sistema 1900)
      var d = new Date(Math.round((v - 25569) * 86400000));
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    var s = texto(v);
    var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    return null;
  }

  function numero(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = texto(v).replace(',', '.');
    return s && !isNaN(s) ? Number(s) : 0;
  }

  function claveMes(d) {
    return d ? d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') : '';
  }

  /** Mapa campo->columna de los encabezados entre las columnas desde y hasta (sin incluir). */
  function mapearColumnas(fila, desde, hasta) {
    var mapa = {}, rango = {};
    for (var c = desde; c < hasta; c++) {
      var h = sinTildes(fila[c]);
      if (!h) continue;
      for (var campo in COLUMNAS) {
        var alias = [].concat(COLUMNAS[campo]).map(sinTildes);
        var k = alias.indexOf(h);
        // Primera columna con ese nombre; entre alias, gana el de mayor prioridad
        if (k !== -1 && (mapa[campo] === undefined || k < rango[campo])) { mapa[campo] = c; rango[campo] = k; }
      }
    }
    return mapa;
  }

  /**
   * Busca la fila de encabezados y devuelve índice de fila + mapa campo->columna de la HU.
   * Si existe el bloque Relaciones (desde "Tipo de relación"), devuelve también su mapa.
   */
  function detectarEncabezados(matriz) {
    for (var r = 0; r < Math.min(matriz.length, 20); r++) {
      var fila = matriz[r] || [];
      var inicioRel = -1;
      for (var c = 0; c < fila.length; c++) if (sinTildes(fila[c]) === 'tipo de relacion') { inicioRel = c; break; }
      var mapa = mapearColumnas(fila, 0, inicioRel === -1 ? fila.length : inicioRel);
      if (mapa.id !== undefined && mapa.estado !== undefined) {
        var rel = inicioRel === -1 ? null : mapearColumnas(fila, inicioRel, fila.length);
        if (rel) rel.tipoRelacion = inicioRel;
        return { fila: r, mapa: mapa, relaciones: rel && rel.id !== undefined ? rel : null };
      }
    }
    return null;
  }

  /** Estado con el que se cuenta la HU: el propio si tiene fase; si no, el de la primera equivalencia que coincida. */
  function estadoEquivalente(estado, equivalencias, fases) {
    if (fases && fases.hasOwnProperty(estado)) return estado;
    for (var i = 0; i < equivalencias.length; i++) {
      if (texto(equivalencias[i][0]) && coincideCriterio(estado, equivalencias[i][0])) return texto(equivalencias[i][1]);
    }
    return estado;
  }

  function mapaFases(estados) {
    var m = {};
    (estados || ESTADOS).forEach(function (e) { m[texto(e[0])] = texto(e[1]); });
    return m;
  }

  /**
   * Convierte la matriz de la Hoja1 (array de arrays) en registros únicos de HU.
   * La base repite la HU por cada relación: se toma la primera aparición del ID
   * (equivalente a Base Calculo!B "Única").
   * @param opciones { estados, equivalencias }
   */
  function parsearHoja(matriz, opciones) {
    if (Array.isArray(opciones)) opciones = {};
    opciones = opciones || {};
    var fases = mapaFases(opciones.estados);
    var equivalencias = opciones.equivalencias || EQUIVALENCIAS;
    var enc = detectarEncabezados(matriz);
    if (!enc) throw new Error('No se encontró la fila de encabezados (se esperan columnas "ID" y "Estado").');
    var m = enc.mapa, mr = enc.relaciones;
    var vistos = {};
    var registros = [];
    var paquetes = {};
    var filasLeidas = 0;
    for (var r = enc.fila + 1; r < matriz.length; r++) {
      var fila = matriz[r] || [];
      var id = texto(fila[m.id]);
      if (!id) continue;
      filasLeidas++;
      if (mr) leerRelacion(fila, mr, vistos[id] || null, id, paquetes);
      if (vistos[id]) continue;
      var get = function (campo) { return m[campo] === undefined ? '' : fila[m[campo]]; };
      var fin = parseFecha(get('fechaFin'));
      var estadoOriginal = texto(get('estado'));
      var estado = estadoEquivalente(estadoOriginal, equivalencias, fases);
      var fase = fases.hasOwnProperty(estado) ? fases[estado] : SIN_MAPEAR;
      var considerada = fase !== EXCLUIDO && fase !== SIN_MAPEAR;
      var registro = {
        id: id,
        asunto: texto(get('asunto')),
        tipo: texto(get('tipo')),
        cliente: texto(get('cliente')) || SIN_CLIENTE,
        estado: estado,
        estadoOriginal: estadoOriginal,
        fase: fase,
        considerada: considerada,
        enFlujo: considerada ? (fase === 'Cerrado' ? 'Cerrado' : 'Abierto') : '',
        trabajo: numero(get('trabajo')),
        trabajoTotal: numero(get('trabajoTotal')),
        asignado: texto(get('asignado')),
        responsable: texto(get('responsable')),
        fechaInicio: parseFecha(get('fechaInicio')),
        fechaFin: fin,
        mesFin: claveMes(fin),
        tipoSolicitud: texto(get('tipoSolicitud')) || SIN_CLASIFICAR,
        sprint: texto(get('sprint')) || SIN_SPRINT,
        precision: texto(get('precision')),
        puntos: numero(get('puntos')),
        tiempo: numero(get('tiempo')),
        relaciones: [],
        fila: r + 1
      };
      registro.modulo = texto(get('modulo')) || SIN_MODULO;
      registros.push(registro);
      vistos[id] = registro;
      if (mr) leerRelacion(fila, mr, registro, id, paquetes);
    }
    var columnas = {};
    Object.keys(m).forEach(function (k) { columnas[k] = true; });
    if (mr) columnas.relaciones = true;
    return { registros: registros, filasLeidas: filasLeidas, filaEncabezado: enc.fila + 1, columnas: columnas };
  }

  /**
   * Agrega a la HU el paquete de trabajo relacionado de la fila (bloque Relaciones).
   * Un mismo paquete relacionado con varias HU es un único objeto.
   */
  function leerRelacion(fila, mr, hu, idHU, paquetes) {
    if (!hu) return; // la primera fila de la HU la lee parsearHoja después de crear el registro
    var id = texto(fila[mr.id]);
    if (!id) return;
    var get = function (campo) { return mr[campo] === undefined ? '' : fila[mr[campo]]; };
    var p = paquetes[id];
    if (!p) {
      p = paquetes[id] = {
        id: id, asunto: texto(get('asunto')), tipo: texto(get('tipo')), sprint: texto(get('sprint')),
        estado: texto(get('estado')), asignado: texto(get('asignado')), modulo: texto(get('modulo')),
        puntos: numero(get('puntos')), tiempo: numero(get('tiempo')),
        fechaInicio: parseFecha(get('fechaInicio')), fechaFin: parseFecha(get('fechaFin')),
        relacion: texto(fila[mr.tipoRelacion]), hus: []
      };
    }
    if (p.hus.indexOf(idHU) === -1) p.hus.push(idHU);
    if (hu.relaciones.indexOf(p) === -1) hu.relaciones.push(p);
  }

  function contar(lista, fn) {
    var n = 0;
    for (var i = 0; i < lista.length; i++) if (fn(lista[i])) n++;
    return n;
  }

  function ratio(num, den) {
    if (num === null || num === undefined || num === '' || !den) return null;
    return num / den;
  }

  function evaluar(resultado, meta, tipoMeta) {
    if (resultado === null) return 'SIN DATOS';
    if (meta === null || meta === undefined || meta === '') return 'SIN META';
    if (tipoMeta === '<=') return resultado <= meta ? 'EN META' : 'FUERA DE META';
    return resultado >= meta ? 'EN META' : 'FUERA DE META';
  }

  function tipoNormalizado(h) {
    return TIPOS_SOLICITUD.indexOf(h.tipoSolicitud) === -1 ? SIN_CLASIFICAR : h.tipoSolicitud;
  }

  /** Tabla cruzada: filas = valores de `campo`, columnas = Tipo de Solicitud. */
  function cruzarPorTipo(lista, campo) {
    var filas = {};
    lista.forEach(function (h) {
      var k = h[campo] || (campo === 'asignado' ? 'Sin asignar' : '(vacío)');
      if (!filas[k]) {
        filas[k] = { nombre: k, total: 0 };
        TIPOS_SOLICITUD.forEach(function (t) { filas[k][t] = 0; });
      }
      filas[k][tipoNormalizado(h)]++;
      filas[k].total++;
    });
    var arr = Object.keys(filas).map(function (k) { return filas[k]; });
    arr.sort(function (a, b) { return b.total - a.total || String(a.nombre).localeCompare(String(b.nombre)); });
    var total = lista.length;
    arr.forEach(function (f) { f.pct = total ? f.total / total : 0; });
    return arr;
  }

  function alFinal(arr, nombre) {
    var i = arr.findIndex(function (f) { return f.nombre === nombre; });
    if (i !== -1) arr.push(arr.splice(i, 1)[0]);
    return arr;
  }

  /**
   * Calcula todos los indicadores.
   * @param registros  registros únicos (ya filtrados por el usuario)
   * @param config     { metas: {id: {meta, tipoMeta}}, parametros: {...}, estados: [[estado, fase]] }
   */
  function calcular(registros, config) {
    config = config || {};
    var metas = config.metas || {};
    var par = Object.assign({}, PARAMETROS, config.parametros || {});
    var estados = config.estados || ESTADOS;

    var consideradas = registros.filter(function (h) { return h.considerada; });
    var nC = consideradas.length;

    var excluidos = estados.filter(function (e) { return e[1] === EXCLUIDO; }).map(function (e) { return e[0]; });
    var base = {
      total: registros.length,
      excluidos: excluidos.map(function (e) {
        return { estado: e, n: contar(registros, function (h) { return h.estado === e; }) };
      }),
      sinMapear: contar(registros, function (h) { return h.fase === SIN_MAPEAR; }),
      consideradas: nC
    };

    var porEstado = estados.filter(function (e) { return e[1] !== EXCLUIDO; }).map(function (e) {
      var n = contar(consideradas, function (h) { return h.estado === e[0]; });
      return { estado: e[0], fase: e[1], n: n, pct: nC ? n / nC : 0 };
    });
    var estadosNoListados = {};
    registros.forEach(function (h) {
      if (h.fase === SIN_MAPEAR) { var k = h.estado || '(vacío)'; estadosNoListados[k] = (estadosNoListados[k] || 0) + 1; }
    });

    var porFase = FASES.map(function (f) {
      var n = contar(consideradas, function (h) { return h.fase === f; });
      return { fase: f, n: n, pct: nC ? n / nC : 0 };
    });
    var fase = {};
    porFase.forEach(function (f) { fase[f.fase] = f.n; });

    var abiertas = contar(consideradas, function (h) { return h.enFlujo === 'Abierto'; });
    var cerradas = contar(consideradas, function (h) { return h.enFlujo === 'Cerrado'; });
    var porFlujo = [
      { nombre: 'Abierto', n: abiertas, pct: nC ? abiertas / nC : 0 },
      { nombre: 'Cerrado', n: cerradas, pct: nC ? cerradas / nC : 0 }
    ];

    var porTipo = ['Mejora Funcional', 'Deuda Tecnica', 'Desarrollo Contratado', 'Garantía', SIN_CLASIFICAR].map(function (t) {
      var n = contar(consideradas, function (h) { return tipoNormalizado(h) === t; });
      return { nombre: t, n: n, pct: nC ? n / nC : 0 };
    });

    var sinExcluido = consideradas.filter(function (h) { return !coincideCriterio(h.tipoSolicitud, par.tipoExcluidoBacklog) || !texto(par.tipoExcluidoBacklog); });
    var horas = consideradas.reduce(function (s, h) { return s + h.trabajoTotal; }, 0);
    var horasContratadas = par.horasContratadas === '' || par.horasContratadas === null || par.horasContratadas === undefined
      ? null : Number(par.horasContratadas);

    // Efectividad: horas invertidas por punto de historia (HU consideradas con puntos)
    var conPuntos = consideradas.filter(function (h) { return h.puntos > 0; });
    var puntos = conPuntos.reduce(function (s, h) { return s + h.puntos; }, 0);
    var tiempo = conPuntos.reduce(function (s, h) { return s + h.tiempo; }, 0);

    var numDen = {
      deuda: [contar(consideradas, function (h) { return h.tipoSolicitud === 'Deuda Tecnica'; }), nC],
      backlog: [contar(sinExcluido, function (h) { return h.enFlujo === 'Cerrado'; }), sinExcluido.length],
      horas: [horas, horasContratadas],
      refinement: [contar(consideradas, function (h) { return h.precision === ''; }), nC],
      alcance: [contar(consideradas, function (h) { return coincideCriterio(h.precision, par.marcaAlcance); }), nC],
      ambiguedad: [contar(consideradas, function (h) { return coincideCriterio(h.precision, par.marcaRetrabajo); }), nC],
      efectividad: [tiempo, puntos],
      analisis: [fase['Análisis'], nC],
      desarrollo: [fase['Desarrollo'], nC],
      pruebas: [fase['Pruebas'], nC],
      abiertas: [abiertas, nC]
    };

    var kpis = KPIS.map(function (k) {
      var nd = numDen[k.id];
      var mt = metas[k.id] || {};
      var meta = mt.meta !== undefined ? mt.meta : k.meta;
      var tipoMeta = mt.tipoMeta !== undefined ? mt.tipoMeta : k.tipoMeta;
      var res = ratio(nd[0], nd[1]);
      return Object.assign({}, k, {
        numerador: nd[0], denominador: nd[1], resultado: res,
        meta: meta, tipoMeta: tipoMeta, estado: evaluar(res, meta, tipoMeta)
      });
    });

    // HU por mes de finalización del año seleccionado (Datos Indicadores!A65:C80)
    var anio = par.anio === 'todos' ? null : Number(par.anio) || new Date().getFullYear();
    var meses = {};
    consideradas.forEach(function (h) {
      if (!h.mesFin) return;
      if (anio && h.fechaFin.getFullYear() !== anio) return;
      if (!meses[h.mesFin]) meses[h.mesFin] = { mes: h.mesFin, cerradas: 0, abiertas: 0 };
      if (h.enFlujo === 'Cerrado') meses[h.mesFin].cerradas++; else meses[h.mesFin].abiertas++;
    });
    var mensual = [];
    var claves = Object.keys(meses).sort();
    var desde, hasta;
    if (anio) { desde = [anio, 1]; hasta = [anio, 12]; }
    else if (claves.length) { desde = claves[0].split('-').map(Number); hasta = claves[claves.length - 1].split('-').map(Number); }
    if (desde) {
      var y = desde[0], mo = desde[1];
      while (y < hasta[0] || (y === hasta[0] && mo <= hasta[1])) {
        var key = y + '-' + String(mo).padStart(2, '0');
        mensual.push(meses[key] || { mes: key, cerradas: 0, abiertas: 0 });
        mo++; if (mo > 12) { mo = 1; y++; }
      }
    }
    var enGrafico = mensual.reduce(function (s, m) { return { cerradas: s.cerradas + m.cerradas, abiertas: s.abiertas + m.abiertas }; }, { cerradas: 0, abiertas: 0 });
    var sinFecha = { cerradas: cerradas - enGrafico.cerradas, abiertas: abiertas - enGrafico.abiertas };
    var anios = {};
    consideradas.forEach(function (h) { if (h.fechaFin) anios[h.fechaFin.getFullYear()] = true; });

    // Indicadores de seguimiento: HU abiertas
    var abiertasLista = consideradas.filter(function (h) { return h.enFlujo === 'Abierto'; });
    var seguimiento = {
      total: abiertasLista.length,
      porCliente: alFinal(cruzarPorTipo(abiertasLista, 'cliente'), SIN_CLIENTE),
      porSprint: alFinal(cruzarPorTipo(abiertasLista, 'sprint'), SIN_SPRINT),
      porModulo: alFinal(cruzarPorTipo(abiertasLista, 'modulo'), SIN_MODULO),
      porAsignado: alFinal(cruzarPorTipo(abiertasLista, 'asignado'), 'Sin asignar')
    };

    return {
      base: base,
      porEstado: porEstado,
      estadosNoListados: estadosNoListados,
      porFase: porFase,
      porFlujo: porFlujo,
      porTipo: porTipo,
      kpis: kpis,
      horas: { ejecutadas: horas, contratadas: horasContratadas },
      efectividad: { horas: tiempo, puntos: puntos, hu: conPuntos.length },
      mensual: mensual,
      anio: anio,
      anios: Object.keys(anios).map(Number).sort(),
      sinFecha: sinFecha,
      seguimiento: seguimiento,
      consideradas: consideradas
    };
  }

  /**
   * Compara las HU de dos informes (A = base, B = posterior) por ID.
   * Solo se tienen en cuenta las HU consideradas de cada informe.
   * Tipos de cambio: nueva, retirada, cerrada, reabierta, estado (cambio de estado sin cambiar En flujo).
   */
  function compararHU(registrosA, registrosB) {
    var mapa = function (lista) {
      var m = {};
      lista.forEach(function (h) { if (h.considerada) m[h.id] = h; });
      return m;
    };
    var A = mapa(registrosA), B = mapa(registrosB);
    var cambios = [];
    Object.keys(B).forEach(function (id) {
      var b = B[id], a = A[id];
      if (!a) { cambios.push({ tipo: 'nueva', id: id, asunto: b.asunto, cliente: b.cliente, de: '', a: b.estado }); return; }
      if (a.estado === b.estado) return;
      var tipo = a.enFlujo === 'Abierto' && b.enFlujo === 'Cerrado' ? 'cerrada'
        : a.enFlujo === 'Cerrado' && b.enFlujo === 'Abierto' ? 'reabierta' : 'estado';
      cambios.push({ tipo: tipo, id: id, asunto: b.asunto, cliente: b.cliente, de: a.estado, a: b.estado, faseDe: a.fase, faseA: b.fase });
    });
    Object.keys(A).forEach(function (id) {
      if (!B[id]) cambios.push({ tipo: 'retirada', id: id, asunto: A[id].asunto, cliente: A[id].cliente, de: A[id].estado, a: '' });
    });
    var orden = { nueva: 0, cerrada: 1, reabierta: 2, estado: 3, retirada: 4 };
    cambios.sort(function (x, y) { return orden[x.tipo] - orden[y.tipo] || (Number(x.id) - Number(y.id)) || String(x.id).localeCompare(String(y.id)); });
    var resumen = { nueva: 0, cerrada: 0, reabierta: 0, estado: 0, retirada: 0 };
    cambios.forEach(function (c) { resumen[c.tipo]++; });
    return { cambios: cambios, resumen: resumen };
  }

  // ---------- Avance por estado ----------
  /** Tabla estado -> avance acumulado según la configuración de acciones. */
  function tablaAvance(cfg) {
    cfg = cfg || AVANCE;
    var acumulado = 0, filas = [];
    cfg.acciones.forEach(function (a, i) {
      var antes = acumulado;
      acumulado += Number(a.esfuerzo) || 0;
      filas.push({ orden: i, accion: a.accion, esfuerzo: Number(a.esfuerzo) || 0, estados: a.estados || [],
        avance: cfg.incluirActual === false ? antes : acumulado });
    });
    return { filas: filas, total: acumulado };
  }

  /** Avance (0 a 1) de una HU y la acción en la que está. */
  function avanceHU(h, tabla) {
    if (h.fase === 'Cerrado') return { avance: 1, accion: 'Ejecutada', orden: 99 };
    var e = normalizar(h.estado);
    for (var i = 0; i < tabla.filas.length; i++) {
      var f = tabla.filas[i];
      if (f.estados.some(function (x) { return normalizar(x) === e; })) return { avance: Math.min(1, f.avance), accion: f.accion, orden: f.orden };
    }
    return { avance: 0, accion: 'Sin avance definido', orden: -1 };
  }

  function calcularAvance(registros, cfg) {
    var tabla = tablaAvance(cfg);
    var lista = registros.filter(function (h) { return h.considerada; }).map(function (h) {
      var a = avanceHU(h, tabla);
      return { hu: h, avance: a.avance, accion: a.accion, orden: a.orden };
    });
    var n = lista.length;
    var suma = lista.reduce(function (s, x) { return s + x.avance; }, 0);
    var porEstado = {};
    lista.forEach(function (x) {
      var k = x.hu.estado;
      if (!porEstado[k]) porEstado[k] = { estado: k, accion: x.accion, orden: x.orden, avance: x.avance, n: 0, originales: {} };
      porEstado[k].n++;
      if (x.hu.estadoOriginal && x.hu.estadoOriginal !== k) porEstado[k].originales[x.hu.estadoOriginal] = (porEstado[k].originales[x.hu.estadoOriginal] || 0) + 1;
    });
    var estados = Object.keys(porEstado).map(function (k) { return porEstado[k]; })
      .sort(function (a, b) { return a.avance - b.avance || a.orden - b.orden || b.n - a.n; });
    estados.forEach(function (e) { e.aporte = n ? e.avance * e.n / n : 0; });
    var tramos = [
      { nombre: 'Sin avance (0%)', min: 0, max: 0 },
      { nombre: '1% a 49%', min: 0.0001, max: 0.4999 },
      { nombre: '50% a 99%', min: 0.5, max: 0.9999 },
      { nombre: 'Ejecutadas (100%)', min: 1, max: 1 }
    ].map(function (t) {
      var c = lista.filter(function (x) { return x.avance >= t.min - 1e-9 && x.avance <= t.max + 1e-9; }).length;
      return { nombre: t.nombre, n: c, pct: n ? c / n : 0 };
    });
    var agrupar = function (campo) {
      var g = {};
      lista.forEach(function (x) {
        var k = x.hu[campo] || (campo === 'asignado' ? 'Sin asignar' : '(vacío)');
        if (!g[k]) g[k] = { nombre: k, n: 0, suma: 0, ejecutadas: 0 };
        g[k].n++; g[k].suma += x.avance; if (x.avance >= 1) g[k].ejecutadas++;
      });
      return Object.keys(g).map(function (k) { var f = g[k]; f.promedio = f.n ? f.suma / f.n : 0; return f; })
        .sort(function (a, b) { return b.n - a.n || a.nombre.localeCompare(b.nombre); });
    };
    return {
      tabla: tabla,
      lista: lista,
      n: n,
      promedio: n ? suma / n : null,
      ejecutadas: lista.filter(function (x) { return x.avance >= 1; }).length,
      enCurso: lista.filter(function (x) { return x.avance > 0 && x.avance < 1; }).length,
      sinAvance: lista.filter(function (x) { return x.avance === 0; }).length,
      porEstado: estados,
      tramos: tramos,
      porCliente: agrupar('cliente'),
      porSprint: agrupar('sprint'),
      porModulo: agrupar('modulo'),
      porAsignado: agrupar('asignado')
    };
  }

  // ---------- Sprint: paquetes Diseño y Desarrollo ----------
  /** Avance (0 a 1) de un paquete Diseño y Desarrollo según su estado. */
  function avancePaquete(p, cfg) {
    cfg = cfg || AVANCE_DYD;
    if ((cfg.terminados || []).some(function (t) { return texto(t) && coincideCriterio(p.estado, t); })) {
      return { avance: 1, accion: 'Terminado', orden: 99 };
    }
    var tabla = tablaAvance(cfg);
    var total = tabla.total || 1;
    var e = normalizar(p.estado);
    for (var i = 0; i < tabla.filas.length; i++) {
      var f = tabla.filas[i];
      if (f.estados.some(function (x) { return normalizar(x) === e; })) return { avance: Math.min(1, f.avance / total), accion: f.accion, orden: f.orden };
    }
    return { avance: 0, accion: 'Sin iniciar', orden: -1 };
  }

  /** Paquetes Diseño y Desarrollo (únicos) relacionados con las HU, con su avance. */
  function paquetesDyD(registros, cfg) {
    cfg = cfg || AVANCE_DYD;
    var vistos = {}, lista = [];
    registros.forEach(function (h) {
      (h.relaciones || []).forEach(function (p) {
        if (vistos[p.id] || !coincideCriterio(p.tipo, cfg.tipo)) return;
        vistos[p.id] = true;
        var a = avancePaquete(p, cfg);
        lista.push({ paquete: p, avance: a.avance, accion: a.accion, orden: a.orden });
      });
    });
    return lista;
  }

  function resumenPaquetes(nombre, items) {
    var n = items.length;
    var puntos = items.reduce(function (s, x) { return s + x.paquete.puntos; }, 0);
    var fin = items.reduce(function (m, x) { return x.paquete.fechaFin && (!m || x.paquete.fechaFin > m) ? x.paquete.fechaFin : m; }, null);
    return {
      nombre: nombre, n: n, items: items,
      avance: n ? items.reduce(function (s, x) { return s + x.avance; }, 0) / n : null,
      terminados: contar(items, function (x) { return x.avance >= 1; }),
      enPruebas: contar(items, function (x) { return x.avance > 0 && x.avance < 1; }),
      sinAvance: contar(items, function (x) { return x.avance === 0; }),
      puntos: puntos,
      puntosTerminados: items.reduce(function (s, x) { return s + (x.avance >= 1 ? x.paquete.puntos : 0); }, 0),
      fin: fin
    };
  }

  /**
   * Avance del sprint y avance por ID a partir de los paquetes Diseño y Desarrollo
   * relacionados con las HU. El sprint es el del paquete (sin sprint no entra al avance del sprint).
   */
  function calcularSprint(registros, cfg) {
    cfg = cfg || AVANCE_DYD;
    var lista = paquetesDyD(registros, cfg);
    var grupos = {}, sinSprint = [];
    lista.forEach(function (x) {
      var s = x.paquete.sprint;
      if (!s) { sinSprint.push(x); return; }
      (grupos[s] = grupos[s] || []).push(x);
    });
    var porSprint = Object.keys(grupos).map(function (k) { return resumenPaquetes(k, grupos[k]); }).sort(function (a, b) { return compararSprint(b.nombre, a.nombre); });
    var porAvance = {};
    lista.forEach(function (x) { porAvance[x.paquete.id] = x; });
    var porHU = registros.map(function (h) {
      var items = (h.relaciones || []).map(function (p) { return porAvance[p.id]; }).filter(Boolean);
      if (!items.length) return null;
      var r = resumenPaquetes(h.id, items);
      r.hu = h;
      return r;
    }).filter(Boolean);
    return { tabla: tablaAvance(cfg), paquetes: lista, porSprint: porSprint, sinSprint: resumenPaquetes('Sin sprint', sinSprint), porHU: porHU, total: resumenPaquetes('Total', lista) };
  }

  /** Orden natural de nombres de sprint (Sprint 4 < SPRINT 5 < Sprint 10). */
  function compararSprint(a, b) {
    return String(a).localeCompare(String(b), 'es', { numeric: true, sensitivity: 'base' });
  }

  /**
   * Burndown de un sprint a partir de varios cortes (los informes cargados como Sprint).
   * @param cortes  [{ corte: 'AAAA-MM-DD', registros }] en cualquier orden
   * Devuelve los puntos pendientes de cada corte, la línea ideal y los ID que salieron o entraron al sprint.
   */
  function calcularBurndown(cortes, sprint, cfg) {
    var orden = cortes.slice().sort(function (a, b) { return String(a.corte).localeCompare(String(b.corte)); });
    var anterior = null, fin = null, movimientos = [];
    var puntos = orden.map(function (c) {
      var lista = paquetesDyD(c.registros, cfg);
      var mapa = {};
      lista.forEach(function (x) { mapa[x.paquete.id] = x; });
      var en = lista.filter(function (x) { return x.paquete.sprint === sprint; });
      en.forEach(function (x) { if (x.paquete.fechaFin && (!fin || x.paquete.fechaFin > fin)) fin = x.paquete.fechaFin; });
      if (anterior) {
        Object.keys(anterior).forEach(function (id) {
          var a = anterior[id], b = mapa[id];
          if (a.paquete.sprint === sprint && b && b.paquete.sprint !== sprint) {
            movimientos.push({ tipo: 'salio', id: id, asunto: b.paquete.asunto, de: sprint, a: b.paquete.sprint || 'Sin sprint', corte: c.corte, puntos: b.paquete.puntos, estado: b.paquete.estado, hus: b.paquete.hus });
          }
        });
        en.forEach(function (b) {
          var a = anterior[b.paquete.id];
          if (a && a.paquete.sprint !== sprint) {
            movimientos.push({ tipo: 'entro', id: b.paquete.id, asunto: b.paquete.asunto, de: a.paquete.sprint || 'Sin sprint', a: sprint, corte: c.corte, puntos: b.paquete.puntos, estado: b.paquete.estado, hus: b.paquete.hus });
          }
        });
      }
      anterior = mapa;
      var r = resumenPaquetes(sprint, en);
      return { corte: c.corte, paquetes: r.n, puntos: r.puntos, pendientes: r.puntos - r.puntosTerminados, paquetesPendientes: r.n - r.terminados };
    });
    var hayPuntos = puntos.some(function (p) { return p.puntos > 0; });
    var iniciales = puntos.length ? puntos[0].pendientes : 0;
    var inicialesPaq = puntos.length ? puntos[0].paquetesPendientes : 0;
    var finISO = fin ? fin.getFullYear() + '-' + String(fin.getMonth() + 1).padStart(2, '0') + '-' + String(fin.getDate()).padStart(2, '0') : null;
    var ultimo = puntos.length ? puntos[puntos.length - 1].corte : null;
    if (!finISO || (ultimo && finISO < ultimo)) finISO = ultimo;
    var dia = function (iso) { var p = iso.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000; };
    // Línea ideal: de lo pendiente en el primer corte a 0 en la fecha de fin del sprint
    var ideal = function (iso, inicial) {
      if (!puntos.length) return null;
      var d0 = dia(puntos[0].corte), d1 = dia(finISO);
      if (d1 <= d0) return iso === finISO ? 0 : inicial;
      return Math.max(0, inicial * (1 - (dia(iso) - d0) / (d1 - d0)));
    };
    puntos.forEach(function (p) { p.ideal = ideal(p.corte, iniciales); p.idealPaquetes = ideal(p.corte, inicialesPaq); });
    return { sprint: sprint, puntos: puntos, fin: finISO, hayPuntos: hayPuntos, iniciales: iniciales, inicialesPaquetes: inicialesPaq, movimientos: movimientos };
  }

  return {
    AVANCE_DYD: AVANCE_DYD, avancePaquete: avancePaquete, paquetesDyD: paquetesDyD, calcularSprint: calcularSprint,
    calcularBurndown: calcularBurndown, compararSprint: compararSprint,
    AVANCE: AVANCE, tablaAvance: tablaAvance, avanceHU: avanceHU, calcularAvance: calcularAvance,
    ESTADOS: ESTADOS, EQUIVALENCIAS: EQUIVALENCIAS, FASES: FASES, TIPOS_SOLICITUD: TIPOS_SOLICITUD, KPIS: KPIS, PARAMETROS: PARAMETROS,
    EXCLUIDO: EXCLUIDO, SIN_MAPEAR: SIN_MAPEAR, SIN_MODULO: SIN_MODULO, SIN_SPRINT: SIN_SPRINT,
    parseFecha: parseFecha, parsearHoja: parsearHoja, calcular: calcular,
    evaluar: evaluar, coincideCriterio: coincideCriterio, compararHU: compararHU
  };
});
