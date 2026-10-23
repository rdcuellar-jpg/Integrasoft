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

  var PARAMETROS = {
    tipoExcluidoBacklog: 'Deuda Tecnica', // J55
    marcaAlcance: 'Cambio de Alcance',    // J58
    marcaRetrabajo: 'Retrabajo*',         // J59 (admite comodines * y ?)
    horasContratadas: null,               // C56
    anio: null                            // E65 (null = año actual)
  };

  // Tabla de módulos por palabra clave (Indicadores Seguimiento!J6:K60).
  // Se aplica la primera coincidencia de arriba hacia abajo.
  var PALABRAS_CLAVE = [
    ['contabilidad -', 'Contabilidad'], ['cesantias', 'Nómina'], ['nomina', 'Nómina'],
    ['nómina', 'Nómina'], ['viaticos', 'Nómina'], ['horas extras', 'Nómina'],
    ['parafiscales', 'Nómina'], ['vacaciones', 'Nómina'], ['salvamento', 'Activos fijos'],
    ['activos', 'Activos fijos'], ['tesorer', 'Tesorería'], ['retenciones bancarias', 'Tesorería'],
    ['caja menor', 'Tesorería'], ['controlblue', 'ControlBlue / Órdenes de trabajo'],
    ['ordenes de trabajo', 'ControlBlue / Órdenes de trabajo'], ['cdp', 'Presupuesto'],
    ['cufe', 'Facturación y Liquidaciones'], ['nota crédito', 'Facturación y Liquidaciones'],
    ['factura', 'Facturación y Liquidaciones'], ['liquidac', 'Facturación y Liquidaciones'],
    ['predial', 'Facturación y Liquidaciones'], ['presupuesto', 'Presupuesto'],
    ['documentos compra', 'Compras'], ['las OPS', 'Contratación'], ['contab', 'Contabilidad'],
    ['libro auxiliar', 'Contabilidad'], ['plan de cuentas', 'Contabilidad'],
    ['nota de cierre', 'Contabilidad'], ['el iva', 'Contabilidad'],
    ['elaboracion de documentos', 'Contabilidad'], ['ctos automaticos', 'Contabilidad']
  ];

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
    precision: 'Precisión HU', responsable: 'Responsable'
  };

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

  /** Busca la fila de encabezados y devuelve índice de fila + mapa campo->columna. */
  function detectarEncabezados(matriz) {
    for (var r = 0; r < Math.min(matriz.length, 20); r++) {
      var fila = matriz[r] || [];
      var mapa = {};
      for (var c = 0; c < fila.length; c++) {
        var h = texto(fila[c]);
        for (var campo in COLUMNAS) {
          if (h === COLUMNAS[campo] && mapa[campo] === undefined) mapa[campo] = c;
        }
      }
      if (mapa.id !== undefined && mapa.estado !== undefined) return { fila: r, mapa: mapa };
    }
    return null;
  }

  function asignarModulo(asunto, palabras) {
    var a = normalizar(asunto);
    for (var i = 0; i < palabras.length; i++) {
      var kw = normalizar(palabras[i][0]);
      if (kw && a.indexOf(kw) !== -1) return palabras[i][1];
    }
    return SIN_MODULO;
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
   * @param opciones { palabras, estados } (también acepta solo el array de palabras)
   */
  function parsearHoja(matriz, opciones) {
    if (Array.isArray(opciones)) opciones = { palabras: opciones };
    opciones = opciones || {};
    var palabras = opciones.palabras || PALABRAS_CLAVE;
    var fases = mapaFases(opciones.estados);
    var equivalencias = opciones.equivalencias || EQUIVALENCIAS;
    var enc = detectarEncabezados(matriz);
    if (!enc) throw new Error('No se encontró la fila de encabezados (se esperan columnas "ID" y "Estado").');
    var m = enc.mapa;
    var vistos = {};
    var registros = [];
    var filasLeidas = 0;
    for (var r = enc.fila + 1; r < matriz.length; r++) {
      var fila = matriz[r] || [];
      var id = texto(fila[m.id]);
      if (!id) continue;
      filasLeidas++;
      if (vistos[id]) continue;
      vistos[id] = true;
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
        fila: r + 1
      };
      registro.modulo = asignarModulo(registro.asunto, palabras);
      registros.push(registro);
    }
    var columnas = {};
    Object.keys(m).forEach(function (k) { columnas[k] = true; });
    return { registros: registros, filasLeidas: filasLeidas, filaEncabezado: enc.fila + 1, columnas: columnas };
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

    var numDen = {
      deuda: [contar(consideradas, function (h) { return h.tipoSolicitud === 'Deuda Tecnica'; }), nC],
      backlog: [contar(sinExcluido, function (h) { return h.enFlujo === 'Cerrado'; }), sinExcluido.length],
      horas: [horas, horasContratadas],
      refinement: [contar(consideradas, function (h) { return h.precision === ''; }), nC],
      alcance: [contar(consideradas, function (h) { return coincideCriterio(h.precision, par.marcaAlcance); }), nC],
      ambiguedad: [contar(consideradas, function (h) { return coincideCriterio(h.precision, par.marcaRetrabajo); }), nC],
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

  return {
    ESTADOS: ESTADOS, EQUIVALENCIAS: EQUIVALENCIAS, FASES: FASES, TIPOS_SOLICITUD: TIPOS_SOLICITUD, KPIS: KPIS, PARAMETROS: PARAMETROS,
    PALABRAS_CLAVE: PALABRAS_CLAVE, EXCLUIDO: EXCLUIDO, SIN_MAPEAR: SIN_MAPEAR, SIN_MODULO: SIN_MODULO, SIN_SPRINT: SIN_SPRINT,
    parseFecha: parseFecha, parsearHoja: parsearHoja, asignarModulo: asignarModulo, calcular: calcular,
    evaluar: evaluar, coincideCriterio: coincideCriterio, compararHU: compararHU
  };
});
