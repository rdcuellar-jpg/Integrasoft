/*
 * Motor de cálculo de indicadores — replica las fórmulas del libro
 * "Base para Indicadores.xlsx" (hojas Datos Indicadores, Indicadores Seguimiento
 * y Ficha de Indicadores) a partir de la fuente manual de la Hoja1.
 *
 * Funciona en el navegador (window.Indicadores) y en Node (module.exports).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Indicadores = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var EXCLUIDOS = ['Nuevo / Identificado', 'Asignación'];

  // Estado -> Fase (Datos Indicadores!A13:B22)
  var ESTADOS = [
    ['Levantamiento Requerimiento', 'Análisis'],
    ['Realizar Mockups', 'Análisis'],
    ['Documentación', 'Análisis'],
    ['Aprobación cliente', 'Análisis'],
    ['Revisión (sol)', 'Análisis'],
    ['En desarrollo', 'Desarrollo'],
    ['Revisión (IDS)', 'Desarrollo'],
    ['Testing', 'Pruebas'],
    ['Entrega', 'Cerrado'],
    ['CIERRE en Desarrollo o Solución', 'Cerrado']
  ];
  var FASES = ['Análisis', 'Desarrollo', 'Pruebas', 'Cerrado'];
  var TIPOS_SOLICITUD = ['Desarrollo Contratado', 'Mejora Funcional', 'Deuda Tecnica', 'Sin clasificar'];
  var SIN_MODULO = 'Sin módulo identificado';
  var SIN_SPRINT = 'Sin sprint';
  var SIN_CLIENTE = 'Sin cliente';

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

  // KPIs (Datos Indicadores!A49:I58 + Ficha de Indicadores)
  var KPIS = [
    { id: 'deuda', nombre: '% Deuda técnica / backlog total', meta: 0.20, tipoMeta: '<=', grupo: 'plantilla',
      definicion: 'Proporción de ítems del backlog clasificados como deuda técnica respecto al total de ítems del backlog.',
      formula: '(N° de ítems de deuda técnica / N° total de ítems del backlog) x 100', frecuencia: 'Mensual',
      responsable: 'Product Owner', criterio: 'HU con Tipo de Solicitud = Deuda Tecnica / total HU consideradas' },
    { id: 'roadmap', nombre: '% Cumplimiento del Roadmap', meta: 0.90, tipoMeta: '>=', grupo: 'plantilla',
      definicion: 'Porcentaje de entregas realizadas cumpliendo la fecha de finalización en el período respecto a las planificadas.',
      formula: '(N° de entregas realizadas / N° de entregas planificadas) x 100', frecuencia: 'Mensual',
      responsable: 'Director de Soluciones / PO', criterio: 'HU con En flujo = Cerrado (Entrega / CIERRE) / total HU consideradas' },
    { id: 'automatizacion', nombre: '% Automatización de pruebas', meta: 0.70, tipoMeta: '>=', grupo: 'plantilla', manual: true,
      definicion: 'Proporción de casos de prueba ejecutados de forma automatizada respecto al total.',
      formula: '(N° de pruebas automatizadas / N° total de pruebas) x 100', frecuencia: 'Mensual',
      responsable: 'Líder de QA', criterio: 'Sin datos en la base: requiere N° de pruebas automatizadas y total de pruebas' },
    { id: 'refinement', nombre: 'Precisión backlog refinement', meta: 0.85, tipoMeta: '>=', grupo: 'plantilla', manual: true,
      definicion: 'Historias refinadas que no requirieron retrabajo o reapertura una vez ingresadas al sprint.',
      formula: '(N° de historias sin retrabajo / N° total de historias refinadas) x 100', frecuencia: 'Mensual',
      responsable: 'Product Owner', criterio: 'Sin datos en la base: requiere marca de retrabajo/reapertura por HU' },
    { id: 'alcance', nombre: '% Cambios de alcance post-firma', meta: 0.10, tipoMeta: '<=', grupo: 'plantilla', manual: true,
      definicion: 'Requerimientos modificados después de su aprobación formal.',
      formula: '(N° de requerimientos con cambio post-firma / N° total firmados) x 100', frecuencia: 'Mensual',
      responsable: 'Analista Funcional Líder', criterio: 'Sin datos en la base: requiere marca de cambio de alcance tras aprobación' },
    { id: 'ambiguedad', nombre: '% Retrabajo por ambigüedad', meta: 0.08, tipoMeta: '<=', grupo: 'plantilla', manual: true,
      definicion: 'Requerimientos retrabajados por especificaciones ambiguas o incompletas.',
      formula: '(N° retrabajados por ambigüedad / N° total de requerimientos) x 100', frecuencia: 'Mensual',
      responsable: 'Analista Funcional Líder', criterio: 'Sin datos en la base: requiere marca de retrabajo por ambigüedad' },
    { id: 'analisis', nombre: '% HU en Análisis', meta: null, tipoMeta: '', grupo: 'flujo',
      definicion: 'HU en Levantamiento, Mockups, Documentación, Aprobación cliente o Revisión (sol).',
      formula: '(N° HU en Análisis / N° HU consideradas) x 100', frecuencia: 'Mensual',
      responsable: 'Analista Funcional Líder', criterio: 'Levantamiento, Mockups, Documentación, Aprobación cliente, Revisión (sol)' },
    { id: 'desarrollo', nombre: '% HU en Desarrollo', meta: null, tipoMeta: '', grupo: 'flujo',
      definicion: 'HU en estado En desarrollo o Revisión (IDS).',
      formula: '(N° HU en Desarrollo / N° HU consideradas) x 100', frecuencia: 'Mensual',
      responsable: 'Líder de Desarrollo', criterio: 'En desarrollo, Revisión (IDS)' },
    { id: 'pruebas', nombre: '% HU en Pruebas', meta: null, tipoMeta: '', grupo: 'flujo',
      definicion: 'HU en estado Testing.',
      formula: '(N° HU en Testing / N° HU consideradas) x 100', frecuencia: 'Mensual',
      responsable: 'Líder de QA', criterio: 'Testing' },
    { id: 'abiertas', nombre: '% HU Abiertas', meta: null, tipoMeta: '', grupo: 'flujo',
      definicion: 'HU con En flujo = Abierto sobre el total considerado.',
      formula: '(N° HU abiertas / N° HU consideradas) x 100', frecuencia: 'Mensual',
      responsable: 'Product Owner', criterio: 'HU con En flujo = Abierto / total HU consideradas' }
  ];

  // Columnas de la Hoja1 (primer bloque "Paquetes de trabajo"; el bloque
  // "Relaciones" repite encabezados y se ignora).
  var COLUMNAS = {
    id: 'ID', asunto: 'Asunto', tipo: 'Tipo', cliente: 'Cliente', estado: 'Estado',
    enFlujo: 'En flujo', trabajoTotal: 'Trabajo total', asignado: 'Asignado a',
    fechaInicio: 'Fecha de inicio', fechaFin: 'Fecha de finalización',
    responsable: 'Responsable', tipoSolicitud: 'Tipo de Solicitud', sprint: 'Sprint'
  };

  function texto(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function normalizar(s) {
    return texto(s).toLowerCase();
  }

  /** Convierte "dd-mm-aaaa", Date o serial de Excel a Date (o null). */
  function parseFecha(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number') {
      // Serial de Excel (sistema 1900)
      var ms = Math.round((v - 25569) * 86400000);
      var d = new Date(ms);
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    var s = texto(v);
    var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    return null;
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

  /**
   * Convierte la matriz de la Hoja1 (array de arrays) en registros únicos de HU.
   * La base repite la HU por cada relación: se toma la primera aparición del ID
   * (equivalente a Aux: Registro único HU).
   */
  function parsearHoja(matriz, palabras) {
    palabras = palabras || PALABRAS_CLAVE;
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
      var estado = texto(get('estado'));
      var registro = {
        id: id,
        asunto: texto(get('asunto')),
        tipo: texto(get('tipo')),
        cliente: texto(get('cliente')) || SIN_CLIENTE,
        estado: estado,
        enFlujo: texto(get('enFlujo')),
        trabajoTotal: get('trabajoTotal') === '' || get('trabajoTotal') == null ? null : Number(get('trabajoTotal')),
        asignado: texto(get('asignado')),
        fechaInicio: parseFecha(get('fechaInicio')),
        fechaFin: fin,
        mesFin: claveMes(fin),
        responsable: texto(get('responsable')),
        tipoSolicitud: texto(get('tipoSolicitud')) || 'Sin clasificar',
        sprint: texto(get('sprint')) || SIN_SPRINT,
        fila: r + 1
      };
      registro.modulo = asignarModulo(registro.asunto, palabras);
      registro.fase = faseDe(estado);
      registro.considerada = EXCLUIDOS.indexOf(estado) === -1;
      registros.push(registro);
    }
    return { registros: registros, filasLeidas: filasLeidas, filaEncabezado: enc.fila + 1 };
  }

  function faseDe(estado) {
    for (var i = 0; i < ESTADOS.length; i++) if (ESTADOS[i][0] === estado) return ESTADOS[i][1];
    if (EXCLUIDOS.indexOf(estado) !== -1) return 'Excluido';
    return 'Sin fase';
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

  /** Tabla cruzada: filas = valores de `campo`, columnas = Tipo de Solicitud. */
  function cruzarPorTipo(lista, campo, ordenFilas) {
    var filas = {};
    lista.forEach(function (h) {
      var k = h[campo];
      if (!filas[k]) {
        filas[k] = { nombre: k, total: 0 };
        TIPOS_SOLICITUD.forEach(function (t) { filas[k][t] = 0; });
      }
      var t = TIPOS_SOLICITUD.indexOf(h.tipoSolicitud) === -1 ? 'Sin clasificar' : h.tipoSolicitud;
      filas[k][t]++;
      filas[k].total++;
    });
    var arr = Object.keys(filas).map(function (k) { return filas[k]; });
    arr.sort(function (a, b) {
      if (ordenFilas) {
        var ia = ordenFilas.indexOf(a.nombre), ib = ordenFilas.indexOf(b.nombre);
        if (ia !== ib) return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
      }
      return b.total - a.total || String(a.nombre).localeCompare(String(b.nombre));
    });
    var total = lista.length;
    arr.forEach(function (f) { f.pct = total ? f.total / total : 0; });
    return arr;
  }

  /**
   * Calcula todos los indicadores.
   * @param registros  registros únicos (ya filtrados por el usuario)
   * @param config     { metas: {id: {meta, tipoMeta}}, manuales: {id: {num, den}} }
   */
  function calcular(registros, config) {
    config = config || {};
    var metas = config.metas || {};
    var manuales = config.manuales || {};

    var consideradas = registros.filter(function (h) { return h.considerada; });
    var nC = consideradas.length;

    var base = {
      total: registros.length,
      excluidasNuevo: contar(registros, function (h) { return h.estado === 'Nuevo / Identificado'; }),
      excluidasAsignacion: contar(registros, function (h) { return h.estado === 'Asignación'; }),
      consideradas: nC
    };

    var porEstado = ESTADOS.map(function (e) {
      var n = contar(consideradas, function (h) { return h.estado === e[0]; });
      return { estado: e[0], fase: e[1], n: n, pct: nC ? n / nC : 0 };
    });
    var sinFase = consideradas.filter(function (h) { return h.fase === 'Sin fase'; });
    var estadosNoListados = {};
    sinFase.forEach(function (h) { estadosNoListados[h.estado || '(vacío)'] = (estadosNoListados[h.estado || '(vacío)'] || 0) + 1; });

    var porFase = FASES.map(function (f) {
      var n = porEstado.filter(function (e) { return e.fase === f; }).reduce(function (s, e) { return s + e.n; }, 0);
      return { fase: f, n: n };
    });
    var totalFases = porFase.reduce(function (s, f) { return s + f.n; }, 0);
    porFase.forEach(function (f) { f.pct = totalFases ? f.n / totalFases : 0; });

    var abiertas = contar(consideradas, function (h) { return h.enFlujo === 'Abierto'; });
    var cerradas = contar(consideradas, function (h) { return h.enFlujo === 'Cerrado'; });
    var totalFlujo = abiertas + cerradas;
    var porFlujo = [
      { nombre: 'Abierto', n: abiertas, pct: totalFlujo ? abiertas / totalFlujo : 0 },
      { nombre: 'Cerrado', n: cerradas, pct: totalFlujo ? cerradas / totalFlujo : 0 }
    ];

    var porTipo = ['Mejora Funcional', 'Deuda Tecnica', 'Desarrollo Contratado', 'Sin clasificar'].map(function (t) {
      var n = contar(consideradas, function (h) {
        return t === 'Sin clasificar' ? TIPOS_SOLICITUD.indexOf(h.tipoSolicitud) === -1 || h.tipoSolicitud === t : h.tipoSolicitud === t;
      });
      return { nombre: t, n: n, pct: nC ? n / nC : 0 };
    });
    var totalTipo = porTipo.reduce(function (s, t) { return s + t.n; }, 0);
    var deuda = porTipo[1].n;

    var fase = {};
    porFase.forEach(function (f) { fase[f.fase] = f.n; });

    var numDen = {
      deuda: [deuda, totalTipo],
      roadmap: [cerradas, totalFlujo],
      analisis: [fase['Análisis'], totalFases],
      desarrollo: [fase['Desarrollo'], totalFases],
      pruebas: [fase['Pruebas'], totalFases],
      abiertas: [abiertas, totalFlujo]
    };

    var kpis = KPIS.map(function (k) {
      var nd = numDen[k.id];
      if (k.manual) {
        var man = manuales[k.id] || {};
        var num = man.num === '' || man.num === undefined || man.num === null ? null : Number(man.num);
        var den = man.den === '' || man.den === undefined || man.den === null ? null : Number(man.den);
        nd = [num, den];
      }
      var mt = metas[k.id] || {};
      var meta = mt.meta !== undefined ? mt.meta : k.meta;
      var tipoMeta = mt.tipoMeta !== undefined ? mt.tipoMeta : k.tipoMeta;
      var res = ratio(nd[0], nd[1]);
      return Object.assign({}, k, {
        numerador: nd[0], denominador: nd[1], resultado: res,
        meta: meta, tipoMeta: tipoMeta, estado: evaluar(res, meta, tipoMeta)
      });
    });

    // HU por mes de finalización
    var meses = {};
    consideradas.forEach(function (h) {
      if (h.enFlujo !== 'Abierto' && h.enFlujo !== 'Cerrado') return;
      var k = h.mesFin || 'sin-fecha';
      if (!meses[k]) meses[k] = { mes: k, cerradas: 0, abiertas: 0 };
      if (h.enFlujo === 'Cerrado') meses[k].cerradas++; else meses[k].abiertas++;
    });
    var sinFecha = meses['sin-fecha'] || { mes: 'sin-fecha', cerradas: 0, abiertas: 0 };
    delete meses['sin-fecha'];
    var claves = Object.keys(meses).sort();
    var mensual = [];
    if (claves.length) {
      // Completa los meses intermedios sin datos
      var p = claves[0].split('-').map(Number), u = claves[claves.length - 1].split('-').map(Number);
      var y = p[0], mo = p[1];
      while (y < u[0] || (y === u[0] && mo <= u[1])) {
        var key = y + '-' + String(mo).padStart(2, '0');
        mensual.push(meses[key] || { mes: key, cerradas: 0, abiertas: 0 });
        mo++; if (mo > 12) { mo = 1; y++; }
      }
    }

    // Indicadores de seguimiento: HU abiertas
    var abiertasLista = consideradas.filter(function (h) { return h.enFlujo === 'Abierto'; });
    var seguimiento = {
      total: abiertasLista.length,
      porCliente: cruzarPorTipo(abiertasLista, 'cliente'),
      porSprint: cruzarPorTipo(abiertasLista, 'sprint'),
      porModulo: cruzarPorTipo(abiertasLista, 'modulo'),
      porResponsable: cruzarPorTipo(abiertasLista, 'responsable')
    };
    // "Sin sprint" / "Sin módulo" al final
    [['porSprint', SIN_SPRINT], ['porModulo', SIN_MODULO]].forEach(function (p) {
      var arr = seguimiento[p[0]];
      var i = arr.findIndex(function (f) { return f.nombre === p[1]; });
      if (i !== -1) arr.push(arr.splice(i, 1)[0]);
    });

    return {
      base: base,
      porEstado: porEstado,
      estadosNoListados: estadosNoListados,
      porFase: porFase,
      porFlujo: porFlujo,
      porTipo: porTipo,
      kpis: kpis,
      mensual: mensual,
      sinFecha: sinFecha,
      seguimiento: seguimiento,
      consideradas: consideradas
    };
  }

  return {
    ESTADOS: ESTADOS, FASES: FASES, TIPOS_SOLICITUD: TIPOS_SOLICITUD, KPIS: KPIS,
    PALABRAS_CLAVE: PALABRAS_CLAVE, EXCLUIDOS: EXCLUIDOS, SIN_MODULO: SIN_MODULO, SIN_SPRINT: SIN_SPRINT,
    parseFecha: parseFecha, parsearHoja: parsearHoja, asignarModulo: asignarModulo, calcular: calcular,
    evaluar: evaluar
  };
});
