/* Interfaz de la aplicación de indicadores. La lógica de cálculo vive en indicadores.js. */
(function () {
  'use strict';
  var I = window.Indicadores;
  var CLAVE = 'indicadores-integrasoft:';

  var estado = {
    matriz: null,
    fuente: '',
    cargado: null,
    registros: [],
    filtros: {},
    resultado: null,
    graficos: {},
    vistaKpi: leer('vistaKpi', 'tabla'),
    orden: { campo: 'id', asc: true },
    config: {
      metas: leer('metas', {}),
      parametros: leer('parametros', {}),
      estados: leer('estados', null) || I.ESTADOS.slice(),
      palabras: leer('palabras', null) || I.PALABRAS_CLAVE.slice()
    }
  };

  // ---------- utilidades ----------
  function leer(k, def) {
    try { var v = localStorage.getItem(CLAVE + k); return v ? JSON.parse(v) : def; } catch (e) { return def; }
  }
  function guardar(k, v) {
    try { localStorage.setItem(CLAVE + k, JSON.stringify(v)); } catch (e) { /* almacenamiento no disponible */ }
  }
  function $(sel, ctx) { return (ctx || document).querySelector(sel); }
  function $$(sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); }
  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  var fmtNum = new Intl.NumberFormat('es-CO');
  function pct(v, dec) {
    if (v === null || v === undefined) return 'N/D';
    return (v * 100).toLocaleString('es-CO', { minimumFractionDigits: dec === undefined ? 1 : dec, maximumFractionDigits: dec === undefined ? 1 : dec }) + '%';
  }
  function fecha(d) {
    if (!d) return '';
    return String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear();
  }
  var MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  function nombreMes(clave) {
    var p = clave.split('-');
    return MESES[+p[1] - 1] + ' ' + p[0];
  }
  function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
  function toast(msg) {
    var t = $('#toast');
    t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, 4000);
  }

  // ---------- carga de datos ----------
  function cargarArchivo(file) {
    if (!file) return;
    if (typeof XLSX === 'undefined') { toast('No se pudo cargar la librería de lectura de Excel (revise la conexión).'); return; }
    var lector = new FileReader();
    lector.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        var nombre = wb.SheetNames.indexOf('Hoja1') !== -1 ? 'Hoja1' : null;
        var matriz = null;
        var candidatas = nombre ? [nombre] : wb.SheetNames;
        for (var i = 0; i < candidatas.length && !matriz; i++) {
          var m = XLSX.utils.sheet_to_json(wb.Sheets[candidatas[i]], { header: 1, raw: true, defval: '' });
          try { I.parsearHoja(m); matriz = m; nombre = candidatas[i]; } catch (err) { /* probar la siguiente hoja */ }
        }
        if (!matriz) throw new Error('El archivo no tiene una hoja con las columnas ID y Estado.');
        // Las fechas como Date no sobreviven a JSON: se pasan a texto dd-mm-aaaa
        matriz = matriz.map(function (fila) {
          return fila.map(function (v) {
            return v instanceof Date ? String(v.getDate()).padStart(2, '0') + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + v.getFullYear() : v;
          });
        });
        usarMatriz(matriz, file.name + ' — ' + nombre, new Date());
        toast('Datos cargados desde ' + file.name);
      } catch (err) {
        toast('Error al leer el archivo: ' + err.message);
      }
    };
    lector.readAsArrayBuffer(file);
  }

  function cargarEjemplo() {
    fetch('data/hoja1-ejemplo.json')
      .then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); })
      .then(function (d) { usarMatriz(d.matriz, d.fuente + ' (ejemplo)', new Date()); toast('Datos de ejemplo cargados'); })
      .catch(function () { toast('No se pudieron leer los datos de ejemplo. Abra la app desde un servidor web (ver README).'); });
  }

  function usarMatriz(matriz, fuente, cuando) {
    var p = I.parsearHoja(matriz, opcionesParseo());
    estado.matriz = matriz;
    estado.fuente = fuente;
    estado.cargado = cuando;
    estado.registros = p.registros;
    estado.filasLeidas = p.filasLeidas;
    estado.columnas = p.columnas;
    guardar('datos', { matriz: matriz, fuente: fuente, cargado: cuando.toISOString() });
    $('#zona-carga').hidden = true;
    $('#app').hidden = false;
    construirFiltros();
    actualizar();
  }

  function opcionesParseo() {
    return { palabras: estado.config.palabras, estados: estado.config.estados };
  }

  function reparsear() {
    if (!estado.matriz) return;
    estado.registros = I.parsearHoja(estado.matriz, opcionesParseo()).registros;
    construirFiltros();
    actualizar();
  }

  // ---------- filtros ----------
  var CAMPOS_FILTRO = ['cliente', 'tipoSolicitud', 'sprint', 'responsable', 'asignado', 'modulo', 'precision'];

  function construirFiltros() {
    // Oculta los filtros y columnas cuya columna no viene en la Hoja1 cargada
    $$('[data-requiere]').forEach(function (el) {
      var campo = el.getAttribute('data-requiere');
      el.hidden = !(estado.columnas && estado.columnas[campo]);
      if (el.hidden) delete estado.filtros[campo];
    });
    CAMPOS_FILTRO.forEach(function (campo) {
      var sel = $('select[data-filtro="' + campo + '"]');
      var valores = {};
      estado.registros.forEach(function (r) { valores[r[campo] || '(vacío)'] = (valores[r[campo] || '(vacío)'] || 0) + 1; });
      var actual = estado.filtros[campo] || '';
      var claves = Object.keys(valores).sort(function (a, b) { return a.localeCompare(b, 'es'); });
      if (actual && !valores[actual]) { actual = ''; delete estado.filtros[campo]; }
      sel.innerHTML = '<option value="">Todos</option>' + claves.map(function (k) {
        return '<option value="' + esc(k) + '"' + (k === actual ? ' selected' : '') + '>' + esc(k) + ' (' + valores[k] + ')</option>';
      }).join('');
      sel.classList.toggle('activo', !!actual);
    });
  }

  function registrosFiltrados() {
    return estado.registros.filter(function (r) {
      for (var campo in estado.filtros) {
        var v = estado.filtros[campo];
        if (v && (r[campo] || '(vacío)') !== v) return false;
      }
      return true;
    });
  }

  // ---------- render principal ----------
  function actualizar() {
    var filtrados = registrosFiltrados();
    estado.filtrados = filtrados;
    estado.resultado = I.calcular(filtrados, estado.config);
    renderFuente();
    renderResumen();
    renderKpis();
    renderDistribucion();
    renderSeguimiento();
    renderDetalle();
    renderFicha();
    renderConfig();
  }

  function renderFuente() {
    var r = estado.resultado;
    var filtros = Object.keys(estado.filtros).filter(function (k) { return estado.filtros[k]; }).length;
    $('#fuente-info').innerHTML = 'Corte cargado desde <strong>' + esc(estado.fuente) + '</strong>' +
      (estado.cargado ? ' el ' + esc(estado.cargado.toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' })) : '') +
      ', con ' + fmtNum.format(estado.filasLeidas) + ' filas y ' + fmtNum.format(estado.registros.length) + ' HU únicas.' +
      (filtros ? ' Las cifras reflejan los filtros aplicados.' : '');
    $('#filtros-activos').textContent = filtros ? String(filtros) : '';
    var no = Object.keys(r.estadosNoListados);
    $('#fuente-alertas').innerHTML = no.length
      ? '<span class="alerta">Hay estados sin fase asignada que no se cuentan como HU consideradas: ' + esc(no.map(function (k) { return k + ' (' + r.estadosNoListados[k] + ')'; }).join(', ')) + '. Asígnales una fase en Configuración.</span>'
      : '';
  }

  function listaNatural(items) {
    if (items.length <= 1) return items.join('');
    return items.slice(0, -1).join(', ') + ' y ' + items[items.length - 1];
  }

  function renderResumen() {
    var r = estado.resultado, b = r.base;
    var plantilla = r.kpis.filter(function (k) { return k.grupo === 'plantilla'; });
    var conDatos = plantilla.filter(function (k) { return k.resultado !== null; });
    var enMeta = conDatos.filter(function (k) { return k.estado === 'EN META'; });
    var fuera = conDatos.filter(function (k) { return k.estado === 'FUERA DE META'; });
    var sinDatos = plantilla.filter(function (k) { return k.resultado === null; });

    $('#veredicto').textContent = conDatos.length
      ? enMeta.length + ' de ' + conDatos.length + ' indicadores con datos están en meta'
      : 'Aún no hay indicadores con datos para este corte';
    var partes = [];
    if (fuera.length) {
      partes.push('Fuera de meta: ' + listaNatural(fuera.map(function (k) { return k.nombre.replace(/^% /, '') + ' (' + pct(k.resultado) + ')'; })) + '.');
    }
    if (sinDatos.length) {
      partes.push((sinDatos.length === 1 ? 'Falta el dato de ' : 'Faltan los datos de ') +
        listaNatural(sinDatos.map(function (k) { return k.unidad === 'h' ? 'horas contratadas' : k.nombre.replace(/^% /, ''); })) + ', que se ingresa en Configuración.');
    }
    $('#veredicto-detalle').textContent = partes.join(' ');

    var excl = b.excluidos.reduce(function (s, e) { return s + e.n; }, 0) + b.sinMapear;
    var cifras = [
      ['HU consideradas', fmtNum.format(b.consideradas), 'de ' + fmtNum.format(b.total) + ' únicas; ' + fmtNum.format(excl) + ' excluidas'],
      ['HU abiertas', fmtNum.format(r.porFlujo[0].n), pct(r.porFlujo[0].pct, 0) + ' de las consideradas'],
      ['HU cerradas', fmtNum.format(r.porFlujo[1].n), pct(r.porFlujo[1].pct, 0) + ' de las consideradas'],
      ['Horas ejecutadas', fmtNum.format(r.horas.ejecutadas) + ' h', r.horas.contratadas ? 'de ' + fmtNum.format(r.horas.contratadas) + ' h contratadas' : 'sin horas contratadas']
    ];
    $('#cifras').innerHTML = cifras.map(function (c) {
      return '<div title="' + esc(c[2]) + '"><dt>' + esc(c[0]) + '</dt><dd>' + esc(c[1]) + '</dd></div>';
    }).join('');
  }

  var ICONOS = {
    'EN META': '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9.5"/><path d="M8 12.5l2.7 2.7L16 9.8"/></svg>',
    'FUERA DE META': '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9.5"/><path d="M12 7.5v5.5m0 3.5h.01"/></svg>',
    'SIN DATOS': '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9.5" stroke-dasharray="3 3"/></svg>',
    'SIN META': '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9.5"/></svg>'
  };
  var ETIQUETAS = { 'EN META': 'En meta', 'FUERA DE META': 'Fuera de meta', 'SIN DATOS': 'Sin datos', 'SIN META': 'Sin meta' };
  function pill(est) {
    var cls = est === 'EN META' ? 'en-meta' : est === 'FUERA DE META' ? 'fuera' : '';
    return '<span class="estado ' + cls + '">' + ICONOS[est] + esc(ETIQUETAS[est]) + '</span>';
  }

  function tieneMeta(k) { return k.meta !== null && k.meta !== undefined && k.meta !== ''; }

  function textoDetalle(k) {
    if (k.unidad === 'h') {
      return k.resultado === null
        ? fmtNum.format(k.numerador) + ' h ejecutadas; faltan las horas contratadas'
        : fmtNum.format(k.numerador) + ' h de ' + fmtNum.format(k.denominador) + ' h';
    }
    if (k.resultado === null) return 'Sin datos en la base';
    return fmtNum.format(k.numerador) + ' de ' + fmtNum.format(k.denominador) + ' HU';
  }

  // Gráfico de bala: valor frente a la meta, con la zona fuera de meta rayada
  function bala(k) {
    var meta = tieneMeta(k) ? k.meta : null;
    var escala = Math.max(1, k.resultado || 0, meta || 0);
    var x = function (v) { return Math.min(100, v / escala * 100); };
    var textoMetaCorto = meta === null ? '' : (k.tipoMeta === '<=' ? '≤ ' : '≥ ') + pct(meta, 0);
    var zona = '';
    if (meta !== null) {
      zona = k.tipoMeta === '<='
        ? '<span class="bala-zona" style="left:' + x(meta) + '%;right:0"></span>'
        : '<span class="bala-zona" style="left:0;width:' + x(meta) + '%"></span>';
    }
    var etiqueta = k.nombre + ': ' + pct(k.resultado) + (meta !== null ? ', meta ' + textoMetaCorto : '');
    return '<div class="bala' + (k.resultado === null ? ' vacia' : '') + '" role="img" aria-label="' + esc(etiqueta) + '">' +
      '<span class="bala-pista">' + zona + '</span>' +
      (k.resultado !== null ? '<span class="bala-valor" style="width:' + x(k.resultado) + '%"></span>' : '') +
      (meta !== null ? '<span class="bala-meta' + (x(meta) > 88 ? ' al-final' : '') + '" style="left:' + x(meta) + '%"><span>' + esc(textoMetaCorto) + '</span></span>' : '') +
      '</div>';
  }

  // ---------- vistas de los indicadores de la plantilla ----------
  var VISTAS = ['tabla', 'distancia', 'estado', 'cliente'];

  function renderKpis() {
    var vista = VISTAS.indexOf(estado.vistaKpi) === -1 ? 'tabla' : estado.vistaKpi;
    $$('[data-vista]').forEach(function (b) { b.setAttribute('aria-checked', String(b.getAttribute('data-vista') === vista)); });
    $('#scorecard').hidden = vista !== 'tabla';
    ['distancia', 'estado', 'cliente'].forEach(function (v) { $('#vista-' + v).hidden = v !== vista; });
    ({ tabla: renderTabla, distancia: renderDistancia, estado: renderPorEstado, cliente: renderPorCliente })[vista]();
  }

  function plantilla(resultado) {
    return (resultado || estado.resultado).kpis.filter(function (x) { return x.grupo === 'plantilla'; });
  }

  // Desviación frente a la meta en puntos porcentuales; positiva = favorable
  function desviacion(k) {
    if (k.resultado === null || !tieneMeta(k)) return null;
    var d = (k.resultado - k.meta) * 100;
    return k.tipoMeta === '<=' ? -d : d;
  }
  var fmtPp = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  function fraseDesviacion(k) {
    var d = desviacion(k);
    if (d === null) return k.unidad === 'h' ? 'Falta ingresar las horas contratadas.' : 'Sin datos en la base.';
    var pp = fmtPp.format(Math.abs(d)) + ' puntos';
    var meta = pct(k.meta, 0);
    if (k.tipoMeta === '<=') return d >= 0 ? pp + ' por debajo del máximo de ' + meta + '.' : pp + ' por encima del máximo de ' + meta + '.';
    return d >= 0 ? pp + ' por encima del mínimo de ' + meta + '.' : pp + ' por debajo del mínimo de ' + meta + '.';
  }

  function renderDistancia() {
    var k = plantilla();
    var con = k.filter(function (x) { return desviacion(x) !== null; });
    var sin = k.filter(function (x) { return desviacion(x) === null; });
    var escala = Math.max(10, Math.max.apply(null, con.map(function (x) { return Math.abs(desviacion(x)); }).concat([0])));
    $('#vista-distancia').innerHTML =
      '<div class="dist-ejes"><span>Peor que la meta</span><span>Meta</span><span>Mejor que la meta</span></div>' +
      con.sort(function (a, b) { return desviacion(a) - desviacion(b); }).map(function (x) {
        var d = desviacion(x), ancho = Math.abs(d) / escala * 34;
        var cls = d >= 0 ? 'favor' : 'contra';
        return '<div class="dist-fila">' +
          '<div class="dist-nombre"><strong>' + esc(x.nombre) + '</strong><small>' + pct(x.resultado) + ' frente a ' + (x.tipoMeta === '<=' ? '≤ ' : '≥ ') + pct(x.meta, 0) + '</small></div>' +
          '<div class="dist-pista" role="img" aria-label="' + esc(x.nombre + ': ' + fraseDesviacion(x)) + '">' +
            '<span class="dist-barra ' + cls + '" style="' + (d >= 0 ? 'left:50%' : 'right:50%') + ';width:' + ancho + '%"></span>' +
            '<span class="dist-valor ' + cls + '" style="' + (d >= 0 ? 'left:calc(50% + ' + ancho + '% + 8px)' : 'right:calc(50% + ' + ancho + '% + 8px)') + '">' + (d >= 0 ? '+' : '−') + fmtPp.format(Math.abs(d)) + ' pp</span>' +
          '</div>' +
          '<div class="dist-estado">' + pill(x.estado) + '</div>' +
          '</div>';
      }).join('') +
      (sin.length ? '<p class="nota">Sin datos para comparar: ' + esc(listaNatural(sin.map(function (x) { return x.nombre + (x.unidad === 'h' ? ' (faltan las horas contratadas)' : ''); }))) + '.</p>' : '') +
      '<p class="nota">pp = puntos porcentuales de diferencia entre el resultado y la meta.</p>';
  }

  function renderPorEstado() {
    var k = plantilla();
    var grupos = [
      ['Requieren atención', 'fuera', k.filter(function (x) { return x.estado === 'FUERA DE META'; })],
      ['En meta', 'en-meta', k.filter(function (x) { return x.estado === 'EN META'; })],
      ['Sin datos', '', k.filter(function (x) { return x.estado === 'SIN DATOS' || x.estado === 'SIN META'; })]
    ];
    $('#vista-estado').innerHTML = '<div class="grupos">' + grupos.map(function (g) {
      return '<section class="grupo ' + g[1] + '"><h3>' + ICONOS[g[1] === 'fuera' ? 'FUERA DE META' : g[1] === 'en-meta' ? 'EN META' : 'SIN DATOS'] + esc(g[0]) + ' <span>' + g[2].length + '</span></h3>' +
        (g[2].length ? '<ul>' + g[2].map(function (x) {
          return '<li><div class="g-linea"><span>' + esc(x.nombre) + '</span><b>' + pct(x.resultado) + '</b></div>' +
            '<p>' + esc(fraseDesviacion(x)) + (x.resultado !== null ? ' Base: ' + esc(textoDetalle(x)) + '.' : '') + '</p></li>';
        }).join('') + '</ul>' : '<p class="g-vacio">Ningún indicador en este grupo.</p>') + '</section>';
    }).join('') + '</div>';
  }

  function renderPorCliente() {
    var ids = ['deuda', 'backlog', 'refinement', 'alcance', 'ambiguedad'];
    var cortos = { deuda: 'Deuda técnica', backlog: 'Cumplimiento backlog', refinement: 'Precisión refinement', alcance: 'Cambios de alcance', ambiguedad: 'Retrabajo' };
    var porCliente = {};
    estado.filtrados.forEach(function (r) { (porCliente[r.cliente] = porCliente[r.cliente] || []).push(r); });
    var filas = Object.keys(porCliente).map(function (c) {
      var res = I.calcular(porCliente[c], estado.config);
      return { cliente: c, n: res.base.consideradas, kpis: res.kpis };
    }).filter(function (f) { return f.n > 0; }).sort(function (a, b) { return b.n - a.n; });
    var celda = function (k) {
      var cls = k.estado === 'EN META' ? 'en-meta' : k.estado === 'FUERA DE META' ? 'fuera' : 'nd';
      return '<td class="mc ' + cls + '" title="' + esc(k.nombre + ': ' + textoDetalle(k)) + '">' +
        (k.resultado === null ? '<span class="mc-nd">sin datos</span>' : ICONOS[k.estado] + pct(k.resultado)) + '</td>';
    };
    var total = plantilla();
    $('#vista-cliente').innerHTML = '<div class="tabla-scroll"><table class="tabla matriz"><thead><tr><th>Cliente</th><th class="num">HU</th>' +
      ids.map(function (id) { var k = total.find(function (x) { return x.id === id; }); return '<th>' + esc(cortos[id]) + '<small>' + (k.tipoMeta === '<=' ? '≤ ' : '≥ ') + pct(k.meta, 0) + '</small></th>'; }).join('') +
      '</tr></thead><tbody>' + filas.map(function (f) {
        return '<tr><td>' + esc(f.cliente) + '</td><td class="num">' + f.n + '</td>' +
          ids.map(function (id) { return celda(f.kpis.find(function (x) { return x.id === id; })); }).join('') + '</tr>';
      }).join('') +
      '<tr class="total"><td>Total</td><td class="num">' + estado.resultado.base.consideradas + '</td>' +
      ids.map(function (id) { return celda(total.find(function (x) { return x.id === id; })); }).join('') + '</tr></tbody></table></div>' +
      '<p class="nota">Cada celda calcula el indicador solo con las HU de ese cliente. El control de horas no se desglosa porque las horas contratadas son globales.</p>';
  }

  function renderTabla() {
    var k = plantilla();
    $('#scorecard').innerHTML =
      '<div class="sc-fila sc-cab" role="row"><span role="columnheader">Indicador</span><span role="columnheader" class="sc-valor-cab" style="text-align:right">Resultado</span>' +
      '<span role="columnheader">Frente a la meta</span><span role="columnheader" class="sc-detalle">Base</span><span role="columnheader">Estado</span></div>' +
      k.map(function (x) {
        return '<div class="sc-fila" role="row">' +
          '<div class="sc-nombre" role="cell"><strong>' + esc(x.nombre) + '</strong><small>' + esc(x.definicion) + '</small></div>' +
          '<div class="sc-valor' + (x.resultado === null ? ' nd' : '') + '" role="cell">' + pct(x.resultado) + '</div>' +
          '<div class="sc-bala" role="cell">' + bala(x) + '</div>' +
          '<div class="sc-detalle" role="cell">' + esc(textoDetalle(x)) + '</div>' +
          '<div class="sc-estado" role="cell">' + pill(x.estado) + '</div>' +
          '</div>';
      }).join('');
  }

  function renderPipeline() {
    var r = estado.resultado, c = colores(), cf = colorFase(c);
    var total = r.base.consideradas;
    $('#flujo-resumen').textContent = fmtNum.format(r.porFlujo[0].n) + ' abiertas y ' + fmtNum.format(r.porFlujo[1].n) +
      ' cerradas de ' + fmtNum.format(total) + ' HU consideradas';
    $('#pipeline').innerHTML =
      '<div class="pp-barra" role="img" aria-label="Distribución de HU por fase">' + r.porFase.map(function (f) {
        return f.n ? '<span class="pp-seg" style="flex:' + f.n + ';background:' + cf[f.fase] + '" title="' + esc(f.fase + ': ' + f.n + ' HU') + '"></span>' : '';
      }).join('') + '</div>' +
      '<div class="pp-cols">' + r.porFase.map(function (f) {
        var estados = r.porEstado.filter(function (e) { return e.fase === f.fase; });
        return '<div class="pp-col">' +
          '<div class="pp-fase"><i style="background:' + cf[f.fase] + '"></i>' + esc(f.fase) + '</div>' +
          '<div class="pp-num">' + fmtNum.format(f.n) + '<small>' + pct(f.pct) + '</small></div>' +
          '<div class="pp-flujo">' + (f.fase === 'Cerrado' ? 'En flujo cerrado' : 'En flujo abierto') + '</div>' +
          '<ul class="pp-estados">' + estados.map(function (e) {
            return '<li class="' + (e.n ? '' : 'cero') + '"><span>' + esc(e.estado) + '</span><b>' + e.n + '</b></li>';
          }).join('') + '</ul></div>';
      }).join('') + '</div>';
  }

  // ---------- gráficos ----------
  function colores() {
    return {
      s: [1, 2, 3, 4, 5, 6, 7, 8].map(function (i) { return cssVar('--series-' + i); }),
      superficie: cssVar('--surface-1'),
      texto: cssVar('--text-secondary'),
      textoFuerte: cssVar('--text-primary'),
      grid: cssVar('--grid'),
      muted: cssVar('--text-muted')
    };
  }
  function colorFase(c) {
    return { 'Análisis': cssVar('--fase-1'), 'Desarrollo': cssVar('--fase-2'), 'Pruebas': cssVar('--fase-3'), 'Cerrado': cssVar('--fase-4') };
  }
  function colorTipo(c) {
    return { 'Desarrollo Contratado': c.s[0], 'Mejora Funcional': c.s[1], 'Deuda Tecnica': c.s[2], 'Sin clasificar': c.muted };
  }

  function configurarChart() {
    var c = colores();
    Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
    Chart.defaults.font.size = 12;
    Chart.defaults.color = c.texto;
    Chart.defaults.borderColor = c.grid;
    Chart.defaults.plugins.legend.labels.usePointStyle = true;
    Chart.defaults.plugins.legend.labels.pointStyle = 'rectRounded';
    Chart.defaults.plugins.legend.labels.boxWidth = 10;
    Chart.defaults.plugins.legend.labels.boxHeight = 10;
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.cornerRadius = 6;
    Chart.defaults.maintainAspectRatio = false;
    Chart.defaults.animation.duration = 300;
  }

  function grafico(id, cfg) {
    if (typeof Chart === 'undefined') return;
    if (estado.graficos[id]) estado.graficos[id].destroy();
    estado.graficos[id] = new Chart(document.getElementById(id), cfg);
  }

  function ejeValores(c, horizontal) {
    var eje = { beginAtZero: true, ticks: { precision: 0 }, grid: { color: c.grid }, border: { display: false } };
    var cat = { grid: { display: false }, border: { color: c.grid } };
    if (horizontal) cat.ticks = { callback: etiquetaCorta };
    return horizontal ? { x: eje, y: cat } : { x: cat, y: eje };
  }

  // Etiqueta directa al final de cada barra (selectiva: solo el total)
  var etiquetasTotales = {
    id: 'etiquetasTotales',
    afterDatasetsDraw: function (chart, args, opts) {
      if (!opts || !opts.activo) return;
      var ctx = chart.ctx, c = colores();
      var horizontal = chart.options.indexAxis === 'y';
      var metas = chart.data.datasets.map(function (d, i) { return chart.getDatasetMeta(i); }).filter(function (m) { return !m.hidden; });
      if (!metas.length) return;
      var n = chart.data.labels.length;
      ctx.save();
      ctx.font = '600 11.5px ' + Chart.defaults.font.family;
      ctx.fillStyle = c.textoFuerte;
      for (var i = 0; i < n; i++) {
        var total = 0, extremo = null;
        metas.forEach(function (m) {
          var v = chart.data.datasets[m.index].data[i] || 0;
          total += v;
          var el = m.data[i];
          if (!el) return;
          var pos = horizontal ? el.x : el.y;
          if (extremo === null || (horizontal ? pos > extremo.x : pos < extremo.y)) extremo = el;
        });
        if (!total || !extremo) continue;
        var txt = opts.formato ? opts.formato(total, i) : String(total);
        if (horizontal) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(txt, extremo.x + 6, extremo.y); }
        else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(txt, extremo.x, extremo.y - 4); }
      }
      ctx.restore();
    }
  };

  // En pantallas angostas abrevia las etiquetas largas del eje de categorías
  function etiquetaCorta(valor) {
    var t = String(this.getLabelForValue(valor));
    var max = window.innerWidth < 600 ? 16 : 34;
    return t.length > max ? t.slice(0, max - 1) + '…' : t;
  }

  function barra(c, color, extra) {
    return Object.assign({
      backgroundColor: color, borderRadius: 4, borderSkipped: 'start',
      borderColor: c.superficie, borderWidth: 0, maxBarThickness: 28, categoryPercentage: 0.7, barPercentage: 0.9
    }, extra || {});
  }

  function renderDistribucion() {
    renderPipeline();
    if (typeof Chart === 'undefined') return;
    configurarChart();
    var r = estado.resultado, c = colores(), cf = colorFase(c);

    // Cerradas con el tono de la fase Cerrado; abiertas planificadas con un tono de fase abierta
    grafico('g-mensual', {
      type: 'bar',
      data: {
        labels: r.mensual.map(function (m) { return r.anio ? MESES[+m.mes.split('-')[1] - 1] : nombreMes(m.mes); }),
        datasets: [
          barra(c, cf['Cerrado'], { label: 'HU cerradas', data: r.mensual.map(function (m) { return m.cerradas; }) }),
          barra(c, cf['Análisis'], { label: 'HU abiertas con fin planificado', data: r.mensual.map(function (m) { return m.abiertas; }) })
        ]
      },
      options: {
        scales: Object.assign(ejeValores(c), { x: { grid: { display: false }, border: { color: c.grid }, ticks: { maxRotation: 0, autoSkipPadding: 8 } } }),
        interaction: { mode: 'index', intersect: false },
        plugins: { legend: { position: 'top', align: 'start' } }
      }
    });
    $('#nota-mensual').textContent = r.sinFecha.cerradas + ' HU cerradas y ' + r.sinFecha.abiertas +
      ' abiertas no tienen fecha de finalización' + (r.anio ? ' en ' + r.anio : '') + ', por eso no aparecen en el gráfico.';
    var anios = r.anios.slice();
    if (r.anio && anios.indexOf(r.anio) === -1) anios.push(r.anio);
    anios.sort();
    $('#sel-anio').innerHTML = anios.map(function (a) {
      return '<option value="' + a + '"' + (a === r.anio ? ' selected' : '') + '>' + a + '</option>';
    }).join('') + '<option value="todos"' + (r.anio ? '' : ' selected') + '>Todos</option>';

    var ct = colorTipo(c);
    grafico('g-tipo', {
      type: 'bar',
      data: {
        labels: r.porTipo.map(function (t) { return t.nombre; }),
        datasets: [barra(c, r.porTipo.map(function (t) { return ct[t.nombre]; }), { label: 'HU', data: r.porTipo.map(function (t) { return t.n; }) })]
      },
      options: {
        indexAxis: 'y', scales: ejeValores(c, true), layout: { padding: { right: 78 } },
        plugins: {
          legend: { display: false },
          etiquetasTotales: { activo: true, formato: function (v, i) { return v + ' (' + pct(r.porTipo[i].pct) + ')'; } },
          tooltip: { callbacks: { label: function (it) { return ' ' + it.raw + ' HU (' + pct(r.porTipo[it.dataIndex].pct) + ')'; } } }
        }
      },
      plugins: [etiquetasTotales]
    });
  }

  function tablaCruzada(id, filas, totalAbiertas, etiqueta) {
    var tipos = I.TIPOS_SOLICITUD;
    var tot = { total: 0 }; tipos.forEach(function (t) { tot[t] = 0; });
    filas.forEach(function (f) { tipos.forEach(function (t) { tot[t] += f[t]; }); tot.total += f.total; });
    $(id).innerHTML = '<thead><tr><th>' + esc(etiqueta) + '</th>' + tipos.map(function (t) { return '<th class="num">' + esc(t) + '</th>'; }).join('') +
      '<th class="num">Total</th><th class="num">% del total</th></tr></thead><tbody>' +
      (filas.length ? filas.map(function (f) {
        return '<tr><td>' + esc(f.nombre || '(vacío)') + '</td>' + tipos.map(function (t) { return '<td class="num">' + f[t] + '</td>'; }).join('') +
          '<td class="num"><strong>' + f.total + '</strong></td><td class="num">' + pct(f.pct) + '</td></tr>';
      }).join('') : '<tr><td colspan="7">Sin HU abiertas para los filtros seleccionados.</td></tr>') +
      '<tr class="total"><td>Total</td>' + tipos.map(function (t) { return '<td class="num">' + tot[t] + '</td>'; }).join('') +
      '<td class="num">' + tot.total + '</td><td class="num">' + pct(totalAbiertas ? tot.total / totalAbiertas : 0) + '</td></tr></tbody>';
  }

  function graficoApilado(id, filas) {
    var c = colores(), ct = colorTipo(c);
    var el = document.getElementById(id).parentNode;
    el.style.height = Math.max(200, filas.length * 34 + 70) + 'px';
    grafico(id, {
      type: 'bar',
      data: {
        labels: filas.map(function (f) { return f.nombre || '(vacío)'; }),
        datasets: I.TIPOS_SOLICITUD.map(function (t, i) {
          return barra(c, ct[t], {
            label: t, data: filas.map(function (f) { return f[t] || null; }), stack: 'a',
            borderWidth: { right: 2, left: 0, top: 0, bottom: 0 }, borderSkipped: false, borderRadius: 0
          });
        })
      },
      options: {
        indexAxis: 'y',
        scales: { x: { stacked: true, beginAtZero: true, ticks: { precision: 0 }, grid: { color: c.grid }, border: { display: false } }, y: { stacked: true, grid: { display: false }, ticks: { callback: etiquetaCorta } } },
        layout: { padding: { right: 36 } },
        interaction: { mode: 'index', intersect: false, axis: 'y' },
        plugins: {
          legend: { position: 'top', align: 'start' },
          etiquetasTotales: { activo: true },
          tooltip: { filter: function (it) { return it.raw; }, callbacks: { footer: function (its) { return 'Total: ' + filas[its[0].dataIndex].total + ' HU'; } } }
        }
      },
      plugins: [etiquetasTotales]
    });
  }

  function renderSeguimiento() {
    var s = estado.resultado.seguimiento;
    [['cliente', 'porCliente', 'Cliente'], ['sprint', 'porSprint', 'Sprint'], ['modulo', 'porModulo', 'Módulo'], ['asignado', 'porAsignado', 'Asignado a']]
      .forEach(function (p) {
        tablaCruzada('#t-seg-' + p[0], s[p[1]], s.total, p[2]);
        if (typeof Chart !== 'undefined') graficoApilado('g-seg-' + p[0], s[p[1]]);
      });
  }

  // ---------- detalle ----------
  var COLS_DETALLE = [
    ['id', 'ID'], ['asunto', 'Asunto'], ['cliente', 'Cliente'], ['estado', 'Estado'], ['fase', 'Fase'],
    ['enFlujo', 'En flujo'], ['tipoSolicitud', 'Tipo de Solicitud'], ['sprint', 'Sprint'], ['modulo', 'Módulo'],
    ['responsable', 'Responsable'], ['asignado', 'Asignado a'], ['trabajoTotal', 'Trabajo total (h)'],
    ['fechaFin', 'Fecha finalización'], ['precision', 'Precisión HU']
  ];
  function colsDetalle() {
    return COLS_DETALLE.filter(function (c) { return c[0] !== 'responsable' || (estado.columnas && estado.columnas.responsable); });
  }

  function filasDetalle() {
    var q = $('#buscar').value.trim().toLowerCase();
    var solo = $('#solo-consideradas').checked;
    var flujo = $('#filtro-flujo').value, fase = $('#filtro-fase').value;
    var o = estado.orden;
    return estado.filtrados.filter(function (r) {
      if (solo && !r.considerada) return false;
      if (flujo && r.enFlujo !== flujo) return false;
      if (fase && r.fase !== fase) return false;
      if (q) {
        var txt = [r.id, r.asunto, r.cliente, r.estado, r.responsable, r.asignado, r.sprint, r.modulo, r.precision].join(' ').toLowerCase();
        if (txt.indexOf(q) === -1) return false;
      }
      return true;
    }).sort(function (a, b) {
      var va = a[o.campo], vb = b[o.campo], res;
      if (va instanceof Date || vb instanceof Date) res = (va ? va.getTime() : Infinity) - (vb ? vb.getTime() : Infinity);
      else if (o.campo === 'id' || o.campo === 'trabajoTotal') res = Number(va) - Number(vb) || String(va).localeCompare(String(vb));
      else res = String(va || '').localeCompare(String(vb || ''), 'es');
      return o.asc ? res : -res;
    });
  }

  function renderDetalle() {
    var filas = filasDetalle();
    var o = estado.orden;
    $('#contador-detalle').textContent = fmtNum.format(filas.length) + ' HU';
    $('#t-detalle').innerHTML = '<thead><tr>' + colsDetalle().map(function (c) {
      var flecha = o.campo === c[0] ? (o.asc ? ' ▲' : ' ▼') : '';
      return '<th><button class="btn-texto" data-orden="' + c[0] + '" style="all:unset;cursor:pointer">' + esc(c[1]) + flecha + '</button></th>';
    }).join('') + '</tr></thead><tbody>' + filas.map(function (r) {
      return '<tr>' + colsDetalle().map(function (c) {
        var v = r[c[0]];
        if (c[0] === 'fechaFin') v = fecha(v);
        var cls = c[0] === 'asunto' ? ' class="asunto"' : c[0] === 'id' || c[0] === 'trabajoTotal' ? ' class="num"' : '';
        if (c[0] === 'enFlujo' || c[0] === 'fase') return '<td><span class="chip">' + esc(v) + '</span></td>';
        return '<td' + cls + '>' + esc(v) + '</td>';
      }).join('') + '</tr>';
    }).join('') + '</tbody>';
  }

  function exportarCsv() {
    var filas = filasDetalle();
    var csv = [colsDetalle().map(function (c) { return c[1]; })].concat(filas.map(function (r) {
      return colsDetalle().map(function (c) { return c[0] === 'fechaFin' ? fecha(r[c[0]]) : r[c[0]]; });
    })).map(function (f) {
      return f.map(function (v) { v = String(v === null || v === undefined ? '' : v); return /[";\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(';');
    }).join('\r\n');
    var blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'detalle-hu.csv';
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  // ---------- ficha ----------
  function renderFicha() {
    var k = estado.resultado.kpis;
    $('#t-ficha').innerHTML = '<thead><tr><th>N°</th><th>Indicador (KPI)</th><th>Definición</th><th>Fórmula de cálculo</th><th>Frecuencia</th><th>Meta</th><th>Responsable</th><th>Cálculo con la base</th><th class="num">Resultado actual</th><th>Estado</th></tr></thead><tbody>' +
      k.map(function (x, i) {
        var meta = x.meta === null || x.meta === undefined || x.meta === '' ? 'Sin meta' : (x.tipoMeta === '<=' ? '≤ ' : '≥ ') + pct(x.meta, 0);
        return '<tr><td>' + (i + 1) + '</td><td><strong>' + esc(x.nombre) + '</strong></td><td>' + esc(x.definicion) + '</td><td>' + esc(x.formula) +
          '</td><td>' + esc(x.frecuencia) + '</td><td>' + meta + '</td><td>' + esc(x.responsable) + '</td><td>' + esc(x.criterio) +
          '</td><td class="num">' + pct(x.resultado) + '</td><td>' + pill(x.estado) + '</td></tr>';
      }).join('') + '</tbody>';
  }

  // ---------- configuración ----------
  function renderConfig() {
    if (renderConfig.hecho) return;
    renderConfig.hecho = true;
    var metas = estado.config.metas, par = Object.assign({}, I.PARAMETROS, estado.config.parametros);
    $('#cfg-metas').innerHTML = '<div class="fila-cfg"><span class="cab">Indicador</span><span class="cab">Tipo</span><span class="cab">Meta (%)</span></div>' +
      I.KPIS.map(function (k) {
        var m = metas[k.id] || {};
        var meta = m.meta !== undefined ? m.meta : k.meta;
        var tipo = m.tipoMeta !== undefined ? m.tipoMeta : k.tipoMeta;
        return '<div class="fila-cfg"><span>' + esc(k.nombre) + '</span>' +
          '<select data-meta-tipo="' + k.id + '" aria-label="Tipo de meta"><option value="">—</option><option' + (tipo === '<=' ? ' selected' : '') + '>&lt;=</option><option' + (tipo === '>=' ? ' selected' : '') + '>&gt;=</option></select>' +
          '<input type="number" min="0" max="1000" step="0.1" data-meta="' + k.id + '" value="' + (meta === null || meta === '' ? '' : +(meta * 100).toFixed(2)) + '" aria-label="Meta"></div>';
      }).join('');
    var campos = [
      ['horasContratadas', 'Horas contratadas (KPI Control de consumo de horas)', 'number'],
      ['tipoExcluidoBacklog', 'Tipo de Solicitud excluido del % Cumplimiento del Backlog', 'text'],
      ['marcaAlcance', "Marca en Precisión HU para '% Cambios de alcance'", 'text'],
      ['marcaRetrabajo', "Marca en Precisión HU para '% Retrabajo por ambigüedad'", 'text']
    ];
    $('#cfg-parametros').innerHTML = campos.map(function (c) {
      var v = par[c[0]];
      return '<div class="fila-cfg param"><span>' + esc(c[1]) + '</span>' +
        '<input type="' + c[2] + '"' + (c[2] === 'number' ? ' min="0" step="1"' : '') + ' data-param="' + c[0] + '" value="' + esc(v === null || v === undefined ? '' : v) + '" aria-label="' + esc(c[1]) + '"></div>';
    }).join('');
    $('#cfg-estados').value = estado.config.estados.map(function (e) { return e[0] + ' => ' + e[1]; }).join('\n');
    $('#cfg-palabras').value = estado.config.palabras.map(function (p) { return p[0] + ' => ' + p[1]; }).join('\n');
  }

  function leerMetas() {
    var metas = {};
    I.KPIS.forEach(function (k) {
      var v = $('[data-meta="' + k.id + '"]').value;
      var t = $('[data-meta-tipo="' + k.id + '"]').value;
      metas[k.id] = { meta: v === '' ? null : Number(v) / 100, tipoMeta: t };
    });
    estado.config.metas = metas;
    guardar('metas', metas);
    actualizar();
  }

  function guardarParametros(cambios) {
    estado.config.parametros = Object.assign({}, estado.config.parametros, cambios);
    guardar('parametros', estado.config.parametros);
    actualizar();
  }

  function leerParametros() {
    var cambios = {};
    $$('[data-param]').forEach(function (inp) {
      var k = inp.getAttribute('data-param');
      cambios[k] = inp.type === 'number' ? (inp.value === '' ? null : Number(inp.value)) : inp.value.trim();
    });
    guardarParametros(cambios);
  }

  function leerReglas(id) {
    return $(id).value.split(/\r?\n/).map(function (l) {
      var p = l.split('=>');
      return p.length === 2 ? [p[0].trim(), p[1].trim()] : null;
    }).filter(function (p) { return p && p[0] && p[1]; });
  }

  function guardarEstados() {
    var reglas = leerReglas('#cfg-estados');
    var validas = I.FASES.concat([I.EXCLUIDO]);
    var malas = reglas.filter(function (r) { return validas.indexOf(r[1]) === -1; });
    if (!reglas.length) { toast('No se encontraron reglas válidas (formato: Estado => Fase).'); return; }
    if (malas.length) { toast('Fase no válida en: ' + malas.map(function (r) { return r[0]; }).join(', ') + '. Use ' + validas.join(', ') + '.'); return; }
    estado.config.estados = reglas;
    guardar('estados', reglas);
    reparsear();
    toast(reglas.length + ' estados aplicados');
  }

  function guardarPalabras() {
    var reglas = leerReglas('#cfg-palabras');
    if (!reglas.length) { toast('No se encontraron reglas válidas (formato: palabra => Módulo).'); return; }
    estado.config.palabras = reglas;
    guardar('palabras', reglas);
    reparsear();
    toast(reglas.length + ' reglas de módulo aplicadas');
  }

  // ---------- tema ----------
  function aplicarTema(t) {
    if (t) document.documentElement.setAttribute('data-theme', t);
    else document.documentElement.removeAttribute('data-theme');
    if (estado.resultado) { renderDistribucion(); renderSeguimiento(); }
  }
  function temaActual() {
    var t = document.documentElement.getAttribute('data-theme');
    if (t) return t;
    return window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  // ---------- eventos ----------
  function enlazar() {
    $('#archivo').addEventListener('change', function (e) { cargarArchivo(e.target.files[0]); e.target.value = ''; });
    $('#btn-ejemplo').addEventListener('click', cargarEjemplo);
    $('[data-accion="ejemplo"]').addEventListener('click', cargarEjemplo);

    ['dragenter', 'dragover'].forEach(function (ev) {
      document.addEventListener(ev, function (e) { e.preventDefault(); $('#zona-carga').classList.add('arrastrando'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      document.addEventListener(ev, function (e) { e.preventDefault(); $('#zona-carga').classList.remove('arrastrando'); });
    });
    document.addEventListener('drop', function (e) { if (e.dataTransfer.files[0]) cargarArchivo(e.dataTransfer.files[0]); });

    $$('select[data-filtro]').forEach(function (sel) {
      sel.addEventListener('change', function () {
        var campo = sel.getAttribute('data-filtro');
        if (sel.value) estado.filtros[campo] = sel.value; else delete estado.filtros[campo];
        sel.classList.toggle('activo', !!sel.value);
        actualizar();
      });
    });
    $$('[data-vista]').forEach(function (b) {
      b.addEventListener('click', function () {
        estado.vistaKpi = b.getAttribute('data-vista');
        guardar('vistaKpi', estado.vistaKpi);
        renderKpis();
      });
    });
    $('#btn-limpiar').addEventListener('click', function () { estado.filtros = {}; construirFiltros(); actualizar(); });

    $$('.tabs [data-tab]').forEach(function (b) {
      b.addEventListener('click', function () {
        var tab = b.getAttribute('data-tab');
        $$('.tabs [data-tab]').forEach(function (x) { x.setAttribute('aria-selected', String(x === b)); });
        $$('[data-panel]').forEach(function (p) { p.hidden = p.getAttribute('data-panel') !== tab; });
        guardar('tab', tab);
        Object.keys(estado.graficos).forEach(function (k) { estado.graficos[k].resize(); });
      });
    });

    ['#buscar', '#solo-consideradas', '#filtro-flujo', '#filtro-fase'].forEach(function (s) {
      $(s).addEventListener('input', renderDetalle);
    });
    $('#t-detalle').addEventListener('click', function (e) {
      var b = e.target.closest('[data-orden]'); if (!b) return;
      var campo = b.getAttribute('data-orden');
      estado.orden = { campo: campo, asc: estado.orden.campo === campo ? !estado.orden.asc : true };
      renderDetalle();
    });
    $('#btn-exportar').addEventListener('click', exportarCsv);

    $('#cfg-metas').addEventListener('change', leerMetas);
    $('#cfg-parametros').addEventListener('change', leerParametros);
    $('#btn-guardar-estados').addEventListener('click', guardarEstados);
    $('#sel-anio').addEventListener('change', function (e) { guardarParametros({ anio: e.target.value === 'todos' ? 'todos' : Number(e.target.value) }); });
    $('#btn-guardar-palabras').addEventListener('click', guardarPalabras);
    $('#btn-restaurar').addEventListener('click', function () {
      estado.config = { metas: {}, parametros: {}, estados: I.ESTADOS.slice(), palabras: I.PALABRAS_CLAVE.slice() };
      ['metas', 'parametros', 'estados', 'palabras', 'manuales'].forEach(function (k) { try { localStorage.removeItem(CLAVE + k); } catch (e) { /* */ } });
      renderConfig.hecho = false;
      reparsear();
      toast('Configuración restaurada');
    });

    $('#btn-tema').addEventListener('click', function () {
      var t = temaActual() === 'dark' ? 'light' : 'dark';
      guardar('tema', t);
      aplicarTema(t);
    });
    if (window.matchMedia) {
      matchMedia('(prefers-color-scheme: dark)').addEventListener('change', function () {
        if (!document.documentElement.getAttribute('data-theme')) aplicarTema(null);
      });
    }
  }

  function iniciar() {
    if (document.fonts && document.fonts.ready) {
      document.fonts.ready.then(function () { if (estado.resultado) { renderDistribucion(); renderSeguimiento(); } });
    }
    var tema = leer('tema', null);
    if (tema) document.documentElement.setAttribute('data-theme', tema);
    enlazar();
    var tab = leer('tab', null);
    if (tab && $('.tabs [data-tab="' + tab + '"]')) $('.tabs [data-tab="' + tab + '"]').click();
    var previo = leer('datos', null);
    if (previo && previo.matriz) {
      try { usarMatriz(previo.matriz, previo.fuente, new Date(previo.cargado)); } catch (e) { /* datos guardados inválidos */ }
    }
  }

  iniciar();
})();
