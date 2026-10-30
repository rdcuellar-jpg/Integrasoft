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
    avGrupo: leer('avGrupo', 'porCliente'),
    informes: [],
    informeId: null,
    seleccion: {},
    cache: {},
    versionConfig: 0,
    comp: { a: null, b: null, tipo: 'todos', kpi: leer('compKpi', 'backlog') },
    orden: { campo: 'id', asc: true },
    config: {
      metas: leer('metas', {}),
      parametros: leer('parametros', {}),
      estados: leer('estados', null) || I.ESTADOS.slice(),
      equivalencias: leer('equivalencias', null) || I.EQUIVALENCIAS.slice(),
      avance: leer('avance', null) || JSON.parse(JSON.stringify(I.AVANCE)),
      dyd: leer('dyd', null) || JSON.parse(JSON.stringify(I.AVANCE_DYD))
    },
    sprint: leer('sprint', null)
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
        var hoja = nombre, cuando = file.lastModified ? new Date(file.lastModified) : new Date();
        pedirTipo(file.name).then(function (tipo) {
          if (tipo) registrarInforme(matriz, file.name, hoja, cuando, tipo);
        });
      } catch (err) {
        toast('Error al leer el archivo: ' + err.message);
      }
    };
    lector.readAsArrayBuffer(file);
  }

  /** Pregunta si el Excel es de indicadores de gestión o de un sprint (null si se cancela). */
  function pedirTipo(archivo) {
    var dlg = $('#dlg-tipo');
    if (!dlg || typeof dlg.showModal !== 'function') return Promise.resolve('gestion');
    $('#dlg-tipo-archivo').textContent = archivo;
    return new Promise(function (resolver) {
      dlg.addEventListener('close', function cerrar() {
        dlg.removeEventListener('close', cerrar);
        var marcado = $('input[name="tipo-carga"]:checked', dlg);
        resolver(dlg.returnValue === 'ok' ? (marcado ? marcado.value : 'gestion') : null);
      });
      dlg.returnValue = '';
      dlg.showModal();
    });
  }

  function cargarEjemplo() {
    fetch('data/hoja1-ejemplo.json')
      .then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); })
      .then(function (d) { registrarInforme(d.matriz, 'Datos de ejemplo', 'Hoja1', new Date(), 'gestion'); })
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
    if (!estado.historicoPersistente) guardar('datos', { matriz: matriz, fuente: fuente, cargado: cuando.toISOString() });
    $('#zona-carga').hidden = true;
    $('#app').hidden = false;
    construirFiltros();
    actualizar();
  }

  function opcionesParseo() {
    return { estados: estado.config.estados, equivalencias: estado.config.equivalencias };
  }

  function reparsear() {
    estado.cache = {};
    estado.versionConfig++;
    if (!estado.matriz) return;
    estado.registros = I.parsearHoja(estado.matriz, opcionesParseo()).registros;
    construirFiltros();
    actualizar();
  }

  // ---------- histórico de informes ----------
  function hoyISO(d) {
    d = d || new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function fechaCorte(iso) {
    if (!iso) return 'sin fecha';
    var p = iso.split('-');
    return (+p[2]) + ' ' + MESES[+p[1] - 1] + ' ' + p[0];
  }
  // Fecha AAAA-MM-DD en el nombre del archivo (las descargas de InteTrack la incluyen)
  function fechaEnNombre(nombre) {
    var m = String(nombre || '').match(/(20\d{2})-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])/);
    return m ? m[1] + '-' + m[2] + '-' + m[3] : null;
  }
  function etiquetaInforme(inf) {
    return fechaCorte(inf.corte) + ', ' + inf.nombre + (inf.tipo === 'sprint' ? ' · Sprint' : '');
  }
  var TIPOS_CARGA = { gestion: 'Indicadores de gestión', sprint: 'Sprint' };
  function informeActual() {
    return estado.informes.find(function (i) { return i.id === estado.informeId; }) || null;
  }

  function refrescarInformes() {
    return Historico.listar().then(function (lista) {
      estado.informes = lista;
      return lista;
    });
  }

  /** Guarda el Excel cargado como informe (si no estaba ya) y lo muestra. */
  function registrarInforme(matriz, archivo, hoja, fechaArchivo, tipo) {
    var h = Historico.huella(matriz);
    tipo = tipo === 'sprint' ? 'sprint' : 'gestion';
    refrescarInformes().then(function (lista) {
      var existente = lista.find(function (i) { return i.huella === h; });
      if (existente) {
        var aviso = 'Este archivo ya estaba guardado como el informe del ' + fechaCorte(existente.corte) + '.';
        if ((existente.tipo || 'gestion') === tipo) { mostrarInforme(existente); toast(aviso); return null; }
        return Historico.actualizar(existente.id, { tipo: tipo }).then(refrescarInformes).then(function (l) {
          mostrarInforme(l.find(function (i) { return i.id === existente.id; }) || existente);
          toast(aviso + ' Ahora queda como ' + TIPOS_CARGA[tipo] + '.');
        });
      }
      var p = I.parsearHoja(matriz, opcionesParseo());
      var inf = {
        id: 'inf-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6),
        nombre: archivo.replace(/\.(xlsx|xls|csv)$/i, ''),
        archivo: archivo,
        hoja: hoja,
        cargado: new Date().toISOString(),
        corte: fechaEnNombre(archivo) || hoyISO(fechaArchivo),
        huella: h,
        tipo: tipo,
        hu: p.registros.length,
        matriz: matriz
      };
      return Historico.guardar(inf).then(function () { return refrescarInformes(); }).then(function () {
        mostrarInforme(inf);
        toast(estado.informes.length > 1
          ? 'Informe guardado con fecha de corte ' + fechaCorte(inf.corte) + '. Ya puedes compararlo con los anteriores en Comparativo de informes.'
          : 'Informe guardado con fecha de corte ' + fechaCorte(inf.corte) + '. Carga el Excel de otra fecha para poder comparar.');
      });
    }).catch(function (err) { toast('No se pudo guardar el informe: ' + err.message); });
  }

  function mostrarInforme(inf) {
    estado.informeId = inf.id;
    guardar('informeActual', inf.id);
    usarMatriz(inf.matriz, inf.nombre + ' — ' + (inf.hoja || 'Hoja1'), new Date(inf.cargado));
  }

  function renderSelectorInforme() {
    var sel = $('#sel-informe');
    sel.innerHTML = estado.informes.map(function (i) {
      return '<option value="' + esc(i.id) + '"' + (i.id === estado.informeId ? ' selected' : '') + '>' + esc(etiquetaInforme(i)) + '</option>';
    }).join('');
    $('#n-informes').textContent = '(' + estado.informes.length + ' guardado' + (estado.informes.length === 1 ? '' : 's') + ')';
    $('#aviso-corte').hidden = !(estado.informes.length && estado.informeId && estado.informes[0].id !== estado.informeId);
  }

  // Registros de un informe del histórico, con la configuración y filtros actuales
  function registrosDeInforme(inf) {
    var c = estado.cache[inf.id];
    if (!c || c.version !== estado.versionConfig) {
      c = estado.cache[inf.id] = { version: estado.versionConfig, registros: I.parsearHoja(inf.matriz, opcionesParseo()).registros };
    }
    return filtrar(c.registros);
  }
  function resultadoDeInforme(inf) {
    return I.calcular(registrosDeInforme(inf), estado.config);
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
    return filtrar(estado.registros);
  }

  function filtrar(lista) {
    return lista.filter(function (r) {
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
    estado.avance = I.calcularAvance(filtrados, estado.config.avance);
    estado.registros.forEach(function (h) { h.avance = null; });
    estado.avance.lista.forEach(function (x) { x.hu.avance = x.avance; });
    estado.sprintRes = I.calcularSprint(filtrados, estado.config.dyd);
    renderFuente();
    renderResumen();
    renderKpis();
    renderDistribucion();
    renderSeguimiento();
    renderDetalle();
    renderFicha();
    renderConfig();
    renderSelectorInforme();
    renderAvance();
    renderSprint();
    renderComparativo();
  }

  // ---------- avance por estado ----------
  function claseAccion(orden) {
    return orden === 99 ? 'acc-ok' : orden >= 0 ? 'acc-' + (orden % 5 + 1) : 'acc-0';
  }
  function barraAvance(v) {
    var p = Math.max(0, Math.min(1, v || 0)) * 100;
    return '<span class="av-barra" role="img" aria-label="Avance ' + pct(v, 0) + '"><span style="width:' + p + '%"></span></span>';
  }

  function renderAvance() {
    var a = estado.avance;
    if (!a || !$('#av-titular')) return;
    $('#av-titular').textContent = a.n
      ? 'Las HU consideradas llevan ' + pct(a.promedio) + ' de avance promedio'
      : 'No hay HU consideradas para calcular el avance';
    var accionActual = estado.config.avance.incluirActual;
    $('#av-detalle').textContent = a.n
      ? fmtNum.format(a.ejecutadas) + ' HU están ejecutadas (Entrega a Operaciones, Entrega cliente o cerradas), ' +
        fmtNum.format(a.enCurso) + ' van en curso y ' + fmtNum.format(a.sinAvance) + ' no tienen avance. ' +
        (accionActual ? 'Cada HU suma la acción de su estado y las anteriores.' : 'Cada HU suma solo las acciones que ya terminó, antes de su estado actual.')
      : '';
    $('#av-cifras').innerHTML = [
      ['Avance promedio', pct(a.promedio)], ['HU ejecutadas', fmtNum.format(a.ejecutadas)],
      ['HU en curso', fmtNum.format(a.enCurso)], ['HU sin avance', fmtNum.format(a.sinAvance)]
    ].map(function (c) { return '<div><dt>' + esc(c[0]) + '</dt><dd>' + esc(c[1]) + '</dd></div>'; }).join('');

    // Tabla de esfuerzo (como la tabla de referencia del área)
    var cuenta = {};
    a.porEstado.forEach(function (e) { cuenta[e.orden] = (cuenta[e.orden] || 0) + e.n; });
    var filas = a.tabla.filas.map(function (f) {
      return '<tr><td><span class="acc ' + claseAccion(f.orden) + '">' + esc(f.accion) + '</span></td><td class="num">' + pct(f.esfuerzo, 0) +
        '</td><td class="num">' + pct(f.avance, 0) + '</td><td>' + esc(f.estados.join(', ')) + '</td><td class="num">' + fmtNum.format(cuenta[f.orden] || 0) + '</td></tr>';
    });
    filas.push('<tr class="total"><td>Total</td><td class="num">' + pct(a.tabla.total, 0) + '</td><td></td><td></td><td></td></tr>');
    filas.push('<tr><td><span class="acc acc-ok">Ejecutadas</span></td><td class="num"></td><td class="num">100%</td><td>Fase Cerrado: Entrega a Operaciones, Entrega cliente, CIERRE, Cierre…</td><td class="num">' + fmtNum.format(a.ejecutadas) + '</td></tr>');
    filas.push('<tr><td><span class="acc acc-0">Otros estados</span></td><td class="num"></td><td class="num">0%</td><td>Backlog, Aprobacion PO, Revisión (IDS)…</td><td class="num">' + fmtNum.format(cuenta[-1] || 0) + '</td></tr>');
    var aviso = Math.abs(a.tabla.total - 1) > 0.0001 ? '<p class="alerta">Los porcentajes de esfuerzo suman ' + pct(a.tabla.total, 0) + '; deberían sumar 100%. Ajústalos en Configuración.</p>' : '';
    $('#t-av-esfuerzo-tabla').innerHTML = '<thead><tr><th>Acción</th><th class="num">% del esfuerzo</th><th class="num">Avance en este estado</th><th>Estados del ERP</th><th class="num">HU</th></tr></thead><tbody>' + filas.join('') + '</tbody>';
    $$('.av-aviso').forEach(function (el) { el.remove(); });
    if (aviso) $('#t-av-esfuerzo-tabla').parentNode.insertAdjacentHTML('beforebegin', aviso.replace('class="alerta"', 'class="alerta av-aviso"'));

    // Tramos
    var colTramo = ['acc-0', 'tr-1', 'tr-2', 'acc-ok'];
    $('#av-tramos-resumen').textContent = fmtNum.format(a.n) + ' HU consideradas';
    $('#av-tramos').innerHTML = '<div class="pp-barra" role="img" aria-label="Distribución de HU por tramo de avance">' + a.tramos.map(function (t, i) {
      return t.n ? '<span class="pp-seg ' + colTramo[i] + '" style="flex:' + t.n + '" title="' + esc(t.nombre + ': ' + t.n + ' HU') + '"></span>' : '';
    }).join('') + '</div><div class="pp-cols">' + a.tramos.map(function (t, i) {
      return '<div class="pp-col"><div class="pp-fase"><i class="' + colTramo[i] + '"></i>' + esc(t.nombre) + '</div><div class="pp-num">' + fmtNum.format(t.n) + '<small>' + pct(t.pct) + '</small></div></div>';
    }).join('') + '</div>';

    // Por estado
    $('#t-av-estado-tabla').innerHTML = '<thead><tr><th>Estado</th><th>Acción</th><th>Avance</th><th class="num">HU</th><th class="num">Aporte al promedio</th></tr></thead><tbody>' +
      a.porEstado.map(function (e) {
        var orig = Object.keys(e.originales).map(function (k) { return k + ' (' + e.originales[k] + ')'; });
        return '<tr><td>' + esc(e.estado) + (orig.length ? '<small class="estado-erp">Incluye ' + esc(orig.join(', ')) + '</small>' : '') + '</td><td><span class="acc ' + claseAccion(e.orden) + '">' + esc(e.accion) +
          '</span></td><td class="av-celda">' + barraAvance(e.avance) + '<b>' + pct(e.avance, 0) + '</b></td><td class="num">' + fmtNum.format(e.n) + '</td><td class="num">' + fmtPp.format(e.aporte * 100) + ' pp</td></tr>';
      }).join('') + '<tr class="total"><td colspan="3">Avance promedio</td><td class="num">' + fmtNum.format(a.n) + '</td><td class="num">' + pct(a.promedio) + '</td></tr></tbody>';

    // Por grupo
    var g = ['porCliente', 'porSprint', 'porModulo', 'porAsignado'].indexOf(estado.avGrupo) === -1 ? 'porCliente' : estado.avGrupo;
    $$('[data-av-grupo]').forEach(function (b) { b.setAttribute('aria-checked', String(b.getAttribute('data-av-grupo') === g)); });
    var etiqueta = { porCliente: 'Cliente', porSprint: 'Sprint', porModulo: 'Módulo', porAsignado: 'Asignado a' }[g];
    $('#t-av-grupo-tabla').innerHTML = '<thead><tr><th>' + etiqueta + '</th><th class="num">HU</th><th class="num">Ejecutadas</th><th>Avance promedio</th></tr></thead><tbody>' +
      a[g].map(function (f) {
        return '<tr><td>' + esc(f.nombre) + '</td><td class="num">' + fmtNum.format(f.n) + '</td><td class="num">' + fmtNum.format(f.ejecutadas) +
          '</td><td class="av-celda">' + barraAvance(f.promedio) + '<b>' + pct(f.promedio) + '</b></td></tr>';
      }).join('') + '</tbody>';

    renderAvanceHU();
  }

  function renderAvanceHU() {
    var a = estado.avance;
    var q = ($('#av-buscar').value || '').trim().toLowerCase();
    var filas = a.lista.filter(function (x) { return x.avance < 1; }).filter(function (x) {
      return !q || [x.hu.id, x.hu.asunto, x.hu.cliente, x.hu.estado, x.hu.asignado].join(' ').toLowerCase().indexOf(q) !== -1;
    }).sort(function (x, y) { return x.avance - y.avance || Number(x.hu.id) - Number(y.hu.id); });
    $('#t-av-hu-tabla').innerHTML = '<thead><tr><th class="num">ID</th><th>Asunto</th><th>Cliente</th><th>Estado</th><th>Asignado a</th><th>Avance</th></tr></thead><tbody>' +
      (filas.length ? filas.map(function (x) {
        return '<tr><td class="num">' + esc(x.hu.id) + '</td><td class="asunto">' + esc(x.hu.asunto) + '</td><td>' + esc(x.hu.cliente) + '</td><td>' + esc(x.hu.estado) +
          '</td><td>' + esc(x.hu.asignado || 'Sin asignar') + '</td><td class="av-celda">' + barraAvance(x.avance) + '<b>' + pct(x.avance, 0) + '</b></td></tr>';
      }).join('') : '<tr><td colspan="6">' + (q ? 'Ninguna HU sin terminar coincide con la búsqueda.' : 'Todas las HU consideradas están ejecutadas.') + '</td></tr>') + '</tbody>';
  }

  // ---------- sprint ----------
  function renderSprint() {
    var sp = estado.sprintRes;
    if (!sp || !$('#sp-titular')) return;
    var cfg = estado.config.dyd;
    $('#sp-sin-rel').hidden = !!(estado.columnas && estado.columnas.relaciones);

    // Sprints disponibles: los del corte actual y los de los cortes cargados como Sprint
    var cortesSprint = estado.informes.filter(function (i) { return i.tipo === 'sprint'; }).map(function (inf) {
      return { inf: inf, corte: inf.corte, registros: registrosDeInforme(inf) };
    });
    var nombres = {};
    sp.porSprint.forEach(function (x) { nombres[x.nombre] = true; });
    cortesSprint.forEach(function (c) { I.paquetesDyD(c.registros, cfg).forEach(function (x) { if (x.sprint) nombres[x.sprint] = true; }); });
    var lista = Object.keys(nombres).sort(function (a, b) { return I.compararSprint(b, a); });
    // Por defecto: el sprint con más paquetes sin terminar; si todos terminaron, el de fecha de fin más reciente
    var porDefecto = sp.porSprint.slice().sort(function (a, b) {
      return (b.n - b.terminados) - (a.n - a.terminados) || (b.fin ? b.fin.getTime() : 0) - (a.fin ? a.fin.getTime() : 0);
    })[0];
    var actual = lista.indexOf(estado.sprint) !== -1 ? estado.sprint : porDefecto ? porDefecto.nombre : lista[0] || null;
    $('#sp-sprint').innerHTML = lista.length ? lista.map(function (n) {
      return '<option' + (n === actual ? ' selected' : '') + '>' + esc(n) + '</option>';
    }).join('') : '<option value="">Sin sprints</option>';
    $('#sp-sprint').disabled = !lista.length;

    var s = sp.porSprint.find(function (x) { return x.nombre === actual; }) || null;
    var total = sp.tabla.total || 1;
    var regla = sp.tabla.filas.map(function (f) { return f.estados.join(' / ') + ' ' + pct(f.avance / total, 0); });
    $('#sp-regla').textContent = 'Paquetes ' + cfg.tipo + ' relacionados con las HU, agrupados por el Sprint de la Historia de Usuario. Avance de cada paquete: ' +
      regla.join(', ') + ', terminados (' + (cfg.terminados || []).join(', ') + ') 100% y otros estados 0%.';
    var sinSprint = sp.sinSprint.n
      ? ' ' + fmtNum.format(sp.sinSprint.n) + ' paquetes ' + cfg.tipo + ' son de HU sin sprint y no entran en el avance del sprint.' : '';
    if (s) {
      $('#sp-titular').textContent = 'Avance del sprint: ' + pct(s.avance);
      $('#sp-detalle').textContent = fmtNum.format(s.terminados) + ' de ' + fmtNum.format(s.n) + ' paquetes ' + cfg.tipo + ' del ' + actual +
        ' están terminados, ' + fmtNum.format(s.enPruebas) + ' en pruebas y ' + fmtNum.format(s.sinAvance) + ' sin avance.' + sinSprint;
      $('#sp-cifras').innerHTML = [
        ['Avance del sprint', pct(s.avance)], ['Paquetes', fmtNum.format(s.n)], ['Terminados', fmtNum.format(s.terminados)],
        ['Puntos pendientes', fmtNum.format(s.puntos - s.puntosTerminados) + ' de ' + fmtNum.format(s.puntos)]
      ].map(function (c) { return '<div><dt>' + esc(c[0]) + '</dt><dd>' + esc(c[1]) + '</dd></div>'; }).join('');
    } else {
      $('#sp-titular').textContent = actual ? 'El ' + actual + ' no tiene paquetes en este corte' : 'No hay paquetes ' + cfg.tipo + ' con sprint en este corte';
      $('#sp-detalle').textContent = (sp.total.n ? fmtNum.format(sp.total.n) + ' paquetes ' + cfg.tipo + ' relacionados con las HU, con ' + pct(sp.total.avance) + ' de avance promedio.' : '') + sinSprint;
      $('#sp-cifras').innerHTML = '';
    }

    // Avance de cada sprint
    var filaSprint = function (x, clase) {
      return '<tr' + (clase ? ' class="' + clase + '"' : '') + '><td>' + (x.nombre === 'Sin sprint' ? esc(x.nombre) : '<button class="btn-texto" type="button" data-sp="' + esc(x.nombre) + '">' + esc(x.nombre) + '</button>') +
        '</td><td class="num">' + fmtNum.format(x.n) + '</td><td class="num">' + fmtNum.format(x.sinAvance) + '</td><td class="num">' + fmtNum.format(x.enPruebas) +
        '</td><td class="num">' + fmtNum.format(x.terminados) + '</td><td class="num">' + fmtNum.format(x.puntosTerminados) + ' / ' + fmtNum.format(x.puntos) +
        '</td><td>' + esc(x.fin ? fecha(x.fin) : '') + '</td><td class="av-celda">' + barraAvance(x.avance) + '<b>' + pct(x.avance) + '</b></td></tr>';
    };
    $('#t-sp-sprints-tabla').innerHTML = '<thead><tr><th>Sprint</th><th class="num">Paquetes</th><th class="num">Sin avance</th><th class="num">En pruebas</th><th class="num">Terminados</th><th class="num">Puntos terminados</th><th>Fin</th><th>Avance</th></tr></thead><tbody>' +
      (sp.porSprint.length ? sp.porSprint.map(function (x) { return filaSprint(x, x.nombre === actual ? 'actual' : ''); }).join('') : '<tr><td colspan="8">Ninguna HU con paquetes ' + esc(cfg.tipo) + ' tiene sprint en este corte.</td></tr>') +
      (sp.sinSprint.n ? filaSprint(sp.sinSprint, 'sep') : '') + '</tbody>';

    estado.sprintActual = actual;
    renderBurndown(cortesSprint, actual);
    renderSprintId();
  }

  function renderBurndown(cortesSprint, sprint) {
    var cfg = estado.config.dyd;
    var mov = $('#t-sp-movidos-tabla');
    if (!cortesSprint.length || !sprint) {
      $('#sp-burndown-resumen').textContent = '';
      $('#sp-burndown-nota').textContent = !sprint ? 'No hay sprints para graficar.'
        : 'Aún no hay Excel cargados como Sprint. Al cargar un Excel elige "Sprint" (o cámbialo en Archivos cargados); con un corte por fecha se arma el Burndown.';
      if (estado.graficos['g-burndown']) { estado.graficos['g-burndown'].destroy(); delete estado.graficos['g-burndown']; }
      mov.innerHTML = '<tbody><tr><td>Se necesitan al menos dos cortes de Sprint para detectar los ID que pasan de un sprint a otro.</td></tr></tbody>';
      return;
    }
    var bd = I.calcularBurndown(cortesSprint, sprint, cfg);
    var porPuntos = bd.hayPuntos;
    var u = porPuntos ? 'puntos' : 'paquetes';
    var ultimo = bd.puntos[bd.puntos.length - 1];
    var salieron = bd.movimientos.filter(function (m) { return m.tipo === 'salio'; }).length;
    var entraron = bd.movimientos.length - salieron;
    $('#sp-burndown-resumen').textContent = fmtNum.format(porPuntos ? ultimo.pendientes : ultimo.paquetesPendientes) + ' ' + u + ' pendientes al ' + fechaCorte(ultimo.corte) +
      ' (de ' + fmtNum.format(porPuntos ? bd.iniciales : bd.inicialesPaquetes) + ' en el primer corte); ' + salieron + ' ID salieron del sprint y ' + entraron + ' entraron.';
    $('#sp-burndown-nota').textContent = 'Cada punto es un Excel cargado como Sprint (' + bd.puntos.length + (bd.puntos.length === 1 ? ' corte' : ' cortes') + '). ' +
      'Pendiente = ' + u + ' de los paquetes ' + cfg.tipo + ' del sprint que no están terminados. La línea ideal va de lo pendiente en el primer corte a 0 el ' + fechaCorte(bd.fin) +
      ' (fecha de finalización más tardía de los paquetes del sprint).' + (porPuntos ? '' : ' Los paquetes no tienen puntos de historia, por eso se cuentan paquetes.') +
      (bd.puntos.length < 2 ? ' Carga otro corte del sprint para ver la tendencia.' : '');

    if (typeof Chart !== 'undefined') {
      configurarChart();
      var c = colores();
      var etiquetas = bd.puntos.map(function (p) { return p.corte; });
      if (bd.fin && etiquetas.indexOf(bd.fin) === -1) etiquetas.push(bd.fin);
      var real = etiquetas.map(function (iso) { var p = bd.puntos.find(function (x) { return x.corte === iso; }); return p ? (porPuntos ? p.pendientes : p.paquetesPendientes) : null; });
      var ideal = etiquetas.map(function (iso) {
        var p = bd.puntos.find(function (x) { return x.corte === iso; });
        if (p) return +(porPuntos ? p.ideal : p.idealPaquetes).toFixed(2);
        return iso === bd.fin ? 0 : null;
      });
      grafico('g-burndown', {
        type: 'line',
        data: {
          labels: etiquetas.map(fechaCorte),
          datasets: [
            { label: 'Pendiente real', data: real, borderColor: cssVar('--navy'), backgroundColor: cssVar('--navy'), borderWidth: 2.5, tension: 0, spanGaps: false, pointRadius: 5, pointHoverRadius: 7, pointBorderColor: c.superficie, pointBorderWidth: 2 },
            { label: 'Ideal', data: ideal, borderColor: cssVar('--coral'), backgroundColor: cssVar('--coral'), borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, pointHoverRadius: 0, spanGaps: true, fill: false }
          ]
        },
        options: {
          interaction: { mode: 'index', intersect: false },
          scales: {
            x: { grid: { display: false }, border: { color: c.grid }, ticks: { maxRotation: 0, autoSkipPadding: 12 } },
            y: { beginAtZero: true, grid: { color: c.grid }, border: { display: false }, ticks: { precision: 0 }, title: { display: true, text: porPuntos ? 'Puntos pendientes' : 'Paquetes pendientes' } }
          },
          plugins: {
            legend: { position: 'top', align: 'start' },
            tooltip: { filter: function (it) { return it.raw !== null; }, callbacks: { label: function (it) { return ' ' + it.dataset.label + ': ' + fmtPp.format(it.raw) + ' ' + u; } } }
          }
        }
      });
    }

    var DESC = { salio: 'Salió', entro: 'Entró' };
    mov.innerHTML = '<thead><tr><th>Movimiento</th><th class="num">ID</th><th>Asunto</th><th>HU</th><th>Sprint anterior</th><th>Sprint nuevo</th><th>Detectado en el corte</th><th class="num">Puntos</th><th>Estado</th></tr></thead><tbody>' +
      (bd.movimientos.length ? bd.movimientos.map(function (m) {
        return '<tr><td><span class="chip cambio-' + (m.tipo === 'salio' ? 'retirada' : 'nueva') + '">' + DESC[m.tipo] + '</span></td><td class="num">' + esc(m.id) + '</td><td class="asunto">' + esc(m.asunto) +
          '</td><td>' + esc(m.hus.join(', ')) + '</td><td>' + esc(m.de) + '</td><td>' + esc(m.a) + '</td><td>' + esc(fechaCorte(m.corte)) + '</td><td class="num">' + fmtNum.format(m.puntos) + '</td><td>' + esc(m.estado) + '</td></tr>';
      }).join('') : '<tr><td colspan="9">' + (bd.puntos.length < 2 ? 'Se necesitan al menos dos cortes de Sprint para detectar los ID que pasan de un sprint a otro.' : 'Ningún ID entró ni salió del ' + esc(sprint) + ' entre los cortes cargados.') + '</td></tr>') + '</tbody>';
  }

  function renderSprintId() {
    var sp = estado.sprintRes;
    if (!sp) return;
    var tipo = estado.config.dyd.tipo, sprint = estado.sprintActual;
    var q = ($('#sp-buscar').value || '').trim().toLowerCase();
    var solo = $('#sp-solo-pendientes').checked;
    // Mismo sprint elegido arriba (el de la Historia de Usuario)
    var delSprint = sp.porHU.filter(function (x) { return !sprint || x.sprint === sprint; });
    var filas = delSprint.filter(function (x) { return !solo || x.avance < 1; }).filter(function (x) {
      return !q || [x.hu.id, x.hu.asunto, x.hu.cliente, x.hu.estado].concat(x.items.map(function (i) { return i.paquete.id + ' ' + i.paquete.asunto; })).join(' ').toLowerCase().indexOf(q) !== -1;
    }).sort(function (a, b) { return a.avance - b.avance || Number(a.hu.id) - Number(b.hu.id); });
    var pendientes = delSprint.filter(function (x) { return x.avance < 1; }).length;
    $('#sp-id-nota').textContent = fmtNum.format(delSprint.length) + ' HU ' + (sprint ? 'del ' + sprint + ' ' : '') + 'tienen paquetes ' + tipo + ' relacionados (' +
      fmtNum.format(pendientes) + ' sin terminar)' + (filas.length !== delSprint.length ? '; se muestran ' + fmtNum.format(filas.length) : '') +
      '. El avance de cada HU es el promedio de sus paquetes; solo cuentan Desarrollo y Pruebas funcionales.';
    var lienzo = $('#sp-id-lienzo');
    lienzo.hidden = !filas.length;
    if (!filas.length) {
      $('#sp-id-nota').textContent += ' ' + (delSprint.length ? 'Ninguna HU coincide con el filtro.' : 'No hay HU con paquetes ' + tipo + ' relacionados en este sprint.');
      if (estado.graficos['g-sp-id']) { estado.graficos['g-sp-id'].destroy(); delete estado.graficos['g-sp-id']; }
      return;
    }
    if (typeof Chart === 'undefined') return;
    configurarChart();
    var c = colores(), navy = cssVar('--navy'), coral = cssVar('--coral');
    lienzo.style.height = Math.max(160, filas.length * 30 + 50) + 'px';
    var corto = function (t) { t = String(t); return t.length > 42 ? t.slice(0, 41) + '…' : t; };
    grafico('g-sp-id', {
      type: 'bar',
      data: {
        labels: filas.map(function (x) { return x.hu.id + ' · ' + corto(x.hu.asunto); }),
        datasets: [barra(c, filas.map(function (x) { return x.avance >= 1 ? navy : x.avance > 0 ? cssVar('--fase-2') : coral; }), {
          label: 'Avance', data: filas.map(function (x) { return +(x.avance * 100).toFixed(1); }), borderRadius: 1, maxBarThickness: 18, minBarLength: 4
        })]
      },
      options: {
        indexAxis: 'y', layout: { padding: { right: 56 } },
        scales: {
          x: { beginAtZero: true, max: 100, grid: { color: c.grid }, border: { display: false }, ticks: { callback: function (v) { return v + '%'; } } },
          y: { grid: { display: false }, border: { color: c.grid }, ticks: { autoSkip: false, callback: etiquetaCorta } }
        },
        plugins: {
          legend: { display: false },
          etiquetasTotales: { activo: true, mostrarCero: true, formato: function (v) { return fmtNum.format(v) + '%'; } },
          tooltip: { callbacks: {
            title: function (its) { var x = filas[its[0].dataIndex]; return 'HU ' + x.hu.id + ' · ' + x.hu.cliente; },
            label: function (it) { return ' Avance: ' + pct(filas[it.dataIndex].avance, 0); },
            afterLabel: function (it) {
              return filas[it.dataIndex].items.map(function (i) { return i.paquete.id + ' ' + (i.paquete.estado || 'Sin estado') + ': ' + pct(i.avance, 0); });
            }
          } }
        }
      },
      plugins: [etiquetasTotales]
    });
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
    var sinModulo = estado.columnas && !estado.columnas.modulo
      ? '<span class="aviso-info">El archivo no tiene la columna Módulo afectado: las HU quedan como Sin módulo identificado.</span>' : '';
    $('#fuente-alertas').innerHTML = sinModulo + (no.length
      ? '<span class="alerta">Hay estados sin fase asignada que no se cuentan como HU consideradas: ' + esc(no.map(function (k) { return k + ' (' + r.estadosNoListados[k] + ')'; }).join(', ')) + '. Asígnales una fase en Configuración.</span>'
      : '');
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
      partes.push('Fuera de meta: ' + listaNatural(fuera.map(function (k) { return k.nombre.replace(/^% /, '') + ' (' + valorKpi(k, k.resultado) + ')'; })) + '.');
    }
    if (sinDatos.length) {
      partes.push((sinDatos.length === 1 ? 'Falta el dato de ' : 'Faltan los datos de ') +
        listaNatural(sinDatos.map(function (k) { return k.unidad === 'h' ? 'horas contratadas (se ingresan en Configuración)' : esHpt(k) ? 'puntos de historia para la Efectividad' : k.nombre.replace(/^% /, ''); })) + '.');
    }
    $('#veredicto-detalle').textContent = partes.join(' ');

    var excl = b.excluidos.reduce(function (s, e) { return s + e.n; }, 0) + b.sinMapear;
    var inf = informeActual();
    $('#tab-seccion').textContent = 'Excelencia operativa, corte del ' + (inf ? fechaCorte(inf.corte) : fechaCorte(hoyISO()));

    // Cortes evaluados (como el selector de años de la referencia)
    $('#lista-cortes').innerHTML = estado.informes.length ? estado.informes.slice(0, 8).map(function (i) {
      var actual = i.id === estado.informeId;
      return '<li><button type="button" data-corte="' + esc(i.id) + '"' + (actual ? ' aria-current="true"' : '') + ' title="' + esc(i.nombre) + '">' + esc(fechaCorte(i.corte)) + '</button></li>';
    }).join('') + (estado.informes.length > 8 ? '<li class="mas">' + (estado.informes.length - 8) + ' más en Comparativo</li>' : '')
      : '<li class="mas">Solo este corte</li>';

    // Tarjetas con cinta
    var tarjetas = [
      ['HU consideradas', fmtNum.format(b.consideradas), 'de ' + fmtNum.format(b.total) + ' únicas; ' + fmtNum.format(excl) + ' excluidas'],
      ['HU abiertas', fmtNum.format(r.porFlujo[0].n), pct(r.porFlujo[0].pct, 0) + ' de las consideradas'],
      ['HU cerradas', fmtNum.format(r.porFlujo[1].n), pct(r.porFlujo[1].pct, 0) + ' de las consideradas'],
      ['Horas ejecutadas', fmtNum.format(r.horas.ejecutadas) + ' h', r.horas.contratadas ? 'de ' + fmtNum.format(r.horas.contratadas) + ' h contratadas' : 'sin horas contratadas']
    ];
    $('#kpis-cinta').innerHTML = tarjetas.map(function (t) {
      return '<div class="caja t-kpi"><span class="cinta">' + esc(t[0]) + '</span><b>' + esc(t[1]) + '</b><small>' + esc(t[2]) + '</small></div>';
    }).join('');

    // Estado de las metas (como la columna lateral de la referencia)
    var av = estado.avance;
    var fila = function (clase, icono, etiqueta, n) {
      return '<div class="m-fila ' + clase + '">' + ICONOS[icono] + '<div><span>' + esc(etiqueta) + '</span><b>' + n + ' de ' + plantilla.length + '</b></div></div>';
    };
    $('#t-metas').innerHTML = '<p class="m-sub">Indicadores de la plantilla</p>' +
      fila('en-meta', 'EN META', 'En meta', enMeta.length) +
      fila('fuera', 'FUERA DE META', 'Fuera de meta', fuera.length) +
      fila('', 'SIN DATOS', 'Sin datos', sinDatos.length) +
      '<div class="m-grande"><span>Avance promedio</span><b>' + (av ? pct(av.promedio) : 'N/D') + '</b></div>';
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

  // Los KPI son porcentajes, salvo Efectividad, que está en horas por punto de historia
  var fmtDec = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  function esHpt(k) { return k.unidad === 'h/pt'; }
  function valorKpi(k, v, dec) {
    if (!esHpt(k)) return pct(v, dec);
    return v === null || v === undefined ? 'N/D' : fmtDec.format(v) + ' h/pt';
  }
  function textoMeta(k) {
    return tieneMeta(k) ? (k.tipoMeta === '<=' ? '≤ ' : '≥ ') + valorKpi(k, k.meta, 0) : 'Sin meta';
  }
  // Escala de las diferencias: puntos porcentuales o h/pt
  function escalaKpi(k) { return esHpt(k) ? 1 : 100; }
  function unidadDif(k) { return esHpt(k) ? ' h/pt' : ' pp'; }

  function textoDetalle(k) {
    if (k.unidad === 'h') {
      return k.resultado === null
        ? fmtNum.format(k.numerador) + ' h ejecutadas; faltan las horas contratadas'
        : fmtNum.format(k.numerador) + ' h de ' + fmtNum.format(k.denominador) + ' h';
    }
    if (esHpt(k)) {
      return k.resultado === null ? 'Sin HU con puntos de historia'
        : fmtDec.format(k.numerador) + ' h en ' + fmtNum.format(k.denominador) + ' puntos';
    }
    if (k.resultado === null) return 'Sin datos en la base';
    return fmtNum.format(k.numerador) + ' de ' + fmtNum.format(k.denominador) + ' HU';
  }

  // Gráfico de bala: valor frente a la meta, con la zona fuera de meta rayada
  function bala(k) {
    var meta = tieneMeta(k) ? k.meta : null;
    // En h/pt se deja margen después de la meta para ver la zona fuera de meta
    var escala = esHpt(k) ? Math.max(k.resultado || 0, meta || 0) * 1.25 || 1 : Math.max(1, k.resultado || 0, meta || 0);
    var x = function (v) { return Math.min(100, v / escala * 100); };
    var textoMetaCorto = meta === null ? '' : textoMeta(k);
    var zona = '';
    if (meta !== null) {
      zona = k.tipoMeta === '<='
        ? '<span class="bala-zona" style="left:' + x(meta) + '%;right:0"></span>'
        : '<span class="bala-zona" style="left:0;width:' + x(meta) + '%"></span>';
    }
    var etiqueta = k.nombre + ': ' + valorKpi(k, k.resultado) + (meta !== null ? ', meta ' + textoMetaCorto : '');
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
    var d = (k.resultado - k.meta) * escalaKpi(k);
    return k.tipoMeta === '<=' ? -d : d;
  }
  var fmtPp = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  function fraseDesviacion(k) {
    var d = desviacion(k);
    if (d === null) return k.unidad === 'h' ? 'Falta ingresar las horas contratadas.' : 'Sin datos en la base.';
    var pp = fmtPp.format(Math.abs(d)) + (esHpt(k) ? ' h/pt' : ' puntos');
    var meta = valorKpi(k, k.meta, 0);
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
          '<div class="dist-nombre"><strong>' + esc(x.nombre) + '</strong><small>' + valorKpi(x, x.resultado) + ' frente a ' + textoMeta(x) + '</small></div>' +
          '<div class="dist-pista" role="img" aria-label="' + esc(x.nombre + ': ' + fraseDesviacion(x)) + '">' +
            '<span class="dist-barra ' + cls + '" style="' + (d >= 0 ? 'left:50%' : 'right:50%') + ';width:' + ancho + '%"></span>' +
            '<span class="dist-valor ' + cls + '" style="' + (d >= 0 ? 'left:calc(50% + ' + ancho + '% + 8px)' : 'right:calc(50% + ' + ancho + '% + 8px)') + '">' + (d >= 0 ? '+' : '−') + fmtPp.format(Math.abs(d)) + unidadDif(x) + '</span>' +
          '</div>' +
          '<div class="dist-estado">' + pill(x.estado) + '</div>' +
          '</div>';
      }).join('') +
      (sin.length ? '<p class="nota">Sin datos para comparar: ' + esc(listaNatural(sin.map(function (x) { return x.nombre + (x.unidad === 'h' ? ' (faltan las horas contratadas)' : ''); }))) + '.</p>' : '') +
      '<p class="nota">pp = puntos porcentuales de diferencia entre el resultado y la meta; h/pt = horas por punto de historia.</p>';
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
          return '<li><div class="g-linea"><span>' + esc(x.nombre) + '</span><b>' + valorKpi(x, x.resultado) + '</b></div>' +
            '<p>' + esc(fraseDesviacion(x)) + (x.resultado !== null ? ' Base: ' + esc(textoDetalle(x)) + '.' : '') + '</p></li>';
        }).join('') + '</ul>' : '<p class="g-vacio">Ningún indicador en este grupo.</p>') + '</section>';
    }).join('') + '</div>';
  }

  function renderPorCliente() {
    var ids = ['deuda', 'backlog', 'refinement', 'alcance', 'ambiguedad', 'efectividad'];
    var cortos = { deuda: 'Deuda técnica', backlog: 'Cumplimiento backlog', refinement: 'Precisión refinement', alcance: 'Cambios de alcance', ambiguedad: 'Retrabajo', efectividad: 'Efectividad' };
    var porCliente = {};
    estado.filtrados.forEach(function (r) { (porCliente[r.cliente] = porCliente[r.cliente] || []).push(r); });
    var filas = Object.keys(porCliente).map(function (c) {
      var res = I.calcular(porCliente[c], estado.config);
      return { cliente: c, n: res.base.consideradas, kpis: res.kpis };
    }).filter(function (f) { return f.n > 0; }).sort(function (a, b) { return b.n - a.n; });
    var celda = function (k) {
      var cls = k.estado === 'EN META' ? 'en-meta' : k.estado === 'FUERA DE META' ? 'fuera' : 'nd';
      return '<td class="mc ' + cls + '" title="' + esc(k.nombre + ': ' + textoDetalle(k)) + '">' +
        (k.resultado === null ? '<span class="mc-nd">sin datos</span>' : ICONOS[k.estado] + valorKpi(k, k.resultado)) + '</td>';
    };
    var total = plantilla();
    $('#vista-cliente').innerHTML = '<div class="tabla-scroll"><table class="tabla matriz"><thead><tr><th>Cliente</th><th class="num">HU</th>' +
      ids.map(function (id) { var k = total.find(function (x) { return x.id === id; }); return '<th>' + esc(cortos[id]) + '<small>' + esc(textoMeta(k)) + '</small></th>'; }).join('') +
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
          '<div class="sc-valor' + (x.resultado === null ? ' nd' : '') + '" role="cell">' + (esHpt(x) && x.resultado !== null ? fmtDec.format(x.resultado) + '<small class="unidad">h/pt</small>' : valorKpi(x, x.resultado)) + '</div>' +
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
    return { 'Desarrollo Contratado': c.s[0], 'Mejora Funcional': c.s[1], 'Deuda Tecnica': c.s[2], 'Garantía': c.s[6], 'Sin clasificar': c.muted };
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
        if ((!total && !opts.mostrarCero) || !extremo) continue;
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

    // ---- Tablero (estilo de la referencia: barras azul marino, dona azul y coral) ----
    var navy = cssVar('--navy'), coral = cssVar('--coral');
    var sinGrid = function (horizontal) {
      var e = ejeValores(c, horizontal);
      (horizontal ? e.x : e.y).grid = { display: false };
      (horizontal ? e.x : e.y).border = { color: c.grid };
      return e;
    };
    var barrasH = function (id, filas, extraTooltip) {
      document.getElementById(id).parentNode.style.height = Math.max(200, filas.length * 30 + 50) + 'px';
      grafico(id, {
        type: 'bar',
        data: { labels: filas.map(function (f) { return f.nombre; }), datasets: [barra(c, navy, { label: 'HU', data: filas.map(function (f) { return f.n; }), borderRadius: 1, maxBarThickness: 16 })] },
        options: {
          indexAxis: 'y', scales: sinGrid(true), layout: { padding: { right: 64 } },
          plugins: {
            legend: { display: false },
            etiquetasTotales: { activo: true, formato: function (v, i) { return extraTooltip ? extraTooltip(v, i) : String(v); } },
            tooltip: { callbacks: { label: function (it) { return ' ' + it.raw + ' HU'; } } }
          }
        },
        plugins: [etiquetasTotales]
      });
    };

    barrasH('g-tipo', r.porTipo.filter(function (t) { return t.n; }).sort(function (x, y) { return y.n - x.n; }), function (v) {
      return v + ' (' + pct(r.base.consideradas ? v / r.base.consideradas : 0) + ')';
    });

    grafico('g-fase', {
      type: 'bar',
      data: { labels: r.porFase.map(function (f) { return f.fase; }), datasets: [barra(c, navy, { label: 'HU', data: r.porFase.map(function (f) { return f.n; }), borderRadius: 1, maxBarThickness: 56, categoryPercentage: 0.8 })] },
      options: {
        scales: sinGrid(false), layout: { padding: { top: 22 } },
        plugins: {
          legend: { display: false },
          etiquetasTotales: { activo: true },
          tooltip: { callbacks: { label: function (it) { return ' ' + it.raw + ' HU (' + pct(r.porFase[it.dataIndex].pct) + ')'; } } }
        }
      },
      plugins: [etiquetasTotales]
    });

    var flujo = r.porFlujo;
    grafico('g-flujo', {
      type: 'doughnut',
      data: { labels: ['Abiertas', 'Cerradas'], datasets: [{ data: flujo.map(function (f) { return f.n; }), backgroundColor: [coral, navy], borderColor: c.superficie, borderWidth: 2, hoverOffset: 4 }] },
      options: {
        cutout: '55%', layout: { padding: 6 },
        plugins: {
          legend: { position: 'bottom', labels: { generateLabels: function (ch) {
            var base = Chart.overrides.doughnut.plugins.legend.labels.generateLabels(ch);
            base.forEach(function (l, i) { l.text = (i ? 'Cerradas ' : 'Abiertas ') + pct(flujo[i].pct) + ' (' + flujo[i].n + ')'; });
            return base;
          } } },
          tooltip: { callbacks: { label: function (it) { return ' ' + it.label + ': ' + it.raw + ' HU (' + pct(flujo[it.dataIndex].pct) + ')'; } } }
        }
      }
    });

    var modulos = estado.resultado.seguimiento.porModulo.map(function (f) { return { nombre: f.nombre, n: f.total }; });
    document.getElementById('g-modulo').parentNode.style.height = '300px';
    grafico('g-modulo', {
      type: 'bar',
      data: { labels: modulos.map(function (f) { return f.nombre; }), datasets: [barra(c, navy, { label: 'HU abiertas', data: modulos.map(function (f) { return f.n; }), borderRadius: 1, maxBarThickness: 18 })] },
      options: {
        scales: Object.assign(sinGrid(false), { x: { grid: { display: false }, border: { color: c.grid }, ticks: { maxRotation: 45, minRotation: 45, autoSkip: false, callback: etiquetaCorta } } }),
        layout: { padding: { top: 18 } },
        plugins: { legend: { display: false }, etiquetasTotales: { activo: true }, tooltip: { callbacks: { label: function (it) { return ' ' + it.raw + ' HU abiertas'; } } } }
      },
      plugins: [etiquetasTotales]
    });

    var clientes = (estado.avance ? estado.avance.porCliente : []).slice(0, 10).map(function (f) { return { nombre: f.nombre, n: f.n }; });
    barrasH('g-cliente', clientes);
  }

  function graficoApilado(id, filas, totalAbiertas) {
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
        layout: { padding: { right: 84 } },
        interaction: { mode: 'index', intersect: false, axis: 'y' },
        plugins: {
          legend: { position: 'top', align: 'start' },
          etiquetasTotales: { activo: true, formato: function (v) { return v + ' (' + pct(totalAbiertas ? v / totalAbiertas : 0) + ')'; } },
          tooltip: { filter: function (it) { return it.raw; }, callbacks: { footer: function (its) { var f = filas[its[0].dataIndex]; return 'Total: ' + f.total + ' HU (' + pct(f.pct) + ' del total)'; } } }
        }
      },
      plugins: [etiquetasTotales]
    });
  }

  function renderSeguimiento() {
    var s = estado.resultado.seguimiento;
    [['cliente', 'porCliente', 'Cliente'], ['sprint', 'porSprint', 'Sprint'], ['modulo', 'porModulo', 'Módulo'], ['asignado', 'porAsignado', 'Asignado a']]
      .forEach(function (p) {
        if (typeof Chart !== 'undefined') graficoApilado('g-seg-' + p[0], s[p[1]], s.total);
      });
  }

  // ---------- detalle ----------
  var COLS_DETALLE = [
    ['id', 'ID'], ['asunto', 'Asunto'], ['cliente', 'Cliente'], ['estado', 'Estado'], ['fase', 'Fase'],
    ['enFlujo', 'En flujo'], ['tipoSolicitud', 'Tipo de Solicitud'], ['sprint', 'Sprint'], ['modulo', 'Módulo'],
    ['responsable', 'Responsable'], ['asignado', 'Asignado a'], ['trabajoTotal', 'Trabajo total (h)'],
    ['puntos', 'Puntos de historia'], ['tiempo', 'Tiempo invertido (h)'], ['avance', 'Avance'],
    ['fechaFin', 'Fecha finalización'], ['precision', 'Precisión HU']
  ];
  function colsDetalle() {
    // Las columnas opcionales solo se muestran si vienen en el Excel
    return COLS_DETALLE.filter(function (c) {
      return ['responsable', 'puntos', 'tiempo'].indexOf(c[0]) === -1 || (estado.columnas && estado.columnas[c[0]]);
    });
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
        var txt = [r.id, r.asunto, r.cliente, r.estado, r.estadoOriginal, r.responsable, r.asignado, r.sprint, r.modulo, r.precision].join(' ').toLowerCase();
        if (txt.indexOf(q) === -1) return false;
      }
      return true;
    }).sort(function (a, b) {
      var va = a[o.campo], vb = b[o.campo], res;
      if (va instanceof Date || vb instanceof Date) res = (va ? va.getTime() : Infinity) - (vb ? vb.getTime() : Infinity);
      else if (o.campo === 'avance') res = (va === null || va === undefined ? -1 : va) - (vb === null || vb === undefined ? -1 : vb);
      else if (['id', 'trabajoTotal', 'puntos', 'tiempo'].indexOf(o.campo) !== -1) res = Number(va) - Number(vb) || String(va).localeCompare(String(vb));
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
        if (c[0] === 'avance') return '<td class="num">' + (v === null || v === undefined ? '—' : pct(v, 0)) + '</td>';
        var cls = c[0] === 'asunto' ? ' class="asunto"' : ['id', 'trabajoTotal', 'puntos', 'tiempo'].indexOf(c[0]) !== -1 ? ' class="num"' : '';
        if (c[0] === 'tiempo') v = v ? fmtDec.format(v) : '';
        if (c[0] === 'puntos') v = v || '';
        if (c[0] === 'enFlujo' || c[0] === 'fase') return '<td><span class="chip">' + esc(v) + '</span></td>';
        if (c[0] === 'estado' && r.estadoOriginal && r.estadoOriginal !== r.estado) {
          return '<td>' + esc(v) + '<small class="estado-erp">En el ERP: ' + esc(r.estadoOriginal) + '</small></td>';
        }
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

  // ---------- comparativo de informes ----------
  var TIPOS_CAMBIO = [
    ['todos', 'Todos'], ['nueva', 'Nuevas'], ['cerrada', 'Cerradas'], ['reabierta', 'Reabiertas'],
    ['estado', 'Cambio de estado'], ['retirada', 'Retiradas']
  ];
  var DESCRIPCION_CAMBIO = {
    nueva: 'Nueva', cerrada: 'Cerrada', reabierta: 'Reabierta', estado: 'Cambio de estado', retirada: 'Retirada'
  };

  function opcionesInformes(seleccionado) {
    return estado.informes.map(function (i) {
      return '<option value="' + esc(i.id) + '"' + (i.id === seleccionado ? ' selected' : '') + '>' + esc(etiquetaInforme(i)) + '</option>';
    }).join('');
  }

  // Variación en puntos; favorable según el tipo de meta
  function celdaVariacion(k, a, b) {
    if (a === null || b === null) return '<td class="num var-nd">sin dato</td>';
    var d = (b - a) * escalaKpi(k);
    if (Math.abs(d) < 0.05) return '<td class="num var-igual">sin cambio</td>';
    var cls = 'var-neutra';
    if (tieneMeta(k)) cls = (k.tipoMeta === '<=' ? d < 0 : d > 0) ? 'var-mejora' : 'var-empeora';
    var flecha = d > 0 ? '▲' : '▼';
    return '<td class="num ' + cls + '"><span aria-hidden="true">' + flecha + '</span> ' + (d > 0 ? '+' : '−') + fmtPp.format(Math.abs(d)) + unidadDif(k) +
      (cls === 'var-mejora' ? '<span class="sr"> (mejora)</span>' : cls === 'var-empeora' ? '<span class="sr"> (empeora)</span>' : '') + '</td>';
  }

  function renderComparativo() {
    if (!$('#comp-contenido')) return;
    var lista = estado.informes;
    renderInformesGuardados();
    var hay = lista.length >= 2;
    $('#comp-vacio').hidden = hay;
    $('#comp-contenido').hidden = !hay;
    if (!hay) return;

    var c = estado.comp;
    var existe = function (id) { return lista.some(function (i) { return i.id === id; }); };
    if (!existe(c.b)) c.b = lista[0].id;
    if (!existe(c.a) || c.a === c.b) c.a = (lista.find(function (i) { return i.id !== c.b; }) || lista[1]).id;
    $('#comp-a').innerHTML = opcionesInformes(c.a);
    $('#comp-b').innerHTML = opcionesInformes(c.b);

    var infA = lista.find(function (i) { return i.id === c.a; });
    var infB = lista.find(function (i) { return i.id === c.b; });
    var rA = resultadoDeInforme(infA), rB = resultadoDeInforme(infB);
    var filtros = Object.keys(estado.filtros).filter(function (k) { return estado.filtros[k]; }).length;
    $('#comp-nota').textContent = 'Se compara el corte del ' + fechaCorte(infA.corte) + ' con el del ' + fechaCorte(infB.corte) + '. ' +
      (filtros ? 'Los filtros activos se aplican a los dos informes. ' : '') +
      'Las metas y parámetros son los de Configuración.';

    // Indicadores
    var mejoran = 0, empeoran = 0;
    var filas = rB.kpis.map(function (kB, i) {
      var kA = rA.kpis[i];
      var d = kA.resultado === null || kB.resultado === null ? null : (kB.resultado - kA.resultado) * escalaKpi(kB) / 100;
      if (d !== null && Math.abs(d) >= 0.0005 && tieneMeta(kB)) {
        if (kB.tipoMeta === '<=' ? d < 0 : d > 0) mejoran++; else empeoran++;
      }
      var meta = textoMeta(kB);
      var cambioEstado = kA.estado === kB.estado ? pill(kB.estado) : pill(kA.estado) + '<span class="comp-a" aria-label="pasa a">→</span>' + pill(kB.estado);
      var sep = i > 0 && rB.kpis[i - 1].grupo !== kB.grupo;
      return '<tr' + (sep ? ' class="sep"' : '') + '><td><strong>' + esc(kB.nombre) + '</strong></td><td>' + meta + '</td>' +
        '<td class="num">' + valorKpi(kA, kA.resultado) + '<small>' + esc(textoDetalle(kA)) + '</small></td>' +
        '<td class="num">' + valorKpi(kB, kB.resultado) + '<small>' + esc(textoDetalle(kB)) + '</small></td>' +
        celdaVariacion(kB, kA.resultado, kB.resultado) + '<td class="comp-estado">' + cambioEstado + '</td></tr>';
    });
    $('#t-comp-kpi-tabla').innerHTML = '<thead><tr><th>Indicador</th><th>Meta</th><th class="num">' + esc(fechaCorte(infA.corte)) + '</th><th class="num">' +
      esc(fechaCorte(infB.corte)) + '</th><th class="num">Variación</th><th>Estado</th></tr></thead><tbody>' + filas.join('') + '</tbody>';
    $('#comp-kpi-resumen').textContent = mejoran + ' mejoran y ' + empeoran + ' empeoran frente a su meta';

    // Volumen de HU
    var fase = function (r, f) { return r.porFase.find(function (x) { return x.fase === f; }).n; };
    var base = [
      ['HU únicas en la base', rA.base.total, rB.base.total],
      ['HU consideradas', rA.base.consideradas, rB.base.consideradas],
      ['HU abiertas', rA.porFlujo[0].n, rB.porFlujo[0].n],
      ['HU cerradas', rA.porFlujo[1].n, rB.porFlujo[1].n]
    ].concat(I.FASES.map(function (f) { return ['En ' + f.toLowerCase(), fase(rA, f), fase(rB, f)]; }))
      .concat([['Horas ejecutadas', rA.horas.ejecutadas, rB.horas.ejecutadas, ' h']]);
    var avA = I.calcularAvance(registrosDeInforme(infA), estado.config.avance).promedio;
    var avB = I.calcularAvance(registrosDeInforme(infB), estado.config.avance).promedio;
    $('#t-comp-base-tabla').innerHTML = '<thead><tr><th></th><th class="num">' + esc(fechaCorte(infA.corte)) + '</th><th class="num">' +
      esc(fechaCorte(infB.corte)) + '</th><th class="num">Diferencia</th></tr></thead><tbody>' + base.map(function (f, i) {
        var d = f[2] - f[1], u = f[3] || '';
        return '<tr' + (i === 4 || i === 8 ? ' class="sep"' : '') + '><td>' + esc(f[0]) + '</td><td class="num">' + fmtNum.format(f[1]) + u + '</td><td class="num">' + fmtNum.format(f[2]) + u +
          '</td><td class="num">' + (d === 0 ? '<span class="var-igual">0</span>' : (d > 0 ? '+' : '−') + fmtNum.format(Math.abs(d)) + u) + '</td></tr>';
      }).join('') + '<tr class="sep"><td>Avance promedio de las HU</td><td class="num">' + pct(avA) + '</td><td class="num">' + pct(avB) + '</td><td class="num">' +
      (avA === null || avB === null ? '' : Math.abs(avB - avA) < 0.0005 ? '<span class="var-igual">sin cambio</span>' : (avB > avA ? '+' : '−') + fmtPp.format(Math.abs(avB - avA) * 100) + ' pp') + '</td></tr></tbody>';

    renderEvolucion(infA, infB);
    renderCambiosHU(infA, infB);
  }

  function renderEvolucion(infA, infB) {
    var sel = $('#comp-kpi-evol');
    var kpiDef = I.KPIS.find(function (k) { return k.id === estado.comp.kpi; }) || I.KPIS[1];
    sel.innerHTML = I.KPIS.map(function (k) {
      return '<option value="' + k.id + '"' + (k.id === kpiDef.id ? ' selected' : '') + '>' + esc(k.nombre) + '</option>';
    }).join('');
    var orden = estado.informes.slice().sort(function (a, b) {
      return (a.corte || '').localeCompare(b.corte || '') || (a.cargado || '').localeCompare(b.cargado || '');
    });
    var puntos = orden.map(function (inf) {
      var k = resultadoDeInforme(inf).kpis.find(function (x) { return x.id === kpiDef.id; });
      return { inf: inf, k: k, v: k.resultado === null ? null : +(k.resultado * escalaKpi(k)).toFixed(2) };
    });
    var k0 = puntos[puntos.length - 1].k;
    var sinDato = puntos.filter(function (p) { return p.v === null; }).length;
    $('#comp-evol-nota').textContent = puntos.length + ' informes guardados, ordenados por fecha de corte. Los puntos más grandes son los dos informes comparados.' +
      (sinDato ? ' ' + sinDato + ' sin dato para este indicador.' : '');
    if (typeof Chart === 'undefined') return;
    configurarChart();
    var c = colores();
    var resaltar = function (p) { return p.inf.id === infA.id || p.inf.id === infB.id; };
    var datasets = [{
      label: kpiDef.nombre, data: puntos.map(function (p) { return p.v; }),
      borderColor: c.s[0], backgroundColor: c.s[0], borderWidth: 2, tension: 0, spanGaps: true,
      pointRadius: puntos.map(function (p) { return resaltar(p) ? 6 : 4; }),
      pointHoverRadius: 7, pointBorderColor: c.superficie, pointBorderWidth: 2
    }];
    if (tieneMeta(k0)) {
      datasets.push({
        label: 'Meta ' + textoMeta(k0), data: puntos.map(function () { return +(k0.meta * escalaKpi(k0)).toFixed(2); }),
        borderColor: c.texto, borderWidth: 1.5, borderDash: [5, 4], pointRadius: 0, pointHoverRadius: 0, fill: false
      });
    }
    grafico('g-evolucion', {
      type: 'line',
      data: { labels: puntos.map(function (p) { return fechaCorte(p.inf.corte); }), datasets: datasets },
      options: {
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { grid: { display: false }, border: { color: c.grid }, ticks: { maxRotation: 0, autoSkipPadding: 12 } },
          y: { beginAtZero: true, grid: { color: c.grid }, border: { display: false }, ticks: { callback: function (v) { return v + (esHpt(kpiDef) ? ' h/pt' : '%'); } } }
        },
        plugins: {
          legend: { position: 'top', align: 'start' },
          tooltip: { callbacks: {
            title: function (its) { return etiquetaInforme(puntos[its[0].dataIndex].inf); },
            label: function (it) { return ' ' + it.dataset.label + ': ' + (it.raw === null ? 'sin dato' : fmtPp.format(it.raw) + (esHpt(kpiDef) ? ' h/pt' : '%')); }
          } }
        }
      }
    });
  }

  function renderCambiosHU(infA, infB) {
    var comp = I.compararHU(registrosDeInforme(infA), registrosDeInforme(infB));
    var total = comp.cambios.length;
    var tipo = estado.comp.tipo;
    $('#comp-tipos').innerHTML = TIPOS_CAMBIO.map(function (t) {
      var n = t[0] === 'todos' ? total : comp.resumen[t[0]];
      return '<button type="button" role="radio" data-cambio="' + t[0] + '" aria-checked="' + (t[0] === tipo) + '">' + esc(t[1]) + ' <b>' + n + '</b></button>';
    }).join('');
    var filas = comp.cambios.filter(function (x) { return tipo === 'todos' || x.tipo === tipo; });
    $('#t-comp-hu-tabla').innerHTML = '<thead><tr><th>Cambio</th><th class="num">ID</th><th>Asunto</th><th>Cliente</th><th>Estado el ' +
      esc(fechaCorte(infA.corte)) + '</th><th>Estado el ' + esc(fechaCorte(infB.corte)) + '</th></tr></thead><tbody>' +
      (filas.length ? filas.map(function (x) {
        return '<tr><td><span class="chip cambio-' + x.tipo + '">' + esc(DESCRIPCION_CAMBIO[x.tipo]) + '</span></td><td class="num">' + esc(x.id) +
          '</td><td class="asunto">' + esc(x.asunto) + '</td><td>' + esc(x.cliente) + '</td><td>' + esc(x.de || 'No estaba') + '</td><td>' + esc(x.a || 'Ya no está') + '</td></tr>';
      }).join('') : '<tr><td colspan="6">No hay HU con este tipo de cambio entre los dos informes.</td></tr>') + '</tbody>';
  }

  function renderInformesGuardados() {
    var lista = estado.informes;
    Historico.persistente().then(function (ok) {
      $('#hist-nota').textContent = (ok
        ? 'Los informes se guardan en este navegador. Para usarlos en otro equipo, exporta el histórico e impórtalo allá. '
        : 'Este navegador no permite guardar datos: el histórico se perderá al cerrar la página. Exporta el histórico para conservarlo. ') +
        'Eliminar un archivo lo quita de la app en este navegador; el Excel original no se modifica.';
    });
    // Limpia de la selección los informes que ya no existen
    Object.keys(estado.seleccion).forEach(function (id) { if (!lista.some(function (i) { return i.id === id; })) delete estado.seleccion[id]; });
    var nSel = Object.keys(estado.seleccion).length;
    var todos = lista.length && nSel === lista.length;
    $('#t-informes-tabla').innerHTML = '<thead><tr><th class="sel"><input type="checkbox" id="sel-todos" aria-label="Seleccionar todos los archivos"' +
      (todos ? ' checked' : '') + (lista.length ? '' : ' disabled') + '></th><th>Fecha de corte</th><th>Nombre</th><th>Tipo</th><th>Archivo</th><th>Cargado</th><th class="num">HU únicas</th><th></th></tr></thead><tbody>' +
      (lista.length ? lista.map(function (i) {
        var actual = i.id === estado.informeId;
        return '<tr' + (actual ? ' class="actual"' : '') + '>' +
          '<td class="sel"><input type="checkbox" data-inf-sel="' + esc(i.id) + '"' + (estado.seleccion[i.id] ? ' checked' : '') + ' aria-label="Seleccionar ' + esc(i.nombre) + '"></td>' +
          '<td><input type="date" data-inf-corte="' + esc(i.id) + '" value="' + esc(i.corte || '') + '" aria-label="Fecha de corte"></td>' +
          '<td><input type="text" data-inf-nombre="' + esc(i.id) + '" value="' + esc(i.nombre) + '" aria-label="Nombre del informe"></td>' +
          '<td><select data-inf-tipo="' + esc(i.id) + '" aria-label="Tipo de carga">' + Object.keys(TIPOS_CARGA).map(function (t) {
            return '<option value="' + t + '"' + ((i.tipo || 'gestion') === t ? ' selected' : '') + '>' + (t === 'gestion' ? 'Gestión' : 'Sprint') + '</option>';
          }).join('') + '</select></td>' +
          '<td>' + esc(i.archivo) + '</td><td>' + esc(new Date(i.cargado).toLocaleDateString('es-CO')) + '</td><td class="num">' + fmtNum.format(i.hu || 0) + '</td>' +
          '<td class="acciones-fila">' + (actual ? '<span class="chip">En pantalla</span>' : '<button class="btn-texto" type="button" data-inf-ver="' + esc(i.id) + '">Ver</button>') +
          '<button class="btn-texto peligro" type="button" data-inf-borrar="' + esc(i.id) + '">Eliminar</button></td></tr>';
      }).join('') : '<tr><td colspan="8">No hay archivos cargados. Usa Cargar otro Excel para agregar uno.</td></tr>') + '</tbody>';
    actualizarSeleccion();
  }

  // Sincroniza contador, botones y casilla general sin redibujar la tabla (conserva el foco)
  function actualizarSeleccion() {
    var lista = estado.informes;
    var nSel = Object.keys(estado.seleccion).length;
    $('#arch-seleccion').textContent = nSel ? nSel + (nSel === 1 ? ' archivo seleccionado' : ' archivos seleccionados') : 'Ningún archivo seleccionado';
    $('#btn-borrar-sel').disabled = !nSel;
    $('#btn-borrar-sel').textContent = nSel ? 'Eliminar seleccionados (' + nSel + ')' : 'Eliminar seleccionados';
    $('#btn-borrar-todos').disabled = !lista.length;
    var todos = $('#sel-todos');
    if (todos) {
      todos.checked = !!lista.length && nSel === lista.length;
      todos.indeterminate = nSel > 0 && nSel < lista.length;
    }
    $$('[data-inf-sel]').forEach(function (c) { c.checked = !!estado.seleccion[c.getAttribute('data-inf-sel')]; });
  }

  function cambiarInforme(id, cambios) {
    Historico.actualizar(id, cambios).then(refrescarInformes).then(function () {
      var inf = informeActual();
      if (inf && inf.id === id && cambios.nombre) estado.fuente = inf.nombre + ' — ' + (inf.hoja || 'Hoja1');
      actualizar();
    }).catch(function (err) { toast(err.message); });
  }

  function borrarInforme(id) { borrarInformes([id]); }

  /** Elimina uno o varios archivos cargados del histórico (pide confirmación). */
  function borrarInformes(ids) {
    var aBorrar = estado.informes.filter(function (i) { return ids.indexOf(i.id) !== -1; });
    if (!aBorrar.length) return;
    var texto = aBorrar.length === 1
      ? '¿Eliminar el archivo del ' + fechaCorte(aBorrar[0].corte) + ' (' + aBorrar[0].nombre + ')?'
      : '¿Eliminar ' + aBorrar.length + ' archivos cargados' + (aBorrar.length === estado.informes.length ? ' (todos)' : '') + '?';
    if (!window.confirm(texto + ' Se borran de este navegador y no se puede deshacer. Si los necesitas después, exporta el histórico antes.')) return;
    var borraActual = aBorrar.some(function (i) { return i.id === estado.informeId; });
    Promise.all(aBorrar.map(function (i) { delete estado.cache[i.id]; delete estado.seleccion[i.id]; return Historico.eliminar(i.id); }))
      .then(refrescarInformes).then(function (lista) {
        toast(aBorrar.length === 1 ? 'Archivo eliminado' : aBorrar.length + ' archivos eliminados');
        if (!borraActual) { actualizar(); return; }
        if (lista.length) { mostrarInforme(lista[0]); return; }
        estado.informeId = null;
        try { localStorage.removeItem(CLAVE + 'informeActual'); localStorage.removeItem(CLAVE + 'datos'); } catch (e) { /* */ }
        $('#app').hidden = true;
        $('#zona-carga').hidden = false;
      }).catch(function (err) { toast('No se pudo eliminar: ' + err.message); });
  }

  function exportarHistorico() {
    var datos = { tipo: 'historico-indicadores', version: 1, exportado: new Date().toISOString(), informes: estado.informes };
    var blob = new Blob([JSON.stringify(datos)], { type: 'application/json' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'historico-indicadores-' + hoyISO() + '.json';
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  function importarHistorico(file) {
    if (!file) return;
    var lector = new FileReader();
    lector.onload = function () {
      var datos;
      try { datos = JSON.parse(lector.result); } catch (e) { toast('El archivo no es un histórico válido (JSON).'); return; }
      var entrada = datos && datos.tipo === 'historico-indicadores' && Array.isArray(datos.informes) ? datos.informes : null;
      if (!entrada) { toast('El archivo no es un histórico exportado desde esta app.'); return; }
      var huellas = estado.informes.map(function (i) { return i.huella; });
      var nuevos = entrada.filter(function (i) {
        if (!i || !Array.isArray(i.matriz) || !i.id) return false;
        try { I.parsearHoja(i.matriz); } catch (e) { return false; }
        i.huella = i.huella || Historico.huella(i.matriz);
        return huellas.indexOf(i.huella) === -1;
      });
      Promise.all(nuevos.map(function (i) {
        return Historico.guardar({
          id: String(i.id), nombre: String(i.nombre || i.archivo || 'Informe importado'), archivo: String(i.archivo || ''),
          hoja: String(i.hoja || 'Hoja1'), cargado: i.cargado || new Date().toISOString(),
          corte: /^\d{4}-\d{2}-\d{2}$/.test(i.corte) ? i.corte : hoyISO(), huella: i.huella, tipo: i.tipo === 'sprint' ? 'sprint' : 'gestion',
          hu: Number(i.hu) || 0, matriz: i.matriz
        });
      })).then(refrescarInformes).then(function (lista) {
        toast(nuevos.length ? nuevos.length + ' informe(s) importado(s); ' + (entrada.length - nuevos.length) + ' ya estaban o no eran válidos.' : 'No había informes nuevos para importar.');
        if (!estado.informeId && lista.length) mostrarInforme(lista[0]); else actualizar();
      }).catch(function (err) { toast('No se pudo importar: ' + err.message); });
    };
    lector.readAsText(file);
  }

  // ---------- ficha ----------
  function renderFicha() {
    var k = estado.resultado.kpis;
    $('#t-ficha').innerHTML = '<thead><tr><th>N°</th><th>Indicador (KPI)</th><th>Definición</th><th>Fórmula de cálculo</th><th>Frecuencia</th><th>Meta</th><th>Responsable</th><th>Cálculo con la base</th><th class="num">Resultado actual</th><th>Estado</th></tr></thead><tbody>' +
      k.map(function (x, i) {
        var meta = textoMeta(x);
        return '<tr><td>' + (i + 1) + '</td><td><strong>' + esc(x.nombre) + '</strong></td><td>' + esc(x.definicion) + '</td><td>' + esc(x.formula) +
          '</td><td>' + esc(x.frecuencia) + '</td><td>' + meta + '</td><td>' + esc(x.responsable) + '</td><td>' + esc(x.criterio) +
          '</td><td class="num">' + valorKpi(x, x.resultado) + '</td><td>' + pill(x.estado) + '</td></tr>';
      }).join('') + '</tbody>';
  }

  // ---------- configuración ----------
  function renderConfig() {
    if (renderConfig.hecho) return;
    renderConfig.hecho = true;
    var metas = estado.config.metas, par = Object.assign({}, I.PARAMETROS, estado.config.parametros);
    $('#cfg-metas').innerHTML = '<div class="fila-cfg"><span class="cab">Indicador</span><span class="cab">Tipo</span><span class="cab">Meta (% o h/pt)</span></div>' +
      I.KPIS.map(function (k) {
        var m = metas[k.id] || {};
        var meta = m.meta !== undefined ? m.meta : k.meta;
        var tipo = m.tipoMeta !== undefined ? m.tipoMeta : k.tipoMeta;
        return '<div class="fila-cfg"><span>' + esc(k.nombre) + '</span>' +
          '<select data-meta-tipo="' + k.id + '" aria-label="Tipo de meta"><option value="">—</option><option' + (tipo === '<=' ? ' selected' : '') + '>&lt;=</option><option' + (tipo === '>=' ? ' selected' : '') + '>&gt;=</option></select>' +
          '<input type="number" min="0" max="1000" step="0.1" data-meta="' + k.id + '" value="' + (meta === null || meta === '' ? '' : +(meta * escalaKpi(k)).toFixed(2)) + '" aria-label="Meta' + (esHpt(k) ? ' en horas por punto' : ' en %') + '"></div>';
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
    $('#cfg-avance').value = estado.config.avance.acciones.map(function (x) { return x.accion + ' => ' + +(x.esfuerzo * 100).toFixed(2) + ' => ' + x.estados.join('; '); }).join('\n');
    $('#cfg-avance-actual').checked = !!estado.config.avance.incluirActual;
    var dyd = estado.config.dyd;
    $('#cfg-dyd-tipo').value = dyd.tipo;
    $('#cfg-dyd').value = dyd.acciones.map(function (x) { return x.accion + ' => ' + +(x.esfuerzo * 100).toFixed(2) + ' => ' + x.estados.join('; '); }).join('\n');
    $('#cfg-dyd-terminados').value = (dyd.terminados || []).join('; ');
    $('#cfg-equivalencias').value = estado.config.equivalencias.map(function (e) { return e[0] + ' => ' + e[1]; }).join('\n');
  }

  function leerMetas() {
    var metas = {};
    I.KPIS.forEach(function (k) {
      var v = $('[data-meta="' + k.id + '"]').value;
      var t = $('[data-meta-tipo="' + k.id + '"]').value;
      metas[k.id] = { meta: v === '' ? null : Number(v) / escalaKpi(k), tipoMeta: t };
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

  function guardarDyd() {
    var acciones = [], error = '';
    $('#cfg-dyd').value.split(/\r?\n/).forEach(function (l) {
      if (!l.trim()) return;
      var p = l.split('=>').map(function (x) { return x.trim(); });
      var n = Number(String(p[1] || '').replace('%', '').replace(',', '.'));
      if (p.length !== 3 || !p[0] || isNaN(n)) { error = error || l; return; }
      acciones.push({ accion: p[0], esfuerzo: n / 100, estados: p[2].split(';').map(function (x) { return x.trim(); }).filter(Boolean) });
    });
    if (error) { toast('Línea no válida: "' + error + '". Formato: Acción => % => estados separados por ;'); return; }
    if (!acciones.length) { toast('Agrega al menos una acción.'); return; }
    var tipo = $('#cfg-dyd-tipo').value.trim();
    if (!tipo) { toast('Indica el tipo de paquete (por ejemplo Diseño y Desarrollo).'); return; }
    estado.config.dyd = {
      tipo: tipo, acciones: acciones, incluirActual: false,
      terminados: $('#cfg-dyd-terminados').value.split(';').map(function (x) { return x.trim(); }).filter(Boolean)
    };
    guardar('dyd', estado.config.dyd);
    actualizar();
    toast('Avance de Diseño y Desarrollo aplicado');
  }

  function guardarAvance() {
    var acciones = [], error = '';
    $('#cfg-avance').value.split(/\r?\n/).forEach(function (l) {
      if (!l.trim()) return;
      var p = l.split('=>').map(function (x) { return x.trim(); });
      var n = Number(String(p[1] || '').replace('%', '').replace(',', '.'));
      if (p.length !== 3 || !p[0] || isNaN(n)) { error = error || l; return; }
      acciones.push({ accion: p[0], esfuerzo: n / 100, estados: p[2].split(';').map(function (x) { return x.trim(); }).filter(Boolean) });
    });
    if (error) { toast('Línea no válida: "' + error + '". Formato: Acción => % => estados separados por ;'); return; }
    if (!acciones.length) { toast('Agrega al menos una acción.'); return; }
    estado.config.avance = { acciones: acciones, incluirActual: $('#cfg-avance-actual').checked };
    guardar('avance', estado.config.avance);
    actualizar();
    var total = acciones.reduce(function (s, a) { return s + a.esfuerzo; }, 0);
    toast(Math.abs(total - 1) > 0.0001 ? 'Avance aplicado, pero los porcentajes suman ' + pct(total, 0) + ' en lugar de 100%.' : 'Avance por estado aplicado');
  }

  function guardarEquivalencias() {
    var reglas = leerReglas('#cfg-equivalencias');
    var conFase = estado.config.estados.map(function (e) { return e[0]; });
    var malas = reglas.filter(function (r) { return conFase.indexOf(r[1]) === -1; });
    if (malas.length) { toast('El estado destino debe estar en la tabla de fases: ' + malas.map(function (r) { return r[1]; }).join(', ') + '.'); return; }
    estado.config.equivalencias = reglas;
    guardar('equivalencias', reglas);
    reparsear();
    toast(reglas.length ? reglas.length + ' equivalencia(s) aplicada(s)' : 'Equivalencias eliminadas');
  }

  // ---------- tema ----------
  function aplicarTema(t) {
    if (t) document.documentElement.setAttribute('data-theme', t);
    else document.documentElement.removeAttribute('data-theme');
    if (estado.resultado) { renderDistribucion(); renderSeguimiento(); renderComparativo(); renderSprint(); }
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
    $('#btn-guardar-equivalencias').addEventListener('click', guardarEquivalencias);
    $('#btn-guardar-avance').addEventListener('click', guardarAvance);
    $('#btn-guardar-dyd').addEventListener('click', guardarDyd);
    $('#sp-sprint').addEventListener('change', function (e) { estado.sprint = e.target.value; guardar('sprint', estado.sprint); renderSprint(); });
    $('#t-sp-sprints-tabla').addEventListener('click', function (e) {
      var b = e.target.closest('[data-sp]'); if (!b) return;
      estado.sprint = b.getAttribute('data-sp'); guardar('sprint', estado.sprint); renderSprint();
    });
    $('#sp-buscar').addEventListener('input', renderSprintId);
    $('#sp-solo-pendientes').addEventListener('change', renderSprintId);
    $('#av-agrupar').addEventListener('click', function (e) {
      var b = e.target.closest('[data-av-grupo]'); if (!b) return;
      estado.avGrupo = b.getAttribute('data-av-grupo'); guardar('avGrupo', estado.avGrupo); renderAvance();
    });
    $('#av-buscar').addEventListener('input', renderAvanceHU);
    $('#btn-restaurar').addEventListener('click', function () {
      estado.config = { metas: {}, parametros: {}, estados: I.ESTADOS.slice(), equivalencias: I.EQUIVALENCIAS.slice(), avance: JSON.parse(JSON.stringify(I.AVANCE)), dyd: JSON.parse(JSON.stringify(I.AVANCE_DYD)) };
      ['metas', 'parametros', 'estados', 'equivalencias', 'avance', 'dyd', 'manuales'].forEach(function (k) { try { localStorage.removeItem(CLAVE + k); } catch (e) { /* */ } });
      renderConfig.hecho = false;
      reparsear();
      toast('Configuración restaurada');
    });

    $('#sel-informe').addEventListener('change', function (e) {
      var inf = estado.informes.find(function (i) { return i.id === e.target.value; });
      if (inf) mostrarInforme(inf);
    });
    $('#lista-cortes').addEventListener('click', function (e) {
      var bt = e.target.closest('[data-corte]'); if (!bt) return;
      var inf = estado.informes.find(function (i) { return i.id === bt.getAttribute('data-corte'); });
      if (inf && inf.id !== estado.informeId) mostrarInforme(inf);
    });
    var irArchivos = function () { $('.tabs [data-tab="archivos"]').click(); window.scrollTo(0, 0); };
    $('#btn-ir-archivos').addEventListener('click', irArchivos);
    $('[data-ir-archivos]').addEventListener('click', irArchivos);
    $('#btn-borrar-sel').addEventListener('click', function () { borrarInformes(Object.keys(estado.seleccion)); });
    $('#btn-borrar-todos').addEventListener('click', function () { borrarInformes(estado.informes.map(function (i) { return i.id; })); });
    $('#btn-ir-comparar').addEventListener('click', function () { $('.tabs [data-tab="comparativo"]').click(); window.scrollTo(0, 0); });
    $('#btn-ultimo').addEventListener('click', function () { if (estado.informes[0]) mostrarInforme(estado.informes[0]); });
    $('#comp-a').addEventListener('change', function (e) { estado.comp.a = e.target.value; renderComparativo(); });
    $('#comp-b').addEventListener('change', function (e) { estado.comp.b = e.target.value; renderComparativo(); });
    $('#comp-invertir').addEventListener('click', function () {
      var a = estado.comp.a; estado.comp.a = estado.comp.b; estado.comp.b = a; renderComparativo();
    });
    $('#comp-kpi-evol').addEventListener('change', function (e) { estado.comp.kpi = e.target.value; guardar('compKpi', e.target.value); renderComparativo(); });
    $('#comp-tipos').addEventListener('click', function (e) {
      var b = e.target.closest('[data-cambio]'); if (!b) return;
      estado.comp.tipo = b.getAttribute('data-cambio'); renderComparativo();
    });
    $('#t-informes-tabla').addEventListener('change', function (e) {
      var t = e.target;
      if (t.id === 'sel-todos') {
        estado.seleccion = {};
        if (t.checked) estado.informes.forEach(function (i) { estado.seleccion[i.id] = true; });
        actualizarSeleccion(); return;
      }
      if (t.hasAttribute('data-inf-sel')) {
        if (t.checked) estado.seleccion[t.getAttribute('data-inf-sel')] = true; else delete estado.seleccion[t.getAttribute('data-inf-sel')];
        actualizarSeleccion(); return;
      }
      if (t.hasAttribute('data-inf-corte') && t.value) cambiarInforme(t.getAttribute('data-inf-corte'), { corte: t.value });
      if (t.hasAttribute('data-inf-nombre') && t.value.trim()) cambiarInforme(t.getAttribute('data-inf-nombre'), { nombre: t.value.trim() });
      if (t.hasAttribute('data-inf-tipo')) cambiarInforme(t.getAttribute('data-inf-tipo'), { tipo: t.value });
    });
    $('#t-informes-tabla').addEventListener('click', function (e) {
      var ver = e.target.closest('[data-inf-ver]'), borrar = e.target.closest('[data-inf-borrar]');
      if (ver) { var inf = estado.informes.find(function (i) { return i.id === ver.getAttribute('data-inf-ver'); }); if (inf) mostrarInforme(inf); }
      if (borrar) borrarInforme(borrar.getAttribute('data-inf-borrar'));
    });
    $('#btn-exportar-hist').addEventListener('click', exportarHistorico);
    $('#importar-hist').addEventListener('change', function (e) { importarHistorico(e.target.files[0]); e.target.value = ''; });

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
    try { localStorage.removeItem(CLAVE + 'palabras'); } catch (e) { /* almacenamiento no disponible */ }
    if (document.fonts && document.fonts.ready) {
      document.fonts.ready.then(function () { if (estado.resultado) { renderDistribucion(); renderSeguimiento(); } });
    }
    var tema = leer('tema', null);
    if (tema) document.documentElement.setAttribute('data-theme', tema);
    enlazar();
    var tab = leer('tab', null);
    if (tab && $('.tabs [data-tab="' + tab + '"]')) $('.tabs [data-tab="' + tab + '"]').click();
    Historico.persistente().then(function (ok) {
      estado.historicoPersistente = ok;
      return refrescarInformes();
    }).then(function (lista) {
      var previo = leer('datos', null);
      // Migración: la última carga de versiones anteriores pasa a ser el primer informe del histórico
      if (!lista.length && previo && previo.matriz) {
        var inf = {
          id: 'inf-' + Date.now().toString(36), nombre: String(previo.fuente || 'Informe').split(' — ')[0].replace(/\.(xlsx|xls|csv)$/i, ''),
          archivo: String(previo.fuente || '').split(' — ')[0], hoja: 'Hoja1', cargado: previo.cargado || new Date().toISOString(),
          corte: hoyISO(new Date(previo.cargado || Date.now())), huella: Historico.huella(previo.matriz),
          hu: 0, matriz: previo.matriz
        };
        try { inf.hu = I.parsearHoja(inf.matriz).registros.length; } catch (e) { return; }
        return Historico.guardar(inf).then(refrescarInformes).then(function (l) {
          if (estado.historicoPersistente) { try { localStorage.removeItem(CLAVE + 'datos'); } catch (e) { /* */ } }
          mostrarInforme(l[0]);
        });
      }
      if (!lista.length) return;
      var id = leer('informeActual', null);
      mostrarInforme(lista.find(function (i) { return i.id === id; }) || lista[0]);
    }).catch(function () { /* histórico no disponible */ });
  }

  iniciar();
})();
