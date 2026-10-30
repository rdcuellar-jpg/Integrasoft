// Verifica que el motor de cálculo reproduzca los valores del libro Excel
// "Base para Indicadores.xlsx" con la Hoja1 de ejemplo.
// Ejecutar: node --test indicadores/tests/*.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const I = require('../js/indicadores.js');
const ejemplo = require('../data/hoja1-ejemplo.json');
const esperado = require('./esperado-seguimiento.json');

const { registros, filasLeidas, columnas } = I.parsearHoja(ejemplo.matriz);
const config = { parametros: { anio: 2026 } };
const r = I.calcular(registros, config);
const kpi = (id, res = r) => res.kpis.find((k) => k.id === id);

test('base de cálculo (Datos Indicadores!B5:B9)', () => {
  assert.equal(filasLeidas, 294);
  assert.ok(columnas.precision && columnas.trabajoTotal);
  assert.equal(r.base.total, 168);
  assert.deepEqual(r.base.excluidos, [{ estado: 'Nuevo / Identificado', n: 46 }, { estado: 'Asignación', n: 0 }]);
  assert.equal(r.base.sinMapear, 0);
  assert.equal(r.base.consideradas, 122);
});

test('indicadores por estado, fase y En flujo', () => {
  const n = Object.fromEntries(r.porEstado.map((e) => [e.estado, e.n]));
  assert.deepEqual(n, {
    'Levantamiento Requerimiento': 8, 'especificaciones Funcionales': 1, 'Aprobación cliente': 0,
    'Aprobacion PO': 1, 'Backlog': 1, 'En desarrollo': 15, 'Revisión (IDS)': 0,
    'Pruebas funcionales y Documentación': 9, 'Entrega a Operaciones': 29, 'Entrega': 0,
    'CIERRE en Desarrollo o Solución': 57, 'Cierre': 1, 'Garantía HU': 0
  });
  assert.deepEqual(r.porFase.map((f) => f.n), [11, 15, 9, 87]);
  assert.deepEqual(r.porFlujo.map((f) => f.n), [35, 87]);
  assert.deepEqual(r.estadosNoListados, {});
});

test('tipo de solicitud', () => {
  const n = Object.fromEntries(r.porTipo.map((t) => [t.nombre, t.n]));
  assert.deepEqual(n, { 'Mejora Funcional': 70, 'Deuda Tecnica': 18, 'Desarrollo Contratado': 32, 'Garantía': 2, 'Sin clasificar': 0 });
});

test('KPIs (Datos Indicadores!A55:G64)', () => {
  const nd = (id) => [kpi(id).numerador, kpi(id).denominador, kpi(id).estado];
  assert.deepEqual(nd('deuda'), [18, 122, 'EN META']);
  assert.deepEqual(nd('backlog'), [79, 104, 'FUERA DE META']);
  assert.deepEqual(nd('horas'), [271, null, 'SIN DATOS']);
  // Precisión del refinement: sin marca en Precisión HU / todas las HU consideradas
  assert.deepEqual(nd('refinement'), [115, 122, 'EN META']);
  assert.deepEqual(nd('alcance'), [3, 122, 'EN META']);
  assert.deepEqual(nd('ambiguedad'), [4, 122, 'EN META']);
  assert.deepEqual(nd('analisis'), [11, 122, 'SIN META']);
  assert.deepEqual(nd('desarrollo'), [15, 122, 'SIN META']);
  assert.deepEqual(nd('pruebas'), [9, 122, 'SIN META']);
  assert.deepEqual(nd('abiertas'), [35, 122, 'SIN META']);
  assert.equal(Math.round(kpi('backlog').resultado * 1000) / 10, 76);
  assert.equal(Math.round(kpi('refinement').resultado * 1000) / 10, 94.3);
});

test('control de horas con horas contratadas ingresadas', () => {
  const r2 = I.calcular(registros, { parametros: { horasContratadas: 300 } });
  assert.equal(Math.round(kpi('horas', r2).resultado * 1000) / 1000, 0.903);
  assert.equal(kpi('horas', r2).estado, 'EN META');
  const r3 = I.calcular(registros, { parametros: { horasContratadas: 200 } });
  assert.equal(kpi('horas', r3).estado, 'FUERA DE META');
});

test('HU por mes de finalización 2026 (Datos Indicadores!A68:C80)', () => {
  assert.equal(r.mensual.length, 12);
  assert.deepEqual(r.mensual.map((m) => m.cerradas), [1, 0, 2, 0, 2, 2, 12, 1, 13, 12, 0, 0]);
  assert.deepEqual(r.mensual.map((m) => m.abiertas), [0, 0, 0, 1, 1, 3, 1, 0, 0, 22, 2, 0]);
  assert.deepEqual(r.sinFecha, { cerradas: 42, abiertas: 5 });
});

// La hoja Indicadores Seguimiento del Excel tiene columnas fijas (sin Garantía) y una lista
// fija de sprints, por eso se comparan solo las filas y columnas que el Excel sí trae.
function comparar(tabla, esperadoTabla) {
  const filas = Object.fromEntries(tabla.map((f) => [f.nombre, f]));
  for (const [nombre, valores] of Object.entries(esperadoTabla)) {
    const f = filas[nombre];
    if (valores[4] === 0) { assert.ok(!f || f.total === f['Garantía'], nombre); continue; }
    assert.ok(f, `falta ${nombre}`);
    assert.deepEqual(
      [f['Desarrollo Contratado'], f['Mejora Funcional'], f['Deuda Tecnica'], f['Sin clasificar'], f.total - f['Garantía']],
      valores, nombre
    );
  }
}

test('seguimiento HU abiertas por cliente y sprint', () => {
  assert.equal(r.seguimiento.total, 35);
  comparar(r.seguimiento.porCliente, esperado.cliente);
  comparar(r.seguimiento.porSprint, esperado.sprint);
  // El módulo ya no se calcula por palabra clave: el ejemplo no trae la columna Módulo afectado
  assert.ok(r.seguimiento.porModulo.every((f) => f.nombre === I.SIN_MODULO));
  // Lo que el Excel no muestra: 1 HU abierta de Garantía y el SPRINT 8 UNIFICADO
  assert.equal(r.seguimiento.porCliente.reduce((s, f) => s + f['Garantía'], 0), 1);
  assert.equal(r.seguimiento.porSprint.find((f) => f.nombre === 'SPRINT 8 UNIFICADO').total, 19);
  assert.equal(r.seguimiento.porCliente.reduce((s, f) => s + f.total, 0), 35);
});

test('estados no mapeados quedan fuera de las HU consideradas', () => {
  const m = ejemplo.matriz.map((f) => f.slice());
  const fila = m.findIndex((f, i) => i > 1 && f[7] === 'Backlog');
  m.forEach((f, i) => { if (i > 1 && f[1] === m[fila][1]) f[7] = 'Estado nuevo'; });
  const r2 = I.calcular(I.parsearHoja(m).registros);
  assert.equal(r2.base.consideradas, r.base.consideradas - 1);
  assert.deepEqual(r2.estadosNoListados, { 'Estado nuevo': 1 });
});

test('coincideCriterio replica COUNTIFS (comodines, mayúsculas)', () => {
  assert.ok(I.coincideCriterio('Retrabajo por ambigüedad', 'Retrabajo*'));
  assert.ok(I.coincideCriterio('cambio de alcance', 'Cambio de Alcance'));
  assert.ok(!I.coincideCriterio('Cambio de Alcance parcial', 'Cambio de Alcance'));
});

test('parseFecha admite texto, Date y serial de Excel', () => {
  assert.equal(I.parseFecha('22-05-2026').getMonth(), 4);
  assert.equal(I.parseFecha(46164).getFullYear(), 2026);
  assert.equal(I.parseFecha(''), null);
});

test('compararHU detecta HU nuevas, cerradas, reabiertas, cambios de estado y retiradas', () => {
  const m = ejemplo.matriz.map((f) => f.slice());
  const col = { id: 1, estado: 7 };
  const ids = (estado) => [...new Set(m.slice(2).filter((f) => f[col.estado] === estado).map((f) => String(f[col.id])))];
  const cambiar = (id, estado) => m.forEach((f, i) => { if (i > 1 && String(f[col.id]) === id) f[col.estado] = estado; });
  const abierta = ids('En desarrollo')[0];
  const cerrada = ids('Entrega a Operaciones')[0];
  const enAnalisis = ids('Levantamiento Requerimiento')[0];
  const retirada = ids('Pruebas funcionales y Documentación')[0];
  cambiar(abierta, 'CIERRE en Desarrollo o Solución');
  cambiar(cerrada, 'En desarrollo');
  cambiar(enAnalisis, 'En desarrollo');
  cambiar(retirada, 'Nuevo / Identificado');
  const nueva = m[2].slice(); nueva[col.id] = '999999'; nueva[col.estado] = 'Backlog'; m.push(nueva);

  const B = I.parsearHoja(m).registros;
  const { resumen, cambios } = I.compararHU(registros, B);
  assert.deepEqual(resumen, { nueva: 1, cerrada: 1, reabierta: 1, estado: 1, retirada: 1 });
  assert.equal(cambios.find((c) => c.tipo === 'cerrada').id, abierta);
  assert.equal(cambios.find((c) => c.tipo === 'reabierta').id, cerrada);
  assert.deepEqual(cambios.find((c) => c.tipo === 'estado'), {
    tipo: 'estado', id: enAnalisis, asunto: cambios.find((c) => c.tipo === 'estado').asunto,
    cliente: cambios.find((c) => c.tipo === 'estado').cliente,
    de: 'Levantamiento Requerimiento', a: 'En desarrollo', faseDe: 'Análisis', faseA: 'Desarrollo'
  });
  assert.equal(cambios.find((c) => c.tipo === 'retirada').id, retirada);
  assert.equal(cambios[0].tipo, 'nueva');
  assert.deepEqual(I.compararHU(registros, registros).resumen, { nueva: 0, cerrada: 0, reabierta: 0, estado: 0, retirada: 0 });
});

test('las HU en Entrega a Clientes se suman a Entrega a Operaciones', () => {
  const m = ejemplo.matriz.map((f) => f.slice());
  const ids = [...new Set(m.slice(2).filter((f) => f[7] === 'En desarrollo').map((f) => String(f[1])))].slice(0, 2);
  m.forEach((f, i) => {
    if (i > 1 && String(f[1]) === ids[0]) f[7] = 'Entrega a Clientes';
    if (i > 1 && String(f[1]) === ids[1]) f[7] = 'ENTREGA CLIENTES ';
  });
  const regs = I.parsearHoja(m).registros;
  const h = regs.find((x) => x.id === ids[0]);
  assert.equal(h.estado, 'Entrega a Operaciones');
  assert.equal(h.estadoOriginal, 'Entrega a Clientes');
  assert.equal(h.fase, 'Cerrado');
  assert.equal(h.enFlujo, 'Cerrado');
  const r2 = I.calcular(regs);
  const n = (res) => res.porEstado.find((e) => e.estado === 'Entrega a Operaciones').n;
  assert.equal(n(r2), n(r) + 2);
  assert.equal(r2.base.consideradas, r.base.consideradas);
  assert.deepEqual(r2.estadosNoListados, {});
  // Sin la equivalencia, el estado queda sin fase y fuera de las consideradas
  const r3 = I.calcular(I.parsearHoja(m, { equivalencias: [] }).registros);
  assert.equal(r3.base.consideradas, r.base.consideradas - 2);
});

test('avance por estado: suma las acciones anteriores; cerradas al 100%; otros estados 0%', () => {
  const tabla = I.tablaAvance();
  assert.deepEqual(tabla.filas.map((f) => Math.round(f.avance * 100)), [0, 10, 25, 30, 75]);
  assert.equal(Math.round(tabla.total * 100), 100);
  const av = (estado, fase = 'Análisis') => Math.round(I.avanceHU({ estado, fase }, tabla).avance * 100);
  assert.equal(av('Levantamiento Requerimiento'), 0);
  assert.equal(av('especificaciones Funcionales'), 10);
  assert.equal(av('Aprobación cliente'), 25);
  assert.equal(av('En desarrollo', 'Desarrollo'), 30);
  assert.equal(av('Pruebas funcionales y Documentación', 'Pruebas'), 75);
  assert.equal(av('Entrega a Operaciones', 'Cerrado'), 100);
  assert.equal(av('CIERRE en Desarrollo o Solución', 'Cerrado'), 100);
  assert.equal(av('Backlog'), 0);
  assert.equal(av('Revisión (IDS)', 'Desarrollo'), 0);
  // Con la opción de incluir el estado actual
  const t2 = I.tablaAvance(Object.assign({}, I.AVANCE, { incluirActual: true }));
  assert.deepEqual(t2.filas.map((f) => Math.round(f.avance * 100)), [10, 25, 30, 75, 100]);

  const a = I.calcularAvance(registros);
  assert.equal(a.n, 122);
  // 87 cerradas al 100% + 1 espec. 10% + 15 desarrollo 30% + 9 pruebas 75%
  const esperado = (87 * 1 + 1 * 0.10 + 15 * 0.30 + 9 * 0.75) / 122;
  assert.ok(Math.abs(a.promedio - esperado) < 1e-9);
  assert.equal(a.ejecutadas, 87);
  assert.equal(a.enCurso, 25);
  assert.equal(a.sinAvance, 10);
  assert.equal(a.tramos.reduce((s, t) => s + t.n, 0), 122);
  assert.equal(a.porCliente.reduce((s, g) => s + g.n, 0), 122);
});

test('avance: Entrega cliente cuenta como ejecutada', () => {
  const m = ejemplo.matriz.map((f) => f.slice());
  const id = String(m.slice(2).find((f) => f[7] === 'En desarrollo')[1]);
  m.forEach((f, i) => { if (i > 1 && String(f[1]) === id) f[7] = 'Entrega cliente'; });
  const a = I.calcularAvance(I.parsearHoja(m).registros);
  assert.equal(a.lista.find((x) => x.hu.id === id).avance, 1);
  assert.equal(a.ejecutadas, 88);
});

test('el módulo se toma de la columna Módulo afectado', () => {
  const m = ejemplo.matriz.map((f) => f.slice());
  m[1].push('Módulo afectado');
  const ids = [...new Set(m.slice(2).map((f) => String(f[1])))];
  const modulos = ['Nómina', 'Contabilidad', ''];
  m.forEach((f, i) => { if (i > 1) f.push(modulos[ids.indexOf(String(f[1])) % 3]); });
  const p = I.parsearHoja(m);
  assert.ok(p.columnas.modulo);
  const regs = p.registros;
  assert.equal(regs.find((h) => h.id === ids[0]).modulo, 'Nómina');
  assert.equal(regs.find((h) => h.id === ids[1]).modulo, 'Contabilidad');
  assert.equal(regs.find((h) => h.id === ids[2]).modulo, I.SIN_MODULO);
  const r2 = I.calcular(regs);
  const total = r2.seguimiento.porModulo.reduce((s, f) => s + f.total, 0);
  assert.equal(total, r2.seguimiento.total);
  assert.equal(r2.seguimiento.porModulo[r2.seguimiento.porModulo.length - 1].nombre, I.SIN_MODULO);
  // Encabezado sin tilde y en mayúsculas también se reconoce
  const m2 = m.map((f) => f.slice()); m2[1][m2[1].length - 1] = 'MODULO AFECTADO';
  assert.equal(I.parsearHoja(m2).registros.find((h) => h.id === ids[0]).modulo, 'Nómina');
});

test('formato InteTrack: usa Módulo Afectado de la HU y no el del bloque Relaciones', () => {
  const m = ejemplo.matriz.map((f) => f.slice());
  // Descarga de InteTrack: "Módulo Afectado" tras Precisión HU y repetido al final en Relaciones
  m[1].push('Módulo Afectado', 'Módulo Afectado');
  const ids = [...new Set(m.slice(2).map((f) => String(f[1])))];
  m.forEach((f, i) => { if (i > 1) f.push(String(f[1]) === ids[0] ? '' : 'Nómina', 'Contabilidad'); });
  const regs = I.parsearHoja(m).registros;
  assert.equal(regs.find((h) => h.id === ids[0]).modulo, I.SIN_MODULO);
  assert.equal(regs.find((h) => h.id === ids[1]).modulo, 'Nómina');
});

// Matriz pequeña con el formato InteTrack de octubre 2026: HU + bloque Relaciones
const ENC = ['', 'ID', 'Asunto', 'Sprint', 'Tipo', 'Tipo de Solicitud', 'Cliente', 'Estado', 'Trabajo', 'Trabajo total', 'Asignado a',
  'Fecha de inicio', 'Fecha de finalización', 'Precisión HU', 'Módulo Afectado', 'Puntos de Historia', 'Tiempo invertido', '',
  'Tipo de relación', 'Retraso', 'Descripción', 'ID', 'Asunto', 'Sprint', 'Tipo', 'Tipo de Solicitud', 'Cliente', 'Estado', 'Trabajo',
  'Trabajo total', 'Asignado a', 'Fecha de inicio', 'Fecha de finalización', 'Precisión HU', 'Módulo Afectado', 'Puntos de Historia', 'Tiempo invertido', ''];
function fila(hu, rel) {
  const f = new Array(ENC.length).fill('');
  [f[1], f[3], f[7], f[14], f[15], f[16]] = [hu.id, hu.sprint || '', hu.estado, hu.modulo || '', hu.puntos || '', hu.tiempo || 0];
  f[4] = 'Historia de Usuario'; f[6] = 'Cliente A';
  if (rel) [f[18], f[21], f[23], f[24], f[27], f[32], f[34], f[35], f[36]] = [rel.relacion || 'elemento principal de', rel.id, rel.sprint || '', rel.tipo, rel.estado, rel.fin || '', 'Otro módulo', rel.puntos || '', 99];
  return f;
}
function hoja(filas) { return [['Paquetes de trabajo'], ENC].concat(filas); }
const DYD = 'Diseño y Desarrollo';
const base = (sprintP, estados) => hoja([
  fila({ id: '1', estado: 'En desarrollo', puntos: '2', tiempo: 10, modulo: 'Nómina' }, { id: '101', tipo: 'Épica', estado: 'En ejecución', relacion: 'elemento secundario de' }),
  fila({ id: '1', estado: 'En desarrollo', puntos: '2', tiempo: 10, modulo: 'Nómina' }, { id: '201', tipo: DYD, estado: estados[0], sprint: sprintP[0], puntos: '3', fin: '06-11-2026' }),
  fila({ id: '1', estado: 'En desarrollo', puntos: '2', tiempo: 10, modulo: 'Nómina' }, { id: '202', tipo: DYD, estado: estados[1], sprint: sprintP[1], puntos: '5' }),
  fila({ id: '2', estado: 'Entrega a Operaciones', puntos: '3', tiempo: 26 }, { id: '203', tipo: DYD, estado: estados[2], sprint: sprintP[2], puntos: '2' }),
  fila({ id: '3', estado: 'Levantamiento Requerimiento' })
]);

test('efectividad: tiempo invertido / puntos de historia de las HU consideradas', () => {
  const p = I.parsearHoja(base(['S1', 'S1', ''], ['En desarrollo', 'Pruebas funcionales y Documentación', 'Cierre DyD']));
  assert.ok(p.columnas.puntos && p.columnas.tiempo && p.columnas.relaciones);
  const hu1 = p.registros.find((h) => h.id === '1');
  assert.equal(hu1.puntos, 2); assert.equal(hu1.tiempo, 10); assert.equal(hu1.modulo, 'Nómina');
  const k = I.calcular(p.registros).kpis.find((x) => x.id === 'efectividad');
  assert.equal(k.numerador, 36); assert.equal(k.denominador, 5);
  assert.ok(Math.abs(k.resultado - 7.2) < 1e-9);
  assert.equal(k.estado, 'EN META'); // meta <= 8 h por punto
});

test('avance del sprint y por ID: solo Desarrollo y Pruebas de los paquetes Diseño y Desarrollo', () => {
  const p = I.parsearHoja(base(['S1', 'S1', ''], ['En desarrollo', 'Pruebas funcionales y Documentación', 'Cierre DyD']));
  const s = I.calcularSprint(p.registros);
  assert.equal(s.paquetes.length, 3); // la Épica no cuenta
  const av = Object.fromEntries(s.paquetes.map((x) => [x.paquete.id, x.avance]));
  assert.equal(av['201'], 0);                              // En desarrollo: aún no termina Desarrollo
  assert.ok(Math.abs(av['202'] - 45 / 70) < 1e-9);        // Pruebas: Desarrollo hecho, escalado a 100%
  assert.equal(av['203'], 1);                              // Cierre DyD: terminado
  assert.equal(s.porSprint.length, 1);
  assert.equal(s.porSprint[0].nombre, 'S1');
  assert.ok(Math.abs(s.porSprint[0].avance - (45 / 70) / 2) < 1e-9);
  assert.equal(s.sinSprint.n, 1);                          // el paquete sin sprint no entra al avance del sprint
  const id1 = s.porHU.find((x) => x.hu.id === '1');
  assert.ok(Math.abs(id1.avance - (45 / 70) / 2) < 1e-9);
  assert.equal(s.porHU.find((x) => x.hu.id === '2').avance, 1);
  assert.equal(s.porHU.find((x) => x.hu.id === '3'), undefined);
});

test('burndown: puntos pendientes por corte e ID que pasan a otro sprint', () => {
  const c1 = I.parsearHoja(base(['S1', 'S1', 'S1'], ['En desarrollo', 'En desarrollo', 'Pruebas funcionales y Documentación'])).registros;
  const c2 = I.parsearHoja(base(['S1', 'S2', 'S1'], ['Pruebas funcionales y Documentación', 'En desarrollo', 'Cierre DyD'])).registros;
  const b = I.calcularBurndown([{ corte: '2026-11-02', registros: c2 }, { corte: '2026-10-26', registros: c1 }], 'S1');
  assert.deepEqual(b.puntos.map((x) => [x.corte, x.pendientes]), [['2026-10-26', 10], ['2026-11-02', 3]]);
  assert.equal(b.fin, '2026-11-06');
  assert.equal(b.puntos[0].ideal, 10);
  assert.ok(Math.abs(b.puntos[1].ideal - 10 * 4 / 11) < 1e-9);
  assert.deepEqual(b.movimientos.map((m) => [m.tipo, m.id, m.de, m.a, m.corte]), [['salio', '202', 'S1', 'S2', '2026-11-02']]);
});
