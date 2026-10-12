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
  assert.equal(filasLeidas, 290);
  assert.ok(columnas.precision && columnas.trabajoTotal);
  assert.equal(r.base.total, 173);
  assert.deepEqual(r.base.excluidos, [{ estado: 'Nuevo / Identificado', n: 20 }, { estado: 'Asignación', n: 2 }]);
  assert.equal(r.base.sinMapear, 0);
  assert.equal(r.base.consideradas, 151);
});

test('indicadores por estado, fase y En flujo', () => {
  const n = Object.fromEntries(r.porEstado.map((e) => [e.estado, e.n]));
  assert.deepEqual(n, {
    'Levantamiento Requerimiento': 41, 'especificaciones Funcionales': 2, 'Aprobación cliente': 1,
    'Aprobacion PO': 2, 'Backlog': 1, 'En desarrollo': 13, 'Revisión (IDS)': 3,
    'Pruebas funcionales y Documentación': 7, 'Entrega a Operaciones': 20, 'Entrega': 1,
    'CIERRE en Desarrollo o Solución': 58, 'Cierre': 1, 'Garantía HU': 1
  });
  assert.deepEqual(r.porFase.map((f) => f.n), [47, 16, 7, 81]);
  assert.deepEqual(r.porFlujo.map((f) => f.n), [70, 81]);
  assert.deepEqual(r.estadosNoListados, {});
});

test('tipo de solicitud', () => {
  const n = Object.fromEntries(r.porTipo.map((t) => [t.nombre, t.n]));
  assert.deepEqual(n, { 'Mejora Funcional': 36, 'Deuda Tecnica': 46, 'Desarrollo Contratado': 17, 'Sin clasificar': 52 });
});

test('KPIs (Datos Indicadores!A54:G63)', () => {
  const nd = (id) => [kpi(id).numerador, kpi(id).denominador, kpi(id).estado];
  assert.deepEqual(nd('deuda'), [46, 151, 'FUERA DE META']);
  assert.deepEqual(nd('backlog'), [78, 105, 'FUERA DE META']);
  assert.deepEqual(nd('horas'), [267, null, 'SIN DATOS']);
  assert.deepEqual(nd('refinement'), [40, 42, 'EN META']);
  assert.deepEqual(nd('alcance'), [2, 151, 'EN META']);
  assert.deepEqual(nd('ambiguedad'), [0, 151, 'EN META']);
  assert.deepEqual(nd('analisis'), [47, 151, 'SIN META']);
  assert.deepEqual(nd('desarrollo'), [16, 151, 'SIN META']);
  assert.deepEqual(nd('pruebas'), [7, 151, 'SIN META']);
  assert.deepEqual(nd('abiertas'), [70, 151, 'SIN META']);
  assert.equal(Math.round(kpi('backlog').resultado * 1000) / 10, 74.3);
});

test('control de horas con horas contratadas ingresadas', () => {
  const r2 = I.calcular(registros, { parametros: { horasContratadas: 300 } });
  assert.equal(kpi('horas', r2).resultado, 0.89);
  assert.equal(kpi('horas', r2).estado, 'EN META');
  const r3 = I.calcular(registros, { parametros: { horasContratadas: 200 } });
  assert.equal(kpi('horas', r3).estado, 'FUERA DE META');
});

test('HU por mes de finalización 2026 (Datos Indicadores!A67:C79)', () => {
  assert.equal(r.mensual.length, 12);
  assert.deepEqual(r.mensual.map((m) => m.cerradas), [1, 0, 2, 0, 1, 2, 11, 1, 13, 4, 0, 0]);
  assert.deepEqual(r.mensual.map((m) => m.abiertas), [0, 0, 0, 3, 3, 4, 2, 0, 0, 21, 1, 3]);
  assert.deepEqual(r.sinFecha, { cerradas: 46, abiertas: 33 });
});

function comparar(tabla, esperadoTabla) {
  const filas = Object.fromEntries(tabla.map((f) => [f.nombre, f]));
  for (const [nombre, valores] of Object.entries(esperadoTabla)) {
    if (valores[4] === 0) { assert.ok(!filas[nombre], nombre); continue; }
    const f = filas[nombre];
    assert.ok(f, `falta ${nombre}`);
    assert.deepEqual(
      [f['Desarrollo Contratado'], f['Mejora Funcional'], f['Deuda Tecnica'], f['Sin clasificar'], f.total],
      valores, nombre
    );
  }
  const total = Object.values(esperadoTabla).reduce((s, v) => s + v[4], 0);
  assert.equal(tabla.reduce((s, f) => s + f.total, 0), total);
}

test('seguimiento HU abiertas por cliente, sprint y módulo', () => {
  assert.equal(r.seguimiento.total, 70);
  comparar(r.seguimiento.porCliente, esperado.cliente);
  comparar(r.seguimiento.porSprint, esperado.sprint);
  comparar(r.seguimiento.porModulo, esperado.modulo);
});

test('estados no mapeados quedan fuera de las HU consideradas', () => {
  const m = ejemplo.matriz.map((f) => f.slice());
  const fila = m.findIndex((f, i) => i > 1 && f[7] === 'Backlog');
  m.forEach((f, i) => { if (i > 1 && f[1] === m[fila][1]) f[7] = 'Estado nuevo'; });
  const r2 = I.calcular(I.parsearHoja(m).registros);
  assert.equal(r2.base.consideradas, 150);
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
