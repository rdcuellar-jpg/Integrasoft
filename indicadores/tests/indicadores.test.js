// Verifica que el motor de cálculo reproduzca los valores del libro Excel
// "Base para Indicadores.xlsx" con la Hoja1 de ejemplo.
// Ejecutar: node --test indicadores/tests
const test = require('node:test');
const assert = require('node:assert/strict');
const I = require('../js/indicadores.js');
const ejemplo = require('../data/hoja1-ejemplo.json');
const esperado = require('./esperado-seguimiento.json');

const { registros, filasLeidas } = I.parsearHoja(ejemplo.matriz);
const r = I.calcular(registros);
const kpi = (id) => r.kpis.find((k) => k.id === id);

test('base de cálculo (Datos Indicadores!B5:B9)', () => {
  assert.equal(filasLeidas, 181);
  assert.equal(r.base.total, 133);
  assert.equal(r.base.excluidasNuevo, 19);
  assert.equal(r.base.excluidasAsignacion, 3);
  assert.equal(r.base.consideradas, 111);
});

test('indicadores por estado y fase', () => {
  const n = Object.fromEntries(r.porEstado.map((e) => [e.estado, e.n]));
  assert.deepEqual(n, {
    'Levantamiento Requerimiento': 10, 'Realizar Mockups': 1, 'Documentación': 2,
    'Aprobación cliente': 2, 'Revisión (sol)': 1, 'En desarrollo': 10, 'Revisión (IDS)': 5,
    'Testing': 5, 'Entrega': 19, 'CIERRE en Desarrollo o Solución': 56
  });
  assert.deepEqual(r.porFase.map((f) => f.n), [16, 15, 5, 75]);
  assert.deepEqual(r.estadosNoListados, {});
});

test('tipo de solicitud', () => {
  const n = Object.fromEntries(r.porTipo.map((t) => [t.nombre, t.n]));
  assert.deepEqual(n, { 'Mejora Funcional': 36, 'Deuda Tecnica': 7, 'Desarrollo Contratado': 15, 'Sin clasificar': 53 });
});

test('KPIs (Datos Indicadores!A49:G58)', () => {
  assert.equal(kpi('deuda').numerador, 7);
  assert.equal(kpi('deuda').denominador, 111);
  assert.equal(kpi('deuda').estado, 'EN META');
  assert.equal(kpi('roadmap').numerador, 75);
  assert.equal(kpi('roadmap').estado, 'FUERA DE META');
  assert.equal(Math.round(kpi('roadmap').resultado * 1000) / 10, 67.6);
  ['automatizacion', 'refinement', 'alcance', 'ambiguedad'].forEach((id) => {
    assert.equal(kpi(id).resultado, null);
    assert.equal(kpi(id).estado, 'SIN DATOS');
  });
  assert.equal(kpi('analisis').numerador, 16);
  assert.equal(kpi('desarrollo').numerador, 15);
  assert.equal(kpi('pruebas').numerador, 5);
  assert.equal(kpi('abiertas').numerador, 36);
  assert.equal(kpi('abiertas').estado, 'SIN META');
});

test('KPI manual con datos ingresados', () => {
  const r2 = I.calcular(registros, { manuales: { automatizacion: { num: 30, den: 40 } } });
  const k = r2.kpis.find((x) => x.id === 'automatizacion');
  assert.equal(k.resultado, 0.75);
  assert.equal(k.estado, 'EN META');
});

test('HU por mes de finalización (Datos Indicadores!A62:C75)', () => {
  const m = Object.fromEntries(r.mensual.map((x) => [x.mes, [x.cerradas, x.abiertas]]));
  assert.deepEqual(m['2026-01'], [1, 0]);
  assert.deepEqual(m['2026-03'], [2, 0]);
  assert.deepEqual(m['2026-04'], [0, 3]);
  assert.deepEqual(m['2026-05'], [0, 4]);
  assert.deepEqual(m['2026-06'], [1, 4]);
  assert.deepEqual(m['2026-07'], [11, 2]);
  assert.deepEqual(m['2026-08'], [1, 0]);
  assert.deepEqual(m['2026-09'], [13, 0]);
  assert.deepEqual(m['2026-10'], [0, 21]);
  assert.deepEqual(m['2026-11'], [0, 1]);
  assert.deepEqual([r.sinFecha.cerradas, r.sinFecha.abiertas], [46, 1]);
});

function comparar(tabla, esperadoTabla) {
  const filas = Object.fromEntries(tabla.map((f) => [f.nombre, f]));
  for (const [nombre, valores] of Object.entries(esperadoTabla)) {
    const total = valores[4];
    if (total === 0) { assert.ok(!filas[nombre] || filas[nombre].total === 0, nombre); continue; }
    const f = filas[nombre];
    assert.ok(f, `falta ${nombre}`);
    assert.deepEqual(
      [f['Desarrollo Contratado'], f['Mejora Funcional'], f['Deuda Tecnica'], f['Sin clasificar'], f.total],
      valores, nombre
    );
  }
}

test('seguimiento HU abiertas por cliente, sprint y módulo', () => {
  assert.equal(r.seguimiento.total, 36);
  comparar(r.seguimiento.porCliente, esperado.cliente);
  comparar(r.seguimiento.porSprint, esperado.sprint);
  comparar(r.seguimiento.porModulo, esperado.modulo);
});

test('parseFecha admite texto, Date y serial de Excel', () => {
  assert.equal(I.parseFecha('22-05-2026').getMonth(), 4);
  assert.equal(I.parseFecha(46164).getFullYear(), 2026);
  assert.equal(I.parseFecha(''), null);
});
