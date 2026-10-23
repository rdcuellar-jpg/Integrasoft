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

test('seguimiento HU abiertas por cliente, sprint y módulo', () => {
  assert.equal(r.seguimiento.total, 35);
  comparar(r.seguimiento.porCliente, esperado.cliente);
  comparar(r.seguimiento.porSprint, esperado.sprint);
  comparar(r.seguimiento.porModulo, esperado.modulo);
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
