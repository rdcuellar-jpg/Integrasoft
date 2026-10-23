/*
 * Histórico de informes cargados. Cada carga de Excel se guarda en el navegador
 * (IndexedDB) para poder consultar y comparar cortes anteriores.
 * Si IndexedDB no está disponible, el histórico vive solo en memoria.
 */
(function (root) {
  'use strict';
  var BD = 'indicadores-integrasoft';
  var ALMACEN = 'informes';
  var memoria = {};
  var abrir = null;

  function conexion() {
    if (abrir) return abrir;
    abrir = new Promise(function (resolver) {
      try {
        if (!root.indexedDB) return resolver(null);
        var pet = root.indexedDB.open(BD, 1);
        pet.onupgradeneeded = function () {
          if (!pet.result.objectStoreNames.contains(ALMACEN)) pet.result.createObjectStore(ALMACEN, { keyPath: 'id' });
        };
        pet.onsuccess = function () { resolver(pet.result); };
        pet.onerror = function () { resolver(null); };
        pet.onblocked = function () { resolver(null); };
      } catch (e) { resolver(null); }
    });
    return abrir;
  }

  function operar(modo, fn) {
    return conexion().then(function (db) {
      if (!db) return fn(null);
      return new Promise(function (resolver, rechazar) {
        var tx = db.transaction(ALMACEN, modo);
        var res = fn(tx.objectStore(ALMACEN));
        tx.oncomplete = function () { resolver(res && res.result !== undefined ? res.result : res); };
        tx.onerror = function () { rechazar(tx.error); };
        tx.onabort = function () { rechazar(tx.error || new Error('Operación cancelada')); };
      });
    });
  }

  /** Huella del contenido para no guardar dos veces el mismo archivo. */
  function huella(matriz) {
    var s = JSON.stringify(matriz), h1 = 0x811c9dc5, h2 = 0;
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      h1 = Math.imul(h1 ^ c, 16777619) >>> 0;
      h2 = (h2 * 31 + c) >>> 0;
    }
    return h1.toString(16) + '-' + h2.toString(16) + '-' + s.length;
  }

  var Historico = {
    huella: huella,

    listar: function () {
      return operar('readonly', function (st) {
        if (!st) return Object.keys(memoria).map(function (k) { return memoria[k]; });
        return st.getAll();
      }).then(function (lista) {
        return (lista || []).slice().sort(function (a, b) {
          return (b.corte || '').localeCompare(a.corte || '') || (b.cargado || '').localeCompare(a.cargado || '');
        });
      });
    },

    obtener: function (id) {
      return operar('readonly', function (st) { return st ? st.get(id) : memoria[id]; });
    },

    guardar: function (informe) {
      return operar('readwrite', function (st) {
        if (!st) { memoria[informe.id] = informe; return informe; }
        st.put(informe);
        return informe;
      });
    },

    actualizar: function (id, cambios) {
      return Historico.obtener(id).then(function (inf) {
        if (!inf) throw new Error('El informe ya no existe.');
        Object.keys(cambios).forEach(function (k) { inf[k] = cambios[k]; });
        return Historico.guardar(inf);
      });
    },

    eliminar: function (id) {
      return operar('readwrite', function (st) {
        if (!st) { delete memoria[id]; return true; }
        st.delete(id);
        return true;
      });
    },

    persistente: function () {
      return conexion().then(function (db) { return !!db; });
    }
  };

  root.Historico = Historico;
})(typeof self !== 'undefined' ? self : this);
