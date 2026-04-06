/**
 * sortable.js — Ordenamiento genérico + fila de totales (on/off)
 * para tablas con clase "sortable".
 * Detecta automáticamente el tipo de dato (texto, número, moneda, porcentaje).
 * Click en el encabezado alterna asc/desc y muestra indicador.
 * Switch "Totales" muestra/oculta una fila resumen al pie de la tabla.
 */
(function () {
  'use strict';

  function parseValue(text) {
    var s = text.trim();
    var cleaned = s.replace(/[$%,]/g, '').trim();
    var num = parseFloat(cleaned);
    if (!isNaN(num) && cleaned !== '') return num;
    return null;
  }

  /* ── Formato numérico con separador de miles ── */
  function fmtNum(n) {
    var s = Math.round(n).toString();
    return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  /* ── Sorting ── */
  function sortTable(th, table) {
    var idx = Array.prototype.indexOf.call(th.parentNode.children, th);
    var tbody = table.querySelector('tbody');
    if (!tbody) return;
    var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
    if (rows.length === 0) return;

    var asc = th.dataset.sortDir !== 'asc';
    Array.prototype.forEach.call(th.parentNode.children, function (h) {
      h.dataset.sortDir = '';
      var icon = h.querySelector('.sort-icon');
      if (icon) icon.className = 'bi bi-arrow-down-up sort-icon ms-1 opacity-50';
    });
    th.dataset.sortDir = asc ? 'asc' : 'desc';
    var icon = th.querySelector('.sort-icon');
    if (icon) {
      icon.className = 'bi sort-icon ms-1';
      icon.classList.add(asc ? 'bi-sort-up' : 'bi-sort-down');
    }

    rows.sort(function (a, b) {
      var cellA = a.children[idx];
      var cellB = b.children[idx];
      if (!cellA || !cellB) return 0;
      var ta = cellA.textContent;
      var tb = cellB.textContent;
      var na = parseValue(ta);
      var nb = parseValue(tb);
      if (na !== null && nb !== null) return asc ? na - nb : nb - na;
      ta = ta.trim().toLowerCase();
      tb = tb.trim().toLowerCase();
      return asc ? ta.localeCompare(tb) : tb.localeCompare(ta);
    });

    rows.forEach(function (row) { tbody.appendChild(row); });
  }

  /* ── Totales ── */
  function getDataRows(table) {
    var tbody = table.querySelector('tbody');
    if (!tbody) return [];
    var all = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
    // Si hay cat-rows (tabla de categorías), solo sumar esas para no duplicar
    var cats = all.filter(function (r) { return r.classList.contains('cat-row'); });
    if (cats.length > 0) return cats;
    // Excluir sub-rows de expansión
    return all.filter(function (r) {
      return !Array.prototype.some.call(r.classList, function (c) { return c.indexOf('sub-') === 0; });
    });
  }

  function buildTfoot(table) {
    var rows = getDataRows(table);
    if (rows.length === 0) return null;
    var numCols = table.querySelectorAll('thead th').length;
    if (numCols < 2) return null;

    var cells = [];
    var hasAnyTotal = false;

    for (var col = 0; col < numCols; col++) {
      var sum = 0, count = 0, hasDollar = false, hasPercent = false, nonNum = 0;

      rows.forEach(function (row) {
        var cell = row.children[col];
        if (!cell) return;
        var txt = cell.textContent.trim();
        if (txt === '' || txt === '-' || txt === 'N/A') return;
        if (txt.indexOf('$') !== -1) hasDollar = true;
        if (txt.indexOf('%') !== -1) hasPercent = true;
        var cleaned = txt.replace(/[$%,\s]/g, '');
        var num = parseFloat(cleaned);
        if (!isNaN(num) && cleaned !== '') { sum += num; count++; }
        else { nonNum++; }
      });

      if (col === 0) {
        cells.push('TOTALES');
      } else if (count === 0 || nonNum > count * 0.5) {
        cells.push('');
      } else if (hasPercent) {
        hasAnyTotal = true;
        cells.push((count > 0 ? (sum / count).toFixed(1) : '0') + '%');
      } else {
        hasAnyTotal = true;
        cells.push((hasDollar ? '$' : '') + fmtNum(sum));
      }
    }

    if (!hasAnyTotal) return null;

    var tfoot = document.createElement('tfoot');
    var tr = document.createElement('tr');
    tr.style.borderTop = '2px solid rgba(255,255,255,0.3)';
    cells.forEach(function (val) {
      var td = document.createElement('td');
      td.className = 'fw-bold';
      td.textContent = val;
      tr.appendChild(td);
    });
    tfoot.appendChild(tr);
    tfoot.style.display = 'none';
    return tfoot;
  }

  /* ── Inicialización ── */
  function init() {
    var tables = document.querySelectorAll('table.sortable');
    tables.forEach(function (table) {
      /* Sorting */
      var headers = table.querySelectorAll('thead th');
      headers.forEach(function (th) {
        if (!th.querySelector('.sort-icon')) {
          var ic = document.createElement('i');
          ic.className = 'bi bi-arrow-down-up sort-icon ms-1 opacity-50';
          th.appendChild(ic);
        }
        th.style.cursor = 'pointer';
        th.style.userSelect = 'none';
        th.addEventListener('click', function () { sortTable(th, table); });
      });

      /* Totales toggle */
      var tfoot = buildTfoot(table);
      if (!tfoot) return;
      table.appendChild(tfoot);

      var uid = 'totals-' + Math.random().toString(36).substr(2, 6);
      var wrap = document.createElement('div');
      wrap.className = 'd-flex justify-content-end mb-2';
      wrap.innerHTML =
        '<div class="form-check form-switch">' +
          '<input class="form-check-input" type="checkbox" role="switch" id="' + uid + '">' +
          '<label class="form-check-label text-light small" for="' + uid + '">Totales</label>' +
        '</div>';

      var target = table.parentNode;
      if (target.classList && target.classList.contains('table-responsive')) {
        target.parentNode.insertBefore(wrap, target);
      } else {
        target.insertBefore(wrap, table);
      }

      wrap.querySelector('input').addEventListener('change', function () {
        tfoot.style.display = this.checked ? '' : 'none';
      });
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
