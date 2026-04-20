/* SGOS UI · toasts, global loader, skeletons
   Loaded on every page from layout.html
   -------------------------------------------------------------- */
(function () {
  'use strict';

  // ---------- Toasts ----------
  const ICONS = {
    success: 'bi-check-circle-fill',
    danger:  'bi-x-circle-fill',
    error:   'bi-x-circle-fill',
    warning: 'bi-exclamation-triangle-fill',
    info:    'bi-info-circle-fill'
  };
  const ROLES = {
    success: 'status',
    info:    'status',
    warning: 'alert',
    danger:  'alert',
    error:   'alert'
  };

  function normalizeCategory(cat) {
    if (!cat) return 'info';
    if (cat === 'error') return 'danger';
    return cat;
  }

  function ensureContainer() {
    let el = document.querySelector('.toast-container.sgos-toasts');
    if (!el) {
      el = document.createElement('div');
      el.className = 'toast-container sgos-toasts';
      el.setAttribute('aria-live', 'polite');
      el.setAttribute('aria-atomic', 'true');
      document.body.appendChild(el);
    }
    return el;
  }

  function dismissToast(toast) {
    if (!toast || toast.classList.contains('is-leaving')) return;
    toast.classList.add('is-leaving');
    const done = () => toast.remove();
    toast.addEventListener('animationend', done, { once: true });
    setTimeout(done, 400);
  }

  function showToast(message, category, delay) {
    const cat = normalizeCategory(category);
    const container = ensureContainer();
    const toast = document.createElement('div');
    toast.className = 'sgos-toast sgos-toast--' + cat;
    toast.setAttribute('role', ROLES[cat] || 'status');

    const icon = document.createElement('i');
    icon.className = 'sgos-toast__icon bi ' + (ICONS[cat] || ICONS.info);
    icon.setAttribute('aria-hidden', 'true');

    const body = document.createElement('div');
    body.className = 'sgos-toast__body';
    body.textContent = message;

    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'sgos-toast__close';
    close.setAttribute('aria-label', 'Cerrar notificación');
    close.innerHTML = '&times;';
    close.addEventListener('click', () => dismissToast(toast));

    toast.appendChild(icon);
    toast.appendChild(body);
    toast.appendChild(close);
    container.appendChild(toast);

    const ms = typeof delay === 'number' ? delay : 5000;
    if (ms > 0) setTimeout(() => dismissToast(toast), ms);
    return toast;
  }

  // Mount server-side flashes
  function mountFlashes() {
    const host = document.getElementById('sgos-flash-data');
    if (!host) return;
    let data = [];
    try { data = JSON.parse(host.textContent || '[]'); } catch (_) { return; }
    data.forEach((m) => showToast(m.message, m.category, 5000));
  }

  // ---------- Global loader ----------
  function showLoader(text) {
    const el = document.getElementById('global-loader');
    if (!el) return;
    if (text) {
      const t = el.querySelector('.global-loader__text');
      if (t) t.textContent = text;
    }
    el.classList.add('is-visible');
    el.removeAttribute('hidden');
  }
  function hideLoader() {
    const el = document.getElementById('global-loader');
    if (!el) return;
    el.classList.remove('is-visible');
    el.setAttribute('hidden', '');
  }

  // Auto-hook forms that upload files or opt-in via data-loading
  function wireForms() {
    const forms = document.querySelectorAll(
      'form[enctype="multipart/form-data"], form[data-loading="true"]'
    );
    forms.forEach((form) => {
      form.addEventListener('submit', function () {
        const msg = form.getAttribute('data-loading-text') || 'Procesando…';
        showLoader(msg);
      });
    });
    // If user hits back from bfcache, hide leftover overlay
    window.addEventListener('pageshow', hideLoader);
  }

  // ---------- Skeletons ----------
  function wireTableSkeletons() {
    // Auto-hide skeleton when DataTables initializes on a child table
    if (!window.jQuery || !window.jQuery.fn || !window.jQuery.fn.DataTable) return;
    const hosts = document.querySelectorAll('.table-host.is-loading');
    hosts.forEach((host) => {
      const table = host.querySelector('table');
      if (!table) return;
      window.jQuery(table).on('init.dt draw.dt', function () {
        host.classList.remove('is-loading');
      });
    });
  }

  // ---------- Public API ----------
  window.SGOS = window.SGOS || {};
  window.SGOS.toast = showToast;
  window.SGOS.showLoader = showLoader;
  window.SGOS.hideLoader = hideLoader;

  // ---------- Boot ----------
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () {
      mountFlashes();
      wireForms();
      wireTableSkeletons();
    });
  } else {
    mountFlashes();
    wireForms();
    wireTableSkeletons();
  }
})();
