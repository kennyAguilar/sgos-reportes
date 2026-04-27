/* SGOS Chart.js theme — Casino Royale palette
   Apply defaults so existing inline Chart.js scripts inherit them.
   -------------------------------------------------------------- */
(function () {
  'use strict';
  if (typeof window.Chart === 'undefined') return;

  const PALETTE = [
    '#D4AF37', // gold
    '#10B981', // emerald
    '#60A5FA', // info blue
    '#F59E0B', // amber
    '#EF4444', // ruby
    '#A78BFA', // violet
    '#EC4899', // pink
    '#22D3EE'  // cyan
  ];

  const theme = {
    palette: PALETTE,
    gold:    '#D4AF37',
    emerald: '#10B981',
    ruby:    '#EF4444',
    amber:   '#F59E0B',
    info:    '#60A5FA',
    textPrimary:   '#F5F5F7',
    textSecondary: '#A8B2CE',
    grid:    'rgba(42, 54, 84, 0.6)',
    border:  'rgba(58, 71, 112, 0.6)',

    apply: function () {
      const Chart = window.Chart;
      Chart.defaults.color = this.textSecondary;
      Chart.defaults.borderColor = this.grid;
      Chart.defaults.font.family = "'Plus Jakarta Sans', system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
      Chart.defaults.font.size = 12;

      // Ajustes para móvil: tipografía más compacta, leyendas resumidas,
      // y maintainAspectRatio=false para que respeten contenedor .chart-wrap.
      const isMobile = document.body && document.body.classList.contains('is-mobile');
      if (isMobile) {
        Chart.defaults.font.size = 11;
        Chart.defaults.maintainAspectRatio = false;
        if (Chart.defaults.plugins && Chart.defaults.plugins.legend) {
          Chart.defaults.plugins.legend.position = 'bottom';
          Chart.defaults.plugins.legend.labels = Chart.defaults.plugins.legend.labels || {};
          Chart.defaults.plugins.legend.labels.boxWidth = 10;
          Chart.defaults.plugins.legend.labels.padding = 8;
          Chart.defaults.plugins.legend.labels.font = { size: 10 };
        }
      }

      if (Chart.defaults.plugins && Chart.defaults.plugins.legend) {
        Chart.defaults.plugins.legend.labels = Chart.defaults.plugins.legend.labels || {};
        Chart.defaults.plugins.legend.labels.color = this.textSecondary;
        Chart.defaults.plugins.legend.labels.boxWidth = 14;
        Chart.defaults.plugins.legend.labels.padding = 12;
      }
      if (Chart.defaults.plugins && Chart.defaults.plugins.tooltip) {
        Object.assign(Chart.defaults.plugins.tooltip, {
          backgroundColor: '#1C2540',
          titleColor: '#F5F5F7',
          bodyColor: '#A8B2CE',
          borderColor: '#2A3654',
          borderWidth: 1,
          padding: 10,
          cornerRadius: 8,
          boxPadding: 4
        });
      }
      // Default scale colors (Chart.js 3+)
      const scaleDefaults = Chart.defaults.scale || {};
      scaleDefaults.grid = Object.assign({}, scaleDefaults.grid, {
        color: this.grid,
        tickColor: this.grid
      });
      scaleDefaults.ticks = Object.assign({}, scaleDefaults.ticks, {
        color: this.textSecondary
      });
      Chart.defaults.scale = scaleDefaults;
    },

    colorAt: function (i) { return PALETTE[i % PALETTE.length]; },
    alpha: function (hex, a) {
      const h = hex.replace('#', '');
      const r = parseInt(h.substring(0, 2), 16);
      const g = parseInt(h.substring(2, 4), 16);
      const b = parseInt(h.substring(4, 6), 16);
      return 'rgba(' + r + ',' + g + ',' + b + ',' + a + ')';
    }
  };

  window.SGOSChartTheme = theme;
  theme.apply();
})();
