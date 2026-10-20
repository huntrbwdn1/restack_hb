// Restack: light or dark and the text size straight away, from last time (or the computer's setting
// the first time), so the panel doesn't flash white in dark mode or jump in size. taskpane.js then
// checks Outlook and the saved choices. (The beta keeps its own; the admin console uses neither size,
// and gets its browser tab's name and icon here.)
(function () {
  var beta = /\/beta\/[^/]*$/.test(location.pathname);
  var admin = /[?&]console(=|&|$)/.test(location.search);
  var prefix = 'thread-attachments.' + (beta ? 'beta.' : '') + (admin ? 'console.' : ''); // as taskpane.js keeps them
  var mode = null;
  var size = null;
  try {
    mode = localStorage.getItem(prefix + 'mode');
    size = localStorage.getItem(prefix + 'size');
  } catch (e) { /* private window */ }
  if (mode !== 'light' && mode !== 'dark') {
    mode = window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
  var root = document.documentElement;
  root.setAttribute('data-mode', mode);
  var zoom = admin ? 0 : ({ medium: 1.12, large: 1.25 })[size] || 0;
  if (zoom) root.style.zoom = String(zoom);
  root.style.setProperty('--page-w', Math.round(window.innerWidth / (zoom || 1)) + 'px');
  root.style.setProperty('--page-h', Math.round(window.innerHeight / (zoom || 1)) + 'px');
  // The admin console's browser tab: its name and the Restack icon (Restack Beta's has the β), here so the
  // tab shows them straight away; pinned, it shows just the icon. (taskpane.js names the tab the same way.)
  if (admin) {
    document.title = 'Admin Console - Restack' + (beta ? ' Beta' : '');
    [16, 32, 64].forEach(function (px) {
      var icon = document.createElement('link');
      icon.rel = 'icon';
      icon.type = 'image/png';
      icon.setAttribute('sizes', px + 'x' + px);
      icon.href = 'icon-' + (beta ? 'beta-' : '') + px + '.png';
      document.head.appendChild(icon);
    });
  }
})();
