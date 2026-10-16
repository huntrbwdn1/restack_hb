// Restack: light or dark straight away, from last time (or the computer's setting the first time),
// so the panel doesn't flash white in dark mode. taskpane.js then checks Outlook and the saved choice.
(function () {
  var mode = null;
  try { mode = localStorage.getItem('thread-attachments.mode'); } catch (e) { /* private window */ }
  if (mode !== 'light' && mode !== 'dark') {
    mode = window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
  document.documentElement.setAttribute('data-mode', mode);
})();
