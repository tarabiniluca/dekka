(async () => {
  const PAGES = [
    { href: 'index.html',                              label: 'Indice' },
    { href: 'DEKKA___Chi_siamo.html',                  label: 'Chi siamo' },
    { href: 'DEKKA___Piani_e_Tariffe.html',            label: 'Piani e Tariffe' },
    { href: 'DEKKA___Copertura_Valore.html',           label: 'Copertura Valore' },
    { href: 'DEKKA___Simulatore_Spese_Importazione.html', label: 'Simulatore' },
    { href: 'DEKKA___Termini_e_Condizioni.html',       label: 'Termini e Condizioni' },
    { href: 'DEKKA___Accordo_di_Mandato_IT.html',      label: 'Mandato IT' },
    { href: 'DEKKA___Accordo_di_Mandato.html',         label: 'Mandato IT·JP' },
  ];

  const rawFile = window.location.pathname.split('/').pop();
  const currentFile = rawFile === '' ? 'index.html' : rawFile;
  const isLocal = window.location.protocol === 'file:';

  async function pageExists(href) {
    if (href === currentFile) return true;
    if (isLocal) return true;
    try {
      const r = await fetch(href, { method: 'HEAD', cache: 'no-store' });
      return r.ok;
    } catch {
      return false;
    }
  }

  const checks = await Promise.all(PAGES.map(p => pageExists(p.href)));
  const visible = PAGES.filter((_, i) => checks[i]);

  const items = visible.map(p => {
    const active = p.href === currentFile;
    return `<a class="dknav-link${active ? ' dknav-link--active' : ''}" href="${p.href}">${p.label}</a>`;
  }).join('');

  const style = document.createElement('style');
  style.textContent = `
    #dknav {
      position: sticky;
      top: 0;
      z-index: 100;
      background: var(--bg-dark);
      border-bottom: 1px solid rgba(255,255,255,0.07);
    }
    .dknav-inner {
      max-width: 900px;
      margin: 0 auto;
      padding: 0 24px;
      display: flex;
      align-items: center;
      gap: 20px;
      height: 44px;
    }
    .dknav-brand {
      font-family: var(--font-mono);
      font-size: 11px;
      letter-spacing: 0.16em;
      text-transform: uppercase;
      color: rgba(255,255,255,0.9);
      flex-shrink: 0;
      text-decoration: none;
    }
    .dknav-sep {
      width: 1px;
      height: 14px;
      background: rgba(255,255,255,0.12);
      flex-shrink: 0;
    }
    .dknav-links {
      display: flex;
      align-items: center;
      gap: 2px;
      overflow-x: auto;
      scrollbar-width: none;
      flex: 1;
    }
    .dknav-links::-webkit-scrollbar { display: none; }
    .dknav-link {
      font-family: var(--font-mono);
      font-size: 10px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: rgba(255,255,255,0.4);
      text-decoration: none;
      padding: 4px 10px;
      border-radius: 3px;
      white-space: nowrap;
      transition: color 0.15s, background 0.15s;
    }
    .dknav-link:hover {
      color: rgba(255,255,255,0.8);
      background: rgba(255,255,255,0.07);
    }
    .dknav-link--active {
      color: rgba(255,255,255,0.9);
      background: rgba(255,255,255,0.1);
    }
  `;
  document.head.appendChild(style);

  const nav = document.createElement('nav');
  nav.id = 'dknav';
  nav.innerHTML = `
    <div class="dknav-inner">
      <a class="dknav-brand" href="index.html">DEKKA</a>
      <div class="dknav-sep"></div>
      <div class="dknav-links">${items}</div>
    </div>`;
  document.body.insertBefore(nav, document.body.firstChild);
})();
