(function () {
  const DISMISS_KEY = 'pwa_install_dismissed_until';
  const ONE_DAY = 24 * 60 * 60 * 1000;
  let deferredPrompt = null;

  function dismissed() {
    const until = Number(localStorage.getItem(DISMISS_KEY) || 0);
    return Number.isFinite(until) && until > Date.now();
  }

  function setDismiss(days = 3) {
    localStorage.setItem(DISMISS_KEY, String(Date.now() + (days * ONE_DAY)));
  }

  function ensureBanner() {
    if (document.getElementById('pwaInstallBanner')) return document.getElementById('pwaInstallBanner');
    const banner = document.createElement('div');
    banner.id = 'pwaInstallBanner';
    banner.className = 'pwa-install-banner';
    banner.innerHTML = `
      <p class="pwa-install-title">Install CMBank RSA App</p>
      <p class="pwa-install-text">Install now for faster access, app-like experience, and easier notifications.</p>
      <div class="pwa-install-actions">
        <button type="button" class="pwa-install-btn" id="pwaInstallLaterBtn">Later</button>
        <button type="button" class="pwa-install-btn primary" id="pwaInstallNowBtn">Install</button>
      </div>
    `;
    document.body.appendChild(banner);
    document.getElementById('pwaInstallLaterBtn')?.addEventListener('click', () => {
      banner.classList.remove('active');
      setDismiss(2);
    });
    document.getElementById('pwaInstallNowBtn')?.addEventListener('click', async () => {
      if (deferredPrompt) {
        deferredPrompt.prompt();
        try { await deferredPrompt.userChoice; } catch (_) {}
        deferredPrompt = null;
        banner.classList.remove('active');
        setDismiss(10);
        return;
      }
      alert('Use your browser menu and choose "Install App" to install this portal.');
    });
    return banner;
  }

  function maybeShowBanner() {
    if (dismissed()) return;
    const banner = ensureBanner();
    banner.classList.add('active');
  }

  function registerServiceWorker() {
    if (!('serviceWorker' in navigator)) return;
    navigator.serviceWorker.register('/service-worker.js').catch(() => {});
  }

  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    maybeShowBanner();
  });

  window.addEventListener('appinstalled', () => {
    setDismiss(30);
    const banner = document.getElementById('pwaInstallBanner');
    if (banner) banner.classList.remove('active');
  });

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      registerServiceWorker();
      setTimeout(maybeShowBanner, 900);
    }, { once: true });
  } else {
    registerServiceWorker();
    setTimeout(maybeShowBanner, 900);
  }
})();
