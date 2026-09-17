document.querySelector('#year').textContent = new Date().getFullYear();

const motionPreference = window.matchMedia('(prefers-reduced-motion: reduce)');

const splashScreen = document.createElement('div');
splashScreen.className = 'app-splash';
splashScreen.setAttribute('aria-hidden', 'true');
splashScreen.innerHTML = `
  <span class="splash-orbit splash-orbit-one" aria-hidden="true"></span>
  <span class="splash-orbit splash-orbit-two" aria-hidden="true"></span>
  <span class="splash-orbit splash-orbit-three" aria-hidden="true"></span>
  <div class="splash-panel" role="status" aria-live="polite">
    <span class="splash-mark"><span>CM</span></span>
    <span class="splash-kicker">Opening secure workspace</span>
    <strong class="splash-title">CMBank</strong>
    <span class="splash-message">Preparing your authorised access</span>
    <span class="splash-loader" aria-hidden="true"></span>
    <span class="splash-progress-text">Please wait</span>
  </div>
`;
document.body.appendChild(splashScreen);

let appLaunchStarted = false;

// Mobile browsers can restore this page with its previous launch state intact.
window.addEventListener('pageshow', () => {
  appLaunchStarted = false;
  splashScreen.classList.remove('is-visible', 'is-loading');
  splashScreen.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('is-launching-app');
});

const restartSplashAnimation = () => {
  splashScreen.classList.remove('is-visible');
  splashScreen.classList.remove('is-loading');
  void splashScreen.offsetWidth;
  splashScreen.classList.add('is-visible', 'is-loading');
};

const showSplashAndOpen = (card) => {
  if (appLaunchStarted) return;
  const link = card.querySelector('.launch-button');
  const target = link.getAttribute('href');
  if (!target || target.startsWith('#')) return;

  appLaunchStarted = true;
  const title = card.querySelector('h3')?.textContent?.trim() || 'Workspace';
  splashScreen.querySelector('.splash-title').textContent = title;
  splashScreen.setAttribute('aria-hidden', 'false');
  restartSplashAnimation();
  document.body.classList.add('is-launching-app');

  window.setTimeout(() => {
    window.location.href = target;
  }, 3000);
};

// A soft pointer highlight enhances the cards without affecting their links.
document.querySelectorAll('.workspace').forEach(card => {
  const body = card.querySelector('.card-body');
  const link = card.querySelector('.launch-button');
  card.tabIndex = 0;
  card.setAttribute('role', 'link');
  card.setAttribute('aria-label', link?.textContent?.trim() || 'Open workspace');

  card.addEventListener('pointermove', event => {
    if (motionPreference.matches || event.pointerType === 'touch') return;
    const bounds = body.getBoundingClientRect();
    body.style.setProperty('--pointer-x', `${event.clientX - bounds.left}px`);
    body.style.setProperty('--pointer-y', `${event.clientY - bounds.top}px`);
  });

  const handleLaunch = event => {
    if (event.defaultPrevented || (typeof event.button === 'number' && event.button !== 0) || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
    event.preventDefault();
    showSplashAndOpen(card);
  };

  // Handle the native link directly as well as taps elsewhere on its card.
  link.addEventListener('click', handleLaunch);
  card.addEventListener('click', handleLaunch);

  card.addEventListener('keydown', event => {
    if (event.key !== 'Enter' && event.key !== ' ') return;
    event.preventDefault();
    showSplashAndOpen(card);
  });
});
