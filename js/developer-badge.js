(function () {
  const BADGE_ID = 'developerCreditBadge';
  const SIDEBAR_BADGE_ID = 'developerCreditSidebar';
  if (document.getElementById(BADGE_ID) || document.getElementById(SIDEBAR_BADGE_ID)) return;

  function badgeMarkup() {
    return ''
      + '<div class="dev-credit-mark" aria-hidden="true"><span>TS</span></div>'
      + '<div class="dev-credit-text">'
      + '  <div class="dev-credit-title">Designed by Tsoft Inc.</div>'
      + '  <div class="dev-credit-meta"><i class="fab fa-whatsapp"></i><span>WhatsApp: +2349066678171</span></div>'
      + '  <div class="dev-credit-copy">&copy; Tsoft Inc.</div>'
      + '</div>';
  }

  function buildSidebarBadge(sidebar) {
    let toolsWrap = sidebar.querySelector('#sidebarBottomTools');
    if (!toolsWrap) {
      toolsWrap = document.createElement('div');
      toolsWrap.id = 'sidebarBottomTools';
      toolsWrap.className = 'sidebar-bottom-tools';
      sidebar.appendChild(toolsWrap);
    }

    const card = document.createElement('a');
    card.id = SIDEBAR_BADGE_ID;
    card.className = 'dev-credit-sidebar';
    card.href = 'https://wa.me/2349066678171';
    card.target = '_blank';
    card.rel = 'noopener noreferrer';
    card.title = 'Chat Tsoft Inc. on WhatsApp';
    card.setAttribute('aria-label', 'Designed by Tsoft Inc. WhatsApp +2349066678171');
    card.innerHTML = badgeMarkup();
    toolsWrap.appendChild(card);
  }

  function buildFloatingBadge() {
    const badge = document.createElement('a');
    badge.id = BADGE_ID;
    badge.className = 'dev-credit-badge';
    badge.href = 'https://wa.me/2349066678171';
    badge.target = '_blank';
    badge.rel = 'noopener noreferrer';
    badge.title = 'Chat Tsoft Inc. on WhatsApp';
    badge.setAttribute('aria-label', 'Designed by Tsoft Inc. WhatsApp +2349066678171');
    badge.innerHTML = badgeMarkup();
    document.body.appendChild(badge);
  }

  function buildBadge() {
    const sidebar = document.querySelector('.sidebar');
    if (sidebar) {
      buildSidebarBadge(sidebar);
      return;
    }
    buildFloatingBadge();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', buildBadge, { once: true });
  } else {
    buildBadge();
  }
})();
