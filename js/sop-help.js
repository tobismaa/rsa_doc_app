(function () {
  function initSopModal() {
    const fab = document.getElementById('helpFab');
    const modal = document.getElementById('sopHelpModal');
    const closeHeader = document.getElementById('closeSopHelpModalBtn');
    const closeFooter = document.getElementById('closeSopHelpModalFooterBtn');

    if (!modal) return;

    const sidebar = document.querySelector('.sidebar');
    const navMenu = document.querySelector('.sidebar .nav-menu');
    if (fab) {
      fab.style.display = 'none';
      fab.setAttribute('aria-hidden', 'true');
    }

    let triggerBtn = fab;
    if (sidebar) {
      let toolsWrap = sidebar.querySelector('#sidebarBottomTools');
      if (!toolsWrap) {
        toolsWrap = document.createElement('div');
        toolsWrap.id = 'sidebarBottomTools';
        toolsWrap.className = 'sidebar-bottom-tools';
        sidebar.appendChild(toolsWrap);
      }
      let item = document.getElementById('sidebarSopNavItem');
      if (!item) {
        item = document.createElement('button');
        item.type = 'button';
        item.id = 'sidebarSopNavItem';
        item.className = 'sidebar-help-fab';
        item.innerHTML = '<i class="fas fa-circle-question"></i><span>Help / SOP</span>';
        toolsWrap.appendChild(item);
      }
      triggerBtn = item;
    } else if (navMenu) {
      let item = document.getElementById('sidebarSopNavItem');
      if (!item) {
        item = document.createElement('button');
        item.type = 'button';
        item.id = 'sidebarSopNavItem';
        item.className = 'nav-item nav-item-button';
        item.innerHTML = '<i class="fas fa-circle-question"></i><span>Help / SOP</span>';
        navMenu.appendChild(item);
      }
      triggerBtn = item;
    }

    if (!triggerBtn) return;

    triggerBtn.addEventListener('click', function () {
      modal.classList.add('active');
    });

    const closeModal = function () {
      modal.classList.remove('active');
    };

    closeHeader && closeHeader.addEventListener('click', closeModal);
    closeFooter && closeFooter.addEventListener('click', closeModal);

    window.addEventListener('click', function (e) {
      if (e.target === modal) closeModal();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSopModal, { once: true });
  } else {
    initSopModal();
  }
})();
