// Global table enhancer
// Adds per-table search + pagination (10 rows) with Prev/Next controls.
(() => {
  const ROWS_PER_PAGE = 10;
  const stateMap = new WeakMap();

  function ensureStyles() {
    if (document.getElementById('tableEnhancerStyles')) return;
    const style = document.createElement('style');
    style.id = 'tableEnhancerStyles';
    style.textContent = `
      .table-enhancer-controls {
        display: flex;
        gap: 10px;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        margin: 0 0 10px;
      }
      .table-enhancer-search {
        min-width: 220px;
        max-width: 320px;
        width: 100%;
        padding: 9px 12px;
        border: 1px solid #d1d9e5;
        border-radius: 8px;
        font-size: 13px;
        background: #fff;
      }
      .table-enhancer-search:focus {
        outline: none;
        border-color: #0f4d8a;
        box-shadow: 0 0 0 3px rgba(15, 77, 138, 0.12);
      }
      .table-enhancer-pager {
        display: inline-flex;
        align-items: center;
        gap: 8px;
      }
      .table-enhancer-btn {
        border: 1px solid #c7d5e3;
        background: #ffffff;
        color: #1f2937;
        border-radius: 8px;
        padding: 7px 10px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
      }
      .table-enhancer-btn:disabled {
        opacity: 0.45;
        cursor: not-allowed;
      }
      .table-enhancer-page {
        font-size: 12px;
        color: #475569;
        min-width: 80px;
        text-align: center;
      }
      .table-enhancer-empty td {
        text-align: center;
        color: #64748b;
        font-size: 13px;
        padding: 12px;
      }
      a[href^="https://wa.me/"] {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        color: #16a34a;
        font-weight: 700;
        text-decoration: none;
        background: rgba(22, 163, 74, 0.1);
        border: 1px solid rgba(22, 163, 74, 0.25);
        border-radius: 999px;
        padding: 4px 10px;
      }
      a[href^="https://wa.me/"]::before {
        content: "\\f232";
        font-family: "Font Awesome 6 Brands";
        font-weight: 400;
      }
      a[href^="https://wa.me/"]:hover {
        background: rgba(22, 163, 74, 0.18);
        border-color: rgba(22, 163, 74, 0.4);
      }
    `;
    document.head.appendChild(style);
  }

  function getDataRows(tbody) {
    return Array.from(tbody.querySelectorAll(':scope > tr')).filter((tr) => {
      if (tr.classList.contains('table-enhancer-empty')) return false;
      if (tr.classList.contains('loading-row')) return false;
      if (tr.classList.contains('no-data')) return false;
      if (tr.querySelector('td.loading-row, td.no-data')) return false;
      return tr.querySelectorAll('td').length > 0;
    });
  }

  function removeEnhancerEmptyRow(tbody) {
    const row = tbody.querySelector(':scope > tr.table-enhancer-empty');
    if (row) row.remove();
  }

  function createEnhancerEmptyRow(tbody, table, message) {
    removeEnhancerEmptyRow(tbody);
    const row = document.createElement('tr');
    row.className = 'table-enhancer-empty';
    const td = document.createElement('td');
    const cols = table.querySelectorAll('thead th').length || 1;
    td.colSpan = cols;
    td.textContent = message;
    row.appendChild(td);
    tbody.appendChild(row);
  }

  function applyPagination(table, state) {
    const tbody = state.tbody;
    if (!tbody) return;

    removeEnhancerEmptyRow(tbody);
    const query = state.search.value.trim().toLowerCase();
    const allRows = getDataRows(tbody);

    // keep original "loading/no-data" rows visible when no data rows exist
    if (!allRows.length) {
      state.controls.style.display = 'none';
      return;
    }

    state.controls.style.display = 'flex';

    const filteredRows = allRows.filter((tr) => {
      if (!query) return true;
      const haystack = String(tr.dataset.search || tr.textContent || '').toLowerCase();
      return haystack.includes(query);
    });

    const totalPages = Math.max(1, Math.ceil(filteredRows.length / ROWS_PER_PAGE));
    if (state.page > totalPages) state.page = totalPages;
    if (state.page < 1) state.page = 1;

    allRows.forEach((tr) => {
      tr.style.display = 'none';
    });

    if (!filteredRows.length) {
      state.pageLabel.textContent = '0 / 0';
      state.prevBtn.disabled = true;
      state.nextBtn.disabled = true;
      createEnhancerEmptyRow(tbody, table, 'No matching records found');
      return;
    }

    const start = (state.page - 1) * ROWS_PER_PAGE;
    const end = start + ROWS_PER_PAGE;
    filteredRows.slice(start, end).forEach((tr) => {
      tr.style.display = '';
    });

    state.pageLabel.textContent = `${state.page} / ${totalPages}`;
    state.prevBtn.disabled = state.page <= 1;
    state.nextBtn.disabled = state.page >= totalPages;
  }

  function setupTable(table, index) {
    if (stateMap.has(table)) return;
    const tbody = table.tBodies && table.tBodies[0];
    if (!tbody) return;

    const container = table.closest('.table-container') || table.parentElement;
    if (!container) return;

    const controls = document.createElement('div');
    controls.className = 'table-enhancer-controls';
    controls.setAttribute('data-table-enhancer', String(index));

    const search = document.createElement('input');
    search.type = 'search';
    search.className = 'table-enhancer-search';
    search.placeholder = 'Search this table...';
    search.setAttribute('aria-label', 'Search table rows');

    const pager = document.createElement('div');
    pager.className = 'table-enhancer-pager';

    const prevBtn = document.createElement('button');
    prevBtn.type = 'button';
    prevBtn.className = 'table-enhancer-btn';
    prevBtn.textContent = 'Previous';

    const pageLabel = document.createElement('span');
    pageLabel.className = 'table-enhancer-page';
    pageLabel.textContent = '1 / 1';

    const nextBtn = document.createElement('button');
    nextBtn.type = 'button';
    nextBtn.className = 'table-enhancer-btn';
    nextBtn.textContent = 'Next';

    pager.appendChild(prevBtn);
    pager.appendChild(pageLabel);
    pager.appendChild(nextBtn);

    controls.appendChild(search);
    controls.appendChild(pager);

    container.parentNode.insertBefore(controls, container);

    const state = {
      tbody,
      controls,
      search,
      prevBtn,
      nextBtn,
      pageLabel,
      page: 1,
      observer: null
    };
    stateMap.set(table, state);

    search.addEventListener('input', () => {
      state.page = 1;
      applyPagination(table, state);
    });

    prevBtn.addEventListener('click', () => {
      state.page -= 1;
      applyPagination(table, state);
    });

    nextBtn.addEventListener('click', () => {
      state.page += 1;
      applyPagination(table, state);
    });

    const observer = new MutationObserver(() => {
      // Keep on current page if possible when table data refreshes
      applyPagination(table, state);
    });
    observer.observe(tbody, { childList: true, subtree: false });
    state.observer = observer;

    applyPagination(table, state);
  }

  function init() {
    ensureStyles();
    const tables = Array.from(document.querySelectorAll('table')).filter((table) => {
      if (String(table.dataset.noEnhance || '').toLowerCase() === 'true') return false;
      const tbody = table.tBodies && table.tBodies[0];
      return !!tbody;
    });
    tables.forEach((table, idx) => setupTable(table, idx + 1));
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  // Catch tables rendered later (dynamic tabs/modals)
  const appObserver = new MutationObserver(() => {
    init();
  });
  appObserver.observe(document.documentElement, { childList: true, subtree: true });
})();
