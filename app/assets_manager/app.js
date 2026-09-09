import { requireAssetsSession, friendlyError } from "./auth.js";
import { listAssets, createAsset, validateAsset } from "./asset-store.js";
import { createManagedUser, listManagedUsers, updateManagedUser } from "./admin.js";

const tabs = document.querySelectorAll(".tab");
const panels = document.querySelectorAll(".tab-panel");
let assets = [];

const currencyFormatter = new Intl.NumberFormat("en-NG", {
  style: "currency",
  currency: "NGN",
  maximumFractionDigits: 0,
});

function activateTab(tab) {
  if (!tab) return;

  const target = tab.dataset.tab;

  tabs.forEach((item) => {
    const isActive = item === tab;
    item.classList.toggle("active", isActive);
    item.setAttribute("aria-selected", String(isActive));
    item.tabIndex = isActive ? 0 : -1;
  });

  panels.forEach((panel) => {
    panel.classList.toggle("active", panel.id === target);
  });

  document.querySelector("#page-label").textContent = tab.textContent.trim();
  tab.scrollIntoView({
    behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth",
    block: "nearest",
    inline: "center",
  });
}

function setMessage(message, isError = false) {
  const target = document.querySelector("#asset-form-message");
  if (!target) return;

  target.textContent = message;
  target.classList.toggle("error-text", isError);
}

function formatCurrency(value) {
  return currencyFormatter.format(Number(value || 0));
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
}

function statusClass(status = "") {
  return status.toLowerCase().replace(/\s+/g, "-");
}

function assetMatchesSearch(asset, searchTerm) {
  if (!searchTerm) return true;

  const haystack = [
    asset.name,
    asset.tag,
    asset.category,
    asset.custodian,
    asset.location,
    asset.status,
  ]
    .join(" ")
    .toLowerCase();

  return haystack.includes(searchTerm.toLowerCase());
}

function renderEmptyRow(tableBody, colspan, message) {
  if (!tableBody) return;

  tableBody.innerHTML = `
    <tr>
      <td class="empty-cell" colspan="${colspan}"><span class="empty-icon" aria-hidden="true">&#9633;</span>${escapeHtml(message)}</td>
    </tr>
  `;
}

function renderAssets() {
  const searchTerm = document.querySelector("#asset-search")?.value.trim() ?? "";
  const visibleAssets = assets.filter((asset) => assetMatchesSearch(asset, searchTerm));
  const assetsTable = document.querySelector("#assets-table-body");
  const recentTable = document.querySelector("#recent-assets-body");

  if (assetsTable) {
    if (!visibleAssets.length) {
      renderEmptyRow(assetsTable, 8, "No assets found.");
    } else {
      assetsTable.innerHTML = visibleAssets
        .map(
          (asset) => `
            <tr>
              <td>${escapeHtml(asset.name)}</td>
              <td>${escapeHtml(asset.tag)}</td>
              <td>${escapeHtml(asset.category)}</td>
              <td>${escapeHtml(asset.location || "-")}</td>
              <td>${escapeHtml(asset.custodian || "-")}</td>
              <td>${formatCurrency(asset.purchase_cost)}</td>
              <td>${formatCurrency(asset.book_value)}</td>
              <td><span class="status ${escapeHtml(statusClass(asset.status))}">${escapeHtml(asset.status)}</span></td>
            </tr>
          `,
        )
        .join("");
    }
  }

  if (recentTable) {
    const recentAssets = visibleAssets.slice(0, 4);
    if (!recentAssets.length) {
      renderEmptyRow(recentTable, 6, "No recent assets found.");
    } else {
      recentTable.innerHTML = recentAssets
        .map(
          (asset) => `
            <tr>
              <td>${escapeHtml(asset.name)}</td>
              <td>${escapeHtml(asset.tag)}</td>
              <td>${escapeHtml(asset.category)}</td>
              <td>${escapeHtml(asset.custodian || "-")}</td>
              <td>${formatCurrency(asset.book_value)}</td>
              <td><span class="status ${escapeHtml(statusClass(asset.status))}">${escapeHtml(asset.status)}</span></td>
            </tr>
          `,
        )
        .join("");
    }
  }

  updateMetrics(assets);
  updateOverview();
}

function updateMetrics(sourceAssets) {
  const totalAssets = sourceAssets.length;
  const totalCost = sourceAssets.reduce(
    (total, asset) => total + Number(asset.purchase_cost || 0),
    0,
  );
  const totalBookValue = sourceAssets.reduce(
    (total, asset) => total + Number(asset.book_value || 0),
    0,
  );
  const monthlyDepreciation = Math.max(totalCost - totalBookValue, 0);

  document.querySelector("#metric-total-assets").textContent = totalAssets;
  document.querySelector("#metric-acquisition-cost").textContent = formatCurrency(totalCost);
  document.querySelector("#metric-book-value").textContent = formatCurrency(totalBookValue);
  document.querySelector("#metric-depreciation").textContent =
    formatCurrency(monthlyDepreciation);

  const carryingValue = totalCost ? Math.round((totalBookValue / totalCost) * 100) : 0;
  document.querySelector("#metric-book-value-note").textContent =
    `${carryingValue}% carrying value`;
}

let loadVersion = 0;
let registerReady = false;

async function loadAssets() {
  const version = ++loadVersion;
  const refresh = document.querySelector("#refresh-assets");
  refresh.disabled = true;
  try {
    const result = await listAssets();
    if (version !== loadVersion) return false;
    assets = result;
    registerReady = true;
    renderAssets();
    setMessage("Your asset register is up to date.");
    document.querySelector("#app-message").hidden = true;
    return true;
  } catch (error) {
    if (version !== loadVersion) return false;
    assets = [];
    registerReady = false;
    renderAssets();
    const message = friendlyError(error);
    renderEmptyRow(document.querySelector("#assets-table-body"), 8, message);
    renderEmptyRow(document.querySelector("#recent-assets-body"), 6, message);
    setMessage(message, true);
    const notice = document.querySelector("#app-message");
    notice.textContent = message;
    notice.hidden = false;
    return false;
  } finally {
    if (version === loadVersion) refresh.disabled = false;
  }
}

async function saveAsset(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const saveButton = document.querySelector('[form="asset-form"]');
  if (saveButton.disabled) return;
  const formData = new FormData(form);
  const payload = {
    name: formData.get("name")?.toString().trim(),
    tag: formData.get("tag")?.toString().trim(),
    category: formData.get("category")?.toString(),
    custodian: formData.get("custodian")?.toString().trim() || null,
    location: formData.get("location")?.toString().trim() || null,
    purchase_cost: Number(formData.get("purchase_cost") || 0),
    book_value: Number(formData.get("book_value") || 0),
    purchase_date: formData.get("purchase_date") || null,
    status: formData.get("status")?.toString() || "Active",
  };

  try {
    validateAsset(payload);
  } catch (error) {
    setMessage(error.message, true);
    return;
  }
  saveButton.disabled = true;
  saveButton.textContent = "Saving...";
  try {
    await createAsset(payload);
    form.reset();
    const refreshed = await loadAssets();
    setMessage(refreshed ? `${payload.name} was added to the register.` : `${payload.name} was saved, but the register could not refresh. Please use Refresh.`, !refreshed);
  } catch (error) {
    setMessage(friendlyError(error), true);
  } finally {
    saveButton.disabled = false;
    saveButton.textContent = "Save Asset";
  }
}

async function initializeAssetManager() {
  const session = await requireAssetsSession();
  if (!session) return;

  const isSuperAdmin = session.assetRole === "super_admin";
  const adminTab = document.querySelector("#tab-admin");
  const adminPanel = document.querySelector("#admin");
  if (isSuperAdmin) {
    adminTab.hidden = false;
    adminPanel.hidden = false;
    document.querySelector("#asset-role-badge").textContent = "Super Admin";
    document.querySelector("#refresh-admin-users")?.addEventListener("click", loadAdminUsers);
    document.querySelector("#admin-user-form")?.addEventListener("submit", saveAdminUser);
    document.querySelector("#admin-users-body")?.addEventListener("click", handleAdminUserAction);
    loadAdminUsers();
  } else {
    document.querySelector("#asset-role-badge").textContent = session.assetRole === "admin" ? "Admin" : "Staff";
  }

  tabs.forEach((tab, index) => {
    tab.addEventListener("click", () => {
      activateTab(tab);
    });

    tab.addEventListener("keydown", (event) => {
      const isNext = event.key === "ArrowRight" || event.key === "ArrowDown";
      const isPrevious = event.key === "ArrowLeft" || event.key === "ArrowUp";

      if (!isNext && !isPrevious) return;

      event.preventDefault();

      const direction = isNext ? 1 : -1;
      const nextIndex = (index + direction + tabs.length) % tabs.length;
      tabs[nextIndex].focus();
      activateTab(tabs[nextIndex]);
    });
  });

  document.querySelector("#asset-form")?.addEventListener("submit", saveAsset);
  document.querySelector("#refresh-assets")?.addEventListener("click", loadAssets);
  document.querySelector("#asset-search")?.addEventListener("input", renderAssets);

  document.querySelectorAll(".primary-button").forEach((button) => {
    if (!button.textContent.includes("Add Asset")) return;

    button.addEventListener("click", () => {
      activateTab(document.querySelector('[data-tab="assets"]'));
      document.querySelector('[name="name"]')?.focus();
    });
  });

  document.querySelector("#dashboard .text-button")?.addEventListener("click", () => {
    activateTab(document.querySelector('[data-tab="assets"]'));
  });

  document.querySelectorAll("[data-export]").forEach(button => button.addEventListener("click", exportAssets));
  document.querySelector("[data-open-register]").addEventListener("click", () => activateTab(document.querySelector('[data-tab="assets"]')));
  document.querySelector("#snapshot-date").textContent = new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "long", year: "numeric" }).format(new Date());
  await loadAssets();
}

function updateOverview() {
  const categories = new Map();
  assets.forEach(asset => categories.set(asset.category || "Uncategorised", (categories.get(asset.category || "Uncategorised") || 0) + Number(asset.purchase_cost || 0)));
  document.querySelector("#category-count").textContent = `Across ${categories.size} asset categories`;
  const total = [...categories.values()].reduce((sum, value) => sum + value, 0);
  document.querySelector("#category-breakdown").innerHTML = categories.size
    ? [...categories.entries()].sort((a, b) => b[1] - a[1]).map(([name, value]) => {
      const percent = total ? Math.round(value / total * 100) : 0;
      return `<div class="category-row"><span>${escapeHtml(name)}</span><div class="bar-track"><div class="bar-fill" style="width:${percent}%"></div></div><strong>${percent}%</strong></div>`;
    }).join("")
    : '<p class="muted">Add your first asset to see your portfolio by category.</p>';
  document.querySelector("#attention-list").innerHTML = [["Service due", "Assets due for service"], ["Review", "Assets marked for review"], ["Disposal", "Assets marked for disposal"]].map(([status, label]) => `<li><span>${label}</span><strong>${assets.filter(asset => asset.status === status).length}</strong></li>`).join("");
}

function renderAdminUsers(users) {
  const target = document.querySelector("#admin-users-body");
  if (!target) return;
  document.querySelector("#admin-user-count").textContent = users.length;
  document.querySelector("#admin-active-count").textContent = users.filter(user => user.active && !user.disabled).length;
  document.querySelector("#admin-super-count").textContent = users.filter(user => user.role === "super_admin").length;
  target.innerHTML = users.length ? users.map(user => `
    <tr>
      <td><strong>${escapeHtml(user.display_name || user.email)}</strong><small class="table-subtext">${escapeHtml(user.email)}</small></td>
      <td><span class="status ${escapeHtml(user.role)}">${escapeHtml(user.role.replaceAll("_", " "))}</span></td>
      <td><span class="status ${user.active && !user.disabled ? "active" : "disposal"}">${user.active && !user.disabled ? "Active" : "Suspended"}</span></td>
      <td>${user.last_sign_in_at ? escapeHtml(new Date(user.last_sign_in_at).toLocaleDateString("en-GB")) : "Never"}</td>
      <td><button class="text-button admin-user-action" data-user-action="toggle" data-user-id="${escapeHtml(user.uid)}" data-user-active="${user.active && !user.disabled}">${user.active && !user.disabled ? "Suspend" : "Activate"}</button></td>
    </tr>
  `).join("") : '<tr><td colspan="5">No managed users found.</td></tr>';
}

async function loadAdminUsers() {
  const target = document.querySelector("#admin-users-body");
  if (target) target.innerHTML = '<tr><td colspan="5">Loading users...</td></tr>';
  try {
    const data = await listManagedUsers();
    renderAdminUsers(data.users || []);
  } catch (error) {
    if (target) target.innerHTML = `<tr><td colspan="5" class="error-text">${escapeHtml(error.message)}</td></tr>`;
  }
}

async function saveAdminUser(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const button = form.querySelector("button[type=submit]");
  const message = document.querySelector("#admin-user-message");
  button.disabled = true;
  button.textContent = "Saving...";
  try {
    const data = Object.fromEntries(new FormData(form).entries());
    const result = await createManagedUser({
      ...data,
      active: form.elements.active.checked,
      resetExistingPassword: form.elements.resetExistingPassword.checked,
    });
    message.textContent = `${result.email} is ${result.created ? "created" : "approved"} as ${String(result.role).replaceAll("_", " ")}.`;
    message.classList.remove("error-text");
    form.reset();
    form.elements.active.checked = true;
    await loadAdminUsers();
  } catch (error) {
    message.textContent = error.message;
    message.classList.add("error-text");
  } finally {
    button.disabled = false;
    button.textContent = "Create or approve user";
  }
}

async function handleAdminUserAction(event) {
  const button = event.target.closest("[data-user-action]");
  if (!button) return;
  const active = button.dataset.userActive === "true";
  if (active && !window.confirm("Suspend this user's Asset Manager access?")) return;
  button.disabled = true;
  try {
    await updateManagedUser(button.dataset.userId, { active: !active });
    await loadAdminUsers();
  } catch (error) {
    const message = document.querySelector("#admin-user-message");
    message.textContent = error.message;
    message.classList.add("error-text");
    button.disabled = false;
  }
}

function exportAssets() {
  if (!registerReady) {
    const notice = document.querySelector("#app-message");
    notice.textContent = "Load the register successfully before exporting.";
    notice.hidden = false;
    return;
  }
  const term = document.querySelector("#asset-search").value.trim();
  const fields = ["name", "tag", "category", "location", "custodian", "purchase_cost", "book_value", "purchase_date", "status"];
  const quote = value => '"' + String(value ?? "").replace(/^[=+@\-\t\r]/, "'$&").replaceAll('"', '""') + '"';
  const rows = [fields, ...assets.filter(asset => assetMatchesSearch(asset, term)).map(asset => fields.map(field => asset[field]))];
  const blob = new Blob(["\uFEFF" + rows.map(row => row.map(quote).join(",")).join("\r\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `cmbank-assets-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

initializeAssetManager().catch(() => {
  const notice = document.querySelector("#app-message");
  notice.textContent = "Unable to load the asset register. Please refresh and try again.";
  notice.hidden = false;
  setMessage("Unable to load the asset register. Please refresh and try again.", true);
  ["#assets-table-body", "#recent-assets-body"].forEach((selector, index) => renderEmptyRow(document.querySelector(selector), index ? 6 : 8, "Unable to load assets. Please refresh and try again."));
});
