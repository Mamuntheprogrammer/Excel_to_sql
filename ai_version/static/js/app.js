/** ExcelSQL Studio 2.0 — application shell and interactions. */

import { api, ApiError } from "./api.js";
import { SqlEditor } from "./editor.js";
import { DataGrid, gridEmpty } from "./grid.js";
import {
  $, $$, el, iconEl, iconHTML, escapeHtml, formatNumber, formatBytes,
  formatDuration, relativeTime, formatDateTime, debounce, sanitizeTableName,
  quoteIdent, toCSV, downloadBlob, timestampForFile, copyText, loadJSON,
  saveJSON, openContextMenu, closeContextMenu, sqlPreview,
} from "./utils.js";

const APP_VERSION = document.querySelector('meta[name="app-version"]')?.content || "2.0.0";
const HISTORY_KEY = "xlsql.history.v1";
const SETTINGS_KEY = "xlsql.settings.v1";
const THEME_KEY = "xlsql.theme";

const state = {
  status: { connected: false, database: null, recent: [], sqliteVersion: "", app: {} },
  schema: { tables: [], views: [] },
  schemaLoading: false,
  tabs: [],
  activeTabId: null,
  queryCounter: 0,
  expanded: new Set(),
  history: loadJSON(HISTORY_KEY, []),
  historySearch: "",
  settings: Object.assign({ maxRows: 5000, pageSize: 100 }, loadJSON(SETTINGS_KEY, {})),
  importSession: null,
  fatalShown: false,
};

const dom = {
  topbarConnection: $("#topbarConnection"),
  connectionCard: $("#connectionCard"),
  schemaTree: $("#schemaTree"),
  schemaSearch: $("#schemaSearch"),
  clearSchemaSearch: $("#clearSchemaSearch"),
  tabList: $("#tabList"),
  tabPanels: $("#tabPanels"),
  statusDot: $("#statusDot"),
  statusConnection: $("#statusConnection"),
  statusMessage: $("#statusMessage"),
  statusTabInfo: $("#statusTabInfo"),
  sqliteVersion: $("#sqliteVersion"),
  modal: $("#appModal"),
  confirmModal: $("#confirmModal"),
  toastStack: $("#toastStack"),
  historyDrawer: $("#historyDrawer"),
  drawerBackdrop: $("#drawerBackdrop"),
  dropOverlay: $("#dropOverlay"),
  sidebar: $("#sidebar"),
};

/* ------------------------------------------------------------------ helpers */
function saveSettings() {
  saveJSON(SETTINGS_KEY, state.settings);
}

function activeTab() {
  return state.tabs.find((tab) => tab.id === state.activeTabId) || null;
}

function dbConnected() {
  return Boolean(state.status.connected && state.status.database);
}

function setStatusDot(mode) {
  if (!dom.statusDot) return;
  dom.statusDot.className = "status-dot";
  if (mode === "busy") dom.statusDot.classList.add("busy");
  else if (mode === "on" || dbConnected()) dom.statusDot.classList.add("on");
}

function setStatusMessage(message) {
  if (dom.statusMessage) dom.statusMessage.textContent = message || "";
}

function setBusy(isBusy) {
  setStatusDot(isBusy ? "busy" : dbConnected() ? "on" : "off");
}

function notifyError(error, title = "Something went wrong") {
  console.error(error);
  if (error instanceof ApiError && error.status === 403) {
    showConnectionLost(error.message);
    return;
  }
  if (error instanceof ApiError) {
    toast({ type: "error", title: error.code === "NO_DATABASE" ? "No database open" : title, message: error.message });
    return;
  }
  toast({ type: "error", title, message: error?.message || String(error) });
}

function showConnectionLost(message) {
  if (state.fatalShown) return;
  state.fatalShown = true;
  const reload = el("button", { class: "btn btn-primary" }, "Reload page");
  reload.addEventListener("click", () => window.location.reload());
  openModal({
    title: "Connection lost",
    subtitle: "The page token no longer matches the local server.",
    icon: "alert",
    body: `<div class="inline-note danger">${iconHTML("alert")}<div>${escapeHtml(message || "The local server was restarted or the page is stale.")}<br>Reload the page to reconnect. Your database file is untouched.</div></div>`,
    footer: reload,
    dismissible: false,
  });
}

function debounceInput(fn, delay) {
  return debounce(fn, delay);
}

/* -------------------------------------------------------------------- toasts */
const TOAST_ICONS = { success: "check", error: "alert", warn: "alert", info: "info" };

function toast({ type = "info", title = "", message = "", duration = 4200 } = {}) {
  const node = document.createElement("div");
  node.className = `toast ${type}`;
  node.innerHTML = `
    <svg class="icon t-icon"><use href="#i-${TOAST_ICONS[type] || "info"}"></use></svg>
    <div class="t-main">
      ${title ? `<div class="t-title">${escapeHtml(title)}</div>` : ""}
      ${message ? `<div class="t-msg">${escapeHtml(message)}</div>` : ""}
    </div>
    <button class="t-close" title="Dismiss"><svg class="icon"><use href="#i-x"></use></svg></button>
  `;
  const remove = () => {
    node.classList.add("out");
    setTimeout(() => node.remove(), 200);
  };
  node.querySelector(".t-close").addEventListener("click", remove);
  dom.toastStack.append(node);
  while (dom.toastStack.children.length > 5) dom.toastStack.firstElementChild.remove();
  if (duration > 0) setTimeout(remove, duration);
  return node;
}

/* ------------------------------------------------------------------ modals */
function openModal({ title, subtitle = "", icon = "info", body = "", footer = null, wide = false, xl = false, dismissible = true }) {
  const dialog = dom.modal;
  if (dialog.open) dialog.close();
  dialog.className = `modal${wide ? " modal-wide" : ""}${xl ? " modal-xl" : ""}`;
  dialog.innerHTML = `
    <div class="modal-header">
      <div class="mh-icon"><svg class="icon"><use href="#i-${icon}"></use></svg></div>
      <div class="modal-title-wrap">
        <h2 class="modal-title">${escapeHtml(title)}</h2>
        ${subtitle ? `<p class="modal-sub">${escapeHtml(subtitle)}</p>` : ""}
      </div>
      ${dismissible ? `<button class="icon-btn modal-close" title="Close"><svg class="icon"><use href="#i-x"></use></svg></button>` : ""}
    </div>
    <div class="modal-body"></div>
    <div class="modal-footer"></div>
  `;
  const bodyEl = dialog.querySelector(".modal-body");
  if (body instanceof Node) bodyEl.append(body);
  else bodyEl.innerHTML = body || "";
  const footerEl = dialog.querySelector(".modal-footer");
  if (footer instanceof Node) footerEl.append(footer);
  else if (typeof footer === "string") footerEl.innerHTML = footer;
  if (!footer) footerEl.remove();

  const close = () => {
    if (dialog.open) dialog.close();
  };
  dialog.querySelector(".modal-close")?.addEventListener("click", close);
  dialog.onclick = (event) => {
    if (event.target === dialog && dismissible) close();
  };
  dialog.showModal();
  requestAnimationFrame(() => {
    const autofocus = dialog.querySelector("[data-autofocus]");
    autofocus?.focus();
  });
  return { dialog, body: bodyEl, footer: footerEl, close };
}

function confirmDialog({ title, message, confirmLabel = "Confirm", cancelLabel = "Cancel", danger = false }) {
  const dialog = dom.confirmModal;
  if (dialog.open) dialog.close();
  dialog.className = "modal confirm-modal";
  dialog.innerHTML = `
    <div class="modal-body">
      <h2 class="modal-title" style="font-size:15px;margin-bottom:7px">${escapeHtml(title)}</h2>
      <div class="confirm-text">${message}</div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" data-cancel>${escapeHtml(cancelLabel)}</button>
      <button class="btn ${danger ? "btn-danger" : "btn-primary"}" data-confirm>${escapeHtml(confirmLabel)}</button>
    </div>
  `;
  return new Promise((resolve) => {
    let settled = false;
    const finish = (value) => {
      if (settled) return;
      settled = true;
      dialog.close();
      resolve(value);
    };
    dialog.querySelector("[data-cancel]").addEventListener("click", () => finish(false));
    dialog.querySelector("[data-confirm]").addEventListener("click", () => finish(true));
    dialog.oncancel = (event) => {
      event.preventDefault();
      finish(false);
    };
    dialog.onclick = (event) => {
      if (event.target === dialog) finish(false);
    };
    dialog.showModal();
  });
}

/* -------------------------------------------------------------------- theme */
function applyTheme(theme) {
  const chosen = theme || localStorage.getItem(THEME_KEY) || (window.matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark");
  document.documentElement.dataset.theme = chosen;
  localStorage.setItem(THEME_KEY, chosen);
  const button = $("#btnTheme");
  if (button) button.title = chosen === "dark" ? "Switch to light theme" : "Switch to dark theme";
}

function toggleTheme() {
  const current = document.documentElement.dataset.theme;
  applyTheme(current === "dark" ? "light" : "dark");
  if (dbConnected()) {
    setStatusMessage(`Theme: ${document.documentElement.dataset.theme}`);
  }
}

/* ------------------------------------------------------------------ status */
function applyStatus(payload) {
  if (!payload) return;
  state.status = payload;
  if (payload.app?.version) {
    // keep in sync if the server was upgraded
    state.status.app = payload.app;
  }
  if (dom.sqliteVersion) dom.sqliteVersion.textContent = payload.sqliteVersion ? `SQLite ${payload.sqliteVersion}` : "";
  renderTopbarConnection();
  renderConnectionCard();
  renderWelcomeIfOpen();
  updateStatusBar();
}

function renderTopbarConnection() {
  if (!dom.topbarConnection) return;
  dom.topbarConnection.innerHTML = "";
  if (!dbConnected()) {
    dom.topbarConnection.innerHTML = `<span class="conn-empty">${iconHTML("database")} No database connected</span>`;
    return;
  }
  const db = state.status.database;
  const pill = el("button", { class: "conn-pill", title: `${db.path}\nClick to switch database` });
  pill.innerHTML = `
    ${iconHTML("database")}
    <span class="conn-name">${escapeHtml(db.name)}</span>
    <span class="conn-path">${escapeHtml(db.path)}</span>
    ${db.readOnly ? '<span class="pill-badge">read only</span>' : ""}
  `;
  pill.addEventListener("click", () => openDatabaseModal());
  dom.topbarConnection.append(pill);
}

function renderConnectionCard() {
  if (!dom.connectionCard) return;
  dom.connectionCard.innerHTML = "";
  if (!dbConnected()) {
    const card = el("div", { class: "conn-empty-card" });
    card.innerHTML = `<div class="empty-title">No database</div><div>Open a SQLite file or import a spreadsheet.</div>`;
    const button = el("button", { class: "btn btn-soft btn-sm" }, iconEl("folder"), "Open database");
    button.addEventListener("click", () => openDatabaseModal());
    card.append(button);
    dom.connectionCard.append(card);
    return;
  }
  const db = state.status.database;
  const card = el("div", { class: "conn-card" });
  card.innerHTML = `
    <div class="conn-icon">${iconHTML("database")}</div>
    <div class="conn-main">
      <div class="conn-title-row">
        <span class="conn-title" title="${escapeHtml(db.path)}">${escapeHtml(db.name)}</span>
        <span class="conn-badge ${db.readOnly ? "ro" : "ok"}">${db.readOnly ? "RO" : "RW"}</span>
      </div>
      <div class="conn-meta">${escapeHtml(formatBytes(db.size))} · ${escapeHtml(formatDateTime(db.modified))}</div>
    </div>
    <div class="conn-actions">
      <button class="icon-btn sm" data-refresh title="Refresh schema">${iconHTML("refresh")}</button>
      <button class="icon-btn sm danger" data-close title="Close database">${iconHTML("x")}</button>
    </div>
  `;
  card.querySelector("[data-refresh]").addEventListener("click", () => refreshSchema({ toast: true }));
  card.querySelector("[data-close]").addEventListener("click", closeDatabase);
  dom.connectionCard.append(card);
}

function updateStatusBar() {
  if (!dom.statusConnection) return;
  if (dbConnected()) {
    const db = state.status.database;
    dom.statusConnection.textContent = `${db.name}${db.readOnly ? " · read only" : ""}`;
    dom.statusConnection.title = db.path;
  } else {
    dom.statusConnection.textContent = "No database connected";
    dom.statusConnection.title = "";
  }
  setStatusDot("auto");
  const tab = activeTab();
  if (!tab) {
    dom.statusTabInfo.textContent = "";
    return;
  }
  if (tab.type === "data") {
    const info = tab.total != null ? `${formatNumber(tab.total)} rows` : tab.dataRows?.length ? `${formatNumber(tab.dataRows.length)} rows` : "";
    dom.statusTabInfo.textContent = `${tab.title}${info ? ` · ${info}` : ""}`;
  } else if (tab.type === "query") {
    dom.statusTabInfo.textContent = tab.lastSummary || "Query editor";
  } else {
    dom.statusTabInfo.textContent = "Workspace";
  }
}

/* ------------------------------------------------------------------ history */
function addHistory(entry) {
  state.history.unshift({
    sql: String(entry.sql || "").slice(0, 20000),
    ts: Date.now(),
    ok: Boolean(entry.ok),
    durationMs: entry.durationMs || 0,
    rows: entry.rows || 0,
    db: entry.db || (state.status.database?.name ?? ""),
    error: entry.error ? String(entry.error).slice(0, 2000) : "",
  });
  state.history = state.history.slice(0, 120);
  saveJSON(HISTORY_KEY, state.history);
  if (!dom.historyDrawer.hidden) renderHistoryDrawer();
}

function openHistoryDrawer() {
  dom.historyDrawer.hidden = false;
  dom.drawerBackdrop.hidden = false;
  renderHistoryDrawer();
  setTimeout(() => dom.historyDrawer.querySelector("#historySearch")?.focus(), 40);
}

function closeHistoryDrawer() {
  dom.historyDrawer.hidden = true;
  dom.drawerBackdrop.hidden = true;
}

function renderHistoryDrawer() {
  dom.historyDrawer.innerHTML = `
    <div class="drawer-header">
      <div class="dh-icon">${iconHTML("history")}</div>
      <div style="flex:1">
        <h2>Query history</h2>
        <div class="drawer-sub">Stored locally in your browser · ${state.history.length} queries</div>
      </div>
      <button class="icon-btn sm" id="historyClose" title="Close">${iconHTML("x")}</button>
    </div>
    <div class="drawer-search">
      <div class="sidebar-search" style="margin:0">
        ${iconHTML("search")}
        <input id="historySearch" type="text" placeholder="Search history" value="${escapeHtml(state.historySearch)}" spellcheck="false">
      </div>
    </div>
    <div class="drawer-list" id="historyList"></div>
    <div class="drawer-footer">
      <span style="font-size:10px;color:var(--faint)">Click a query to load it into the editor.</span>
      <span style="flex:1"></span>
      <button class="btn btn-subtle btn-sm" id="historyClear" ${state.history.length ? "" : "disabled"}>${iconHTML("trash")}Clear</button>
    </div>
  `;
  const list = dom.historyDrawer.querySelector("#historyList");
  const renderItems = () => {
    const search = state.historySearch.toLowerCase();
    const filtered = state.history.filter((entry) => !search || entry.sql.toLowerCase().includes(search) || (entry.error || "").toLowerCase().includes(search));
    list.innerHTML = "";
    if (!filtered.length) {
      list.innerHTML = `<div class="drawer-empty">${iconHTML("history")}<div>${state.history.length ? "No queries match your search." : "No queries yet. Run SQL and it will show up here."}</div></div>`;
      return;
    }
    for (const entry of filtered) {
      const button = el("button", { class: "history-item", title: entry.sql });
      button.innerHTML = `
        <div class="h-top">
          <span class="h-dot ${entry.ok ? "" : "fail"}"></span>
          <span class="h-time">${relativeTime(entry.ts)}</span>
          <span class="h-db">${escapeHtml(entry.db || "")}</span>
        </div>
        <div class="h-sql">${escapeHtml(entry.sql)}</div>
        <div class="h-meta">
          <span>${entry.ok ? "success" : "failed"}</span>
          <span>${formatDuration(entry.durationMs)}</span>
          ${entry.rows ? `<span>${formatNumber(entry.rows)} rows</span>` : ""}
        </div>
      `;
      button.addEventListener("click", () => {
        closeHistoryDrawer();
        loadSqlIntoEditor(entry.sql);
      });
      list.append(button);
    }
  };
  renderItems();
  dom.historyDrawer.querySelector("#historyClose").addEventListener("click", closeHistoryDrawer);
  dom.historyDrawer.querySelector("#historySearch").addEventListener("input", debounce((event) => {
    state.historySearch = event.target.value;
    renderItems();
  }, 160));
  dom.historyDrawer.querySelector("#historyClear").addEventListener("click", async () => {
    const ok = await confirmDialog({
      title: "Clear query history?",
      message: "This removes all locally stored queries from this browser. It does not touch your database.",
      confirmLabel: "Clear history",
      danger: true,
    });
    if (!ok) return;
    state.history = [];
    saveJSON(HISTORY_KEY, []);
    renderHistoryDrawer();
    toast({ type: "success", title: "History cleared" });
  });
}

function loadSqlIntoEditor(sql) {
  let tab = activeTab();
  if (!tab || tab.type !== "query") {
    tab = openQueryTab(sql);
    return;
  }
  tab.editor.setValue(sql, true);
  setStatusMessage("Query loaded from history");
}

/* --------------------------------------------------------------- schema tree */
async function refreshSchema({ toastMessage = false, silent = false } = {}) {
  if (!dbConnected()) {
    state.schema = { tables: [], views: [] };
    renderSchemaTree();
    return;
  }
  state.schemaLoading = true;
  renderSchemaTree();
  try {
    const schema = await api.schema();
    state.schema = { tables: schema.tables || [], views: schema.views || [] };
    for (const tab of state.tabs) {
      if (tab.type === "query" && tab.editor) tab.editor.setSchema(state.schema);
    }
    renderSchemaTree();
    renderWelcomeIfOpen();
    if (toastMessage) {
      toast({
        type: "success",
        title: "Schema refreshed",
        message: `${state.schema.tables.length} tables · ${state.schema.views.length} views`,
      });
    }
  } catch (error) {
    state.schemaLoading = false;
    renderSchemaTree();
    if (!silent) notifyError(error, "Could not load schema");
  }
  state.schemaLoading = false;
}

function updateSchemaFilterClear() {
  if (!dom.clearSchemaSearch) return;
  dom.clearSchemaSearch.hidden = !(dom.schemaSearch?.value || "");
}

function schemaMatches(name, columns, query) {
  if (!query) return true;
  if (name.toLowerCase().includes(query)) return true;
  return (columns || []).some((column) => column.name.toLowerCase().includes(query));
}

function renderSchemaTree() {
  if (!dom.schemaTree) return;
  dom.schemaTree.innerHTML = "";
  if (!dbConnected()) {
    dom.schemaTree.innerHTML = `
      <div class="tree-empty">
        ${iconHTML("database")}
        <strong>No connection</strong>
        Open a .db, .sqlite or .sqlite3 file.<br>Or import an Excel workbook.
      </div>`;
    return;
  }
  if (state.schemaLoading && !state.schema.tables.length) {
    dom.schemaTree.innerHTML = `<div class="tree-loading">
      <div class="skeleton-line" style="width:60%"></div>
      <div class="skeleton-line" style="width:82%"></div>
      <div class="skeleton-line" style="width:46%"></div>
      <div class="skeleton-line" style="width:70%"></div>
    </div>`;
    return;
  }
  const filter = (dom.schemaSearch?.value || "").trim().toLowerCase();
  const showTables = state.schema.tables.filter((table) => schemaMatches(table.name, table.columns, filter));
  const showViews = state.schema.views.filter((view) => schemaMatches(view.name, view.columns, filter));

  if (!state.schema.tables.length && !state.schema.views.length) {
    dom.schemaTree.innerHTML = `
      <div class="tree-empty">
        ${iconHTML("table")}
        <strong>Database is empty</strong>
        Import an Excel or CSV file to create tables.
      </div>`;
    return;
  }
  if (filter && !showTables.length && !showViews.length) {
    dom.schemaTree.innerHTML = `<div class="tree-empty">${iconHTML("search")}<strong>No matches</strong>Nothing matches “${escapeHtml(filter)}”.</div>`;
    return;
  }

  const renderGroup = (label, items, kind) => {
    if (!items.length) return null;
    const group = el("div", { class: "tree-group" });
    const head = el("button", { class: "tree-group-head" });
    head.innerHTML = `${iconHTML("chevron")}<span>${label}</span><span class="count">${items.length}</span>`;
    const container = el("div", { class: "tree-items" });
    head.addEventListener("click", () => group.classList.toggle("collapsed"));
    group.append(head, container);

    for (const item of items) {
      const itemEl = el("div", { class: `tree-item${state.expanded.has(item.name) ? " open" : ""}` });
      const row = el("button", { class: `tree-row ${kind}` });
      row.innerHTML = `
        ${iconHTML("chevron")}
        ${iconHTML(kind === "view" ? "view" : "table", "icon obj-icon")}
        <span class="obj-name">${escapeHtml(item.name)}</span>
        <span class="obj-meta">${item.columns.length}</span>
      `;
      const columns = el("div", { class: "tree-columns" });
      for (const column of item.columns) {
        if (filter && !item.name.toLowerCase().includes(filter) && !column.name.toLowerCase().includes(filter)) continue;
        const columnEl = el("button", { class: "tree-column", title: `${column.name} · ${column.type || "TEXT"}${column.pk ? " · primary key" : ""}` });
        columnEl.innerHTML = `
          ${column.pk ? iconHTML("key", "icon pk-icon") : iconHTML("column", "icon col-icon")}
          <span class="col-name">${escapeHtml(column.name)}</span>
          <span class="col-type">${escapeHtml(column.type || "TEXT")}</span>
        `;
        columnEl.addEventListener("click", () => insertIdentifier(column.name));
        columnEl.addEventListener("contextmenu", (event) => {
          event.preventDefault();
          openContextMenu(event.clientX, event.clientY, [
            { label: column.name, icon: "column" },
            { separator: true },
            { label: `Copy column name`, icon: "copy", action: () => copyText(column.name) },
            { label: `Insert into editor`, icon: "code", action: () => insertIdentifier(column.name) },
            {
              label: "Generate SELECT",
              icon: "play",
              action: () => openQueryTab(`SELECT ${quoteIdent(column.name)}\nFROM ${quoteIdent(item.name)}\nLIMIT 100;`),
            },
          ]);
        });
        columns.append(columnEl);
      }
      row.addEventListener("click", () => {
        const open = !itemEl.classList.contains("open");
        itemEl.classList.toggle("open", open);
        if (open) state.expanded.add(item.name);
        else state.expanded.delete(item.name);
      });
      row.addEventListener("dblclick", () => openDataTab(item.name));
      row.addEventListener("contextmenu", (event) => {
        event.preventDefault();
        openContextMenu(event.clientX, event.clientY, [
          { label: item.name, icon: kind === "view" ? "view" : "table" },
          { separator: true },
          { label: "View data", icon: "table", action: () => openDataTab(item.name) },
          { label: "Generate SELECT", icon: "play", action: () => openQueryTab(buildSelect(item.name)) },
          { label: "Copy name", icon: "copy", action: () => copyText(item.name) },
          { separator: true },
          {
            label: kind === "view" ? "Export view to CSV" : "Export table to CSV",
            icon: "download",
            action: () => exportTable(item.name),
          },
        ]);
      });
      itemEl.append(row, columns);
      container.append(itemEl);
    }
    return group;
  };

  const tableGroup = renderGroup("Tables", showTables, "table");
  const viewGroup = renderGroup("Views", showViews, "view");
  if (tableGroup) dom.schemaTree.append(tableGroup);
  if (viewGroup) dom.schemaTree.append(viewGroup);
}

function buildSelect(name, limit = 100) {
  return `SELECT *\nFROM ${quoteIdent(name)}\nLIMIT ${limit};`;
}

function insertIdentifier(name) {
  let tab = activeTab();
  if (!tab || tab.type !== "query") {
    tab = openQueryTab();
    tab.editor.insertAtCursor(quoteIdent(name));
  } else {
    tab.editor.insertAtCursor(quoteIdent(name));
  }
  setStatusMessage(`Inserted ${name}`);
}

async function exportTable(name) {
  try {
    setBusy(true);
    setStatusMessage(`Exporting ${name}…`);
    await api.exportTable(name);
    toast({ type: "success", title: "Export ready", message: `${name}.csv was downloaded.` });
    setStatusMessage("Export complete");
  } catch (error) {
    notifyError(error, "Export failed");
  } finally {
    setBusy(false);
    updateStatusBar();
  }
}

/* ------------------------------------------------------------------- welcome */
function renderWelcomeIfOpen() {
  const tab = state.tabs.find((entry) => entry.type === "welcome");
  if (tab) renderWelcome(tab);
}

function renderWelcome(tab) {
  const db = state.status.database;
  const recents = state.status.recent || [];
  const container = tab.contentEl;
  if (!container) return;
  const tables = state.schema.tables.length;
  const views = state.schema.views.length;
  const connectedChips = dbConnected()
    ? `<span>${iconHTML("table")} ${formatNumber(tables)} tables</span><span>${iconHTML("view")} ${formatNumber(views)} views</span><span>${iconHTML("database")} ${escapeHtml(db.name)}</span>`
    : "";
  container.innerHTML = `
    <div class="welcome-inner">
      <section class="welcome-hero">
        <div>
          <div class="hero-eyebrow">${iconHTML("sparkle")} SQLite & spreadsheet workbench</div>
          <h1 class="hero-title">Query your data.<br><span class="grad">Beautifully.</span></h1>
          <p class="hero-sub">A local-first studio for SQLite databases and spreadsheets. Open a <strong>.db</strong> file, import <strong>Excel</strong> or <strong>CSV</strong>, browse tables, and run SQL with a modern, keyboard-friendly interface. Nothing leaves your machine.</p>
          <div class="hero-actions">
            <button class="btn btn-primary btn-lg" id="welcomeOpen">${iconHTML("database")} Open database</button>
            <button class="btn btn-soft btn-lg" id="welcomeImport">${iconHTML("upload")} Import spreadsheet</button>
            <button class="btn btn-ghost btn-lg" id="welcomeQuery">${iconHTML("code")} New query</button>
          </div>
          <div class="hero-meta">
            <span>${iconHTML("check")} Local-first</span>
            <span>${iconHTML("check")} No installation or cloud</span>
            <span>${iconHTML("check")} Excel, CSV, SQLite</span>
            ${connectedChips}
          </div>
        </div>
        <div class="hero-visual" aria-hidden="true">
          <div class="hero-code-card">
            <div class="code-card-head"><i></i><i></i><i></i><span>workspace.sql</span></div>
            <div class="code-card-body">
              <div><span class="ln">1</span><span class="k">SELECT</span> c.name,</div>
              <div><span class="ln">2</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <span class="f">SUM</span>(o.total) <span class="k">AS</span> revenue</div>
              <div><span class="ln">3</span><span class="k">FROM</span> customers c</div>
              <div><span class="ln">4</span><span class="k">JOIN</span> orders o <span class="k">ON</span> o.customer_id = c.id</div>
              <div><span class="ln">5</span><span class="k">WHERE</span> o.status = <span class="s">'completed'</span></div>
              <div><span class="ln">6</span><span class="k">GROUP BY</span> c.name <span class="k">ORDER BY</span> revenue <span class="k">DESC</span></div>
              <div><span class="ln">7</span><span class="k">LIMIT</span> <span class="n">10</span>;</div>
            </div>
          </div>
          <div class="hero-float">${iconHTML("check")} 161 orders · 3 tables · 12 ms</div>
        </div>
      </section>

      <section class="welcome-section">
        <div class="section-head"><h2>Start here</h2><span class="hint">Three ways to get going</span></div>
        <div class="action-grid">
          <button class="action-card" id="cardOpen">
            <span class="ac-icon">${iconHTML("database")}</span>
            <span class="ac-title">Open a database</span>
            <span class="ac-desc">Load an existing SQLite file. Browse its schema and run SQL without changing your data.</span>
            <svg class="icon ac-arrow"><use href="#i-chevron"></use></svg>
          </button>
          <button class="action-card ac-import" id="cardImport">
            <span class="ac-icon">${iconHTML("upload")}</span>
            <span class="ac-title">Import a spreadsheet</span>
            <span class="ac-desc">Turn Excel sheets or CSV files into queryable SQLite tables in seconds.</span>
            <svg class="icon ac-arrow"><use href="#i-chevron"></use></svg>
          </button>
          <button class="action-card ac-query" id="cardQuery">
            <span class="ac-icon">${iconHTML("code")}</span>
            <span class="ac-title">Write a query</span>
            <span class="ac-desc">A syntax-highlighted editor with autocomplete, history and one-click export.</span>
            <svg class="icon ac-arrow"><use href="#i-chevron"></use></svg>
          </button>
        </div>
      </section>

      ${recents.length ? `
      <section class="welcome-section">
        <div class="section-head"><h2>Recent databases</h2><span class="hint">Stored on this machine</span></div>
        <div class="recent-list" id="welcomeRecents"></div>
      </section>` : ""}

      <section class="welcome-section">
        <div class="demo-banner">
          <div class="db-illus">${iconHTML("database")}</div>
          <div class="demo-text">
            <strong>No data handy? Try the demo database.</strong>
            <span>Customers, products and orders — ready for joins and aggregations.</span>
          </div>
          <button class="btn btn-soft btn-sm" id="welcomeDemo">${iconHTML("play")} Load demo</button>
        </div>
      </section>

      <div class="welcome-foot">
        <span class="kbd-hint"><kbd>Ctrl</kbd><kbd>Enter</kbd> Run query</span>
        <span class="kbd-hint"><kbd>Ctrl</kbd><kbd>O</kbd> Open database</span>
        <span class="kbd-hint"><kbd>Ctrl</kbd><kbd>T</kbd> New query</span>
        <span style="margin-left:auto">ExcelSQL Studio v${escapeHtml(APP_VERSION)} · SQLite ${escapeHtml(state.status.sqliteVersion || "")}</span>
      </div>
    </div>
  `;
  const bind = (id, handler) => container.querySelector(id)?.addEventListener("click", handler);
  bind("#welcomeOpen", () => openDatabaseModal());
  bind("#welcomeImport", () => openImportModal());
  bind("#welcomeQuery", () => openQueryTab());
  bind("#cardOpen", () => openDatabaseModal());
  bind("#cardImport", () => openImportModal());
  bind("#cardQuery", () => openQueryTab());
  bind("#welcomeDemo", () => loadDemoDatabase());

  const recentList = container.querySelector("#welcomeRecents");
  if (recentList) {
    for (const entry of recents) {
      const row = el("button", { class: `recent-row${entry.exists ? "" : " missing"}`, title: entry.path });
      row.innerHTML = `
        <span class="r-icon">${iconHTML("database")}</span>
        <span class="r-main">
          <span class="r-name">${escapeHtml(entry.name)}</span>
          <span class="r-path">${escapeHtml(entry.path)}</span>
        </span>
        <span class="r-time">${entry.exists ? relativeTime(entry.lastOpened) : "missing"}</span>
        <svg class="icon r-open"><use href="#i-chevron"></use></svg>
      `;
      row.addEventListener("click", () => {
        if (!entry.exists) {
          toast({ type: "warn", title: "File not found", message: entry.path });
          return;
        }
        openDatabasePath(entry.path, false);
      });
      recentList.append(row);
    }
  }
}

async function loadDemoDatabase() {
  try {
    setBusy(true);
    setStatusMessage("Loading demo database…");
    const response = await api.demoDatabase();
    applyStatus(response.status);
    await refreshSchema();
    toast({ type: "success", title: "Demo database loaded", message: "customers, products and orders are ready." });
    openDataTab("customers");
    setStatusMessage("Demo database ready");
  } catch (error) {
    notifyError(error, "Could not load demo");
  } finally {
    setBusy(false);
    updateStatusBar();
  }
}

/* ---------------------------------------------------------------------- tabs */
function createTabShell({ type, title, icon, closeable = true }) {
  const id = `tab-${Math.random().toString(36).slice(2, 10)}`;
  const tab = { id, type, title, icon, closeable, el: null };
  const panel = el("div", { class: "tab-panel", id: `panel-${id}`, role: "tabpanel" });
  panel.hidden = true;
  const content = el("div", { class: "tab-content", style: { display: "contents" } });
  panel.append(content);
  tab.el = panel;
  tab.contentEl = content;
  dom.tabPanels.append(panel);
  state.tabs.push(tab);
  updateTabBar();
  return tab;
}

function updateTabBar() {
  dom.tabList.innerHTML = "";
  for (const tab of state.tabs) {
    const button = el("button", {
      class: `tab${tab.id === state.activeTabId ? " active" : ""}`,
      title: tab.tooltip || tab.title,
      role: "tab",
    });
    button.dataset.tabId = tab.id;
    button.innerHTML = `
      <svg class="icon tab-icon"><use href="#i-${tab.icon || "code"}"></use></svg>
      <span class="tab-title">${escapeHtml(tab.title)}</span>
      ${tab.closeable ? `<span class="tab-close" title="Close tab"><svg class="icon"><use href="#i-x"></use></svg></span>` : ""}
    `;
    button.addEventListener("click", (event) => {
      if (event.target.closest(".tab-close")) {
        event.stopPropagation();
        closeTab(tab.id);
        return;
      }
      activateTab(tab.id);
    });
    button.addEventListener("auxclick", (event) => {
      if (event.button === 1 && tab.closeable) closeTab(tab.id);
    });
    dom.tabList.append(button);
  }
  requestAnimationFrame(() => {
    dom.tabList.querySelector(".tab.active")?.scrollIntoView({ block: "nearest", inline: "nearest" });
  });
}

function activateTab(id) {
  const tab = state.tabs.find((entry) => entry.id === id);
  if (!tab) return;
  state.activeTabId = id;
  for (const entry of state.tabs) {
    entry.el.hidden = entry.id !== id;
  }
  updateTabBar();
  updateStatusBar();
  if (tab.type === "query") {
    requestAnimationFrame(() => tab.editor?.focus());
  }
  if (tab.type === "data" && !tab.loaded && tab.dataRefs) {
    tab.loaded = true;
    loadTableData(tab);
  }
  if (tab.type === "welcome") renderWelcome(tab);
}

async function closeTab(id) {
  const tab = state.tabs.find((entry) => entry.id === id);
  if (!tab) return;
  if (tab.type === "query" && tab.editor?.getValue().trim()) {
    const ok = await confirmDialog({
      title: "Close query tab?",
      message: `The SQL in <strong>${escapeHtml(tab.title)}</strong> will be discarded.`,
      confirmLabel: "Close tab",
      danger: true,
    });
    if (!ok) return;
  }
  tab.el.remove();
  state.tabs = state.tabs.filter((entry) => entry.id !== id);
  if (state.activeTabId === id) {
    const next = state.tabs[Math.max(0, state.tabs.length - 1)];
    if (next) activateTab(next.id);
    else {
      state.activeTabId = null;
      openWelcomeTab();
    }
  } else {
    updateTabBar();
  }
  if (tab.type === "data") updateStatusBar();
}

function openWelcomeTab() {
  let tab = state.tabs.find((entry) => entry.type === "welcome");
  if (!tab) {
    tab = createTabShell({ type: "welcome", title: "Welcome", icon: "sparkle" });
    renderWelcome(tab);
  }
  activateTab(tab.id);
  return tab;
}

function openQueryTab(sql = "") {
  state.queryCounter += 1;
  const tab = createTabShell({
    type: "query",
    title: `Query ${state.queryCounter}`,
    icon: "code",
  });
  createQueryView(tab, sql);
  activateTab(tab.id);
  return tab;
}

function openDataTab(name, { activate = true } = {}) {
  if (!dbConnected()) {
    toast({ type: "warn", title: "No database open", message: "Open a SQLite file or import data first." });
    return null;
  }
  const existing = state.tabs.find((entry) => entry.type === "data" && entry.tableName === name);
  if (existing) {
    if (activate) activateTab(existing.id);
    return existing;
  }
  const tab = createTabShell({ type: "data", title: name, icon: "table", closeable: true });
  createDataView(tab, name);
  if (activate) activateTab(tab.id);
  return tab;
}

function openDatabaseFromRecent(path, readOnly = false) {
  openDatabasePath(path, readOnly);
}

async function openDatabasePath(path, readOnly = false) {
  if (!path) return;
  try {
    setBusy(true);
    setStatusMessage(`Opening ${path}…`);
    const response = await api.openDatabase(path, readOnly);
    applyStatus(response.status);
    dom.schemaSearch.value = "";
    updateSchemaFilterClear();
    await refreshSchema();
    renderWelcomeIfOpen();
    toast({ type: "success", title: "Database opened", message: response.database.name + (readOnly ? " (read-only)" : "") });
    openWelcomeTab();
    setStatusMessage("Ready");
  } catch (error) {
    notifyError(error, "Could not open database");
  } finally {
    setBusy(false);
    updateStatusBar();
  }
}

async function closeDatabase() {
  const ok = await confirmDialog({
    title: "Close database?",
    message: "Your file stays exactly where it is. Nothing is deleted.",
    confirmLabel: "Close database",
  });
  if (!ok) return;
  try {
    setBusy(true);
    const response = await api.closeDatabase();
    applyStatus(response.status);
    state.schema = { tables: [], views: [] };
    state.expanded.clear();
    renderSchemaTree();
    dom.schemaSearch.value = "";
    updateSchemaFilterClear();
    for (const tab of [...state.tabs]) {
      if (tab.type === "data") closeTab(tab.id);
    }
    renderWelcomeIfOpen();
    toast({ type: "info", title: "Database closed" });
    setStatusMessage("No database connected");
    updateStatusBar();
  } catch (error) {
    notifyError(error, "Could not close database");
  } finally {
    setBusy(false);
  }
}

/* ------------------------------------------------------------- open database */
function pickFile(accept) {
  return new Promise((resolve) => {
    const input = el("input", { type: "file", accept, style: { display: "none" } });
    document.body.append(input);
    let settled = false;
    const finish = (file) => {
      if (settled) return;
      settled = true;
      input.remove();
      resolve(file || null);
    };
    input.addEventListener("change", () => finish(input.files?.[0] || null));
    window.addEventListener("focus", () => {
      setTimeout(() => {
        if (!settled && !input.files?.length) finish(null);
      }, 400);
    }, { once: true });
    input.click();
  });
}

function openDatabaseModal(droppedFile = null) {
  let mode = "open";
  let readOnly = false;

  const body = el("div");
  body.innerHTML = `
    <div class="segmented" style="margin-bottom:16px" data-mode>
      <button data-mode-open class="active">Open existing</button>
      <button data-mode-new>New database</button>
    </div>
    <div data-panel-open>
      <div class="field">
        <div class="field-label">SQLite file <span class="opt">.db · .sqlite · .sqlite3</span></div>
        <div class="input-row">
          <input class="input mono" data-path placeholder="Path to a database file…" spellcheck="false">
          <button class="btn btn-soft" data-browse>${iconHTML("folder")} Browse</button>
        </div>
        <div class="field-hint" data-hint>Type an absolute path on this machine, or upload a copy below.</div>
      </div>
      <label class="check-row" style="margin-bottom:14px">
        <input type="checkbox" data-readonly>
        <span class="check-text">Open in read-only mode <small>Prevents any write or schema change. Recommended for production files.</small></span>
      </label>
      <div class="divider">or open a copy</div>
      <div class="dropzone" data-dropzone>
        <div class="dz-icon">${iconHTML("database")}</div>
        <div class="dz-title">Drop a SQLite file here</div>
        <div class="dz-sub">or click to browse — a writable copy is placed in the workspace</div>
        <div class="dz-formats"><span class="file-chip">.db</span><span class="file-chip">.sqlite</span><span class="file-chip">.sqlite3</span></div>
      </div>
      <div class="progress" data-progress hidden><div class="bar"></div></div>
      <div class="field-hint" data-upload-status></div>
      <div data-recent-wrap style="margin-top:16px">
        <div class="field-label" style="margin-bottom:7px">Recent <span class="opt">This machine only</span></div>
        <div class="modal-recent" data-recents></div>
      </div>
    </div>
    <div data-panel-new hidden>
      <div class="field">
        <div class="field-label">New database file</div>
        <div class="input-row">
          <input class="input mono" data-new-path placeholder="e.g. sales-2026.db or /home/you/data/sales.db" spellcheck="false">
          <button class="btn btn-soft" data-browse-new>${iconHTML("folder")} Browse</button>
        </div>
        <div class="field-hint">A relative name is created inside the app workspace. The parent folder must exist.</div>
      </div>
      <div class="inline-note">${iconHTML("info")}<div>A new, empty SQLite database will be created and opened. You can then import Excel or CSV sheets into it.</div></div>
    </div>
  `;

  const footer = el("div", { style: { display: "contents" } });
  const footerLeft = el("div", { class: "footer-left" });
  footerLeft.innerHTML = `${iconHTML("info")} Local only · 127.0.0.1`;
  const demo = el("button", { class: "btn btn-subtle btn-sm" }, iconEl("play"), "Try demo");
  footerLeft.append(demo);
  const cancel = el("button", { class: "btn btn-ghost" }, "Cancel");
  const submit = el("button", { class: "btn btn-primary" }, iconEl("folder"), "Open database");
  footer.append(footerLeft, cancel, submit);

  const modal = openModal({
    title: "Open database",
    subtitle: "Connect to a SQLite file, or start with a fresh database.",
    icon: "database",
    body,
    footer,
    wide: true,
  });
  const close = () => modal.close();

  const pathInput = body.querySelector("[data-path]");
  const newPathInput = body.querySelector("[data-new-path]");
  const readOnlyInput = body.querySelector("[data-readonly]");
  const hint = body.querySelector("[data-hint]");
  const uploadStatus = body.querySelector("[data-upload-status]");
  const progress = body.querySelector("[data-progress]");
  const progressBar = progress.querySelector(".bar");
  const dropzone = body.querySelector("[data-dropzone]");
  const panelOpen = body.querySelector("[data-panel-open]");
  const panelNew = body.querySelector("[data-panel-new]");
  const recentsWrap = body.querySelector("[data-recent-wrap]");
  const recentsEl = body.querySelector("[data-recents]");

  const setMode = (nextMode) => {
    mode = nextMode;
    body.querySelector("[data-mode-open]").classList.toggle("active", mode === "open");
    body.querySelector("[data-mode-new]").classList.toggle("active", mode === "new");
    panelOpen.hidden = mode !== "open";
    panelNew.hidden = mode !== "new";
    submit.innerHTML = mode === "new" ? `${iconHTML("plus")} Create database` : `${iconHTML("folder")} Open database`;
    setTimeout(() => (mode === "new" ? newPathInput : pathInput).focus(), 20);
  };
  body.querySelector("[data-mode-open]").addEventListener("click", () => setMode("open"));
  body.querySelector("[data-mode-new]").addEventListener("click", () => setMode("new"));
  cancel.addEventListener("click", close);

  const setInlineError = (target, message) => {
    target.textContent = message || "";
    target.classList.toggle("error", Boolean(message));
  };

  const browse = async (isNew) => {
    const buttons = [body.querySelector("[data-browse]"), body.querySelector("[data-browse-new]")];
    const thisButton = isNew ? buttons[1] : buttons[0];
    const original = thisButton.innerHTML;
    thisButton.disabled = true;
    thisButton.innerHTML = `${iconHTML("spinner")} Waiting…`;
    try {
      const response = await api.browse(isNew ? "save" : "open", isNew ? newPathInput.value.trim() : "");
      if (response.path) {
        if (isNew) newPathInput.value = response.path;
        else pathInput.value = response.path;
      }
    } catch (error) {
      setInlineError(hint, error.message || String(error));
    } finally {
      thisButton.disabled = false;
      thisButton.innerHTML = original;
    }
  };
  body.querySelector("[data-browse]").addEventListener("click", () => browse(false));
  body.querySelector("[data-browse-new]").addEventListener("click", () => browse(true));

  const renderRecents = () => {
    const recents = state.status.recent || [];
    recentsEl.innerHTML = recents.length ? "" : `<div class="field-hint">No recent databases yet.</div>`;
    recents.slice(0, 6).forEach((entry) => {
      const row = el("button", { class: `modal-recent-row${entry.exists ? "" : " missing"}`, title: entry.path });
      row.innerHTML = `${iconHTML("database")}<span class="mr-main"><span class="mr-name">${escapeHtml(entry.name)}</span><span class="mr-path">${escapeHtml(entry.path)}</span></span><span style="font-size:9.5px;color:var(--faint)">${entry.exists ? relativeTime(entry.lastOpened) : "missing"}</span>`;
      row.addEventListener("click", async () => {
        if (!entry.exists) {
          setInlineError(hint, "That file no longer exists.");
          return;
        }
        pathInput.value = entry.path;
        await submitOpen();
      });
      recentsEl.append(row);
    });
  };
  renderRecents();

  let uploading = false;
  const uploadFile = async (file) => {
    if (uploading) return;
    if (!file) return;
    uploading = true;
    progress.hidden = false;
    progressBar.style.width = "0%";
    uploadStatus.textContent = `Uploading ${file.name}…`;
    uploadStatus.classList.remove("error");
    try {
      const response = await api.uploadDatabase(file, (percent) => { progressBar.style.width = `${percent}%`; });
      applyStatus(response.status);
      await refreshSchema();
      renderWelcomeIfOpen();
      close();
      toast({
        type: "success",
        title: "Database copy opened",
        message: `${response.originalName} was copied into the workspace. Changes apply to the copy.`,
      });
      openWelcomeTab();
      setStatusMessage("Ready");
    } catch (error) {
      uploadStatus.textContent = error.message || String(error);
      uploadStatus.classList.add("error");
      toast({ type: "error", title: "Could not open database", message: error.message });
    } finally {
      uploading = false;
      progress.hidden = true;
    }
  };

  dropzone.addEventListener("click", async () => {
    const file = await pickFile(".db,.sqlite,.sqlite3,.db3");
    if (file) uploadFile(file);
  });
  dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
    const file = event.dataTransfer?.files?.[0];
    if (file) uploadFile(file);
  });

  const submitOpen = async () => {
    const path = pathInput.value.trim();
    if (!path) {
      setInlineError(hint, "Choose a database file first.");
      pathInput.classList.add("invalid");
      pathInput.focus();
      return;
    }
    pathInput.classList.remove("invalid");
    setInlineError(hint, "");
    submit.disabled = true;
    submit.innerHTML = `${iconHTML("spinner")} Opening…`;
    try {
      const response = await api.openDatabase(path, readOnlyInput.checked);
      applyStatus(response.status);
      await refreshSchema();
      renderWelcomeIfOpen();
      close();
      toast({
        type: "success",
        title: "Database opened",
        message: `${response.database.name}${response.database.readOnly ? " (read-only)" : ""}`,
      });
      openWelcomeTab();
      setStatusMessage("Ready");
    } catch (error) {
      setInlineError(hint, error.message || String(error));
    } finally {
      submit.disabled = false;
      submit.innerHTML = `${iconHTML("folder")} Open database`;
    }
  };

  const submitNew = async () => {
    const path = newPathInput.value.trim();
    if (!path) {
      newPathInput.classList.add("invalid");
      newPathInput.focus();
      setInlineError(hint, "Enter a file name or path.");
      return;
    }
    newPathInput.classList.remove("invalid");
    submit.disabled = true;
    submit.innerHTML = `${iconHTML("spinner")} Creating…`;
    try {
      const response = await api.newDatabase(path);
      applyStatus(response.status);
      await refreshSchema();
      renderWelcomeIfOpen();
      close();
      toast({ type: "success", title: "Database created", message: response.database.path });
      openWelcomeTab();
      openQueryTab();
      setStatusMessage("Empty database ready — import a sheet or write SQL.");
    } catch (error) {
      notifyError(error, "Could not create database");
    } finally {
      submit.disabled = false;
      submit.innerHTML = `${iconHTML("plus")} Create database`;
    }
  };

  submit.addEventListener("click", () => {
    if (mode === "new") submitNew();
    else submitOpen();
  });
  pathInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") submitOpen();
  });
  newPathInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") submitNew();
  });
  demo.addEventListener("click", async () => {
    close();
    await loadDemoDatabase();
  });
  readOnlyInput.addEventListener("change", () => {
    readOnly = readOnlyInput.checked;
  });

  if (droppedFile) {
    uploadFile(droppedFile);
  } else {
    setTimeout(() => pathInput.focus(), 40);
  }
}

/* ------------------------------------------------------------------- importer */
function formatFileSize(bytes) {
  return formatBytes(bytes);
}

async function openImportModal(droppedFile = null) {
  const session = {
    file: droppedFile,
    upload: null,
    sheets: [],
    activeSheet: null,
    hasHeader: true,
    mode: "replace",
    previews: new Map(),
    busy: false,
    imported: null,
  };
  state.importSession = session;

  const body = el("div");
  const footer = el("div", { style: { display: "contents" } });
  const footerNote = el("div", { class: "footer-left" });
  footerNote.innerHTML = dbConnected()
    ? `${iconHTML("database")} ${escapeHtml(state.status.database.name)}`
    : `${iconHTML("info")} Imports create an “Imported Data” database`;
  const cancel = el("button", { class: "btn btn-ghost" }, "Cancel");
  const runButton = el("button", { class: "btn btn-primary" }, iconEl("upload"), "Import");
  footer.append(footerNote, cancel, runButton);

  const modal = openModal({
    title: "Import spreadsheet",
    subtitle: "Excel sheets and delimited files become SQLite tables.",
    icon: "upload",
    body,
    footer,
    xl: true,
  });
  const close = () => modal.close();
  cancel.addEventListener("click", close);

  const renderChooser = () => {
    body.innerHTML = `
      <div class="dropzone" data-importer-drop>
        <div class="dz-icon">${iconHTML("upload")}</div>
        <div class="dz-title">Drop an Excel or CSV file here</div>
        <div class="dz-sub">or click to browse — .xlsx, .xlsm, .csv, .tsv and .txt are supported</div>
        <div class="dz-formats"><span class="file-chip">.xlsx</span><span class="file-chip">.xlsm</span><span class="file-chip">.csv</span><span class="file-chip">.tsv</span><span class="file-chip">.txt</span></div>
      </div>
      <div class="progress" data-upload-progress hidden><div class="bar"></div></div>
      <div class="field-hint" data-upload-label style="margin-top:8px"></div>
      <div class="inline-note" style="margin-top:16px">${iconHTML("info")}<div><strong>Legacy .xls workbooks are not supported.</strong> Re-save them as .xlsx or export to CSV. Imported data is written to the active database, or to a new “Imported Data” database if none is open.</div></div>
    `;
    runButton.disabled = true;
    runButton.innerHTML = `${iconHTML("upload")} Import`;
    footerNote.innerHTML = dbConnected()
      ? `${iconHTML("database")} ${escapeHtml(state.status.database.name)}`
      : `${iconHTML("info")} Imports create an “Imported Data” database`;
    const drop = body.querySelector("[data-importer-drop]");
    const progress = body.querySelector("[data-upload-progress]");
    const bar = progress.querySelector(".bar");
    const label = body.querySelector("[data-upload-label]");
    drop.addEventListener("click", async () => {
      const file = await pickFile(".xlsx,.xlsm,.csv,.tsv,.txt,.xls");
      if (file) handleFile(file);
    });
    drop.addEventListener("dragover", (event) => { event.preventDefault(); drop.classList.add("dragover"); });
    drop.addEventListener("dragleave", () => drop.classList.remove("dragover"));
    drop.addEventListener("drop", (event) => {
      event.preventDefault();
      drop.classList.remove("dragover");
      const file = event.dataTransfer?.files?.[0];
      if (file) handleFile(file);
    });
    const handleFile = async (file) => {
      session.file = file;
      progress.hidden = false;
      bar.style.width = "0%";
      label.textContent = `Uploading ${file.name}…`;
      label.classList.remove("error");
      try {
        const upload = await api.importUpload(file, (percent) => { bar.style.width = `${percent}%`; });
        session.upload = upload;
        session.sheets = upload.sheets.map((name) => ({
          name,
          selected: true,
          table: sanitizeTableName(name),
        }));
        session.activeSheet = session.sheets[0]?.name || null;
        renderSheets();
      } catch (error) {
        label.textContent = error.message || String(error);
        label.classList.add("error");
      } finally {
        progress.hidden = true;
      }
    };
    session.startWithFile = handleFile;
  };

  const renderSheets = () => {
    const upload = session.upload;
    if (!upload) return;
    body.innerHTML = `
      <div class="import-file-head">
        <div class="if-icon">${iconHTML("file")}</div>
        <div class="if-main"><div class="if-name">${escapeHtml(upload.name)}</div><div class="if-meta">${formatFileSize(upload.size)} · ${upload.sheets.length} sheet${upload.sheets.length === 1 ? "" : "s"}</div></div>
        <button class="icon-btn sm" data-change-file title="Choose another file">${iconHTML("x")}</button>
      </div>
      <div class="import-layout">
        <div class="import-sheets" data-sheet-list></div>
        <div class="import-preview">
          <div class="import-preview-head">
            <span data-preview-name>Preview</span>
            <span class="hdr-badge" data-hdr-badge>First row = header</span>
            <span class="grow"></span>
            <label class="check-row" style="padding:0"><input type="checkbox" data-header-toggle checked><span class="check-text">Header row</span></label>
          </div>
          <div class="preview-scroll" data-preview></div>
        </div>
      </div>
      <div class="import-options">
        <div class="opt-group">Existing tables:
          <select class="select-xs" data-mode style="height:27px">
            <option value="replace">Replace</option>
            <option value="append">Append</option>
          </select>
        </div>
        <span class="grow"></span>
        <span class="field-hint">Table names are editable. Duplicate names get a suffix.</span>
      </div>
      <div data-import-feedback></div>
    `;
    body.querySelector("[data-change-file]").addEventListener("click", () => {
      session.upload = null;
      session.sheets = [];
      session.previews.clear();
      renderChooser();
    });
    const headerToggle = body.querySelector("[data-header-toggle]");
    headerToggle.checked = session.hasHeader;
    headerToggle.addEventListener("change", () => {
      session.hasHeader = headerToggle.checked;
      body.querySelector("[data-hdr-badge]").hidden = !session.hasHeader;
      renderPreview();
    });
    const modeSelect = body.querySelector("[data-mode]");
    modeSelect.value = session.mode;
    modeSelect.addEventListener("change", () => { session.mode = modeSelect.value; });

    const list = body.querySelector("[data-sheet-list]");
    list.innerHTML = "";
    session.sheets.forEach((sheet) => {
      const row = el("div", { class: `sheet-row${sheet.name === session.activeSheet ? " active" : ""}` });
      row.innerHTML = `
        <input type="checkbox" ${sheet.selected ? "checked" : ""} data-select>
        <div class="sh-main">
          <div class="sh-name" title="${escapeHtml(sheet.name)}">${escapeHtml(sheet.name)}</div>
          <div class="sh-meta">sheet</div>
          <input class="sh-table-input" value="${escapeHtml(sheet.table)}" spellcheck="false" data-table>
        </div>
      `;
      row.addEventListener("click", (event) => {
        if (event.target.closest("input")) return;
        session.activeSheet = sheet.name;
        list.querySelectorAll(".sheet-row").forEach((node) => node.classList.remove("active"));
        row.classList.add("active");
        loadPreview(sheet.name);
      });
      row.querySelector("[data-select]").addEventListener("change", (event) => {
        sheet.selected = event.target.checked;
        updateImportButton();
      });
      row.querySelector("[data-select]").addEventListener("click", (event) => event.stopPropagation());
      row.querySelector("[data-table]").addEventListener("click", (event) => event.stopPropagation());
      row.querySelector("[data-table]").addEventListener("input", (event) => {
        sheet.table = event.target.value;
      });
      list.append(row);
    });
    updateImportButton();
    if (session.activeSheet) loadPreview(session.activeSheet);
  };

  const updateImportButton = () => {
    const selected = session.sheets.filter((sheet) => sheet.selected);
    runButton.disabled = session.busy || !selected.length;
    runButton.innerHTML = `${iconHTML("upload")} Import ${selected.length} sheet${selected.length === 1 ? "" : "s"}`;
  };

  const loadPreview = async (sheetName) => {
    const previewEl = body.querySelector("[data-preview]");
    const nameEl = body.querySelector("[data-preview-name]");
    if (!previewEl) return;
    nameEl.textContent = sheetName;
    if (session.previews.has(sheetName)) {
      renderPreview(session.previews.get(sheetName));
      return;
    }
    previewEl.innerHTML = `<div class="loading-card" style="margin:16px auto"><span class="spinner"></span>Loading preview…</div>`;
    try {
      const preview = await api.importPreview(session.upload.fileId, sheetName);
      session.previews.set(sheetName, preview);
      if (session.activeSheet === sheetName) renderPreview(preview);
    } catch (error) {
      previewEl.innerHTML = `<div class="grid-empty"><p>${escapeHtml(error.message)}</p></div>`;
    }
  };

  const renderPreview = (previewOverride) => {
    const previewEl = body.querySelector("[data-preview]");
    if (!previewEl) return;
    const preview = previewOverride || session.previews.get(session.activeSheet);
    if (!preview) return;
    const rows = preview.rows || [];
    const width = preview.width || rows[0]?.length || 1;
    const headerRow = session.hasHeader ? rows[0] : null;
    const dataRows = session.hasHeader ? rows.slice(1) : rows;
    const names = Array.from({ length: width }, (_, index) => headerRow?.[index] ?? `column_${index + 1}`);
    const head = names.map((name) => `<th class="${session.hasHeader ? "is-header" : ""}">${escapeHtml(name == null ? "NULL" : name)}</th>`).join("");
    const bodyRows = dataRows.map((row, index) => `
      <tr><td class="p-index">${index + 1}</td>${Array.from({ length: width }, (_, col) => {
        const value = row[col];
        return `<td class="${value == null ? "p-null" : ""}">${value == null ? "NULL" : escapeHtml(String(value))}</td>`;
      }).join("")}</tr>
    `).join("");
    previewEl.innerHTML = `<table class="preview-table"><thead><tr><th class="p-index">#</th>${head}</tr></thead><tbody>${bodyRows}</tbody></table>`;
  };

  runButton.addEventListener("click", async () => {
    if (session.busy) return;
    const selected = session.sheets.filter((sheet) => sheet.selected);
    if (!selected.length) {
      toast({ type: "warn", title: "Select a sheet", message: "Choose at least one sheet or table to import." });
      return;
    }
    session.busy = true;
    updateImportButton();
    runButton.disabled = true;
    runButton.innerHTML = `${iconHTML("spinner")} Importing…`;
    const feedback = body.querySelector("[data-import-feedback]");
    try {
      const response = await api.importRun({
        fileId: session.upload.fileId,
        sheets: selected.map((sheet) => ({ sheet: sheet.name, table: sheet.table || sheet.name })),
        hasHeader: session.hasHeader,
        mode: session.mode,
      });
      session.imported = response;
      if (response.status) applyStatus(response.status);
      await refreshSchema();
      renderWelcomeIfOpen();
      if (response.errors?.length) {
        feedback.innerHTML = `
          <div class="import-summary">
            ${response.imported.map((entry) => `<div class="import-summary-row">${iconHTML("check")}<span class="row-name">${escapeHtml(entry.table)}</span><span class="row-meta">${formatNumber(entry.rows)} rows · ${entry.mode}</span></div>`).join("")}
          </div>
          <div class="import-error-list">
            ${response.errors.map((entry) => `<div class="import-error-row">${iconHTML("alert")}<div><strong>${escapeHtml(entry.sheet)}</strong> — ${escapeHtml(entry.error)}</div></div>`).join("")}
          </div>
        `;
        if (!response.imported.length) {
          toast({ type: "error", title: "Import failed", message: "Check the errors in the import window." });
        } else {
          toast({ type: "warn", title: "Partially imported", message: `${response.imported.length} imported · ${response.errors.length} failed` });
        }
        runButton.disabled = false;
        return;
      }
      const first = response.imported[0];
      close();
      if (response.createdDatabase) {
        toast({ type: "success", title: "Imported Data database created", message: `${formatNumber(response.imported.reduce((sum, entry) => sum + entry.rows, 0))} rows imported.` });
      } else {
        toast({ type: "success", title: "Import complete", message: `${response.imported.length} table${response.imported.length === 1 ? "" : "s"} · ${formatNumber(response.imported.reduce((sum, entry) => sum + entry.rows, 0))} rows` });
      }
      if (first) openDataTab(first.table);
      setStatusMessage(`Imported ${response.imported.length} table(s)`);
    } catch (error) {
      notifyError(error, "Import failed");
      if (feedback) feedback.innerHTML = `<div class="inline-note danger" style="margin-top:12px">${iconHTML("alert")}<div>${escapeHtml(error.message || String(error))}</div></div>`;
    } finally {
      session.busy = false;
      updateImportButton();
    }
  });

  renderChooser();
  if (droppedFile) session.startWithFile?.(droppedFile);
  setTimeout(() => body.querySelector("[data-importer-drop]")?.focus(), 40);
}

/* --------------------------------------------------------------- cell inspector */
function inspectCell(value, column, rowNumber) {
  const name = column?.name ?? "value";
  const type = column?.type || (value == null ? "NULL" : typeof value === "number" ? "NUMBER" : "TEXT");
  const display = value == null ? "NULL" : String(value);
  const body = el("div");
  body.innerHTML = `
    <pre class="cell-value">${value == null ? '<span style="color:var(--faint);font-style:italic">NULL</span>' : escapeHtml(display)}</pre>
    <div class="cell-meta">
      <span><strong>Column</strong> ${escapeHtml(name)}</span>
      <span><strong>Row</strong> ${escapeHtml(String(rowNumber ?? "—"))}</span>
      <span><strong>Type</strong> ${escapeHtml(String(type).toUpperCase())}</span>
      <span><strong>Length</strong> ${formatNumber(display.length)}</span>
    </div>
  `;
  const footer = el("div", { style: { display: "contents" } });
  const copy = el("button", { class: "btn btn-soft" }, iconEl("copy"), "Copy value");
  const close = el("button", { class: "btn btn-primary" }, "Close");
  footer.append(copy, close);
  const modal = openModal({
    title: name,
    subtitle: `Cell inspector${rowNumber ? ` · row ${rowNumber}` : ""}`,
    icon: "search",
    body,
    footer,
    wide: true,
  });
  copy.addEventListener("click", async () => {
    await copyText(value ?? "");
    toast({ type: "success", title: "Value copied" });
  });
  close.addEventListener("click", () => modal.close());
}

/* --------------------------------------------------------------------- help */
function openHelpModal() {
  const body = el("div");
  body.innerHTML = `
    <div class="help-grid">
      <section class="help-section">
        <h3>Keyboard</h3>
        <div class="shortcut-list">
          <div class="shortcut-row"><span>Run query / selection</span><span class="keys"><kbd>Ctrl</kbd><kbd>Enter</kbd></span></div>
          <div class="shortcut-row"><span>Autocomplete</span><span class="keys"><kbd>Ctrl</kbd><kbd>Space</kbd></span></div>
          <div class="shortcut-row"><span>New query tab</span><span class="keys"><kbd>Ctrl</kbd><kbd>T</kbd></span></div>
          <div class="shortcut-row"><span>Open database</span><span class="keys"><kbd>Ctrl</kbd><kbd>O</kbd></span></div>
          <div class="shortcut-row"><span>Import spreadsheet</span><span class="keys"><kbd>Ctrl</kbd><kbd>I</kbd></span></div>
          <div class="shortcut-row"><span>Close active tab</span><span class="keys"><kbd>Alt</kbd><kbd>W</kbd></span></div>
          <div class="shortcut-row"><span>Query history</span><span class="keys"><kbd>Ctrl</kbd><kbd>Shift</kbd><kbd>H</kbd></span></div>
          <div class="shortcut-row"><span>Focus schema filter</span><span class="keys"><kbd>/</kbd></span></div>
        </div>
      </section>
      <section class="help-section">
        <h3>Editor</h3>
        <div class="shortcut-list">
          <div class="shortcut-row"><span>Indent / dedent</span><span class="keys"><kbd>Tab</kbd><kbd>Shift Tab</kbd></span></div>
          <div class="shortcut-row"><span>Toggle comment</span><span class="keys"><kbd>Ctrl</kbd><kbd>/</kbd></span></div>
          <div class="shortcut-row"><span>Accept suggestion</span><span class="keys"><kbd>Enter</kbd></span></div>
          <div class="shortcut-row"><span>Dismiss suggestions</span><span class="keys"><kbd>Esc</kbd></span></div>
        </div>
        <h3 style="margin-top:16px">Supported files</h3>
        <div class="help-about">
          <strong>Databases:</strong> .db, .sqlite, .sqlite3, .db3<br>
          <strong>Spreadsheets:</strong> .xlsx, .xlsm, .csv, .tsv, .txt<br>
          <span style="color:var(--warning)">Legacy .xls is not supported — save as .xlsx first.</span>
        </div>
      </section>
      <section class="help-section">
        <h3>About</h3>
        <div class="help-about">
          <div class="about-brand">
            <svg class="icon about-mark" viewBox="0 0 32 32" style="width:30px;height:30px"><rect width="32" height="32" rx="8" fill="var(--accent)"/><ellipse cx="16" cy="10" rx="8" ry="3.2" fill="none" stroke="#fff" stroke-width="2"/><path d="M8 10v11c0 1.8 3.6 3.2 8 3.2s8-1.4 8-3.2V10" fill="none" stroke="#fff" stroke-width="2"/></svg>
            <div><strong>ExcelSQL Studio</strong> v${escapeHtml(APP_VERSION)}<br><span style="font-size:10px">Modern SQLite & spreadsheet workbench</span></div>
          </div>
          A local-first rewrite inspired by DBeaver and the original ExcelSQL Studio. Everything runs on your machine: the Flask server binds to 127.0.0.1 and every API call is protected by a session token. No telemetry, no uploads to third parties.
        </div>
      </section>
      <section class="help-section">
        <h3>Tips</h3>
        <div class="help-about">
          <strong>Read-only mode</strong> is ideal for inspecting production databases. Open the file with the checkbox enabled.<br><br>
          <strong>Drag & drop</strong> a database or spreadsheet anywhere on the window to open it.<br><br>
          <strong>Double-click a cell</strong> to inspect its full value. Right-click for copy options.<br><br>
          <strong>Query results</strong> are capped by the toolbar limit (5,000 by default) to keep the UI responsive.
        </div>
      </section>
    </div>
  `;
  const footer = el("div", { style: { display: "contents" } });
  const close = el("button", { class: "btn btn-primary" }, "Close");
  footer.append(close);
  const modal = openModal({
    title: "Help & shortcuts",
    subtitle: "A quick tour of the studio.",
    icon: "help",
    body,
    footer,
    xl: true,
  });
  close.addEventListener("click", () => modal.close());
}

/* ------------------------------------------------------------------- resizers */
function setupSidebarResizer() {
  const handle = $("#sidebarResizer");
  if (!handle) return;
  const savedWidth = Number(state.settings.sidebarWidth);
  if (savedWidth >= 210 && savedWidth <= 480) {
    document.documentElement.style.setProperty("--sidebar-w", `${savedWidth}px`);
  }
  handle.addEventListener("mousedown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    const startX = event.clientX;
    const startWidth = dom.sidebar.getBoundingClientRect().width;
    handle.classList.add("dragging");
    document.body.classList.add("is-resizing");
    const move = (moveEvent) => {
      const width = Math.max(220, Math.min(480, startWidth + moveEvent.clientX - startX));
      document.documentElement.style.setProperty("--sidebar-w", `${width}px`);
    };
    const up = () => {
      handle.classList.remove("dragging");
      document.body.classList.remove("is-resizing");
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
      state.settings.sidebarWidth = Math.round(dom.sidebar.getBoundingClientRect().width);
      saveSettings();
    };
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
  });
}

/* ------------------------------------------------------------------- dropzone */
function setupGlobalDrop() {
  let depth = 0;
  const isFileDrag = (event) => [...(event.dataTransfer?.types || [])].includes("Files");
  window.addEventListener("dragenter", (event) => {
    if (!isFileDrag(event)) return;
    event.preventDefault();
    depth += 1;
    dom.dropOverlay.hidden = false;
  });
  window.addEventListener("dragover", (event) => {
    if (!isFileDrag(event)) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  });
  window.addEventListener("dragleave", (event) => {
    if (!isFileDrag(event)) return;
    depth = Math.max(0, depth - 1);
    if (depth === 0) dom.dropOverlay.hidden = true;
  });
  window.addEventListener("drop", (event) => {
    if (!isFileDrag(event)) return;
    event.preventDefault();
    depth = 0;
    dom.dropOverlay.hidden = true;
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    const ext = (file.name.split(".").pop() || "").toLowerCase();
    if (["db", "sqlite", "sqlite3", "db3"].includes(ext)) {
      openDatabaseModal(file);
    } else if (["xlsx", "xlsm", "csv", "tsv", "txt"].includes(ext)) {
      openImportModal(file);
    } else if (ext === "xls") {
      toast({ type: "warn", title: "Legacy .xls not supported", message: "Save the workbook as .xlsx or export it to CSV." });
    } else {
      toast({ type: "warn", title: "Unsupported file", message: "Open a SQLite database, Excel workbook or CSV file." });
    }
  });
}

/* ------------------------------------------------------------------- events */
function bindEvents() {
  $("#btnOpenDb")?.addEventListener("click", () => openDatabaseModal());
  $("#btnOpenDbSide")?.addEventListener("click", () => openDatabaseModal());
  $("#btnImport")?.addEventListener("click", () => openImportModal());
  $("#btnHistory")?.addEventListener("click", openHistoryDrawer);
  $("#btnTheme")?.addEventListener("click", toggleTheme);
  $("#btnHelp")?.addEventListener("click", openHelpModal);
  $("#btnNewQuery")?.addEventListener("click", () => openQueryTab());
  $("#btnHomeTab")?.addEventListener("click", openWelcomeTab);
  $("#brandHome")?.addEventListener("click", openWelcomeTab);
  $("#btnRefreshSchema")?.addEventListener("click", () => refreshSchema({ toastMessage: true }));
  $("#btnToggleSidebar")?.addEventListener("click", () => dom.sidebar.classList.toggle("mobile-open"));
  dom.schemaSearch?.addEventListener("input", debounce(() => {
    updateSchemaFilterClear();
    renderSchemaTree();
  }, 160));
  dom.clearSchemaSearch?.addEventListener("click", () => {
    dom.schemaSearch.value = "";
    updateSchemaFilterClear();
    renderSchemaTree();
    dom.schemaSearch.focus();
  });
  dom.drawerBackdrop?.addEventListener("click", closeHistoryDrawer);

  document.addEventListener("pointerdown", (event) => {
    closeContextMenu();
    if (window.innerWidth <= 760 && dom.sidebar.classList.contains("mobile-open") && !dom.sidebar.contains(event.target) && !event.target.closest("#btnToggleSidebar")) {
      dom.sidebar.classList.remove("mobile-open");
    }
  });
  document.addEventListener("keydown", (event) => {
    const target = event.target;
    const editable = target.closest?.("input, textarea, select, [contenteditable='true']");
    const mod = event.ctrlKey || event.metaKey;
    if (event.key === "Escape") {
      if (!dom.historyDrawer.hidden) { closeHistoryDrawer(); return; }
      if (dom.dropOverlay && !dom.dropOverlay.hidden) { dom.dropOverlay.hidden = true; return; }
    }
    if (mod && !event.shiftKey && event.key.toLowerCase() === "o") {
      event.preventDefault();
      openDatabaseModal();
      return;
    }
    if (mod && !event.shiftKey && event.key.toLowerCase() === "i") {
      event.preventDefault();
      openImportModal();
      return;
    }
    if (mod && !event.shiftKey && event.key.toLowerCase() === "t") {
      event.preventDefault();
      openQueryTab();
      return;
    }
    if (mod && event.shiftKey && event.key.toLowerCase() === "h") {
      event.preventDefault();
      openHistoryDrawer();
      return;
    }
    if (event.altKey && event.key.toLowerCase() === "w") {
      event.preventDefault();
      if (state.activeTabId) closeTab(state.activeTabId);
      return;
    }
    if (event.key === "F1") {
      event.preventDefault();
      openHelpModal();
      return;
    }
    if (event.key === "?" && !editable) {
      event.preventDefault();
      openHelpModal();
      return;
    }
    if (event.key === "/" && !editable && !dom.modal.open && !dom.confirmModal.open) {
      event.preventDefault();
      dom.schemaSearch?.focus();
    }
  });
  setupSidebarResizer();
  setupGlobalDrop();

  window.addEventListener("resize", () => {
    if (window.innerWidth > 760) dom.sidebar.classList.remove("mobile-open");
  });
  window.addEventListener("unhandledrejection", (event) => {
    if (event.reason instanceof ApiError) notifyError(event.reason);
  });
}

/* --------------------------------------------------------------------- boot */
async function boot() {
  applyTheme();
  bindEvents();
  openWelcomeTab();
  setStatusMessage("Connecting to local server…");
  try {
    const status = await api.status();
    applyStatus(status);
    if (status.connected) {
      await refreshSchema({ silent: true });
      setStatusMessage("Database ready");
    } else {
      renderSchemaTree();
      setStatusMessage("Ready");
    }
  } catch (error) {
    notifyError(error, "Could not connect");
    setStatusMessage("Server unavailable");
  }
  renderTopbarConnection();
  renderConnectionCard();
  updateStatusBar();
}

boot();

console.info(`[ExcelSQL Studio] shell v${APP_VERSION} ready`);

function createQueryView(tab, initialSql = "") {
  tab.sql = initialSql;
  tab.results = [];
  tab.activeResult = 0;
  tab.running = false;

  const view = el("div", { class: "query-view" });
  view.innerHTML = `
    <div class="query-toolbar">
      <button class="btn btn-success run-btn" data-run>${iconHTML("play")}<span>Run</span><span class="run-hint">Ctrl↵</span></button>
      <div class="toolbar-divider"></div>
      <label class="toolbar-field">Limit
        <select class="select-xs" data-limit>
          <option value="1000">1,000 rows</option>
          <option value="5000">5,000 rows</option>
          <option value="20000">20,000 rows</option>
          <option value="100000">100,000 rows</option>
        </select>
      </label>
      <div class="toolbar-spacer"></div>
      <button class="icon-btn" data-save title="Download SQL">${iconHTML("download")}</button>
      <button class="icon-btn" data-history title="Query history">${iconHTML("history")}</button>
      <button class="icon-btn" data-clear title="Clear results">${iconHTML("trash")}</button>
    </div>
    <div class="editor-wrap" data-editor-wrap></div>
    <div class="editor-resizer" data-editor-resizer title="Drag to resize editor"></div>
    <div class="results-pane">
      <div class="results-head">
        <span class="results-title">${iconHTML("table")} Results</span>
        <div class="result-tabs" data-result-tabs></div>
        <span class="grow"></span>
        <div class="results-actions" data-result-actions></div>
      </div>
      <div class="result-body" data-result-body></div>
      <div class="result-footer" data-result-footer></div>
    </div>
  `;
  tab.contentEl.append(view);
  tab.view = view;
  tab.refs = {
    editorWrap: view.querySelector("[data-editor-wrap]"),
    editorResizer: view.querySelector("[data-editor-resizer]"),
    resultsPane: view.querySelector(".results-pane"),
    resultTabs: view.querySelector("[data-result-tabs]"),
    resultActions: view.querySelector("[data-result-actions]"),
    resultBody: view.querySelector("[data-result-body]"),
    resultFooter: view.querySelector("[data-result-footer]"),
    runButton: view.querySelector("[data-run]"),
    limitSelect: view.querySelector("[data-limit]"),
  };

  tab.editor = new SqlEditor(tab.refs.editorWrap, {
    value: initialSql,
    placeholder: "SELECT * FROM table_name LIMIT 100;",
    onRun: (sql) => runQuery(tab, sql),
    onChange: (value) => { tab.sql = value; },
    onCursor: () => updateStatusBar(),
  });
  tab.editor.setSchema(state.schema);

  if (![...tab.refs.limitSelect.options].some((option) => Number(option.value) === Number(state.settings.maxRows))) {
    const option = el("option", { value: state.settings.maxRows }, `${formatNumber(state.settings.maxRows)} rows`);
    tab.refs.limitSelect.append(option);
  }
  tab.refs.limitSelect.value = String(state.settings.maxRows);
  tab.refs.limitSelect.addEventListener("change", () => {
    state.settings.maxRows = Number(tab.refs.limitSelect.value) || 5000;
    saveSettings();
    toast({ type: "info", title: "Result limit", message: `Queries will return up to ${formatNumber(state.settings.maxRows)} rows.` });
  });

  tab.refs.runButton.addEventListener("click", () => tab.editor.runSelectionOrAll());
  view.querySelector("[data-history]").addEventListener("click", openHistoryDrawer);
  view.querySelector("[data-clear]").addEventListener("click", () => {
    tab.results = [];
    tab.activeResult = 0;
    setQueryEmpty(tab, "Results cleared", "Run the query to fetch data again.");
    tab.refs.resultTabs.innerHTML = "";
    tab.refs.resultActions.innerHTML = "";
    tab.refs.resultFooter.innerHTML = "";
    tab.lastSummary = "Query editor";
    updateStatusBar();
  });
  view.querySelector("[data-save]").addEventListener("click", () => {
    const sql = tab.editor.getValue().trim();
    if (!sql) {
      toast({ type: "warn", title: "Nothing to save", message: "Write some SQL first." });
      return;
    }
    downloadBlob(sql + "\n", `query-${timestampForFile()}.sql`, "application/sql;charset=utf-8");
    toast({ type: "success", title: "Query saved" });
  });

  setupEditorResizer(tab);
  setQueryEmpty(
    tab,
    "Ready to query",
    "Write SQL above and press Ctrl+Enter. Select text to run only that statement.",
    "play",
  );
  tab.refs.resultFooter.innerHTML = `<span class="footer-chip">${iconHTML("info")} Query results appear here</span><span class="grow"></span><span>${escapeHtml((state.status.database?.name) || "No database")}</span>`;
}

function setupEditorResizer(tab) {
  const handle = tab.refs.editorResizer;
  const editorWrap = tab.refs.editorWrap;
  handle.addEventListener("mousedown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    const startY = event.clientY;
    const startHeight = editorWrap.getBoundingClientRect().height;
    const totalHeight = tab.view.getBoundingClientRect().height;
    handle.classList.add("dragging");
    document.body.classList.add("is-row-resizing");
    const move = (moveEvent) => {
      const height = Math.max(90, Math.min(totalHeight - 140, startHeight + moveEvent.clientY - startY));
      editorWrap.style.flex = `0 0 ${height}px`;
      tab.refs.resultsPane.style.flex = "1 1 0";
    };
    const up = () => {
      handle.classList.remove("dragging");
      document.body.classList.remove("is-row-resizing");
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
    };
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
  });
  handle.addEventListener("dblclick", () => {
    editorWrap.style.flex = "";
    tab.refs.resultsPane.style.flex = "";
  });
}

function setQueryEmpty(tab, title, message, icon = "table") {
  tab.refs.resultBody.innerHTML = `
    <div class="result-empty">
      <div class="re-icon">${iconHTML(icon)}</div>
      <h3>${escapeHtml(title)}</h3>
      <p>${escapeHtml(message)}</p>
      <div class="empty-shortcuts"><span><kbd>Ctrl</kbd> <kbd>Enter</kbd> Run</span><span><kbd>Ctrl</kbd> <kbd>Space</kbd> Autocomplete</span></div>
    </div>
  `;
}

async function runQuery(tab, sql) {
  if (tab.running) return;
  if (!dbConnected()) {
    setQueryError(tab, "No database open", "Open a SQLite file or import data before running a query.", "");
    toast({ type: "warn", title: "No database open", message: "Use Open database or Import data to get started." });
    return;
  }
  if (!sql.trim()) {
    toast({ type: "warn", title: "Nothing to run", message: "The editor is empty." });
    return;
  }
  tab.running = true;
  tab.refs.runButton.classList.add("running");
  tab.refs.runButton.disabled = true;
  tab.refs.runButton.innerHTML = `${iconHTML("spinner")}<span>Running</span>`;
  setBusy(true);
  setStatusMessage("Running query…");
  tab.refs.resultBody.innerHTML = `<div class="grid-loading-overlay" style="position:static;flex:1"><div class="loading-card"><span class="spinner"></span>Executing SQL…</div></div>`;
  const started = performance.now();
  try {
    const response = await api.query(sql, state.settings.maxRows);
    const duration = response.elapsedMs ?? (performance.now() - started);
    renderQueryResults(tab, response);
    const totalRows = (response.results || [])
      .filter((entry) => entry.kind === "rows")
      .reduce((sum, entry) => sum + (entry.rowCount || 0), 0);
    addHistory({
      sql,
      ok: response.ok,
      durationMs: duration,
      rows: totalRows,
      db: state.status.database?.name,
      error: response.error?.message,
    });
    if (response.ok) {
      const summary = response.results?.some((entry) => entry.kind === "rows")
        ? `${formatNumber(totalRows)} row${totalRows === 1 ? "" : "s"} · ${formatDuration(duration)}`
        : `Statement executed · ${formatDuration(duration)}`;
      tab.lastSummary = summary;
      setStatusMessage(`Query finished · ${summary}`);
    } else {
      tab.lastSummary = "Query failed";
      setStatusMessage(`Query failed · ${formatDuration(duration)}`);
    }
    updateStatusBar();
  } catch (error) {
    setQueryError(tab, "Query could not be executed", error.message || String(error), "");
    addHistory({ sql, ok: false, durationMs: performance.now() - started, db: state.status.database?.name, error: error.message });
    notifyError(error, "Query failed");
    tab.lastSummary = "Query failed";
    setStatusMessage("Query failed");
  } finally {
    tab.running = false;
    tab.refs.runButton.classList.remove("running");
    tab.refs.runButton.disabled = false;
    tab.refs.runButton.innerHTML = `${iconHTML("play")}<span>Run</span><span class="run-hint">Ctrl↵</span>`;
    setBusy(false);
    updateStatusBar();
  }
}

function setQueryError(tab, title, message, statement) {
  tab.refs.resultTabs.innerHTML = "";
  tab.refs.resultActions.innerHTML = "";
  tab.refs.resultBody.innerHTML = `
    <div class="result-error">
      <div class="result-error-head">${iconHTML("alert")}<span>${escapeHtml(title)}</span><span class="grow"></span></div>
      <pre>${escapeHtml(message)}</pre>
      ${statement ? `<pre class="err-stmt">${escapeHtml(statement)}</pre>` : ""}
      ${/read.?only|query_only/i.test(message) ? `<div class="result-hint">This connection is read-only. Reopen the database without the read-only option to make changes.</div>` : ""}
    </div>
  `;
  tab.refs.resultFooter.innerHTML = `<span class="footer-chip warn">${iconHTML("alert")} Not executed</span><span class="grow"></span><span>${escapeHtml(state.status.database?.name || "")}</span>`;
}

function renderQueryResults(tab, response) {
  const entries = [...(response.results || [])];
  if (response.error) {
    entries.push({ kind: "error", error: response.error });
  }
  tab.results = entries;
  if (!entries.length) {
    setQueryEmpty(tab, "Nothing returned", "The statement completed without a result set.", "check");
    tab.refs.resultFooter.innerHTML = `<span class="footer-chip ok">${iconHTML("check")} Executed</span><span class="grow"></span><span>${formatDuration(response.elapsedMs)}</span>`;
    tab.resultEntries = [];
    tab.refs.resultTabs.innerHTML = "";
    tab.refs.resultActions.innerHTML = "";
    return;
  }
  tab.activeResult = response.error ? entries.length - 1 : 0;
  renderResultTabs(tab);
  renderActiveResult(tab);
}

function renderResultTabs(tab) {
  const entries = tab.results;
  const total = entries.length;
  tab.refs.resultTabs.innerHTML = "";
  if (total <= 1) return;
  entries.forEach((entry, index) => {
    const label = entry.kind === "error" ? "Error" : entry.kind === "rows" ? `Result ${index + 1}` : `Statement ${index + 1}`;
    const button = el("button", { class: `result-tab${index === tab.activeResult ? " active" : ""}` }, label);
    button.addEventListener("click", () => {
      tab.activeResult = index;
      renderResultTabs(tab);
      renderActiveResult(tab);
    });
    tab.refs.resultTabs.append(button);
  });
}

function renderActiveResult(tab) {
  const entry = tab.results[tab.activeResult];
  if (!entry) return;
  tab.refs.resultActions.innerHTML = "";
  tab.refs.resultFooter.innerHTML = "";

  if (entry.kind === "rows") {
    const wrap = el("div", { class: "grid-wrap" });
    tab.refs.resultBody.innerHTML = "";
    tab.refs.resultBody.append(wrap);
    const grid = new DataGrid(wrap, {
      onInspect: (value, column, row) => inspectCell(value, column, row + 1),
    });
    grid.setData({
      columns: entry.columns,
      rows: entry.rows,
      sortable: false,
    });
    createResultActions(tab, entry);
    const colCount = entry.columns?.length || 0;
    tab.refs.resultFooter.innerHTML = `
      <span class="footer-chip ok">${iconHTML("check")} ${formatNumber(entry.rowCount)} rows</span>
      <span class="footer-chip">${iconHTML("column")} ${formatNumber(colCount)} columns</span>
      <span class="footer-chip">${iconHTML("clock")} ${formatDuration(entry.elapsedMs)}</span>
      ${entry.truncated ? `<span class="footer-chip warn">${iconHTML("alert")} Showing first ${formatNumber(entry.rowCount)} rows</span>` : ""}
      <span class="grow"></span>
      ${entry.truncated ? `<span>Increase the limit in the toolbar to fetch more.</span>` : ""}
    `;
    tab.lastSummary = `${formatNumber(entry.rowCount)} rows · ${formatDuration(entry.elapsedMs)}`;
    return;
  }

  if (entry.kind === "exec") {
    tab.refs.resultBody.innerHTML = `
      <div class="exec-summary">
        <div class="exec-card">
          <div class="exec-icon">${iconHTML("check")}</div>
          <h3>Statement executed</h3>
          <p>The database accepted the change.</p>
          <div class="exec-stats">
            <div class="exec-stat"><div class="v">${formatNumber(entry.changes)}</div><div class="l">Rows affected</div></div>
            <div class="exec-stat"><div class="v">${formatDuration(entry.elapsedMs)}</div><div class="l">Elapsed</div></div>
            <div class="exec-stat"><div class="v">${entry.lastRowId ?? "—"}</div><div class="l">Last row id</div></div>
          </div>
        </div>
      </div>
    `;
    tab.refs.resultFooter.innerHTML = `
      <span class="footer-chip ok">${iconHTML("check")} Executed successfully</span>
      <span class="footer-chip">${formatNumber(entry.changes)} rows affected</span>
      <span class="grow"></span>
      <span>${formatDuration(entry.elapsedMs)}</span>
    `;
    tab.lastSummary = `Statement executed · ${formatNumber(entry.changes)} rows affected`;
    return;
  }

  if (entry.kind === "error") {
    const error = entry.error;
    tab.refs.resultBody.innerHTML = `
      <div class="result-error">
        <div class="result-error-head">${iconHTML("alert")}<span>SQL error</span><span class="grow"></span>
          <button class="btn btn-sm btn-ghost" data-copy-error>${iconHTML("copy")} Copy error</button>
        </div>
        <pre>${escapeHtml(error.message)}</pre>
        ${error.statement ? `<pre class="err-stmt">${escapeHtml(error.statement)}</pre>` : ""}
        ${/read.?only|query_only/i.test(error.message) ? `<div class="result-hint">This connection is read-only. Reopen the database without the read-only option to make changes.</div>` : ""}
      </div>
    `;
    tab.refs.resultBody.querySelector("[data-copy-error]").addEventListener("click", () => {
      copyText(error.message);
      toast({ type: "success", title: "Error copied" });
    });
    tab.refs.resultFooter.innerHTML = `
      <span class="footer-chip warn">${iconHTML("alert")} Query failed after ${formatDuration(error.elapsedMs)}</span>
      <span class="grow"></span>
      <span>${escapeHtml(state.status.database?.name || "")}</span>
    `;
    tab.lastSummary = "Query failed";
  }
}

function createResultActions(tab, entry) {
  const exportButton = el("button", { class: "btn btn-sm btn-soft", title: "Export result as CSV" });
  exportButton.innerHTML = `${iconHTML("download")}<span>CSV</span>`;
  exportButton.addEventListener("click", () => {
    const csv = toCSV(entry.columns, entry.rows);
    downloadBlob(csv, `query-result-${timestampForFile()}.csv`, "text/csv;charset=utf-8");
    toast({ type: "success", title: "Result exported", message: "CSV download started." });
  });
  const copyButton = el("button", { class: "icon-btn", title: "Copy result to clipboard" });
  copyButton.innerHTML = iconHTML("copy");
  copyButton.addEventListener("click", async () => {
    const header = entry.columns.map((column) => column.name).join("\t");
    const body = entry.rows.map((row) => row.map((value) => value == null ? "" : value).join("\t")).join("\n");
    await copyText(`${header}\n${body}`);
    toast({ type: "success", title: "Copied", message: "Result copied as tab-separated values." });
  });
  tab.refs.resultActions.append(exportButton, copyButton);
}

/* ---------------------------------------------------------------- data view */
function createDataView(tab, tableName) {
  tab.tableName = tableName;
  tab.pageSize = Number(state.settings.pageSize) || 100;
  tab.offset = 0;
  tab.orderBy = null;
  tab.orderDir = "asc";
  tab.search = "";
  tab.loaded = false;
  tab.requestId = 0;
  tab.relation = null;

  const relation = [...state.schema.tables, ...state.schema.views].find((entry) => entry.name === tableName);
  const kind = relation?.type || "table";
  tab.kind = kind;

  const view = el("div", { class: "data-view" });
  view.innerHTML = `
    <div class="data-toolbar">
      <div class="data-title">
        ${iconHTML(kind === "view" ? "view" : "table", "icon obj-icon")}
        <span class="data-name" title="${escapeHtml(tableName)}">${escapeHtml(tableName)}</span>
        <span class="data-kind">${kind}</span>
        <span class="data-count" data-count></span>
      </div>
      <span class="grow"></span>
      <div class="data-search">
        ${iconHTML("search")}
        <input type="text" placeholder="Search this ${kind}…" data-search autocomplete="off" spellcheck="false">
      </div>
      <button class="icon-btn" data-refresh title="Refresh data">${iconHTML("refresh")}</button>
      <button class="btn btn-sm btn-soft" data-query title="Open in SQL editor">${iconHTML("code")} Query</button>
      <button class="btn btn-sm btn-soft" data-export title="Export full table to CSV">${iconHTML("download")} CSV</button>
    </div>
    <div class="grid-wrap" data-grid></div>
    <div class="pagination" data-pagination>
      <span class="page-info" data-page-info>Loading…</span>
      <span class="grow"></span>
      <label class="toolbar-field">Rows
        <select class="select-xs page-size" data-page-size>
          <option value="50">50</option>
          <option value="100">100</option>
          <option value="250">250</option>
          <option value="500">500</option>
        </select>
      </label>
      <button class="pager-btn" data-first title="First page">${iconHTML("chevron")}</button>
      <button class="pager-btn" data-prev title="Previous page">${iconHTML("chevron")}</button>
      <button class="pager-btn" data-next title="Next page">${iconHTML("chevron")}</button>
      <button class="pager-btn" data-last title="Last page">${iconHTML("chevron")}</button>
    </div>
  `;
  tab.contentEl.append(view);
  tab.view = view;
  tab.dataRefs = {
    gridWrap: view.querySelector("[data-grid]"),
    count: view.querySelector("[data-count]"),
    search: view.querySelector("[data-search]"),
    pageInfo: view.querySelector("[data-page-info]"),
    pageSize: view.querySelector("[data-page-size]"),
    first: view.querySelector("[data-first]"),
    prev: view.querySelector("[data-prev]"),
    next: view.querySelector("[data-next]"),
    last: view.querySelector("[data-last]"),
    export: view.querySelector("[data-export]"),
    refresh: view.querySelector("[data-refresh]"),
  };
  tab.dataRefs.pageSize.value = String(tab.pageSize);
  tab.dataRefs.first.querySelector(".icon").style.transform = "rotate(180deg)";
  tab.dataRefs.prev.querySelector(".icon").style.transform = "rotate(180deg)";

  tab.grid = new DataGrid(tab.dataRefs.gridWrap, {
    onSort: ({ name, dir }) => {
      tab.orderBy = name;
      tab.orderDir = dir;
      tab.offset = 0;
      loadTableData(tab, { resetScroll: true });
    },
    onInspect: (value, column, row) => inspectCell(value, column, tab.offset + row + 1),
  });

  const onSearch = debounce(() => {
    tab.search = tab.dataRefs.search.value.trim();
    tab.offset = 0;
    loadTableData(tab, { resetScroll: true });
  }, 320);
  tab.dataRefs.search.addEventListener("input", onSearch);
  tab.dataRefs.search.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      tab.dataRefs.search.value = "";
      onSearch();
    }
  });

  tab.dataRefs.pageSize.addEventListener("change", () => {
    tab.pageSize = Number(tab.dataRefs.pageSize.value) || 100;
    state.settings.pageSize = tab.pageSize;
    saveSettings();
    tab.offset = 0;
    loadTableData(tab, { resetScroll: true });
  });
  tab.dataRefs.refresh.addEventListener("click", () => loadTableData(tab, { resetScroll: false }));
  tab.dataRefs.first.addEventListener("click", () => goToPage(tab, 0));
  tab.dataRefs.prev.addEventListener("click", () => goToPage(tab, Math.max(0, tab.offset - tab.pageSize)));
  tab.dataRefs.next.addEventListener("click", () => {
    if (!canGoNext(tab)) return;
    goToPage(tab, tab.offset + tab.pageSize);
  });
  tab.dataRefs.last.addEventListener("click", () => {
    if (tab.total == null) return;
    const lastOffset = Math.max(0, Math.floor((tab.total - 1) / tab.pageSize) * tab.pageSize);
    goToPage(tab, lastOffset);
  });
  tab.dataRefs.export.addEventListener("click", async () => {
    try {
      tab.dataRefs.export.disabled = true;
      setBusy(true);
      setStatusMessage(`Exporting ${tableName}…`);
      await api.exportTable(tableName);
      toast({ type: "success", title: "Export ready", message: `${tableName}.csv was downloaded.` });
      setStatusMessage("Export complete");
    } catch (error) {
      notifyError(error, "Export failed");
    } finally {
      tab.dataRefs.export.disabled = false;
      setBusy(false);
    }
  });
  view.querySelector("[data-query]").addEventListener("click", () => {
    openQueryTab(buildSelect(tableName, 200));
  });

  tab.dataRefs.gridWrap.innerHTML = `<div class="grid-loading-overlay" style="position:static"><div class="loading-card"><span class="spinner"></span>Loading ${escapeHtml(tableName)}…</div></div>`;
}

function goToPage(tab, offset) {
  if (offset === tab.offset && tab.loaded) return;
  tab.offset = Math.max(0, offset);
  loadTableData(tab, { resetScroll: true });
}

function canGoNext(tab) {
  if (tab.total != null) return tab.offset + tab.pageSize < tab.total;
  return (tab.dataRows?.length || 0) >= tab.pageSize;
}

async function loadTableData(tab, { resetScroll = false } = {}) {
  const requestId = ++tab.requestId;
  tab.dataRefs.gridWrap.querySelector(".grid-loading-overlay")?.remove();
  const overlay = el("div", { class: "grid-loading-overlay", style: "position:absolute" });
  overlay.innerHTML = `<div class="loading-card"><span class="spinner"></span>Loading rows…</div>`;
  if (!tab.dataRows) {
    tab.dataRefs.gridWrap.innerHTML = "";
    tab.dataRefs.gridWrap.append(overlay);
  } else {
    tab.dataRefs.gridWrap.append(overlay);
  }
  setBusy(true);
  try {
    const response = await api.tableData({
      name: tab.tableName,
      limit: tab.pageSize,
      offset: tab.offset,
      search: tab.search,
      orderBy: tab.orderBy,
      orderDir: tab.orderDir,
    });
    if (requestId !== tab.requestId) return;
    tab.dataRows = response.rows;
    tab.columns = response.columns;
    tab.total = response.total;
    tab.loaded = true;
    if (resetScroll) tab.dataRefs.gridWrap.scrollTop = 0;
    overlay.remove();
    if (!response.rows.length) {
      gridEmpty(
        tab.dataRefs.gridWrap,
        tab.search ? "search" : "table",
        tab.search ? "No matching rows" : "This table is empty",
        tab.search ? `Nothing matches “${tab.search}”.` : "Import data or insert rows to populate it.",
      );
    } else {
      tab.grid.setData({
        columns: response.columns,
        rows: response.rows,
        offset: response.offset,
        sortable: true,
        orderBy: tab.orderBy,
        orderDir: tab.orderDir,
        resetScroll: false,
      });
    }
    const count = response.total == null ? `${formatNumber(response.rows.length)}+ rows` : `${formatNumber(response.total)} rows`;
    tab.dataRefs.count.textContent = `· ${count}`;
    tab.total = response.total;
    tab.offset = response.offset;
    updatePagination(tab);
    setStatusMessage(`${tab.tableName} · ${count} · ${formatDuration(response.elapsedMs)}`);
    updateStatusBar();
  } catch (error) {
    if (requestId !== tab.requestId) return;
    overlay.remove();
    gridEmpty(tab.dataRefs.gridWrap, "alert", "Could not load data", error.message || String(error));
    notifyError(error, "Could not load table");
  } finally {
    if (requestId === tab.requestId) setBusy(false);
  }
}

function updatePagination(tab) {
  const shownFrom = tab.dataRows?.length ? tab.offset + 1 : 0;
  const shownTo = tab.offset + (tab.dataRows?.length || 0);
  const totalText = tab.total == null ? "" : ` of ${formatNumber(tab.total)}`;
  tab.dataRefs.pageInfo.textContent = tab.dataRows?.length
    ? `${formatNumber(shownFrom)}–${formatNumber(shownTo)}${totalText}`
    : tab.search ? "No matches" : "No rows";
  tab.dataRefs.first.disabled = tab.offset <= 0;
  tab.dataRefs.prev.disabled = tab.offset <= 0;
  tab.dataRefs.next.disabled = !canGoNext(tab);
  tab.dataRefs.last.disabled = tab.total == null || tab.offset + tab.pageSize >= tab.total;
  const pages = tab.total == null ? null : Math.max(1, Math.ceil(tab.total / tab.pageSize));
  tab.dataRefs.pageInfo.title = pages ? `Page ${Math.floor(tab.offset / tab.pageSize) + 1} of ${pages}` : "";
}



