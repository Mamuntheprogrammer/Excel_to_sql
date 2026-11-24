/** Small DOM, formatting and file helpers shared by the app modules. */

export const $ = (selector, root = document) => root.querySelector(selector);
export const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

export function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs || {})) {
    if (value == null || value === false) continue;
    if (key === "class") node.className = value;
    else if (key === "dataset") Object.assign(node.dataset, value);
    else if (key === "style") {
      if (typeof value === "string") node.style.cssText = value;
      else Object.assign(node.style, value);
    }
    else if (key === "html") node.innerHTML = value;
    else if (key.startsWith("on") && typeof value === "function") {
      node.addEventListener(key.slice(2).toLowerCase(), value);
    } else if (value === true) node.setAttribute(key, "");
    else node.setAttribute(key, String(value));
  }
  for (const child of children.flat(Infinity)) {
    if (child == null || child === false) continue;
    node.append(child instanceof Node ? child : document.createTextNode(String(child)));
  }
  return node;
}

const ICONS = new Set([
  "database", "table", "view", "column", "key", "folder", "file", "play",
  "upload", "download", "search", "refresh", "x", "plus", "sun", "moon",
  "help", "history", "chevron", "external", "copy", "check", "alert", "info",
  "trash", "clock", "code", "filter", "home", "more", "arrow-up", "arrow-down",
  "spinner", "sparkle",
]);

export function iconEl(name, cls = "icon") {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("class", cls);
  svg.setAttribute("aria-hidden", "true");
  const use = document.createElementNS("http://www.w3.org/2000/svg", "use");
  use.setAttribute("href", `#i-${ICONS.has(name) ? name : "info"}`);
  svg.append(use);
  return svg;
}

export function iconHTML(name) {
  return `<svg class="icon" aria-hidden="true"><use href="#i-${ICONS.has(name) ? name : "info"}"></use></svg>`;
}

export function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function formatNumber(value) {
  if (value == null || Number.isNaN(Number(value))) return String(value ?? "");
  return Number(value).toLocaleString("en-US");
}

export function formatBytes(bytes) {
  const n = Number(bytes) || 0;
  if (n < 1024) return `${n} B`;
  const units = ["KB", "MB", "GB", "TB"];
  let value = n / 1024;
  let index = 0;
  while (value >= 1024 && index < units.length - 1) {
    value /= 1024;
    index += 1;
  }
  return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[index]}`;
}

export function formatDuration(ms) {
  const n = Number(ms) || 0;
  if (n < 1) return "<1 ms";
  if (n < 1000) return `${Math.round(n)} ms`;
  if (n < 60000) return `${(n / 1000).toFixed(n < 10000 ? 1 : 0)} s`;
  const minutes = Math.floor(n / 60000);
  const seconds = Math.round((n % 60000) / 1000);
  return `${minutes}m ${seconds}s`;
}

export function relativeTime(timestamp) {
  const delta = Date.now() - Number(timestamp || 0);
  if (!Number.isFinite(delta) || delta < 0) return "just now";
  const seconds = Math.round(delta / 1000);
  if (seconds < 45) return "just now";
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  const days = Math.round(hours / 24);
  if (days < 30) return `${days}d ago`;
  return new Date(timestamp).toLocaleDateString();
}

export function formatDateTime(timestamp) {
  if (!timestamp) return "—";
  return new Date(timestamp).toLocaleString(undefined, {
    year: "numeric", month: "short", day: "numeric",
    hour: "2-digit", minute: "2-digit",
  });
}

export function debounce(fn, delay = 250) {
  let timer = null;
  const wrapped = (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
  wrapped.cancel = () => clearTimeout(timer);
  return wrapped;
}

export function sanitizeTableName(value, fallback = "imported") {
  let name = String(value || "").trim();
  name = name.replace(/[^\p{L}\p{N}_$]+/gu, "_").replace(/_+/g, "_").replace(/^_+|_+$/g, "");
  if (!name) name = fallback;
  if (/^\d/.test(name)) name = `t_${name}`;
  return name.slice(0, 60);
}

export function quoteIdent(name) {
  return `"${String(name).replace(/"/g, '""')}"`;
}

export function needsQuote(name) {
  return !/^[A-Za-z_$][\w$]*$/.test(String(name));
}

export function sqlPreview(sql, max = 42) {
  const cleaned = String(sql || "").replace(/\s+/g, " ").trim();
  return cleaned.length > max ? `${cleaned.slice(0, max - 1)}…` : cleaned;
}

export function csvCell(value) {
  if (value == null) return "";
  const text = typeof value === "object" ? JSON.stringify(value) : String(value);
  return /[",\n\r;]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

export function toCSV(columns, rows) {
  const header = columns.map((col) => csvCell(col.name ?? col)).join(",");
  const body = rows.map((row) => row.map(csvCell).join(",")).join("\n");
  return `\uFEFF${header}\n${body}${body ? "\n" : ""}`;
}

export function downloadBlob(content, filename, type = "text/plain;charset=utf-8") {
  const blob = content instanceof Blob ? content : new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

export function timestampForFile(date = new Date()) {
  const pad = (n) => String(n).padStart(2, "0");
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

export async function copyText(text) {
  try {
    await navigator.clipboard.writeText(String(text ?? ""));
    return true;
  } catch {
    const area = document.createElement("textarea");
    area.value = String(text ?? "");
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.append(area);
    area.select();
    let ok = false;
    try { ok = document.execCommand("copy"); } catch { ok = false; }
    area.remove();
    return ok;
  }
}

let contextMenu = null;

export function closeContextMenu() {
  contextMenu?.remove();
  contextMenu = null;
}

export function openContextMenu(x, y, items) {
  closeContextMenu();
  const menu = el("div", { class: "context-menu", role: "menu" });
  for (const item of items) {
    if (item.separator) {
      menu.append(el("div", { class: "cm-sep" }));
      continue;
    }
    if (item.label && !item.action) {
      menu.append(el("div", { class: "cm-label", html: escapeHtml(item.label) }));
      continue;
    }
    const button = el("button", { class: item.danger ? "danger" : "", role: "menuitem" });
    button.append(iconEl(item.icon || "chevron"));
    button.append(document.createTextNode(item.label));
    button.addEventListener("click", () => {
      closeContextMenu();
      item.action?.();
    });
    menu.append(button);
  }
  document.body.append(menu);
  const rect = menu.getBoundingClientRect();
  const left = Math.min(x, window.innerWidth - rect.width - 8);
  const top = Math.min(y, window.innerHeight - rect.height - 8);
  menu.style.left = `${Math.max(8, left)}px`;
  menu.style.top = `${Math.max(8, top)}px`;
  contextMenu = menu;
  setTimeout(() => {
    const handler = (event) => {
      if (!menu.contains(event.target)) {
        closeContextMenu();
        document.removeEventListener("pointerdown", handler, true);
      }
    };
    document.addEventListener("pointerdown", handler, true);
  }, 0);
  return menu;
}

export function loadJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

export function saveJSON(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* ignore quota */ }
}

export function debounceAsync(fn, delay = 200) {
  return debounce(fn, delay);
}
