/** Scrollable data grid used by query results and the table browser. */

import {
  escapeHtml,
  formatNumber,
  iconEl,
  copyText,
  openContextMenu,
} from "./utils.js";

const CHUNK = 250;

function displayCell(value) {
  if (value == null) return { text: "NULL", cls: "null" };
  if (typeof value === "number") {
    const text = Number.isInteger(value) ? formatNumber(value) : String(value);
    return { text, cls: "num" };
  }
  const text = String(value);
  if (/^<BLOB \d[\d,]* bytes>$/.test(text)) return { text, cls: "blob" };
  const clipped = text.length > 600 ? `${text.slice(0, 600)}…` : text;
  return { text: clipped, cls: "" };
}

function inferTypes(columns, rows) {
  return columns.map((column, index) => {
    if (column.type) return column.type.toUpperCase();
    let type = "";
    for (let r = 0; r < Math.min(rows.length, 40); r += 1) {
      const value = rows[r]?.[index];
      if (value == null) continue;
      if (typeof value === "number") {
        type = Number.isInteger(value) ? "INTEGER" : "REAL";
        break;
      }
      if (typeof value === "string" && /^<BLOB /.test(value)) {
        type = "BLOB";
        break;
      }
      type = "TEXT";
      break;
    }
    return type;
  });
}

export class DataGrid {
  constructor(wrap, options = {}) {
    this.wrap = wrap;
    this.options = options;
    this.columns = [];
    this.rows = [];
    this.types = [];
    this.rendered = 0;
    this.offset = 0;
    this.orderBy = null;
    this.orderDir = "asc";
    this.sortable = false;
    this.selectedCell = null;
    this.rendering = false;
    this.wrap.tabIndex = 0;
    this.wrap.addEventListener("scroll", () => this._onScroll(), { passive: true });
    this.wrap.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c" && this.selectedCell) {
        const { row, col } = this.selectedCell;
        copyText(this.rows[row]?.[col] ?? "");
      }
    });
    this.wrap.addEventListener("click", (event) => {
      const td = event.target.closest?.("td[data-row]");
      if (td) this._selectCell(td);
    });
    this.wrap.addEventListener("dblclick", (event) => {
      const td = event.target.closest?.("td[data-row]");
      if (!td) return;
      const row = Number(td.dataset.row);
      const col = Number(td.dataset.col);
      this.options.onInspect?.(this.rows[row]?.[col], this.columns[col], row, col);
    });
    this.wrap.addEventListener("contextmenu", (event) => {
      const td = event.target.closest?.("td[data-row]");
      if (!td) return;
      event.preventDefault();
      this._selectCell(td);
      const row = Number(td.dataset.row);
      const col = Number(td.dataset.col);
      const value = this.rows[row]?.[col];
      openContextMenu(event.clientX, event.clientY, [
        { label: "Inspect cell", icon: "search", action: () => this.options.onInspect?.(value, this.columns[col], row, col) },
        { label: "Copy value", icon: "copy", action: () => copyText(value ?? "") },
        { separator: true },
        {
          label: "Copy row as JSON",
          icon: "code",
          action: () => {
            const object = {};
            this.columns.forEach((column, index) => {
              object[column.name ?? index] = this.rows[row]?.[index] ?? null;
            });
            copyText(JSON.stringify(object, null, 2));
          },
        },
      ]);
    });
  }

  setData({ columns, rows, offset = 0, sortable = false, orderBy = null, orderDir = "asc", resetScroll = true }) {
    this.columns = columns || [];
    this.rows = rows || [];
    this.types = inferTypes(this.columns, this.rows);
    this.offset = offset;
    this.sortable = Boolean(sortable);
    this.orderBy = orderBy;
    this.orderDir = orderDir;
    this.rendered = 0;
    this.selectedCell = null;

    const head = this._buildHead();
    this.wrap.innerHTML = `
      <table class="data-grid">
        <thead>${head}</thead>
        <tbody></tbody>
      </table>
    `;
    this.tbody = this.wrap.querySelector("tbody");
    this.wrap.querySelectorAll("th.sortable").forEach((th) => {
      th.addEventListener("click", (event) => {
        event.preventDefault();
        const name = th.dataset.name;
        let dir = "asc";
        if (this.orderBy === name) dir = this.orderDir === "asc" ? "desc" : "asc";
        this.options.onSort?.({ name, dir });
      });
    });
    this._renderRows();
    if (resetScroll) this.wrap.scrollTop = 0;
  }

  _buildHead() {
    const cells = this.columns.map((column, index) => {
      const name = column.name ?? String(column);
      const type = this.types[index] || "";
      const active = this.orderBy === name;
      let sortIcon = "";
      if (this.sortable && active) {
        sortIcon = `<svg class="icon sort-ind" aria-hidden="true"><use href="#i-${this.orderDir === "asc" ? "arrow-up" : "arrow-down"}"></use></svg>`;
      }
      const pk = column.pk ? `<svg class="icon pk-ind" aria-hidden="true"><use href="#i-key"></use></svg>` : "";
      const title = type ? `${name} · ${type}` : name;
      return `<th class="${this.sortable ? "sortable" : ""}" data-name="${escapeHtml(name)}" title="${escapeHtml(title)}">
        <span class="th-inner">
          <span class="th-top"><span class="th-name">${escapeHtml(name)}</span>${pk}${sortIcon}</span>
          <span class="th-type">${escapeHtml(type)}</span>
        </span>
      </th>`;
    }).join("");
    return `<th class="row-index" title="Row">#</th>${cells}`;
  }

  _renderRows() {
    if (!this.tbody) return;
    const start = this.rendered;
    const end = Math.min(start + CHUNK, this.rows.length);
    if (start >= end) return;
    const html = [];
    for (let r = start; r < end; r += 1) {
      const row = this.rows[r];
      const cells = row.map((value, col) => {
        const cell = displayCell(value);
        return `<td class="${cell.cls}" data-row="${r}" data-col="${col}" title="Double-click to inspect"><span class="cell-inner">${escapeHtml(cell.text)}</span></td>`;
      }).join("");
      html.push(`<tr><td class="row-index">${formatNumber(this.offset + r + 1)}</td>${cells}</tr>`);
    }
    this.tbody.insertAdjacentHTML("beforeend", html.join(""));
    this.rendered = end;
    this._selectCell(this.wrap.querySelector("td.cell-selected"));
  }

  _selectCell(td) {
    if (!td) return;
    this.wrap.querySelectorAll("td.cell-selected").forEach((cell) => cell.classList.remove("cell-selected"));
    td.classList.add("cell-selected");
    this.selectedCell = {
      row: Number(td.dataset.row),
      col: Number(td.dataset.col),
    };
  }

  _onScroll() {
    if (this.rendering) return;
    const { scrollTop, clientHeight, scrollHeight } = this.wrap;
    if (scrollTop + clientHeight >= scrollHeight - 600 && this.rendered < this.rows.length) {
      this.rendering = true;
      requestAnimationFrame(() => {
        this._renderRows();
        this.rendering = false;
      });
    }
  }

  scrollToTop() {
    this.wrap.scrollTop = 0;
  }

  get hasMore() {
    return this.rendered < this.rows.length;
  }
}

export function gridEmpty(container, icon, title, message) {
  container.innerHTML = `<div class="grid-empty">
    <div class="ge-icon"><svg class="icon"><use href="#i-${icon}"></use></svg></div>
    <h3>${escapeHtml(title)}</h3>
    <p>${escapeHtml(message)}</p>
  </div>`;
}
