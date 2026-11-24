/** Lightweight SQL editor: syntax highlighting, line numbers, autocomplete. */

import { escapeHtml, quoteIdent, needsQuote } from "./utils.js";

const KEYWORDS = `SELECT FROM WHERE AND OR NOT NULL IS IN LIKE GLOB REGEXP MATCH ESCAPE
BETWEEN EXISTS CASE WHEN THEN ELSE END CAST AS DISTINCT ALL ANY SOME GROUP BY ORDER
ASC DESC HAVING LIMIT OFFSET JOIN INNER LEFT RIGHT FULL OUTER CROSS ON USING NATURAL
UNION INTERSECT EXCEPT INSERT INTO VALUES UPDATE SET DELETE CREATE TABLE VIEW INDEX
TRIGGER VIRTUAL TEMP TEMPORARY IF REPLACE ALTER RENAME ADD COLUMN DROP PRAGMA BEGIN
COMMIT ROLLBACK SAVEPOINT RELEASE TRANSACTION EXPLAIN QUERY PLAN WITH RECURSIVE OVER
PARTITION WINDOW FILTER ROWS RANGE PRECEDING FOLLOWING CURRENT ROW UNBOUNDED GROUPS
EXCLUDE TIES NO OTHERS RETURNING CONFLICT DO NOTHING ABORT FAIL IGNORE DEFERRED
IMMEDIATE EXCLUSIVE AUTOINCREMENT PRIMARY KEY FOREIGN REFERENCES UNIQUE CHECK DEFAULT
COLLATE WITHOUT ROWID STRICT GENERATED ALWAYS STORED MATERIALIZED TRUE FALSE
INTEGER TEXT REAL BLOB NUMERIC BOOLEAN DATE DATETIME VARCHAR CHAR CLOB DECIMAL DOUBLE
FLOAT BIGINT SMALLINT INT TINYINT NVARCHAR NCHAR`.split(/\s+/);

const FUNCTIONS = `count sum avg min max total group_concat coalesce ifnull nullif
length lower upper trim ltrim rtrim replace substr substring instr printf format round
abs random date time datetime julianday strftime unixepoch json json_array json_object
json_extract json_insert json_replace json_remove json_set json_type json_valid
json_group_array json_group_object json_each json_tree row_number rank dense_rank ntile
lag lead first_value last_value nth_value cume_dist percent_rank typeof hex quote
zeroblob last_insert_rowid changes total_changes sqlite_version char unicode concat
concat_ws iif likelihood likely unlikely`.split(/\s+/);

const KEYWORD_SET = new Set(KEYWORDS.map((word) => word.toLowerCase()));
const FUNCTION_SET = new Set(FUNCTIONS.map((word) => word.toLowerCase()));

const WORD_START = /[\p{L}_$]/u;
const WORD_CHAR = /[\p{L}\p{N}_$]/u;

function tokenHTML(sql) {
  let html = "";
  let i = 0;
  const length = sql.length;

  while (i < length) {
    const ch = sql[i];
    const next = sql[i + 1];

    if (ch === "-" && next === "-") {
      let end = sql.indexOf("\n", i);
      if (end === -1) end = length;
      html += `<span class="tok-comment">${escapeHtml(sql.slice(i, end))}</span>`;
      i = end;
      continue;
    }
    if (ch === "/" && next === "*") {
      let end = sql.indexOf("*/", i + 2);
      end = end === -1 ? length : end + 2;
      html += `<span class="tok-comment">${escapeHtml(sql.slice(i, end))}</span>`;
      i = end;
      continue;
    }
    if (ch === "'") {
      let end = i + 1;
      while (end < length) {
        if (sql[end] === "'") {
          if (sql[end + 1] === "'") { end += 2; continue; }
          end += 1;
          break;
        }
        end += 1;
      }
      html += `<span class="tok-string">${escapeHtml(sql.slice(i, end))}</span>`;
      i = end;
      continue;
    }
    if (ch === '"' || ch === "`") {
      const quote = ch;
      let end = i + 1;
      while (end < length) {
        if (sql[end] === quote) {
          if (sql[end + 1] === quote) { end += 2; continue; }
          end += 1;
          break;
        }
        end += 1;
      }
      html += `<span class="tok-quoted">${escapeHtml(sql.slice(i, end))}</span>`;
      i = end;
      continue;
    }
    if (ch === "[") {
      let end = sql.indexOf("]", i + 1);
      end = end === -1 ? length : end + 1;
      html += `<span class="tok-quoted">${escapeHtml(sql.slice(i, end))}</span>`;
      i = end;
      continue;
    }
    if (ch === "?" || ((ch === ":" || ch === "@" || ch === "$") && WORD_START.test(next || ""))) {
      let end = i + 1;
      while (end < length && /[\w$]/.test(sql[end])) end += 1;
      if (end === i + 1) end = i + 1;
      html += `<span class="tok-param">${escapeHtml(sql.slice(i, end))}</span>`;
      i = end;
      continue;
    }
    if (/[0-9]/.test(ch) || (ch === "." && /[0-9]/.test(next || ""))) {
      const match = /^(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?/.exec(sql.slice(i));
      const text = match ? match[0] : ch;
      html += `<span class="tok-number">${escapeHtml(text)}</span>`;
      i += text.length;
      continue;
    }
    if (WORD_START.test(ch)) {
      let end = i + 1;
      while (end < length && WORD_CHAR.test(sql[end])) end += 1;
      const word = sql.slice(i, end);
      const lower = word.toLowerCase();
      let look = end;
      while (look < length && /\s/.test(sql[look])) look += 1;
      const isCall = sql[look] === "(";
      let cls = "tok-ident";
      if (KEYWORD_SET.has(lower)) cls = "tok-keyword";
      else if (FUNCTION_SET.has(lower) || isCall) cls = "tok-function";
      html += `<span class="${cls}">${escapeHtml(word)}</span>`;
      i = end;
      continue;
    }
    if ("(),.;".includes(ch)) {
      html += `<span class="tok-punct">${escapeHtml(ch)}</span>`;
      i += 1;
      continue;
    }
    if ("=<>!+-*/%|&^~".includes(ch)) {
      html += `<span class="tok-operator">${escapeHtml(ch)}</span>`;
      i += 1;
      continue;
    }
    html += escapeHtml(ch);
    i += 1;
  }
  return html;
}

function contextAtCursor(text, position) {
  let state = "code";
  for (let i = 0; i < position; i += 1) {
    const ch = text[i];
    const next = text[i + 1];
    if (state === "code") {
      if (ch === "-" && next === "-") { state = "line-comment"; i += 1; }
      else if (ch === "/" && next === "*") { state = "block-comment"; i += 1; }
      else if (ch === "'") state = "string";
      else if (ch === '"') state = "double";
      else if (ch === "`") state = "backtick";
      else if (ch === "[") state = "bracket";
    } else if (state === "line-comment") {
      if (ch === "\n") state = "code";
    } else if (state === "block-comment") {
      if (ch === "*" && next === "/") { state = "code"; i += 1; }
    } else if (state === "string") {
      if (ch === "'") {
        if (next === "'") i += 1;
        else state = "code";
      }
    } else if (state === "double") {
      if (ch === '"') {
        if (next === '"') i += 1;
        else state = "code";
      }
    } else if (state === "backtick") {
      if (ch === "`") state = "code";
    } else if (state === "bracket") {
      if (ch === "]") state = "code";
    }
  }
  return state;
}

export class SqlEditor {
  constructor(container, options = {}) {
    this.container = container;
    this.options = options;
    this.schemaTables = [];
    this.schemaViews = [];
    this.schemaColumns = [];
    this.acItems = [];
    this.acIndex = -1;
    this.acPrefix = "";
    this.acStart = 0;
    this.acOpen = false;
    this._build();
    if (options.value) this.setValue(options.value);
    this.refresh();
  }

  _build() {
    this.container.classList.add("sql-editor-host");
    this.root = document.createElement("div");
    this.root.className = "sql-editor";

    this.gutter = document.createElement("div");
    this.gutter.className = "sql-gutter";
    this.gutterInner = document.createElement("div");
    this.gutterInner.className = "sql-gutter-inner";
    this.gutter.append(this.gutterInner);

    this.code = document.createElement("div");
    this.code.className = "sql-code";
    this.highlight = document.createElement("pre");
    this.highlight.className = "sql-highlight";
    this.highlight.setAttribute("aria-hidden", "true");
    this.highlightCode = document.createElement("code");
    this.highlight.append(this.highlightCode);

    this.textarea = document.createElement("textarea");
    this.textarea.className = "sql-input";
    this.textarea.spellcheck = false;
    this.textarea.autocomplete = "off";
    this.textarea.autocapitalize = "off";
    this.textarea.setAttribute("aria-label", "SQL editor");
    this.textarea.placeholder = this.options.placeholder || "Write SQL here…";

    this.code.append(this.highlight, this.textarea);
    this.root.append(this.gutter, this.code);
    this.container.append(this.root);

    this.textarea.addEventListener("input", () => {
      this.refresh();
      this.options.onChange?.(this.getValue());
      this._autoSuggest();
    });
    this.textarea.addEventListener("scroll", () => this._syncScroll());
    this.textarea.addEventListener("keydown", (event) => this._onKeyDown(event));
    this.textarea.addEventListener("click", () => this._hideAutocomplete());
    this.textarea.addEventListener("blur", () => {
      setTimeout(() => this._hideAutocomplete(), 120);
    });
    this.textarea.addEventListener("keyup", (event) => {
      if (["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End"].includes(event.key)) {
        this.options.onCursor?.();
      }
    });
    this.container.addEventListener("mousedown", (event) => {
      if (event.target === this.container) this.textarea.focus();
    });
  }

  setSchema(schema) {
    const tables = schema?.tables || [];
    const views = schema?.views || [];
    this.schemaTables = tables.map((table) => ({
      name: table.name,
      hint: `${(table.columns || []).length} cols`,
    }));
    this.schemaViews = views.map((view) => ({
      name: view.name,
      hint: `${(view.columns || []).length} cols`,
    }));
    const columns = new Map();
    for (const table of [...tables, ...views]) {
      for (const column of table.columns || []) {
        const key = column.name.toLowerCase();
        if (!columns.has(key)) {
          columns.set(key, { name: column.name, tables: new Set([table.name]) });
        } else {
          columns.get(key).tables.add(table.name);
        }
      }
    }
    this.schemaColumns = [...columns.values()].map((entry) => {
      const owners = [...entry.tables];
      return {
        name: entry.name,
        hint: owners.slice(0, 2).join(", ") + (owners.length > 2 ? "…" : ""),
      };
    });
  }

  getValue() {
    return this.textarea.value;
  }

  setValue(value, focus = false) {
    this.textarea.value = value ?? "";
    this.refresh();
    this.options.onChange?.(this.getValue());
    if (focus) this.focus();
  }

  focus() {
    this.textarea.focus();
    const end = this.textarea.value.length;
    this.textarea.setSelectionRange(end, end);
    this._syncScroll();
  }

  getSelection() {
    const { selectionStart, selectionEnd, value } = this.textarea;
    if (selectionStart === selectionEnd) return "";
    return value.slice(selectionStart, selectionEnd);
  }

  insertAtCursor(text) {
    const { selectionStart, selectionEnd, value } = this.textarea;
    this.textarea.value = value.slice(0, selectionStart) + text + value.slice(selectionEnd);
    const position = selectionStart + text.length;
    this.textarea.setSelectionRange(position, position);
    this.refresh();
    this.options.onChange?.(this.getValue());
    this.focus();
    this._hideAutocomplete();
  }

  runSelectionOrAll() {
    const selected = this.getSelection();
    const sql = (selected || this.getValue()).trim();
    if (sql) this.options.onRun?.(sql);
  }

  getCursorInfo() {
    const position = this.textarea.selectionStart;
    const before = this.textarea.value.slice(0, position);
    const lines = before.split("\n");
    return { line: lines.length, column: lines[lines.length - 1].length + 1 };
  }

  refresh() {
    const value = this.textarea.value;
    this.highlightCode.innerHTML = tokenHTML(value) + (value.endsWith("\n") ? " " : "");
    const lineCount = value.split("\n").length;
    if (this._lineCount !== lineCount) {
      this.gutterInner.textContent = Array.from({ length: lineCount }, (_, i) => i + 1).join("\n");
      this._lineCount = lineCount;
    }
    this._syncScroll();
  }

  _syncScroll() {
    const { scrollTop, scrollLeft } = this.textarea;
    this.highlight.style.transform = `translate(${-scrollLeft}px, ${-scrollTop}px)`;
    this.gutterInner.style.transform = `translateY(${-scrollTop}px)`;
  }

  // -------------------------------------------------------------- suggestions
  setSchemaColumnHint() { /* compatibility no-op */ }

  _autoSuggest() {
    const value = this.textarea.value;
    const position = this.textarea.selectionStart;
    const before = value.slice(0, position);
    const word = /[\p{L}\p{N}_$]*$/u.exec(before)?.[0] || "";
    if (word.length < 2) {
      this._hideAutocomplete();
      return;
    }
    if (contextAtCursor(value, position) !== "code") {
      this._hideAutocomplete();
      return;
    }
    this._showAutocomplete(word, false);
  }

  _suggestions(prefix, forceAll) {
    const lower = prefix.toLowerCase();
    const candidates = [];
    const seen = new Set();
    const push = (name, kind, hint) => {
      const key = `${kind}:${name.toLowerCase()}`;
      if (seen.has(key)) return;
      seen.add(key);
      candidates.push({ name, kind, hint });
    };

    for (const column of this.schemaColumns) {
      if (!lower || column.name.toLowerCase().startsWith(lower)) push(column.name, "column", column.hint);
    }
    for (const table of this.schemaTables) {
      if (!lower || table.name.toLowerCase().startsWith(lower)) push(table.name, "table", table.hint);
    }
    for (const view of this.schemaViews) {
      if (!lower || view.name.toLowerCase().startsWith(lower)) push(view.name, "view", view.hint);
    }
    for (const keyword of KEYWORDS) {
      if (!lower || keyword.toLowerCase().startsWith(lower)) push(keyword, "keyword", "");
    }
    for (const fn of FUNCTIONS) {
      if (!lower || fn.toLowerCase().startsWith(lower)) push(fn, "function", "");
    }

    if (lower) {
      const order = { column: 0, table: 1, view: 2, keyword: 3, function: 4 };
      candidates.sort((a, b) => {
        if (order[a.kind] !== order[b.kind]) return order[a.kind] - order[b.kind];
        return a.name.localeCompare(b.name);
      });
    }
    return candidates.slice(0, 80);
  }

  _showAutocomplete(prefix = "", forceAll = false) {
    const position = this.textarea.selectionStart;
    const before = this.textarea.value.slice(0, position);
    let start = position;
    while (start > 0 && /[\p{L}\p{N}_$]/u.test(before[start - 1])) start -= 1;
    const actualPrefix = before.slice(start, position);
    if (!forceAll && actualPrefix.length < 2) {
      this._hideAutocomplete();
      return;
    }
    const items = this._suggestions(actualPrefix, forceAll);
    if (!items.length) {
      this._hideAutocomplete();
      return;
    }

    this.acItems = items;
    this.acIndex = 0;
    this.acPrefix = actualPrefix;
    this.acStart = start;
    this.acOpen = true;

    if (!this.dropdown) {
      this.dropdown = document.createElement("div");
      this.dropdown.className = "autocomplete";
      this.dropdown.setAttribute("role", "listbox");
      this.root.append(this.dropdown);
      this.dropdown.addEventListener("mousedown", (event) => {
        const item = event.target.closest(".ac-item");
        if (!item) return;
        event.preventDefault();
        const index = Number(item.dataset.index);
        if (Number.isInteger(index)) this._acceptAutocomplete(index);
      });
    }
    this._renderAutocomplete();
    this.dropdown.hidden = false;
    this._positionAutocomplete();
  }

  _renderAutocomplete() {
    const kindLabel = { table: "TBL", view: "VW", column: "COL", keyword: "KW", function: "FN" };
    const rows = this.acItems.map((item, index) => `
      <button type="button" class="ac-item k-${item.kind} ${index === this.acIndex ? "active" : ""}" data-index="${index}" role="option">
        <span class="ac-kind">${kindLabel[item.kind] || "·"}</span>
        <span class="ac-text">${escapeHtml(item.name)}</span>
        ${item.hint ? `<span class="ac-hint">${escapeHtml(item.hint)}</span>` : ""}
      </button>
    `).join("");
    this.dropdown.innerHTML = `<div class="autocomplete-head"><span>Suggestions</span><span>Tab to insert</span></div>${rows}`;
    const active = this.dropdown.querySelector(".ac-item.active");
    active?.scrollIntoView({ block: "nearest" });
  }

  _positionAutocomplete() {
    const { selectionStart } = this.textarea;
    const mirror = document.createElement("div");
    const style = getComputedStyle(this.textarea);
    const properties = [
      "fontFamily", "fontSize", "fontWeight", "letterSpacing", "lineHeight",
      "paddingTop", "paddingRight", "paddingBottom", "paddingLeft",
      "borderTopWidth", "borderRightWidth", "borderBottomWidth", "borderLeftWidth",
      "whiteSpace", "wordSpacing", "tabSize", "textIndent",
    ];
    for (const property of properties) mirror.style[property] = style[property];
    mirror.style.position = "absolute";
    mirror.style.visibility = "hidden";
    mirror.style.left = "-99999px";
    mirror.style.top = "0";
    mirror.style.width = `${this.textarea.offsetWidth}px`;
    mirror.style.whiteSpace = "pre";
    mirror.style.overflow = "hidden";
    mirror.textContent = this.textarea.value.slice(0, selectionStart);
    const marker = document.createElement("span");
    marker.textContent = "\u200b";
    mirror.append(marker);
    document.body.append(mirror);
    const markerRect = marker.getBoundingClientRect();
    mirror.remove();

    const rootRect = this.root.getBoundingClientRect();
    const lineHeight = parseFloat(style.lineHeight) || 20;
    let left = markerRect.left - rootRect.left - this.textarea.scrollLeft;
    let top = markerRect.top - rootRect.top - this.textarea.scrollTop + lineHeight + 2;
    const dropdownWidth = this.dropdown.offsetWidth || 300;
    const dropdownHeight = this.dropdown.offsetHeight || 240;
    left = Math.max(8, Math.min(left, this.root.clientWidth - dropdownWidth - 8));
    if (top + dropdownHeight > this.root.clientHeight) {
      top = Math.max(8, markerRect.top - rootRect.top - this.textarea.scrollTop - dropdownHeight - 4);
    }
    this.dropdown.style.left = `${left}px`;
    this.dropdown.style.top = `${top}px`;
  }

  _hideAutocomplete() {
    if (this.dropdown) this.dropdown.hidden = true;
    this.acOpen = false;
    this.acItems = [];
    this.acIndex = -1;
  }

  _acceptAutocomplete(index = this.acIndex) {
    const item = this.acItems[index];
    if (!item) return;
    const value = this.textarea.value;
    const position = this.textarea.selectionStart;
    let insertion = item.name;
    if ((item.kind === "table" || item.kind === "view" || item.kind === "column") && needsQuote(insertion)) {
      insertion = quoteIdent(insertion);
    }
    const nextValue = value.slice(0, this.acStart) + insertion + value.slice(position);
    this.textarea.value = nextValue;
    const caret = this.acStart + insertion.length;
    this.textarea.setSelectionRange(caret, caret);
    this._hideAutocomplete();
    this.refresh();
    this.options.onChange?.(this.getValue());
    this.textarea.focus();
  }

  _moveAutocomplete(delta) {
    if (!this.acOpen) return false;
    this.acIndex = (this.acIndex + delta + this.acItems.length) % this.acItems.length;
    this._renderAutocomplete();
    return true;
  }

  // -------------------------------------------------------------- key handler
  _onKeyDown(event) {
    const mod = event.ctrlKey || event.metaKey;

    if (mod && event.key === "Enter") {
      event.preventDefault();
      this._hideAutocomplete();
      this.runSelectionOrAll();
      return;
    }
    if (mod && event.code === "Space") {
      event.preventDefault();
      this._showAutocomplete("", true);
      return;
    }
    if (this.acOpen) {
      if (event.key === "ArrowDown") { event.preventDefault(); this._moveAutocomplete(1); return; }
      if (event.key === "ArrowUp") { event.preventDefault(); this._moveAutocomplete(-1); return; }
      if (event.key === "Enter" || event.key === "Tab") {
        event.preventDefault();
        this._acceptAutocomplete();
        return;
      }
      if (event.key === "Escape") { event.preventDefault(); this._hideAutocomplete(); return; }
    }
    if (event.key === "Tab") {
      event.preventDefault();
      this._handleTab(event.shiftKey);
      return;
    }
    if (mod && event.key === "/") {
      event.preventDefault();
      this._toggleComment();
      return;
    }
    if (event.key === "Enter" && !event.shiftKey) {
      const value = this.textarea.value;
      const position = this.textarea.selectionStart;
      if (contextAtCursor(value, position) === "code") {
        const lineStart = value.lastIndexOf("\n", position - 1) + 1;
        const indent = /^[ \t]*/.exec(value.slice(lineStart, position))?.[0] || "";
        if (indent) {
          event.preventDefault();
          const insertion = `\n${indent}`;
          this.insertAtCursor(insertion);
          return;
        }
      }
    }
    if (event.key === "Escape") this._hideAutocomplete();
  }

  _handleTab(shift) {
    const { selectionStart, selectionEnd, value } = this.textarea;
    const hasSelection = selectionStart !== selectionEnd;
    const multiLine = hasSelection && value.slice(selectionStart, selectionEnd).includes("\n");

    if (!multiLine) {
      if (shift) {
        const lineStart = value.lastIndexOf("\n", selectionStart - 1) + 1;
        const beforeCaret = value.slice(lineStart, selectionStart);
        const removable = /^ {1,2}/.exec(beforeCaret);
        if (removable) {
          const count = removable[0].length;
          this.textarea.value = value.slice(0, lineStart) + value.slice(lineStart + count);
          const nextStart = Math.max(lineStart, selectionStart - count);
          const nextEnd = Math.max(nextStart, selectionEnd - count);
          this.textarea.setSelectionRange(nextStart, nextEnd);
          this.refresh();
          this.options.onChange?.(this.getValue());
        }
        return;
      }
      this.textarea.setRangeText("  ", selectionStart, selectionEnd, "end");
      this.refresh();
      this.options.onChange?.(this.getValue());
      return;
    }

    const start = value.lastIndexOf("\n", selectionStart - 1) + 1;
    let end = value.indexOf("\n", selectionEnd);
    if (end === -1) end = value.length;
    const block = value.slice(start, end);
    const lines = block.split("\n");
    let deltaStart = 0;
    let deltaEnd = 0;
    const changed = lines.map((line, index) => {
      if (shift) {
        const removable = line.match(/^ {1,2}/)?.[0].length || 0;
        if (index === 0) deltaStart -= removable;
        deltaEnd -= removable;
        return line.slice(removable);
      }
      if (index === 0) deltaStart += 2;
      deltaEnd += 2;
      return `  ${line}`;
    }).join("\n");
    this.textarea.value = value.slice(0, start) + changed + value.slice(end);
    this.textarea.setSelectionRange(selectionStart + deltaStart, selectionEnd + deltaEnd);
    this.refresh();
    this.options.onChange?.(this.getValue());
  }

  _toggleComment() {
    const { selectionStart, selectionEnd, value } = this.textarea;
    const start = value.lastIndexOf("\n", selectionStart - 1) + 1;
    let end = value.indexOf("\n", selectionEnd);
    if (end === -1) end = value.length;
    const block = value.slice(start, end);
    const lines = block.split("\n");
    const allCommented = lines.every((line) => !line.trim() || line.trimStart().startsWith("--"));
    const changed = lines.map((line) => {
      if (!line.trim()) return line;
      if (allCommented) return line.replace(/^(\s*)--\s?/, "$1");
      return line.replace(/^(\s*)/, "$1-- ");
    }).join("\n");
    this.textarea.value = value.slice(0, start) + changed + value.slice(end);
    this.textarea.setSelectionRange(start, start + changed.length);
    this.refresh();
    this.options.onChange?.(this.getValue());
  }
}

export const SQL_KEYWORDS = KEYWORDS;
