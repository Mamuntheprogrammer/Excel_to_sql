/** Thin wrapper around the local Flask API. */

const TOKEN = document.querySelector('meta[name="app-token"]')?.content || "";

export class ApiError extends Error {
  constructor(message, status = 0, code = "ERROR", payload = null) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.code = code;
    this.payload = payload;
  }
}

async function request(path, options = {}) {
  const headers = new Headers(options.headers || {});
  headers.set("X-App-Token", TOKEN);
  let body = options.body;
  if (body && !(body instanceof FormData) && typeof body !== "string") {
    headers.set("Content-Type", "application/json");
    body = JSON.stringify(body);
  }
  let response;
  try {
    response = await fetch(path, { ...options, headers, body });
  } catch (error) {
    throw new ApiError(
      "Cannot reach the local server. Is ExcelSQL Studio still running?",
      0,
      "NETWORK_ERROR",
    );
  }
  const contentType = response.headers.get("content-type") || "";
  let payload = null;
  if (contentType.includes("application/json")) {
    try { payload = await response.json(); } catch { payload = null; }
  } else if (options.parseText) {
    payload = await response.text();
  }
  if (!response.ok) {
    throw new ApiError(
      payload?.error || `Request failed (${response.status})`,
      response.status,
      payload?.code || "HTTP_ERROR",
      payload,
    );
  }
  return payload;
}

function upload(path, file, onProgress) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", path);
    xhr.setRequestHeader("X-App-Token", TOKEN);
    xhr.responseType = "json";
    xhr.upload.addEventListener("progress", (event) => {
      if (event.lengthComputable && onProgress) {
        onProgress(Math.round((event.loaded / event.total) * 100));
      }
    });
    xhr.addEventListener("load", () => {
      const payload = xhr.response;
      if (xhr.status >= 200 && xhr.status < 300) {
        resolve(payload);
      } else {
        reject(new ApiError(
          payload?.error || `Upload failed (${xhr.status})`,
          xhr.status,
          payload?.code || "UPLOAD_FAILED",
          payload,
        ));
      }
    });
    xhr.addEventListener("error", () => reject(new ApiError("Upload failed: network error.", 0, "NETWORK_ERROR")));
    xhr.addEventListener("abort", () => reject(new ApiError("Upload cancelled.", 0, "ABORTED")));
    const form = new FormData();
    form.append("file", file, file.name);
    xhr.send(form);
  });
}

async function download(path, filename) {
  const response = await fetch(path, { headers: { "X-App-Token": TOKEN } });
  if (!response.ok) {
    let message = `Download failed (${response.status})`;
    try {
      const payload = await response.json();
      message = payload?.error || message;
    } catch { /* ignore */ }
    throw new ApiError(message, response.status, "DOWNLOAD_FAILED");
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

export const api = {
  token: TOKEN,
  status: () => request("/api/status"),
  schema: () => request("/api/schema"),
  openDatabase: (path, readOnly) => request("/api/database/open", { method: "POST", body: { path, readOnly } }),
  newDatabase: (path) => request("/api/database/new", { method: "POST", body: { path } }),
  demoDatabase: () => request("/api/database/demo", { method: "POST", body: {} }),
  closeDatabase: () => request("/api/database/close", { method: "POST", body: {} }),
  uploadDatabase: (file, onProgress) => upload("/api/database/upload", file, onProgress),
  removeRecent: (path) => request("/api/database/recents/remove", { method: "POST", body: { path } }),
  browse: (mode, suggestedName) => request("/api/dialog", { method: "POST", body: { mode, suggestedName } }),
  tableData: (params) => {
    const query = new URLSearchParams({
      name: params.name,
      limit: params.limit ?? 100,
      offset: params.offset ?? 0,
      search: params.search ?? "",
      orderBy: params.orderBy ?? "",
      orderDir: params.orderDir ?? "asc",
    });
    return request(`/api/table/data?${query}`);
  },
  query: (sql, maxRows) => request("/api/query", { method: "POST", body: { sql, maxRows } }),
  exportTable: (name) => download(
    `/api/export/table?name=${encodeURIComponent(name)}`,
    `${name}.csv`,
  ),
  importUpload: (file, onProgress) => upload("/api/import/upload", file, onProgress),
  importPreview: (fileId, sheet) => request("/api/import/preview", { method: "POST", body: { fileId, sheet } }),
  importRun: (payload) => request("/api/import/run", { method: "POST", body: payload }),
  forgetImport: (fileId) => request("/api/import/forget", { method: "POST", body: { fileId } }),
};
