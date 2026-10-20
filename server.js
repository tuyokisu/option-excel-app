"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = Number(process.env.PORT) || 8765;
const APP_DIR = __dirname;
const SAVE_DIR = path.join(APP_DIR, "saveデータ(読み込むよう）");
const EXCEL_DIR = path.join(APP_DIR, "読み込み用excel");
const EXCEL_WRITE_DIR = path.join(APP_DIR, "書き込むエクセル");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".css": "text/css; charset=utf-8",
};

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Private-Network": "true",
    "Access-Control-Allow-Methods": "GET, PUT, DELETE, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

function send(res, status, body, type) {
  const bytes = Buffer.from(body);
  res.writeHead(status, Object.assign(corsHeaders(), {
    "Content-Type": type || "text/plain; charset=utf-8",
    "Content-Length": bytes.length,
    "Cache-Control": "no-store",
  }));
  res.end(bytes);
}

function sendJson(res, status, data) {
  send(res, status, JSON.stringify(data), "application/json; charset=utf-8");
}

function safeSegment(part) {
  const name = String(part || "");
  if (!name || name === "." || name === ".." || name.includes("..")) {
    return "";
  }
  if (!/^[^\u0000-\u001f\\/:*?"<>|]+$/.test(name)) {
    return "";
  }
  return name;
}

function safeRel(raw, ext) {
  let name = decodeURIComponent(raw || "").replace(/\\/g, "/");
  if (name.includes("..")) {
    return "";
  }
  const parts = name.split("/").filter(Boolean);
  if (parts.length < 1 || parts.length > 3) {
    return "";
  }
  const checked = parts.map(safeSegment);
  if (checked.some((part) => !part)) {
    return "";
  }
  const suffixes = Array.isArray(ext) ? ext : [ext];
  const last = checked[checked.length - 1];
  const ok = suffixes.some((item) => new RegExp(item.replace(".", "\\.") + "$", "i").test(last));
  if (!ok) {
    return "";
  }
  return checked.join("/");
}

function listByExt(dir, ext, prefix, depth) {
  if (!fs.existsSync(dir)) {
    return [];
  }
  const suffixes = Array.isArray(ext) ? ext : [ext];
  const names = [];
  const level = depth || 0;
  fs.readdirSync(dir, { withFileTypes: true }).forEach((entry) => {
    if (entry.isDirectory() && level < 2) {
      names.push(...listByExt(path.join(dir, entry.name), suffixes, prefix + entry.name + "/", level + 1));
      return;
    }
    if (
      entry.isFile() &&
      !entry.name.startsWith("~$") &&
      suffixes.some((item) => entry.name.toLowerCase().endsWith(item))
    ) {
      names.push(prefix + entry.name);
    }
  });
  return names.sort((a, b) => a.localeCompare(b, "ja"));
}

function collectBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function listImmediateFolders(dir) {
  if (!fs.existsSync(dir)) {
    return [];
  }
  return fs
    .readdirSync(dir, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort((a, b) => a.localeCompare(b, "ja"));
}

function listSaves() {
  return listByExt(SAVE_DIR, ".json", "");
}

function listExcels() {
  return listByExt(EXCEL_DIR, [".xlsx", ".xlsm", ".xls"], "");
}

function listExcelFolders() {
  return listImmediateFolders(EXCEL_DIR);
}

function removeIfEmptyDir(dir) {
  try {
    if (fs.existsSync(dir) && fs.statSync(dir).isDirectory() && fs.readdirSync(dir).length === 0) {
      fs.rmdirSync(dir);
    }
  } catch (error) {
    // ignore
  }
}

function handleApi(req, res, url) {
  if (url.pathname === "/api/saves" && req.method === "GET") {
    sendJson(res, 200, { files: listSaves(), folders: listImmediateFolders(SAVE_DIR) });
    return;
  }

  if (url.pathname === "/api/excels" && req.method === "GET") {
    sendJson(res, 200, { files: listExcels(), folders: listExcelFolders() });
    return;
  }

  const excelMatch = url.pathname.match(/^\/api\/excels\/(.+)$/);
  if (excelMatch) {
    const filename = safeRel(excelMatch[1], [".xlsx", ".xlsm", ".xls"]);
    if (!filename) {
      sendJson(res, 400, { error: "invalid filename" });
      return;
    }
    const excelReadPath = path.join(EXCEL_DIR, ...filename.split("/"));
    const excelWritePath = path.join(EXCEL_WRITE_DIR, ...filename.split("/"));
    if (req.method === "GET") {
      if (!fs.existsSync(excelReadPath)) {
        sendJson(res, 404, { error: "file not found" });
        return;
      }
      const bytes = fs.readFileSync(excelReadPath);
      res.writeHead(200, Object.assign(corsHeaders(), {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Length": bytes.length,
        "Cache-Control": "no-store",
      }));
      res.end(bytes);
      return;
    }
    if (req.method !== "PUT") {
      sendJson(res, 405, { error: "method not allowed" });
      return;
    }
    collectBody(req).then((bytes) => {
      if (!bytes.length) {
        sendJson(res, 400, { error: "empty file" });
        return;
      }
      if (!fs.existsSync(EXCEL_WRITE_DIR)) {
        sendJson(res, 404, { error: "write folder not found" });
        return;
      }
      fs.mkdirSync(path.dirname(excelWritePath), { recursive: true });
      fs.writeFileSync(excelWritePath, bytes);
      sendJson(res, 200, { ok: true, name: filename });
    });
    return;
  }

  const match = url.pathname.match(/^\/api\/saves\/(.+)$/);
  if (!match) {
    sendJson(res, 404, { error: "not found" });
    return;
  }

  const filename = safeRel(match[1], ".json");
  if (!filename) {
    sendJson(res, 400, { error: "invalid filename" });
    return;
  }
  const filePath = path.join(SAVE_DIR, ...filename.split("/"));

  if (req.method === "GET") {
    if (!fs.existsSync(filePath)) {
      sendJson(res, 404, { error: "file not found" });
      return;
    }
    send(res, 200, fs.readFileSync(filePath), "application/json; charset=utf-8");
    return;
  }

  if (req.method === "PUT") {
    collectBody(req).then((bytes) => {
      const text = bytes.toString("utf8");
      try {
        JSON.parse(text);
      } catch (error) {
        sendJson(res, 400, { error: "invalid json" });
        return;
      }
      const parent = path.dirname(filePath);
      if (!fs.existsSync(SAVE_DIR) || !fs.existsSync(parent)) {
        sendJson(res, 404, { error: "save folder not found" });
        return;
      }
      fs.writeFileSync(filePath, text, "utf8");
      sendJson(res, 200, { ok: true, name: filename });
    });
    return;
  }

  if (req.method === "DELETE") {
    if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
      sendJson(res, 404, { error: "file not found" });
      return;
    }
    fs.unlinkSync(filePath);
    if (filename.includes("/")) {
      removeIfEmptyDir(path.dirname(filePath));
    }
    sendJson(res, 200, { ok: true, name: filename });
    return;
  }

  sendJson(res, 405, { error: "method not allowed" });
}

function handleStatic(req, res, url) {
  let rel = decodeURIComponent(url.pathname);
  if (rel === "/") {
    rel = "/index.html";
  }
  const target = path.normalize(path.join(APP_DIR, rel));
  if (!target.startsWith(APP_DIR + path.sep) && target !== APP_DIR) {
    send(res, 403, "forbidden");
    return;
  }
  if (!fs.existsSync(target) || !fs.statSync(target).isFile()) {
    send(res, 404, "not found");
    return;
  }
  const ext = path.extname(target).toLowerCase();
  send(res, 200, fs.readFileSync(target), TYPES[ext] || "application/octet-stream");
}

const server = http.createServer((req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, corsHeaders());
    res.end();
    return;
  }
  const url = new URL(req.url, "http://127.0.0.1");
  if (url.pathname.startsWith("/api/")) {
    handleApi(req, res, url);
    return;
  }
  if (req.method !== "GET") {
    send(res, 405, "method not allowed");
    return;
  }
  handleStatic(req, res, url);
});

server.listen(PORT, "127.0.0.1", () => {
  console.log("http://127.0.0.1:" + PORT + "/");
  console.log("途中保存: " + SAVE_DIR);
  console.log("Excel読み込み: " + EXCEL_DIR);
  console.log("Excel作成: " + EXCEL_WRITE_DIR);
});
