"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = Number(process.env.PORT) || 8765;
const APP_DIR = __dirname;
const SAVE_DIR = path.join(APP_DIR, "saveデータ");
const EXCEL_DIR = path.join(APP_DIR, "仕様書一覧");
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
    "Access-Control-Allow-Methods": "GET, PUT, OPTIONS",
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

function safeName(raw, ext) {
  let name = decodeURIComponent(raw || "");
  name = path.basename(name);
  const escaped = ext.replace(".", "\\.");
  if (!new RegExp("^[^\\u0000-\\u001f\\\\/:*?\"<>|]+" + escaped + "$", "i").test(name)) {
    return "";
  }
  if (name === "." || name === ".." || name.includes("..")) {
    return "";
  }
  return name;
}

function collectBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function listSaves() {
  fs.mkdirSync(SAVE_DIR, { recursive: true });
  return fs
    .readdirSync(SAVE_DIR)
    .filter((name) => name.toLowerCase().endsWith(".json"))
    .sort((a, b) => a.localeCompare(b, "ja"));
}

function handleApi(req, res, url) {
  if (url.pathname === "/api/saves" && req.method === "GET") {
    sendJson(res, 200, { files: listSaves() });
    return;
  }

  const excelMatch = url.pathname.match(/^\/api\/excels\/(.+)$/);
  if (excelMatch) {
    const filename = safeName(excelMatch[1], ".xlsx");
    if (!filename) {
      sendJson(res, 400, { error: "invalid filename" });
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
      fs.mkdirSync(EXCEL_DIR, { recursive: true });
      fs.writeFileSync(path.join(EXCEL_DIR, filename), bytes);
      sendJson(res, 200, { ok: true, name: filename });
    });
    return;
  }

  const match = url.pathname.match(/^\/api\/saves\/(.+)$/);
  if (!match) {
    sendJson(res, 404, { error: "not found" });
    return;
  }

  const filename = safeName(match[1], ".json");
  if (!filename) {
    sendJson(res, 400, { error: "invalid filename" });
    return;
  }
  const filePath = path.join(SAVE_DIR, filename);

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
      fs.mkdirSync(SAVE_DIR, { recursive: true });
      fs.writeFileSync(filePath, text, "utf8");
      sendJson(res, 200, { ok: true, name: filename });
    });
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

fs.mkdirSync(SAVE_DIR, { recursive: true });
fs.mkdirSync(EXCEL_DIR, { recursive: true });
server.listen(PORT, "127.0.0.1", () => {
  console.log("http://127.0.0.1:" + PORT + "/");
  console.log("途中保存: " + SAVE_DIR);
  console.log("Excel保存: " + EXCEL_DIR);
});
