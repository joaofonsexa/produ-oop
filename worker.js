import { SCRIPT_JS, STYLES_CSS } from "./asset-bundle.js";

const SESSION_NAME = "pulse_session";
const DB_PATH = "./data/db.json";
const D1_STATE_ID = 1;
const D1_SCHEMA = `
CREATE TABLE IF NOT EXISTS app_state (
  id INTEGER PRIMARY KEY CHECK (id = 1),
  data TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
`;
const INDEX_HTML = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Pulse Ops</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="/styles.css">
</head>
<body>
  <div id="app"></div>
  <script src="/script.js"></script>
</body>
</html>`;
const SECRET_KEY = "pulse-ops-local-secret";
const DEFAULT_PASSWORD = "Trocar@01";
const STATIC_FILES = {
  "/": "index.html",
  "/index.html": "index.html",
  "/styles.css": "styles.css",
  "/script.js": "script.js",
  "/logos_KR-02.png": "logos_KR-02.png",
};

const isNode = typeof process !== "undefined" && !!process.versions?.node;
const isLocalNodeRuntime =
  isNode &&
  typeof process !== "undefined" &&
  Array.isArray(process.argv) &&
  typeof process.argv[1] === "string" &&
  typeof import.meta?.url === "string" &&
  import.meta.url.startsWith("file:");

let nodeModulesPromise;
let storageCache = null;

async function nodeModules() {
  if (!isLocalNodeRuntime) return null;
  if (!nodeModulesPromise) {
    nodeModulesPromise = Promise.all([
      import("node:fs/promises"),
      import("node:path"),
      import("node:http"),
      import("node:crypto"),
      import("node:url"),
      import("node:child_process"),
    ]).then(([fs, path, http, crypto, url, childProcess]) => ({
      fs,
      path,
      http,
      crypto,
      fileURLToPath: url.fileURLToPath,
      execFile: childProcess.execFile,
    }));
  }
  return nodeModulesPromise;
}

function nowIso() {
  return new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
}

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function monthRef(dateValue) {
  return String(dateValue).slice(0, 7);
}

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replaceAll(" ", "_")
    .replaceAll("-", "_")
    .replaceAll("/", "_")
    .replaceAll(".", "")
    .replaceAll("ã", "a")
    .replaceAll("á", "a")
    .replaceAll("â", "a")
    .replaceAll("é", "e")
    .replaceAll("ê", "e")
    .replaceAll("í", "i")
    .replaceAll("ó", "o")
    .replaceAll("ô", "o")
    .replaceAll("õ", "o")
    .replaceAll("ú", "u")
    .replaceAll("ç", "c");
}

function toInt(value) {
  if (value === null || value === undefined || value === "") return 0;
  const text = String(value).trim().replace(",", ".");
  const parsed = Number(text);
  return Number.isFinite(parsed) ? Math.trunc(parsed) : 0;
}

function toFloat(value) {
  if (value === null || value === undefined || value === "") return 0;
  const text = String(value).trim().replace(",", ".");
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDate(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return todayIso();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
    const [day, month, year] = raw.split("/");
    return `${year}-${month}-${day}`;
  }
  if (/^\d{2}-\d{2}-\d{4}$/.test(raw)) {
    const [day, month, year] = raw.split("-");
    return `${year}-${month}-${day}`;
  }
  throw new Error(`Data invalida: ${raw}`);
}

function isValidMonthRef(value) {
  return /^\d{4}-\d{2}$/.test(String(value || "").trim());
}

function normalizeComparable(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function serializeUser(user) {
  return {
    id: user.id,
    full_name: user.full_name,
    login: user.login,
    role: user.role,
    platform_0800_id: user.platform_0800_id,
    nuvidio_id: user.nuvidio_id,
    must_change_password: !!user.must_change_password,
    is_active: user.is_active,
  };
}

function normalizeUserRecord(user, roleFallback = "operator") {
  return {
    ...user,
    role: user.role || roleFallback,
    must_change_password: Boolean(user.must_change_password),
    is_active: user.is_active !== false,
  };
}

function normalizeDbState(db) {
  db.users = (db.users || []).map((user, index) => normalizeUserRecord(user, index === 0 ? "manager" : "operator"));
  db.dailyMetrics = db.dailyMetrics || [];
  db.qualityScores = db.qualityScores || [];
  db.importLogs = db.importLogs || [];
  db.counters = db.counters || { users: 2, dailyMetrics: 1, qualityScores: 1, importLogs: 1 };
  return db;
}

function calculateEffectiveness(metric) {
  const actionable =
    toInt(metric.calls_0800_approved) +
    toInt(metric.calls_0800_rejected) +
    toInt(metric.calls_0800_pending) +
    toInt(metric.calls_nuvidio_approved) +
    toInt(metric.calls_nuvidio_rejected);
  const total = actionable + toInt(metric.calls_0800_no_action) + toInt(metric.calls_nuvidio_no_action);
  if (!total) return 0;
  return Number(((actionable / total) * 100).toFixed(2));
}

async function hashPassword(password) {
  const salt = randomHex(16);
  const hash = await digestHex(`${salt}:${password}`);
  return `${salt}$${hash}`;
}

async function verifyPassword(password, storedHash) {
  const [salt, hash] = String(storedHash || "").split("$");
  if (!salt || !hash) return false;
  const candidate = await digestHex(`${salt}:${password}`);
  return timingSafeEqual(hash, candidate);
}

async function digestHex(input) {
  const data = new TextEncoder().encode(input);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function randomHex(length) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return [...bytes].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i += 1) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return result === 0;
}

async function signSession(payload) {
  const raw = JSON.stringify(payload);
  const signature = await digestHex(`${SECRET_KEY}:${raw}`);
  return `${btoa(raw)}.${signature}`;
}

function resolveSecretKey(env) {
  return String(env?.SECRET_KEY || SECRET_KEY);
}

async function signSessionWithEnv(payload, env) {
  const raw = JSON.stringify(payload);
  const signature = await digestHex(`${resolveSecretKey(env)}:${raw}`);
  return `${btoa(raw)}.${signature}`;
}

async function verifySession(token, env) {
  if (!token || !token.includes(".")) return null;
  const [rawBase64, signature] = token.split(".");
  const raw = atob(rawBase64);
  const expected = await digestHex(`${resolveSecretKey(env)}:${raw}`);
  if (!timingSafeEqual(signature, expected)) return null;
  const payload = JSON.parse(raw);
  if (!payload.expires_at || new Date(payload.expires_at) < new Date()) return null;
  return payload;
}

function seedState() {
  return {
    counters: {
      users: 2,
      dailyMetrics: 1,
      qualityScores: 1,
      importLogs: 1,
    },
    users: [
      {
        id: 1,
        full_name: "Administrador",
        login: "admin",
        password_hash: null,
        role: "manager",
        platform_0800_id: "GESTOR-001",
        nuvidio_id: "NUVIDIO-001",
        must_change_password: false,
        is_active: true,
        created_at: nowIso(),
        updated_at: nowIso(),
      },
    ],
    dailyMetrics: [],
    qualityScores: [],
    importLogs: [],
  };
}

async function ensureStorage(env = {}) {
  if (env?.DB) {
    await env.DB.exec(D1_SCHEMA);
    const row = await env.DB.prepare("SELECT data FROM app_state WHERE id = ?").bind(D1_STATE_ID).first();
    if (row?.data) return normalizeDbState(JSON.parse(row.data));
    const seeded = normalizeDbState(seedState());
    seeded.users[0].password_hash = await hashPassword("admin123");
    await persistStorage(seeded, env);
    return seeded;
  }
  if (storageCache) return storageCache;
  if (isLocalNodeRuntime) {
    const { fs, path } = await nodeModules();
    await fs.mkdir(path.dirname(DB_PATH), { recursive: true });
    try {
      const raw = await fs.readFile(DB_PATH, "utf8");
      storageCache = normalizeDbState(JSON.parse(raw));
    } catch {
      storageCache = normalizeDbState(seedState());
      storageCache.users[0].password_hash = await hashPassword("admin123");
      await persistStorage(storageCache, env);
    }
  } else {
    storageCache = normalizeDbState(seedState());
    storageCache.users[0].password_hash = await hashPassword("admin123");
  }
  return storageCache;
}

async function persistStorage(db, env = {}) {
  if (env?.DB) {
    await env.DB.exec(D1_SCHEMA);
    await env.DB.prepare(`
      INSERT INTO app_state (id, data, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        data = excluded.data,
        updated_at = excluded.updated_at
    `).bind(D1_STATE_ID, JSON.stringify(db), nowIso()).run();
    return;
  }
  if (!isLocalNodeRuntime || !db) {
    storageCache = db;
    return;
  }
  storageCache = db;
  const { fs } = await nodeModules();
  await fs.writeFile(DB_PATH, JSON.stringify(db, null, 2), "utf8");
}

function nextId(db, key) {
  const id = db.counters[key];
  db.counters[key] += 1;
  return id;
}

function jsonResponse(payload, status = 200, headers = {}) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      ...headers,
    },
  });
}

async function getCurrentUser(request, db, env) {
  const cookieHeader = request.headers.get("cookie") || "";
  const token = parseCookies(cookieHeader)[SESSION_NAME];
  if (!token) return null;
  const session = await verifySession(token, env);
  if (!session) return null;
  return db.users.find((user) => user.id === session.user_id && user.is_active) || null;
}

function parseCookies(cookieHeader) {
  return cookieHeader
    .split(";")
    .map((item) => item.trim())
    .filter(Boolean)
    .reduce((acc, item) => {
      const index = item.indexOf("=");
      if (index === -1) return acc;
      acc[item.slice(0, index)] = decodeURIComponent(item.slice(index + 1));
      return acc;
    }, {});
}

async function requireAuth(request, db, env) {
  const user = await getCurrentUser(request, db, env);
  if (!user) {
    return { error: jsonResponse({ error: "Nao autenticado" }, 401) };
  }
  return { user };
}

async function requireManager(request, db, env) {
  const auth = await requireAuth(request, db, env);
  if (auth.error) return auth;
  if (auth.user.role !== "manager") {
    return { error: jsonResponse({ error: "Acesso negado" }, 403) };
  }
  return auth;
}

function matchUser(users, row) {
  const identifiers = {
    name: String(row.nome || row.operador || "").trim().toLowerCase(),
    login: String(row.login || row.usuario || "").trim().toLowerCase(),
    id0800: String(row.id_0800 || row.usuario_0800 || row.plataforma_0800 || "").trim().toLowerCase(),
    idNuvidio: String(row.id_nuvidio || row.usuario_nuvidio || row.plataforma_nuvidio || "").trim().toLowerCase(),
  };
  return users.find((user) => {
    return (
      (identifiers.id0800 && String(user.platform_0800_id || "").trim().toLowerCase() === identifiers.id0800) ||
      (identifiers.idNuvidio && String(user.nuvidio_id || "").trim().toLowerCase() === identifiers.idNuvidio) ||
      (identifiers.login && String(user.login || "").trim().toLowerCase() === identifiers.login) ||
      (identifiers.name && String(user.full_name || "").trim().toLowerCase() === identifiers.name)
    );
  });
}

function matchUserByName(users, name) {
  const target = normalizeComparable(name);
  if (!target) return null;
  return users.find((user) => normalizeComparable(user.full_name) === target) || null;
}

function parseCsv(text) {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      row.push(current);
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(current);
      if (row.some((cell) => cell.length)) rows.push(row);
      row = [];
      current = "";
    } else {
      current += char;
    }
  }
  if (current.length || row.length) {
    row.push(current);
    if (row.some((cell) => cell.length)) rows.push(row);
  }
  if (!rows.length) return [];
  const headers = rows[0].map((header) => normalizeHeader(header));
  return rows.slice(1).map((cells) => {
    const entry = {};
    headers.forEach((header, index) => {
      entry[header] = cells[index] ?? "";
    });
    return entry;
  });
}

async function runPythonScript(args) {
  const modules = await nodeModules();
  const pythonPath = "C:\\Users\\joao.fonseca.KRCONSULTORIA\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
  return await new Promise((resolve, reject) => {
    modules.execFile(pythonPath, args, { cwd: process.cwd(), encoding: "utf8", maxBuffer: 10 * 1024 * 1024 }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(stderr?.trim() || error.message));
        return;
      }
      resolve(stdout);
    });
  });
}

async function generateQualityTemplate(users) {
  const { fs, path } = await nodeModules();
  const tempDir = path.join(process.cwd(), "data", "tmp");
  await fs.mkdir(tempDir, { recursive: true });
  const stamp = Date.now();
  const jsonPath = path.join(tempDir, `quality-users-${stamp}.json`);
  const outputPath = path.join(tempDir, `modelo-monitoria-${stamp}.xlsx`);
  const scriptPath = path.join(process.cwd(), "tools", "quality_xlsx.py");
  await fs.writeFile(jsonPath, JSON.stringify(users.map((user) => user.full_name), null, 2), "utf8");
  try {
    await runPythonScript([scriptPath, "export", jsonPath, outputPath]);
    return await fs.readFile(outputPath);
  } finally {
    await fs.unlink(jsonPath).catch(() => {});
    await fs.unlink(outputPath).catch(() => {});
  }
}

async function parseQualityWorkbook(file) {
  const { fs, path } = await nodeModules();
  const tempDir = path.join(process.cwd(), "data", "tmp");
  await fs.mkdir(tempDir, { recursive: true });
  const stamp = Date.now();
  const inputPath = path.join(tempDir, `quality-upload-${stamp}.xlsx`);
  const scriptPath = path.join(process.cwd(), "tools", "quality_xlsx.py");
  const buffer = Buffer.from(await file.arrayBuffer());
  await fs.writeFile(inputPath, buffer);
  try {
    const stdout = await runPythonScript([scriptPath, "import", inputPath]);
    return JSON.parse(stdout);
  } finally {
    await fs.unlink(inputPath).catch(() => {});
  }
}

function upsertDailyMetric(db, userId, metricDate, values, sourceName) {
  let metric = db.dailyMetrics.find((item) => item.user_id === userId && item.metric_date === metricDate);
  if (!metric) {
    metric = {
      id: nextId(db, "dailyMetrics"),
      user_id: userId,
      metric_date: metricDate,
      production: 0,
      calls_0800_approved: 0,
      calls_0800_rejected: 0,
      calls_0800_pending: 0,
      calls_0800_no_action: 0,
      calls_nuvidio_approved: 0,
      calls_nuvidio_rejected: 0,
      calls_nuvidio_no_action: 0,
      import_source: sourceName,
      created_at: nowIso(),
      updated_at: nowIso(),
    };
    db.dailyMetrics.push(metric);
  }
  Object.assign(metric, values, {
    import_source: sourceName ?? metric.import_source,
    updated_at: nowIso(),
  });
  return metric;
}

function buildOverview(db, user, url) {
  const dateValue = url.searchParams.get("date") || todayIso();
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const scopedMetrics = db.dailyMetrics.filter((metric) => user.role === "manager" || metric.user_id === user.id);
  const todayRows = scopedMetrics.filter((metric) => metric.metric_date === dateValue);
  const monthRows = db.qualityScores.filter((score) => score.reference_month === monthRef(dateValue) && (user.role === "manager" || score.user_id === user.id));
  const trendRows = scopedMetrics
    .filter((metric) => metric.metric_date >= start && metric.metric_date <= end)
    .sort((a, b) => a.metric_date.localeCompare(b.metric_date));
  const cards = {
    production: todayRows.reduce((sum, row) => sum + toInt(row.production), 0),
    effectiveness: average(todayRows.map(calculateEffectiveness)),
    quality: average(monthRows.map((item) => toFloat(item.score))),
  };
  const trendMap = new Map();
  for (const row of trendRows) {
    const current = trendMap.get(row.metric_date) || { date: row.metric_date, production: 0, effectivenessParts: [] };
    current.production += toInt(row.production);
    current.effectivenessParts.push(calculateEffectiveness(row));
    trendMap.set(row.metric_date, current);
  }
  const trend = [...trendMap.values()].map((item) => ({
    date: item.date,
    production: item.production,
    effectiveness: average(item.effectivenessParts),
  }));
  const operators = todayRows.map((row) => ({
    name: db.users.find((entry) => entry.id === row.user_id)?.full_name || "Operador",
    production: row.production,
    effectiveness: calculateEffectiveness(row),
  }));
  return { cards, trend, operators };
}

function buildAnalysis(db, user, url) {
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const users = db.users.filter((entry) => entry.is_active && (user.role === "manager" || entry.id === user.id));
  const ranking = users.map((entry) => {
    const metrics = db.dailyMetrics.filter((row) => row.user_id === entry.id && row.metric_date >= start && row.metric_date <= end);
    const quality = db.qualityScores.find((score) => score.user_id === entry.id && score.reference_month === monthRef(end));
    return {
      user_id: entry.id,
      name: entry.full_name,
      avg_production: metrics.length ? Number((metrics.reduce((sum, row) => sum + toInt(row.production), 0) / metrics.length).toFixed(2)) : 0,
      total_production: metrics.reduce((sum, row) => sum + toInt(row.production), 0),
      effectiveness: average(metrics.map(calculateEffectiveness)),
      quality: toFloat(quality?.score),
      active_days: metrics.length,
    };
  }).sort((a, b) => b.total_production - a.total_production || a.name.localeCompare(b.name));
  return { ranking, period: { start, end } };
}

function buildHistory(db, user, url) {
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const requestedUserId = Number(url.searchParams.get("user_id") || user.id);
  const targetUserId = user.role === "manager" ? requestedUserId : user.id;
  const history = db.dailyMetrics
    .filter((row) => row.user_id === targetUserId && row.metric_date >= start && row.metric_date <= end)
    .sort((a, b) => b.metric_date.localeCompare(a.metric_date))
    .map((row) => ({ ...row, effectiveness: calculateEffectiveness(row) }));
  const quality = db.qualityScores
    .filter((item) => item.user_id === targetUserId)
    .sort((a, b) => b.reference_month.localeCompare(a.reference_month));
  return { history, quality };
}

function average(values) {
  if (!values.length) return 0;
  return Number((values.reduce((sum, value) => sum + Number(value || 0), 0) / values.length).toFixed(2));
}

function shiftDate(days) {
  const date = new Date();
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function cloudQualityUnsupported() {
  return jsonResponse({
    error: "No Cloudflare, a monitoria em XLSX ainda precisa ser processada localmente. Use CSV ou integre uma rotina externa antes do deploy final.",
  }, 501);
}

async function handleApi(request, url, db, env = {}) {
  if (url.pathname === "/api/auth/me" && request.method === "GET") {
    const user = await getCurrentUser(request, db, env);
    return jsonResponse({ user: user ? serializeUser(user) : null });
  }

  if (url.pathname === "/api/auth/login" && request.method === "POST") {
    const payload = await request.json();
    const user = db.users.find((entry) => entry.login === String(payload.login || "").trim() && entry.is_active);
    if (!user || !(await verifyPassword(String(payload.password || ""), user.password_hash))) {
      return jsonResponse({ error: "Credenciais invalidas" }, 401);
    }
    const token = await signSessionWithEnv({
      user_id: user.id,
      expires_at: new Date(Date.now() + 7 * 86400000).toISOString(),
    }, env);
    return jsonResponse(
      { user: serializeUser(user) },
      200,
      { "set-cookie": `${SESSION_NAME}=${token}; Path=/; HttpOnly; SameSite=Lax` },
    );
  }

  if (url.pathname === "/api/auth/logout" && request.method === "POST") {
    return jsonResponse({ ok: true }, 200, {
      "set-cookie": `${SESSION_NAME}=; Path=/; Expires=Thu, 01 Jan 1970 00:00:00 GMT; HttpOnly; SameSite=Lax`,
    });
  }

  if (url.pathname === "/api/auth/password" && request.method === "POST") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    const currentPassword = String(payload.current_password || "");
    const newPassword = String(payload.new_password || "").trim();
    const confirmPassword = String(payload.confirm_password || "").trim();

    if (!auth.user.must_change_password && !(await verifyPassword(currentPassword, auth.user.password_hash))) {
      return jsonResponse({ error: "Senha atual invalida" }, 400);
    }
    if (newPassword.length < 4) {
      return jsonResponse({ error: "A nova senha deve ter pelo menos 4 caracteres" }, 400);
    }
    if (newPassword !== confirmPassword) {
      return jsonResponse({ error: "A confirmacao da senha nao confere" }, 400);
    }

    auth.user.password_hash = await hashPassword(newPassword);
    auth.user.must_change_password = false;
    auth.user.updated_at = nowIso();
    await persistStorage(db, env);
    return jsonResponse({ ok: true, user: serializeUser(auth.user) });
  }

  if (url.pathname === "/api/dashboard/overview" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse(buildOverview(db, auth.user, url));
  }

  if (url.pathname === "/api/analysis" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse(buildAnalysis(db, auth.user, url));
  }

  if (url.pathname === "/api/history" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse(buildHistory(db, auth.user, url));
  }

  if (url.pathname === "/api/admin/users" && request.method === "GET") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse({ users: db.users.map(serializeUser) });
  }

  if (url.pathname === "/api/admin/users" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    const required = ["full_name", "login", "role"];
    const missing = required.filter((field) => !String(payload[field] || "").trim());
    if (missing.length) {
      return jsonResponse({ error: `Campos obrigatorios: ${missing.join(", ")}` }, 400);
    }
    if (!["operator", "manager"].includes(payload.role)) {
      return jsonResponse({ error: "Perfil invalido" }, 400);
    }
    if (db.users.some((entry) => entry.login === String(payload.login).trim())) {
      return jsonResponse({ error: "Login ja cadastrado" }, 409);
    }
    const user = {
      id: nextId(db, "users"),
      full_name: String(payload.full_name).trim(),
      login: String(payload.login).trim(),
      password_hash: await hashPassword(DEFAULT_PASSWORD),
      role: payload.role,
      platform_0800_id: String(payload.platform_0800_id || "").trim(),
      nuvidio_id: String(payload.nuvidio_id || "").trim(),
      must_change_password: true,
      is_active: true,
      created_at: nowIso(),
      updated_at: nowIso(),
    };
    db.users.push(user);
    await persistStorage(db, env);
    return jsonResponse({ user: serializeUser(user), default_password: DEFAULT_PASSWORD }, 201);
  }

  if (url.pathname.startsWith("/api/admin/users/") && request.method === "PUT") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const userId = Number(url.pathname.split("/").pop());
    const user = db.users.find((entry) => entry.id === userId);
    if (!user) return jsonResponse({ error: "Usuario nao encontrado" }, 404);
    const payload = await request.json();
    if (payload.role && !["operator", "manager"].includes(payload.role)) {
      return jsonResponse({ error: "Perfil invalido" }, 400);
    }
    if (payload.login && db.users.some((entry) => entry.login === String(payload.login).trim() && entry.id !== user.id)) {
      return jsonResponse({ error: "Login ja cadastrado" }, 409);
    }
    user.full_name = String(payload.full_name || user.full_name).trim();
    user.login = String(payload.login || user.login).trim();
    user.role = payload.role || user.role;
    user.platform_0800_id = String(payload.platform_0800_id ?? user.platform_0800_id).trim();
    user.nuvidio_id = String(payload.nuvidio_id ?? user.nuvidio_id).trim();
    user.updated_at = nowIso();
    if (String(payload.password || "").trim()) {
      user.password_hash = await hashPassword(String(payload.password).trim());
    }
    await persistStorage(db, env);
    return jsonResponse({ user: serializeUser(user) });
  }

  if (url.pathname === "/api/admin/quality" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    const userId = Number(payload.user_id);
    if (!userId || !String(payload.reference_month || "").trim()) {
      return jsonResponse({ error: "user_id e reference_month sao obrigatorios" }, 400);
    }
    let score = db.qualityScores.find((item) => item.user_id === userId && item.reference_month === payload.reference_month);
    if (!score) {
      score = {
        id: nextId(db, "qualityScores"),
        user_id: userId,
        reference_month: payload.reference_month,
        score: toFloat(payload.score),
        notes: String(payload.notes || "").trim(),
        created_at: nowIso(),
        updated_at: nowIso(),
      };
      db.qualityScores.push(score);
    } else {
      score.score = toFloat(payload.score);
      score.notes = String(payload.notes || "").trim();
      score.updated_at = nowIso();
    }
    await persistStorage(db, env);
    return jsonResponse({ quality: score });
  }

  if (url.pathname === "/api/admin/quality/template" && request.method === "GET") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    if (!isNode) {
      const operators = db.users.filter((user) => user.is_active && user.role === "operator");
      const csv = [
        "Nome do Operador,Monitoria 1,Monitoria 2,Monitoria 3,Monitoria 4",
        ...operators.map((user) => `"${String(user.full_name || "").replaceAll('"', '""')}",,,,`),
      ].join("\n");
      return new Response(csv, {
        status: 200,
        headers: {
          "content-type": "text/csv; charset=utf-8",
          "content-disposition": 'attachment; filename="modelo-monitoria.csv"',
        },
      });
    }
    const operators = db.users.filter((user) => user.is_active && user.role === "operator");
    const buffer = await generateQualityTemplate(operators);
    return new Response(buffer, {
      status: 200,
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": 'attachment; filename="modelo-monitoria.xlsx"',
      },
    });
  }

  if (url.pathname === "/api/admin/quality/import" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const form = await request.formData();
    const file = form.get("file");
    const referenceMonth = String(form.get("reference_month") || "").trim();
    if (!(file instanceof File)) {
      return jsonResponse({ error: "Arquivo de monitoria nao enviado" }, 400);
    }
    if (!isValidMonthRef(referenceMonth)) {
      return jsonResponse({ error: "Informe o mes de referencia no formato AAAA-MM" }, 400);
    }
    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith(".xlsx") && !lowerName.endsWith(".csv")) {
      return jsonResponse({ error: "Use um arquivo XLSX ou CSV para a monitoria" }, 400);
    }

    if (!isNode && lowerName.endsWith(".xlsx")) {
      return cloudQualityUnsupported();
    }

    const rows = lowerName.endsWith(".csv")
      ? parseCsv(await file.text()).map((row) => ({
          name: row.nome_do_operador || row.nome_operador || row.operador || row.nome || "",
          monitoria_1: row.monitoria_1 ?? "",
          monitoria_2: row.monitoria_2 ?? "",
          monitoria_3: row.monitoria_3 ?? "",
          monitoria_4: row.monitoria_4 ?? "",
        }))
      : await parseQualityWorkbook(file);
    const operators = db.users.filter((user) => user.is_active && user.role === "operator");
    const errors = [];
    let processed = 0;

    for (const row of rows) {
      try {
        const user = matchUserByName(operators, row.name);
        if (!user) throw new Error(`Operador nao encontrado: ${row.name}`);
        const monitorias = [row.monitoria_1, row.monitoria_2, row.monitoria_3, row.monitoria_4];
        const validMonitorias = monitorias.filter((value) => value !== null && value !== undefined && value !== "");
        for (const value of validMonitorias) {
          const numeric = toFloat(value);
          if (numeric < 0 || numeric > 100) throw new Error(`Monitoria fora da faixa 0-100 para ${row.name}`);
        }
        const scoreValue = validMonitorias.length
          ? Number((validMonitorias.reduce((sum, value) => sum + toFloat(value), 0) / validMonitorias.length).toFixed(2))
          : 0;

        let score = db.qualityScores.find((item) => item.user_id === user.id && item.reference_month === referenceMonth);
        if (!score) {
          score = {
            id: nextId(db, "qualityScores"),
            user_id: user.id,
            reference_month: referenceMonth,
            score: scoreValue,
            notes: "",
            monitoria_1: row.monitoria_1,
            monitoria_2: row.monitoria_2,
            monitoria_3: row.monitoria_3,
            monitoria_4: row.monitoria_4,
            created_at: nowIso(),
            updated_at: nowIso(),
          };
          db.qualityScores.push(score);
        } else {
          score.score = scoreValue;
          score.monitoria_1 = row.monitoria_1;
          score.monitoria_2 = row.monitoria_2;
          score.monitoria_3 = row.monitoria_3;
          score.monitoria_4 = row.monitoria_4;
          score.updated_at = nowIso();
        }
        processed += 1;
      } catch (error) {
        errors.push(error.message);
      }
    }

    await persistStorage(db, env);
    return jsonResponse({ processed, rejected: errors.length, errors: errors.slice(0, 20) });
  }

  if (url.pathname.startsWith("/api/admin/daily-metrics/") && request.method === "PUT") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const metricId = Number(url.pathname.split("/").pop());
    const metric = db.dailyMetrics.find((entry) => entry.id === metricId);
    if (!metric) return jsonResponse({ error: "Registro nao encontrado" }, 404);
    const payload = await request.json();
    metric.production = toInt(payload.production ?? metric.production);
    metric.calls_0800_approved = toInt(payload.calls_0800_approved ?? metric.calls_0800_approved);
    metric.calls_0800_rejected = toInt(payload.calls_0800_rejected ?? metric.calls_0800_rejected);
    metric.calls_0800_pending = toInt(payload.calls_0800_pending ?? metric.calls_0800_pending);
    metric.calls_0800_no_action = toInt(payload.calls_0800_no_action ?? metric.calls_0800_no_action);
    metric.calls_nuvidio_approved = toInt(payload.calls_nuvidio_approved ?? metric.calls_nuvidio_approved);
    metric.calls_nuvidio_rejected = toInt(payload.calls_nuvidio_rejected ?? metric.calls_nuvidio_rejected);
    metric.calls_nuvidio_no_action = toInt(payload.calls_nuvidio_no_action ?? metric.calls_nuvidio_no_action);
    metric.updated_at = nowIso();
    await persistStorage(db, env);
    return jsonResponse({ metric: { ...metric, effectiveness: calculateEffectiveness(metric) } });
  }

  if (url.pathname === "/api/admin/import" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof File)) {
      return jsonResponse({ error: "Arquivo nao enviado" }, 400);
    }
    if (!file.name.toLowerCase().endsWith(".csv")) {
      return jsonResponse({ error: "No formato atual, use CSV para o teste local." }, 400);
    }
    const text = await file.text();
    const rows = parseCsv(text);
    if (!rows.length) return jsonResponse({ error: "A planilha esta vazia." }, 400);
    const required = [
      "data",
      "producao",
      "0800_aprovado",
      "0800_cancelada",
      "0800_pendenciada",
      "0800_sem_acao",
      "nuvidio_aprovada",
      "nuvidio_reprovada",
      "nuvidio_sem_acao",
    ];
    const first = rows[0];
    const missing = required.filter((field) => !(field in first));
    if (missing.length) {
      return jsonResponse({ error: `Colunas obrigatorias ausentes: ${missing.join(", ")}` }, 400);
    }
    let processed = 0;
    let rejected = 0;
    const errors = [];
    rows.forEach((row, index) => {
      try {
        const user = matchUser(db.users.filter((entry) => entry.is_active), row);
        if (!user) throw new Error("Operador nao encontrado para os identificadores informados.");
        upsertDailyMetric(
          db,
          user.id,
          parseDate(row.data),
          {
            production: toInt(row.producao),
            calls_0800_approved: toInt(row["0800_aprovado"]),
            calls_0800_rejected: toInt(row["0800_cancelada"]),
            calls_0800_pending: toInt(row["0800_pendenciada"]),
            calls_0800_no_action: toInt(row["0800_sem_acao"]),
            calls_nuvidio_approved: toInt(row["nuvidio_aprovada"]),
            calls_nuvidio_rejected: toInt(row["nuvidio_reprovada"]),
            calls_nuvidio_no_action: toInt(row["nuvidio_sem_acao"]),
          },
          file.name,
        );
        processed += 1;
      } catch (error) {
        rejected += 1;
        errors.push({ row: index + 2, error: error.message });
      }
    });
    db.importLogs.push({
      id: nextId(db, "importLogs"),
      imported_by: auth.user.id,
      source_name: file.name,
      imported_at: nowIso(),
      rows_processed: processed,
      rows_rejected: rejected,
    });
    await persistStorage(db, env);
    return jsonResponse({ processed, rejected, errors: errors.slice(0, 20) });
  }

  return jsonResponse({ error: "Rota nao encontrada" }, 404);
}

async function serveStatic(request, env = {}) {
  const pathname = new URL(request.url).pathname;
  if (!isLocalNodeRuntime && env?.ASSETS) {
    if (pathname === "/" || pathname === "/index.html") {
      return new Response(INDEX_HTML, {
        status: 200,
        headers: { "content-type": "text/html; charset=utf-8" },
      });
    }
    if (pathname === "/script.js") {
      return new Response(SCRIPT_JS, {
        status: 200,
        headers: { "content-type": "application/javascript; charset=utf-8" },
      });
    }
    if (pathname === "/styles.css") {
      return new Response(STYLES_CSS, {
        status: 200,
        headers: { "content-type": "text/css; charset=utf-8" },
      });
    }
    return env.ASSETS.fetch(request);
  }
  if (!isLocalNodeRuntime) {
    return new Response("Not found", { status: 404 });
  }
  const { fs, path } = await nodeModules();
  const directFile = STATIC_FILES[pathname];
  const fallbackFile = pathname.startsWith("/") ? pathname.slice(1) : pathname;
  let fileName = directFile;
  if (!fileName) {
    const candidate = path.resolve(fallbackFile);
    try {
      const stat = await fs.stat(candidate);
      if (stat.isFile()) fileName = fallbackFile;
    } catch {
      fileName = null;
    }
  }
  if (!fileName) {
    return new Response("Not found", { status: 404 });
  }
  const content = await fs.readFile(path.resolve(fileName));
  const type = contentType(fileName);
  return new Response(content, { status: 200, headers: { "content-type": type } });
}

function contentType(fileName) {
  if (fileName.endsWith(".html")) return "text/html; charset=utf-8";
  if (fileName.endsWith(".css")) return "text/css; charset=utf-8";
  if (fileName.endsWith(".js")) return "application/javascript; charset=utf-8";
  if (fileName.endsWith(".png")) return "image/png";
  return "application/octet-stream";
}

async function handleRequest(request, env = {}) {
  const url = new URL(request.url);
  if (url.pathname.startsWith("/api/")) {
    const db = await ensureStorage(env);
    return handleApi(request, url, db, env);
  }
  return serveStatic(request, env);
}

export default {
  fetch: async (request, env) => {
    try {
      return await handleRequest(request, env);
    } catch (error) {
      console.error("Worker error", error);
      const url = new URL(request.url);
      if (url.pathname.startsWith("/api/")) {
        return jsonResponse(
          {
            error: "Erro interno do worker",
            details: error?.message || String(error),
          },
          500,
        );
      }
      return new Response(
        `Erro interno do worker: ${error?.message || String(error)}`,
        {
          status: 500,
          headers: { "content-type": "text/plain; charset=utf-8" },
        },
      );
    }
  },
};

if (isLocalNodeRuntime) {
  const modules = await nodeModules();
  const __filename = modules.fileURLToPath(import.meta.url);
  if (process.argv[1] === __filename) {
    const server = modules.http.createServer(async (req, res) => {
      const bodyChunks = [];
      req.on("data", (chunk) => bodyChunks.push(chunk));
      req.on("end", async () => {
        const body = Buffer.concat(bodyChunks);
        const origin = `http://${req.headers.host}`;
        const request = new Request(new URL(req.url, origin), {
          method: req.method,
          headers: req.headers,
          body: req.method === "GET" || req.method === "HEAD" ? undefined : body,
        });
        try {
          const response = await handleRequest(request, {});
          res.statusCode = response.status;
          response.headers.forEach((value, key) => {
            res.setHeader(key, value);
          });
          const arrayBuffer = await response.arrayBuffer();
          res.end(Buffer.from(arrayBuffer));
        } catch (error) {
          res.statusCode = 500;
          res.setHeader("content-type", "application/json; charset=utf-8");
          res.end(JSON.stringify({ error: "Erro interno do servidor", details: error.message }));
        }
      });
    });
    const port = Number(process.env.PORT || 8787);
    server.listen(port, () => {
      console.log(`Pulse Ops disponível em http://localhost:${port}`);
    });
  }
}
