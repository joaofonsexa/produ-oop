import { SCRIPT_JS, STYLES_CSS } from "./asset-bundle.js";

const SESSION_NAME = "pulse_session";
const DB_PATH = "./data/db.json";
const D1_STATE_ID = 1;
const FAVICON_DATA = "";
const D1_SCHEMA_STATEMENTS = [
  "CREATE TABLE IF NOT EXISTS app_state (id INTEGER PRIMARY KEY, data TEXT NOT NULL, updated_at TEXT NOT NULL)",
  "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, full_name TEXT NOT NULL, login TEXT NOT NULL UNIQUE, password_hash TEXT NOT NULL, password_plain TEXT, role TEXT NOT NULL, platform_0800_id TEXT, nuvidio_id TEXT, must_change_password INTEGER NOT NULL DEFAULT 1, is_active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, updated_at TEXT NOT NULL)",
  "CREATE INDEX IF NOT EXISTS idx_users_login ON users(login)",
];
const INDEX_HTML = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PORTAL DE RESULTADOS</title>
  <link rel="icon" type="image/png" href="/logos_KR-02.png">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="/styles.css">
</head>
<body>
  <div id="app"></div>
  <script>window.__brandLogo="";</script>
  <script src="/script.js"></script>
</body>
</html>`;
const SECRET_KEY = "pulse-ops-local-secret";
const DEFAULT_PASSWORD = "Trocar@01";
const PRIMARY_MANAGER_HASH = "f18a137a143dc89817660f864bc973b0$6e3ebbd96a5a02981368b64d2b039dd93a52d901d44f73bc5f1d96797760aec7";
const R2_MAX_FILE_BYTES = 8 * 1024 * 1024;
const R2_MAX_ROWS_PER_RUN = 12000;
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
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
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

function normalizeIdentifier(value) {
  return normalizeComparable(value)
    .replace(/[^a-z0-9@._-]/g, "")
    .trim();
}

function loginLocalPart(value) {
  const normalized = normalizeIdentifier(value);
  if (!normalized) return "";
  const atIndex = normalized.indexOf("@");
  return atIndex >= 0 ? normalized.slice(0, atIndex) : normalized;
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
    preferred_theme: user.preferred_theme,
    last_route: user.last_route,
  };
}

function normalizeRouteForRole(route, role = "operator") {
  const safeRoute = String(route || "").trim();
  const allowed = role === "manager"
    ? ["overview", "analysis", "history", "admin"]
    : ["overview", "analysis", "history"];
  return allowed.includes(safeRoute) ? safeRoute : "overview";
}

function normalizeUserRecord(user, roleFallback = "operator") {
  const role = user.role || roleFallback;
  return {
    ...user,
    role,
    must_change_password: Boolean(user.must_change_password),
    is_active: user.is_active !== false,
    preferred_theme: user.preferred_theme === "contrast" ? "contrast" : "dark",
    last_route: normalizeRouteForRole(user.last_route, role),
  };
}

function normalizeDbState(db) {
  db.users = (db.users || []).map((user, index) => normalizeUserRecord(user, index === 0 ? "manager" : "operator"));
  if (!db.users.some((user) => String(user.login).trim().toLowerCase() === "joao.fonseca")) {
    const nextUserId = Math.max(0, ...db.users.map((user) => Number(user.id) || 0)) + 1;
    db.users.push({
      id: nextUserId,
      full_name: "JoÃ£o Fonseca",
      login: "joao.fonseca",
      password_hash: PRIMARY_MANAGER_HASH,
      password_plain: "Krsa@2026",
      role: "manager",
      platform_0800_id: "",
      nuvidio_id: "",
      must_change_password: false,
      is_active: true,
      preferred_theme: "dark",
      last_route: "overview",
      created_at: nowIso(),
      updated_at: nowIso(),
    });
  }
  db.dailyMetrics = (db.dailyMetrics || []).map((metric) => {
    const normalizedMetric = {
      production_0800: 0,
      production_nuvidio: 0,
      ...metric,
    };
    if (normalizedMetric.production_0800 || normalizedMetric.production_nuvidio) {
      normalizedMetric.production = toInt(normalizedMetric.production_0800) + toInt(normalizedMetric.production_nuvidio);
    } else {
      normalizedMetric.production = toInt(normalizedMetric.production);
    }
    return normalizedMetric;
  });
  db.qualityScores = db.qualityScores || [];
  db.importLogs = db.importLogs || [];
  db.r2ProcessedKeys = Array.isArray(db.r2ProcessedKeys) ? db.r2ProcessedKeys : [];
  db.appSettings = {
    maintenance_for_operators: false,
    maintenance_message: "Portal em manutenção. Tente novamente em instantes.",
    ...(db.appSettings || {}),
  };
  const nextUserCounter = Math.max(1, ...db.users.map((user) => Number(user.id) || 0)) + 1;
  db.counters = db.counters || { users: nextUserCounter, dailyMetrics: 1, qualityScores: 1, importLogs: 1 };
  db.counters.users = Math.max(db.counters.users || 1, nextUserCounter);
  return db;
}

function normalizeD1UserRow(row) {
  return normalizeUserRecord({
    id: Number(row.id),
    full_name: String(row.full_name || "").trim(),
    login: String(row.login || "").trim(),
    password_hash: String(row.password_hash || "").trim(),
    password_plain: String(row.password_plain || "").trim(),
    role: String(row.role || "operator").trim(),
    platform_0800_id: String(row.platform_0800_id || "").trim(),
    nuvidio_id: String(row.nuvidio_id || "").trim(),
    must_change_password: Number(row.must_change_password) === 1,
    is_active: Number(row.is_active) === 1,
    preferred_theme: String(row.preferred_theme || "dark").trim(),
    last_route: String(row.last_route || "overview").trim(),
    created_at: row.created_at || nowIso(),
    updated_at: row.updated_at || nowIso(),
  });
}

async function ensureDefaultPasswords(db) {
  let changed = false;
  const admin = db.users.find((user) => String(user.login).trim().toLowerCase() === "admin");
  if (admin && !String(admin.password_hash || "").trim()) {
    admin.password_hash = await hashPassword("admin123");
    admin.password_plain = "admin123";
    admin.updated_at = nowIso();
    changed = true;
  }
  const joao = db.users.find((user) => String(user.login).trim().toLowerCase() === "joao.fonseca");
  if (joao && !String(joao.password_hash || "").trim()) {
    joao.password_hash = PRIMARY_MANAGER_HASH;
    joao.password_plain = "Krsa@2026";
    joao.updated_at = nowIso();
    changed = true;
  }
  return changed;
}

async function loadUsersFromD1(connection) {
  const result = await connection.prepare(`
    SELECT
      id,
      full_name,
      login,
      password_hash,
      password_plain,
      role,
      platform_0800_id,
      nuvidio_id,
      must_change_password,
      is_active,
      preferred_theme,
      last_route,
      created_at,
      updated_at
    FROM users
    ORDER BY id
  `).all();
  return (result?.results || []).map(normalizeD1UserRow);
}

async function persistUsersToD1(connection, users) {
  for (const user of users) {
    await connection.prepare(`
      INSERT INTO users (
        id,
        full_name,
        login,
        password_hash,
        password_plain,
        role,
        platform_0800_id,
        nuvidio_id,
        must_change_password,
        is_active,
        preferred_theme,
        last_route,
        created_at,
        updated_at
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        full_name = excluded.full_name,
        login = excluded.login,
        password_hash = excluded.password_hash,
        password_plain = excluded.password_plain,
        role = excluded.role,
        platform_0800_id = excluded.platform_0800_id,
        nuvidio_id = excluded.nuvidio_id,
        must_change_password = excluded.must_change_password,
        is_active = excluded.is_active,
        preferred_theme = excluded.preferred_theme,
        last_route = excluded.last_route,
        created_at = excluded.created_at,
        updated_at = excluded.updated_at
    `).bind(
      Number(user.id),
      String(user.full_name || "").trim(),
      String(user.login || "").trim(),
      String(user.password_hash || "").trim(),
      String(user.password_plain || "").trim(),
      String(user.role || "operator").trim(),
      String(user.platform_0800_id || "").trim(),
      String(user.nuvidio_id || "").trim(),
      user.must_change_password ? 1 : 0,
      user.is_active ? 1 : 0,
      user.preferred_theme === "contrast" ? "contrast" : "dark",
      normalizeRouteForRole(user.last_route, user.role),
      user.created_at || nowIso(),
      user.updated_at || nowIso(),
    ).run();
  }

  if (users.length) {
    const placeholders = users.map(() => "?").join(", ");
    await connection.prepare(`DELETE FROM users WHERE id NOT IN (${placeholders})`)
      .bind(...users.map((user) => Number(user.id)))
      .run();
    return;
  }

  await connection.prepare("DELETE FROM users").run();
}

async function ensureD1Schema(connection) {
  for (const statement of D1_SCHEMA_STATEMENTS) {
    await connection.prepare(statement).run();
  }
  const tableInfo = await connection.prepare("PRAGMA table_info(users)").all();
  const columns = new Set((tableInfo?.results || []).map((row) => String(row.name)));
  if (!columns.has("platform_0800_id")) {
    await connection.exec("ALTER TABLE users ADD COLUMN platform_0800_id TEXT");
  }
  if (!columns.has("nuvidio_id")) {
    await connection.exec("ALTER TABLE users ADD COLUMN nuvidio_id TEXT");
  }
  if (!columns.has("must_change_password")) {
    await connection.exec("ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 1");
  }
  if (!columns.has("password_plain")) {
    await connection.exec("ALTER TABLE users ADD COLUMN password_plain TEXT");
  }
  if (!columns.has("is_active")) {
    await connection.exec("ALTER TABLE users ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1");
  }
  if (!columns.has("preferred_theme")) {
    await connection.exec("ALTER TABLE users ADD COLUMN preferred_theme TEXT NOT NULL DEFAULT 'dark'");
  }
  if (!columns.has("last_route")) {
    await connection.exec("ALTER TABLE users ADD COLUMN last_route TEXT NOT NULL DEFAULT 'overview'");
  }
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
        password_plain: "admin123",
        role: "manager",
        platform_0800_id: "GESTOR-001",
        nuvidio_id: "NUVIDIO-001",
        must_change_password: false,
        is_active: true,
        preferred_theme: "dark",
        last_route: "overview",
        created_at: nowIso(),
        updated_at: nowIso(),
      },
    ],
    dailyMetrics: [],
    qualityScores: [],
    importLogs: [],
    appSettings: {
      maintenance_for_operators: false,
      maintenance_message: "Portal em manutenção. Tente novamente em instantes.",
    },
  };
}

function serializeAppSettings(db) {
  return {
    maintenance_for_operators: Boolean(db?.appSettings?.maintenance_for_operators),
    maintenance_message: String(
      db?.appSettings?.maintenance_message || "Portal em manutenção. Tente novamente em instantes.",
    ),
  };
}

function isOperatorBlockedByMaintenance(user, db) {
  return Boolean(user && user.role !== "manager" && db?.appSettings?.maintenance_for_operators);
}

async function ensureStorage(env = {}) {
  if (env?.DB) {
    await ensureD1Schema(env.DB);
    const row = await env.DB.prepare("SELECT data FROM app_state WHERE id = ?").bind(D1_STATE_ID).first();
    const baseState = row?.data ? normalizeDbState(JSON.parse(row.data)) : normalizeDbState(seedState());
    const loadedUsers = await loadUsersFromD1(env.DB);
    if (loadedUsers.length) {
      baseState.users = loadedUsers;
    } else if (!row?.data) {
      baseState.users[0].password_hash = await hashPassword("admin123");
    }
    const db = normalizeDbState(baseState);
    const repairedPasswords = await ensureDefaultPasswords(db);
    const shouldPersist =
      !row?.data ||
      !loadedUsers.length ||
      repairedPasswords ||
      db.users.length !== loadedUsers.length;
    if (shouldPersist) {
      await persistStorage(db, env);
    }
    storageCache = db;
    return db;
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
    await ensureD1Schema(env.DB);
    await persistUsersToD1(env.DB, db.users || []);
    await env.DB.prepare(`
      INSERT INTO app_state (id, data, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        data = excluded.data,
        updated_at = excluded.updated_at
    `).bind(D1_STATE_ID, JSON.stringify(db), nowIso()).run();
    storageCache = db;
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
  const rawName = String(row.nome || row.operador || "").trim();
  const nameBase = rawName.split(" - ")[0]?.split(" / ")[0]?.trim() || rawName;
  const identifiers = {
    name: normalizeComparable(nameBase),
    login: normalizeIdentifier(String(row.login || row.usuario || "").trim()),
    id0800: normalizeIdentifier(String(row.id_0800 || row.usuario_0800 || row.plataforma_0800 || "").trim()),
    idNuvidio: normalizeIdentifier(String(row.id_nuvidio || row.usuario_nuvidio || row.plataforma_nuvidio || "").trim()),
  };
  const loginBase = loginLocalPart(identifiers.login);
  const id0800Base = loginLocalPart(identifiers.id0800);
  const idNuvidioBase = loginLocalPart(identifiers.idNuvidio);

  return users.find((user) => {
    const userName = normalizeComparable(user.full_name);
    const userLogin = normalizeIdentifier(user.login);
    const userLoginBase = loginLocalPart(userLogin);
    const user0800 = normalizeIdentifier(user.platform_0800_id || "");
    const user0800Base = loginLocalPart(user0800);
    const userNuvidio = normalizeIdentifier(user.nuvidio_id || "");
    const userNuvidioBase = loginLocalPart(userNuvidio);

    return (
      (identifiers.id0800 && (user0800 === identifiers.id0800 || user0800Base === id0800Base || userLogin === identifiers.id0800 || userLoginBase === id0800Base)) ||
      (identifiers.idNuvidio && (userNuvidio === identifiers.idNuvidio || userNuvidioBase === idNuvidioBase || userLogin === identifiers.idNuvidio || userLoginBase === idNuvidioBase)) ||
      (identifiers.login && (userLogin === identifiers.login || userLoginBase === loginBase || user0800 === identifiers.login || userNuvidio === identifiers.login)) ||
      (identifiers.name && userName === identifiers.name)
    );
  });
}

function matchUserByName(users, name) {
  const target = normalizeComparable(name);
  if (!target) return null;
  return users.find((user) => normalizeComparable(user.full_name) === target) || null;
}

function parseCsv(text, maxRows = Number.POSITIVE_INFINITY) {
  const sanitizedText = String(text || "").replace(/^\uFEFF/, "");
  const firstLine = sanitizedText.split(/\r?\n/, 1)[0] || "";
  const semicolonCount = (firstLine.match(/;/g) || []).length;
  const commaCount = (firstLine.match(/,/g) || []).length;
  const separator = semicolonCount > commaCount ? ";" : ",";
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;
  for (let i = 0; i < sanitizedText.length; i += 1) {
    const char = sanitizedText[i];
    const next = sanitizedText[i + 1];
    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === separator && !inQuotes) {
      row.push(current);
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(current);
      if (row.some((cell) => cell.length)) {
        rows.push(row);
        if (rows.length >= maxRows) break;
      }
      row = [];
      current = "";
    } else {
      current += char;
    }
  }
  if ((current.length || row.length) && rows.length < maxRows) {
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

function hasFields(row, fields) {
  return fields.every((field) => field in row);
}

function detectImportSchema(row, fileName = "") {
  const lowerName = String(fileName || "").toLowerCase();
  const normalizedFields = [
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
  if (hasFields(row, normalizedFields)) return "normalized";
  if (hasFields(row, ["motivo", "numero_do_protocolo", "data_abertura_ocorrencia", "usuario_de_abertura_da_ocorrencia"])) return "0800";
  if (hasFields(row, ["protocolo", "data_abreviada", "email_do_atendente", "tag"])) return "nuvidio";
  if (hasFields(row, ["usuario_nuvidio", "nuvidio_aprovadas", "nuvidio_reprovadas", "nuvidio_sem_acao", "data"])) return "nuvidio_summary";
  if (hasFields(row, ["usuario_0800", "0800_aprovadas", "0800_reprovadas", "0800_pendenciadas", "0800_sem_acao", "data"])) return "0800_summary";
  if (lowerName.includes("0800")) return "0800";
  if (lowerName.includes("nuvidio")) return "nuvidio";
  return "unknown";
}

function normalize0800Reason(value) {
  return normalizeComparable(value).toUpperCase();
}

function normalizeNuvidioTag(value) {
  return normalizeComparable(value);
}

function pickFirstFilled(row, keys) {
  for (const key of keys) {
    const value = row?.[key];
    if (value !== null && value !== undefined && String(value).trim()) return String(value).trim();
  }
  return "";
}

function resolve0800OperatorRaw(row) {
  return pickFirstFilled(row, [
    "usuario_de_abertura_da_ocorrencia",
    "usuario_abertura_ocorrencia",
    "usuario_abertura",
    "usuario",
    "operador",
    "atendente",
    "nome_do_operador",
    "nome_operador",
    "agente",
    "email_do_atendente",
    "email",
    "login",
  ]);
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
      production_0800: 0,
      production_nuvidio: 0,
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
  if ("production_0800" in values || "production_nuvidio" in values) {
    metric.production = toInt(metric.production_0800) + toInt(metric.production_nuvidio);
  } else if ("production" in values) {
    metric.production = toInt(values.production);
  }
  return metric;
}

function importNormalizedRows(db, users, rows, sourceName) {
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
  const first = rows[0] || {};
  const missing = required.filter((field) => !(field in first));
  if (missing.length) {
    throw new Error(`Colunas obrigatorias ausentes: ${missing.join(", ")}`);
  }
  let processed = 0;
  let rejected = 0;
  const errors = [];
  rows.forEach((row, index) => {
    try {
      const user = matchUser(users, row);
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
        sourceName,
      );
      processed += 1;
    } catch (error) {
      rejected += 1;
      errors.push({ row: index + 2, error: error.message });
    }
  });
  return { processed, rejected, errors };
}

function import0800Rows(db, users, rows, sourceName) {
  const aggregates = new Map();
  const errors = [];
  rows.forEach((row, index) => {
    try {
      const rawOperator = resolve0800OperatorRaw(row);
      const user = matchUser(users, {
        nome: rawOperator,
        usuario: rawOperator,
        login: rawOperator,
        plataforma_0800: rawOperator,
        id_0800: rawOperator,
      });
      if (!user) throw new Error(`Operador do 0800 nao encontrado no cadastro: ${rawOperator || "sem identificador"}`);
      const metricDate = parseDate(row.data_abertura_ocorrencia);
      const key = `${user.id}::${metricDate}`;
      const current = aggregates.get(key) || {
        userId: user.id,
        metricDate,
        production_0800: 0,
        calls_0800_approved: 0,
        calls_0800_rejected: 0,
        calls_0800_pending: 0,
        calls_0800_no_action: 0,
      };
      current.production_0800 += 1;
      const reason = normalize0800Reason(row.motivo);
      if (reason === "APROVADA" || reason === "APROVADO") current.calls_0800_approved += 1;
      else if (reason === "CANCELADA" || reason === "REPROVADA") current.calls_0800_rejected += 1;
      else if (reason === "PENDENCIADA" || reason === "PENDENCIADO") current.calls_0800_pending += 1;
      else current.calls_0800_no_action += 1;
      aggregates.set(key, current);
    } catch (error) {
      errors.push({ row: index + 2, error: error.message });
    }
  });

  for (const aggregate of aggregates.values()) {
    upsertDailyMetric(
      db,
      aggregate.userId,
      aggregate.metricDate,
      {
        production_0800: aggregate.production_0800,
        calls_0800_approved: aggregate.calls_0800_approved,
        calls_0800_rejected: aggregate.calls_0800_rejected,
        calls_0800_pending: aggregate.calls_0800_pending,
        calls_0800_no_action: aggregate.calls_0800_no_action,
      },
      sourceName,
    );
  }

  return { processed: aggregates.size, rejected: errors.length, errors };
}

function importNuvidioRows(db, users, rows, sourceName) {
  const aggregates = new Map();
  const errors = [];
  rows.forEach((row, index) => {
    try {
      const user = matchUser(users, {
        login: row.email_do_atendente,
        usuario: row.email_do_atendente,
        plataforma_nuvidio: row.email_do_atendente,
        id_nuvidio: row.email_do_atendente,
      });
      if (!user) throw new Error("Operador da Nuvidio nao encontrado no cadastro.");
      const metricDate = parseDate(row.data_abreviada);
      const key = `${user.id}::${metricDate}`;
      const current = aggregates.get(key) || {
        userId: user.id,
        metricDate,
        production_nuvidio: 0,
        calls_nuvidio_approved: 0,
        calls_nuvidio_rejected: 0,
        calls_nuvidio_no_action: 0,
      };
      current.production_nuvidio += 1;
      const tag = normalizeNuvidioTag(row.tag);
      if (tag === "aprovada" || tag === "aprovado") current.calls_nuvidio_approved += 1;
      else if (tag === "reprovada" || tag === "reprovado") current.calls_nuvidio_rejected += 1;
      else current.calls_nuvidio_no_action += 1;
      aggregates.set(key, current);
    } catch (error) {
      errors.push({ row: index + 2, error: error.message });
    }
  });

  for (const aggregate of aggregates.values()) {
    upsertDailyMetric(
      db,
      aggregate.userId,
      aggregate.metricDate,
      {
        production_nuvidio: aggregate.production_nuvidio,
        calls_nuvidio_approved: aggregate.calls_nuvidio_approved,
        calls_nuvidio_rejected: aggregate.calls_nuvidio_rejected,
        calls_nuvidio_no_action: aggregate.calls_nuvidio_no_action,
      },
      sourceName,
    );
  }

  return { processed: aggregates.size, rejected: errors.length, errors };
}

function importNuvidioSummaryRows(db, users, rows, sourceName) {
  let processed = 0;
  let rejected = 0;
  const errors = [];

  rows.forEach((row, index) => {
    try {
      const user = matchUser(users, {
        login: row.usuario_nuvidio,
        usuario: row.usuario_nuvidio,
        plataforma_nuvidio: row.usuario_nuvidio,
        id_nuvidio: row.usuario_nuvidio,
      });
      if (!user) throw new Error("Operador da Nuvidio não encontrado no cadastro.");
      const metricDate = parseDate(row.data);
      upsertDailyMetric(
        db,
        user.id,
        metricDate,
        {
          production_nuvidio:
            toInt(row.nuvidio_aprovadas) +
            toInt(row.nuvidio_reprovadas) +
            toInt(row.nuvidio_sem_acao),
          calls_nuvidio_approved: toInt(row.nuvidio_aprovadas),
          calls_nuvidio_rejected: toInt(row.nuvidio_reprovadas),
          calls_nuvidio_no_action: toInt(row.nuvidio_sem_acao),
        },
        sourceName,
      );
      processed += 1;
    } catch (error) {
      rejected += 1;
      errors.push({ row: index + 2, error: error.message });
    }
  });

  return { processed, rejected, errors };
}

function registerImportLog(db, userId, sourceName, processed, rejected) {
  db.importLogs.push({
    id: nextId(db, "importLogs"),
    imported_by: userId,
    source_name: sourceName,
    imported_at: nowIso(),
    rows_processed: processed,
    rows_rejected: rejected,
  });
}

function buildBaseTemplateCsv(users = [], model = "nuvidio") {
  const is0800 = String(model || "").toLowerCase() === "0800";
  const header = is0800
    ? "Usuário 0800;0800 Aprovadas;0800 Reprovadas;0800 Pendenciadas;0800 Sem ação;Data"
    : "Usuário Nuvidio;Nuvidio Aprovadas;Nuvidio Reprovadas;Nuvidio Sem ação;Data";
  const rows = users.map((user) => {
    if (is0800) {
      const id0800 = String(user.platform_0800_id || user.login || "").replaceAll('"', '""');
      return `"${id0800}";0;0;0;0;2026-01-01`;
    }
    const nuvidio = String(user.nuvidio_id || user.login || "").replaceAll('"', '""');
    return `"${nuvidio}";0;0;0;2026-01-01`;
  });
  return [header, ...rows].join("\n");
}

function normalizeOperationTag(operation, tag) {
  const op = normalizeComparable(operation);
  const normalizedTag = normalizeComparable(tag);
  return { op, normalizedTag };
}

function buildManualValues(operation, tag, quantity) {
  const qty = Math.max(0, toInt(quantity));
  const { op, normalizedTag } = normalizeOperationTag(operation, tag);
  const values = {};
  if (op.includes("0800") || op === "integrall") {
    values.production_0800 = qty;
    if (normalizedTag.includes("aprova")) values.calls_0800_approved = qty;
    else if (normalizedTag.includes("reprova") || normalizedTag.includes("cancel")) values.calls_0800_rejected = qty;
    else if (normalizedTag.includes("pendenc")) values.calls_0800_pending = qty;
    else values.calls_0800_no_action = qty;
    return values;
  }

  values.production_nuvidio = qty;
  if (normalizedTag.includes("aprova")) values.calls_nuvidio_approved = qty;
  else if (normalizedTag.includes("reprova")) values.calls_nuvidio_rejected = qty;
  else values.calls_nuvidio_no_action = qty;
  return values;
}

function import0800SummaryRows(db, users, rows, sourceName) {
  let processed = 0;
  let rejected = 0;
  const errors = [];

  rows.forEach((row, index) => {
    try {
      const user = matchUser(users, {
        usuario: row.usuario_0800,
        plataforma_0800: row.usuario_0800,
        id_0800: row.usuario_0800,
        login: row.usuario_0800,
      });
      if (!user) throw new Error("Operador do 0800 não encontrado no cadastro.");
      const metricDate = parseDate(row.data);
      const approved = toInt(row["0800_aprovadas"]);
      const rejectedValue = toInt(row["0800_reprovadas"]);
      const pending = toInt(row["0800_pendenciadas"]);
      const noAction = toInt(row["0800_sem_acao"]);
      upsertDailyMetric(
        db,
        user.id,
        metricDate,
        {
          production_0800: approved + rejectedValue + pending + noAction,
          calls_0800_approved: approved,
          calls_0800_rejected: rejectedValue,
          calls_0800_pending: pending,
          calls_0800_no_action: noAction,
        },
        sourceName,
      );
      processed += 1;
    } catch (error) {
      rejected += 1;
      errors.push({ row: index + 2, error: error.message });
    }
  });

  return { processed, rejected, errors };
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
  const requestedRawUserId = String(url.searchParams.get("user_id") || user.id).trim().toLowerCase();
  const includeAllUsers = user.role === "manager" && requestedRawUserId === "all";
  const requestedUserId = Number(requestedRawUserId || user.id);
  const targetUserId = user.role === "manager" ? requestedUserId : user.id;
  const history = db.dailyMetrics
    .filter((row) => {
      if (row.metric_date < start || row.metric_date > end) return false;
      if (includeAllUsers) return true;
      return row.user_id === targetUserId;
    })
    .sort((a, b) => b.metric_date.localeCompare(a.metric_date))
    .map((row) => ({ ...row, effectiveness: calculateEffectiveness(row) }));
  const quality = db.qualityScores
    .filter((item) => {
      if (includeAllUsers) return true;
      return item.user_id === targetUserId;
    })
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
    return jsonResponse({
      user: user ? serializeUser(user) : null,
      app_settings: serializeAppSettings(db),
    });
  }

  if (url.pathname === "/api/auth/login" && request.method === "POST") {
    const payload = await request.json();
    const user = db.users.find((entry) => entry.login === String(payload.login || "").trim() && entry.is_active);
    const inputPassword = String(payload.password || "");
    const hashMatches = user ? await verifyPassword(inputPassword, user.password_hash) : false;
    const plainMatches = user ? String(user.password_plain || "") === inputPassword : false;
    if (!user || (!hashMatches && !plainMatches)) {
      return jsonResponse({ error: "Credenciais invalidas" }, 401);
    }
    if (plainMatches && !hashMatches) {
      user.password_hash = await hashPassword(inputPassword);
      user.updated_at = nowIso();
      await persistStorage(db, env);
    }
    const token = await signSessionWithEnv({
      user_id: user.id,
      expires_at: new Date(Date.now() + 7 * 86400000).toISOString(),
    }, env);
    return jsonResponse(
      {
        user: serializeUser(user),
        app_settings: serializeAppSettings(db),
      },
      200,
      { "set-cookie": `${SESSION_NAME}=${token}; Path=/; HttpOnly; SameSite=Lax` },
    );
  }

  if (url.pathname === "/api/auth/logout" && request.method === "POST") {
    return jsonResponse({ ok: true }, 200, {
      "set-cookie": `${SESSION_NAME}=; Path=/; Expires=Thu, 01 Jan 1970 00:00:00 GMT; HttpOnly; SameSite=Lax`,
    });
  }

  if (url.pathname === "/api/auth/preferences" && request.method === "PATCH") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    if (payload.preferred_theme !== undefined) {
      auth.user.preferred_theme = payload.preferred_theme === "contrast" ? "contrast" : "dark";
    }
    if (payload.last_route !== undefined) {
      auth.user.last_route = normalizeRouteForRole(payload.last_route, auth.user.role);
    }
    auth.user.updated_at = nowIso();
    await persistStorage(db, env);
    return jsonResponse({ user: serializeUser(auth.user) });
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
    auth.user.password_plain = newPassword;
    auth.user.must_change_password = false;
    auth.user.updated_at = nowIso();
    await persistStorage(db, env);
    return jsonResponse({ ok: true, user: serializeUser(auth.user) });
  }

  if (url.pathname === "/api/admin/settings" && request.method === "GET") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse({ app_settings: serializeAppSettings(db) });
  }

  if (url.pathname === "/api/admin/settings" && request.method === "PATCH") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    if (payload.maintenance_for_operators !== undefined) {
      db.appSettings.maintenance_for_operators = Boolean(payload.maintenance_for_operators);
    }
    if (payload.maintenance_message !== undefined) {
      db.appSettings.maintenance_message = String(payload.maintenance_message || "").trim()
        || "Portal em manutenção. Tente novamente em instantes.";
    }
    await persistStorage(db, env);
    return jsonResponse({ app_settings: serializeAppSettings(db) });
  }

  if (url.pathname === "/api/dashboard/overview" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    if (isOperatorBlockedByMaintenance(auth.user, db)) {
      return jsonResponse({ error: "Portal em manutenção para operadores." }, 503);
    }
    return jsonResponse(buildOverview(db, auth.user, url));
  }

  if (url.pathname === "/api/analysis" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    if (isOperatorBlockedByMaintenance(auth.user, db)) {
      return jsonResponse({ error: "Portal em manutenção para operadores." }, 503);
    }
    return jsonResponse(buildAnalysis(db, auth.user, url));
  }

  if (url.pathname === "/api/history" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    if (isOperatorBlockedByMaintenance(auth.user, db)) {
      return jsonResponse({ error: "Portal em manutenção para operadores." }, 503);
    }
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
      password_plain: DEFAULT_PASSWORD,
      role: payload.role,
      platform_0800_id: String(payload.platform_0800_id || "").trim(),
      nuvidio_id: String(payload.nuvidio_id || "").trim(),
      must_change_password: true,
      is_active: true,
      preferred_theme: "dark",
      last_route: "overview",
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
    const wantsActive = payload.is_active === undefined ? user.is_active : Boolean(payload.is_active);
    if (!wantsActive && user.id === auth.user.id) {
      return jsonResponse({ error: "Voce nao pode desativar o proprio usuario" }, 400);
    }
    user.full_name = String(payload.full_name || user.full_name).trim();
    user.login = String(payload.login || user.login).trim();
    user.role = payload.role || user.role;
    user.platform_0800_id = String(payload.platform_0800_id ?? user.platform_0800_id).trim();
    user.nuvidio_id = String(payload.nuvidio_id ?? user.nuvidio_id).trim();
    user.is_active = wantsActive;
    user.updated_at = nowIso();
    if (String(payload.password || "").trim()) {
      user.password_hash = await hashPassword(String(payload.password).trim());
      user.password_plain = String(payload.password).trim();
      user.must_change_password = true;
    }
    await persistStorage(db, env);
    return jsonResponse({ user: serializeUser(user) });
  }

  if (url.pathname.startsWith("/api/admin/users/") && request.method === "DELETE") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const userId = Number(url.pathname.split("/").pop());
    if (userId === auth.user.id) {
      return jsonResponse({ error: "Voce nao pode apagar o proprio usuario" }, 400);
    }
    const userIndex = db.users.findIndex((entry) => entry.id === userId);
    if (userIndex === -1) return jsonResponse({ error: "Usuario nao encontrado" }, 404);
    db.users.splice(userIndex, 1);
    db.dailyMetrics = db.dailyMetrics.filter((entry) => entry.user_id !== userId);
    db.qualityScores = db.qualityScores.filter((entry) => entry.user_id !== userId);
    await persistStorage(db, env);
    return jsonResponse({ ok: true });
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
        "Nome do Operador;Monitoria 1;Monitoria 2;Monitoria 3;Monitoria 4",
        ...operators.map((user) => `"${String(user.full_name || "").replaceAll('"', '""')}";;;;`),
      ].join("\n");
      return new Response(`\uFEFF${csv}`, {
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

  if (url.pathname === "/api/admin/import/template" && request.method === "GET") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const model = String(url.searchParams.get("model") || "nuvidio").toLowerCase();
    const operators = db.users.filter((user) => user.is_active && user.role === "operator");
    const csv = buildBaseTemplateCsv(operators, model);
    return new Response(`\uFEFF${csv}`, {
      status: 200,
      headers: {
        "content-type": "text/csv; charset=utf-8",
        "content-disposition": `attachment; filename="modelo-base-${model === "0800" ? "0800" : "nuvidio"}.csv"`,
      },
    });
  }

  if (url.pathname === "/api/admin/manual-tag" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    const userId = Number(payload.user_id);
    const metricDate = parseDate(payload.date || todayIso());
    const operation = String(payload.operation || "").trim().toLowerCase();
    const approved = Math.max(0, toInt(payload.approved));
    const rejected = Math.max(0, toInt(payload.rejected));
    const noAction = Math.max(0, toInt(payload.no_action));
    const pending = Math.max(0, toInt(payload.pending));
    if (!userId || !operation) {
      return jsonResponse({ error: "Preencha operador, data e operação." }, 400);
    }
    const user = db.users.find((entry) => entry.id === userId && entry.is_active && entry.role === "operator");
    if (!user) return jsonResponse({ error: "Operador não encontrado." }, 404);
    const values = operation === "0800"
      ? {
          production_0800: approved + rejected + pending + noAction,
          calls_0800_approved: approved,
          calls_0800_rejected: rejected,
          calls_0800_pending: pending,
          calls_0800_no_action: noAction,
        }
      : {
          production_nuvidio: approved + rejected + noAction,
          calls_nuvidio_approved: approved,
          calls_nuvidio_rejected: rejected,
          calls_nuvidio_no_action: noAction,
        };
    upsertDailyMetric(db, user.id, metricDate, values, "manual_tag");
    await persistStorage(db, env);
    return jsonResponse({ ok: true });
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

  if (url.pathname.startsWith("/api/admin/daily-metrics/") && request.method === "DELETE") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const metricId = Number(url.pathname.split("/").pop());
    const index = db.dailyMetrics.findIndex((entry) => entry.id === metricId);
    if (index === -1) return jsonResponse({ error: "Registro nao encontrado" }, 404);
    db.dailyMetrics.splice(index, 1);
    await persistStorage(db, env);
    return jsonResponse({ ok: true });
  }

  if (url.pathname === "/api/admin/import/r2" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    if (!env.IMPORTS_BUCKET || typeof env.IMPORTS_BUCKET.list !== "function") {
      return jsonResponse({ error: "Bucket R2 nao configurado para este ambiente." }, 400);
    }

    const listing = await env.IMPORTS_BUCKET.list({ limit: 500 });
    const objects = (listing.objects || [])
      .filter((item) => /\.(csv|xlsx)$/i.test(item.key))
      .sort((a, b) => new Date(b.uploaded || 0).getTime() - new Date(a.uploaded || 0).getTime());

    if (!objects.length) {
      return jsonResponse({ error: "Nenhum arquivo CSV ou XLSX encontrado no R2." }, 404);
    }

    const processedKeys = new Set((db.r2ProcessedKeys || []).map((key) => String(key)));
    const parsedCandidates = objects.filter((item) => /parsed\//i.test(item.key) && /\.csv$/i.test(item.key));
    const unprocessedParsed = parsedCandidates.filter((item) => !processedKeys.has(item.key));
    const unprocessedCsv = objects.filter((item) => /\.csv$/i.test(item.key) && !processedKeys.has(item.key));

    const pickByOperation = (list, operationRegex) =>
      list.find((item) => operationRegex.test(item.key));

    const parsed0800 = pickByOperation(unprocessedParsed, /0800/i);
    const parsedNuvidio = pickByOperation(unprocessedParsed, /nuvidio/i);
    const raw0800 = pickByOperation(unprocessedCsv, /0800/i);
    const rawNuvidio = pickByOperation(unprocessedCsv, /nuvidio/i);

    const uniqueByKey = (items) => {
      const map = new Map();
      for (const item of items) {
        if (item?.key && !map.has(item.key)) map.set(item.key, item);
      }
      return [...map.values()];
    };

    const toProcess = uniqueByKey([
      parsed0800 || raw0800,
      parsedNuvidio || rawNuvidio,
    ]).slice(0, 2);
    const summary = {
      processedFiles: 0,
      processedRows: 0,
      rejectedRows: 0,
      files: [],
      skipped: [],
      mode: "incremental_single_file",
    };

    for (const item of toProcess) {
      const object = await env.IMPORTS_BUCKET.get(item.key);
      if (!object) {
        summary.skipped.push(`${item.key}: arquivo nao encontrado no bucket.`);
        continue;
      }

      if (/\.xlsx$/i.test(item.key)) {
        summary.skipped.push(`${item.key}: leitura automatica de XLSX no R2 ainda nao esta habilitada. Use CSV para a base operacional.`);
        continue;
      }

      if (Number(object.size || 0) > R2_MAX_FILE_BYTES) {
        summary.skipped.push(`${item.key}: arquivo acima do limite de ${Math.round(R2_MAX_FILE_BYTES / (1024 * 1024))}MB para processamento em Worker.`);
        continue;
      }

      const text = await object.text();
      const rows = parseCsv(text, R2_MAX_ROWS_PER_RUN);
      if (!rows.length) {
        summary.skipped.push(`${item.key}: arquivo vazio.`);
        continue;
      }
      if (rows.length >= R2_MAX_ROWS_PER_RUN) {
        summary.skipped.push(`${item.key}: limite de ${R2_MAX_ROWS_PER_RUN} linhas por sincronização atingido (rode novamente para continuar).`);
      }

      const schema = detectImportSchema(rows[0], item.key);
      let result;
      if (schema === "normalized") result = importNormalizedRows(db, db.users.filter((entry) => entry.is_active), rows, item.key);
      else if (schema === "0800") result = import0800Rows(db, db.users.filter((entry) => entry.is_active), rows, item.key);
      else if (schema === "nuvidio") result = importNuvidioRows(db, db.users.filter((entry) => entry.is_active), rows, item.key);
      else if (schema === "nuvidio_summary") result = importNuvidioSummaryRows(db, db.users.filter((entry) => entry.is_active), rows, item.key);
      else {
        summary.skipped.push(`${item.key}: formato nao reconhecido para importacao automatica.`);
        continue;
      }

      registerImportLog(db, auth.user.id, item.key, result.processed, result.rejected);
      if (result.processed > 0 && !db.r2ProcessedKeys.includes(item.key)) db.r2ProcessedKeys.push(item.key);
      summary.processedFiles += 1;
      summary.processedRows += result.processed;
      summary.rejectedRows += result.rejected;
      summary.files.push({
        key: item.key,
        schema,
        processed: result.processed,
        rejected: result.rejected,
      });
      result.errors.slice(0, 10).forEach((entry) => {
        summary.skipped.push(`${item.key} linha ${entry.row}: ${entry.error}`);
      });
    }

    if (!toProcess.length) {
      summary.skipped.push("Nenhum arquivo novo para processar no R2.");
    }

    await persistStorage(db, env);
    return jsonResponse(summary);
  }

  if (url.pathname === "/api/admin/import" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const form = await request.formData();
    const file = form.get("file");
    const model = String(form.get("model") || "").toLowerCase();
    if (!(file instanceof File)) {
      return jsonResponse({ error: "Arquivo nao enviado" }, 400);
    }
    if (!file.name.toLowerCase().endsWith(".csv")) {
      return jsonResponse({ error: "No formato atual, use CSV para o teste local." }, 400);
    }
    const text = await file.text();
    const rows = parseCsv(text);
    if (!rows.length) return jsonResponse({ error: "A planilha esta vazia." }, 400);
    const first = rows[0];
    const schema = model === "0800"
      ? "0800_summary"
      : model === "nuvidio"
        ? "nuvidio_summary"
        : detectImportSchema(first, file.name);
    let outcome;
    if (schema === "normalized") outcome = importNormalizedRows(db, db.users.filter((entry) => entry.is_active), rows, file.name);
    else if (schema === "0800") outcome = import0800Rows(db, db.users.filter((entry) => entry.is_active), rows, file.name);
    else if (schema === "nuvidio") outcome = importNuvidioRows(db, db.users.filter((entry) => entry.is_active), rows, file.name);
    else if (schema === "0800_summary") outcome = import0800SummaryRows(db, db.users.filter((entry) => entry.is_active), rows, file.name);
    else if (schema === "nuvidio_summary") outcome = importNuvidioSummaryRows(db, db.users.filter((entry) => entry.is_active), rows, file.name);
    else return jsonResponse({ error: "Formato de planilha não reconhecido." }, 400);
    const { processed, rejected, errors } = outcome;
    registerImportLog(db, auth.user.id, file.name, processed, rejected);
    await persistStorage(db, env);
    return jsonResponse({ processed, rejected, errors: errors.slice(0, 20) });
  }

  return jsonResponse({ error: "Rota nao encontrada" }, 404);
}

async function serveStatic(request, env = {}) {
  const pathname = new URL(request.url).pathname;
  if (!isLocalNodeRuntime) {
    if (env?.ASSETS && typeof env.ASSETS.fetch === "function") {
      const assetResponse = await env.ASSETS.fetch(request);
      if (assetResponse.status !== 404) return assetResponse;
    }
    if (pathname === "/logos_KR-02.png") {
      return new Response("Not found", { status: 404 });
    }
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
            details: error?.stack || error?.message || String(error),
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
      console.log(`Pulse Ops disponÃ­vel em http://localhost:${port}`);
    });
  }
}




