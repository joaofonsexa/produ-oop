import { SCRIPT_JS, STYLES_CSS } from "./asset-bundle.js";

const SESSION_NAME = "pulse_session";
const DB_PATH = "./data/db.json";
const D1_STATE_ID = 1;
const FAVICON_DATA = "";
const D1_SCHEMA_STATEMENTS = [
  "CREATE TABLE IF NOT EXISTS app_state (id INTEGER PRIMARY KEY, data TEXT NOT NULL, updated_at TEXT NOT NULL)",
  "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, full_name TEXT NOT NULL, login TEXT NOT NULL UNIQUE, password_hash TEXT NOT NULL, password_plain TEXT, role TEXT NOT NULL, platform_0800_id TEXT, nuvidio_id TEXT, must_change_password INTEGER NOT NULL DEFAULT 1, is_active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, updated_at TEXT NOT NULL)",
  "CREATE INDEX IF NOT EXISTS idx_users_login ON users(login)",
  "CREATE TABLE IF NOT EXISTS daily_metrics (id INTEGER PRIMARY KEY, user_id INTEGER NOT NULL, metric_date TEXT NOT NULL, production INTEGER NOT NULL DEFAULT 0, production_0800 INTEGER NOT NULL DEFAULT 0, production_nuvidio INTEGER NOT NULL DEFAULT 0, calls_0800_approved INTEGER NOT NULL DEFAULT 0, calls_0800_rejected INTEGER NOT NULL DEFAULT 0, calls_0800_pending INTEGER NOT NULL DEFAULT 0, calls_0800_no_action INTEGER NOT NULL DEFAULT 0, calls_nuvidio_approved INTEGER NOT NULL DEFAULT 0, calls_nuvidio_rejected INTEGER NOT NULL DEFAULT 0, calls_nuvidio_no_action INTEGER NOT NULL DEFAULT 0, calls_nuvidio_empty INTEGER NOT NULL DEFAULT 0, import_source TEXT, created_at TEXT NOT NULL, updated_at TEXT NOT NULL)",
  "CREATE INDEX IF NOT EXISTS idx_daily_metrics_user_date ON daily_metrics(user_id, metric_date)",
  "CREATE TABLE IF NOT EXISTS quality_scores (id INTEGER PRIMARY KEY, user_id INTEGER NOT NULL, reference_month TEXT NOT NULL, score REAL NOT NULL DEFAULT 0, monitoria_1 REAL, monitoria_2 REAL, monitoria_3 REAL, monitoria_4 REAL, notes TEXT, created_at TEXT NOT NULL, updated_at TEXT NOT NULL)",
  "CREATE INDEX IF NOT EXISTS idx_quality_scores_user_month ON quality_scores(user_id, reference_month)",
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
let storageCacheLoadedAt = 0;
const STORAGE_CACHE_TTL_MS = 1500;

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

function isSaturdayIsoDate(dateValue) {
  const raw = String(dateValue || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return false;
  return new Date(`${raw}T00:00:00`).getDay() === 6;
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

function toOptionalMonitoria(value) {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  if (!text) return null;
  const parsed = Number(text.replace(",", "."));
  if (!Number.isFinite(parsed)) return null;
  return parsed;
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

function repairTextEncoding(value) {
  const raw = String(value ?? "");
  if (!raw || !/[ÃÂâ]/.test(raw)) return raw;
  try {
    const bytes = Uint8Array.from([...raw].map((char) => char.charCodeAt(0) & 0xff));
    const decoded = new TextDecoder("utf-8").decode(bytes);
    const originalNoise = (raw.match(/[ÃÂâ]/g) || []).length;
    const decodedNoise = (decoded.match(/[ÃÂâ�]/g) || []).length;
    return decodedNoise < originalNoise ? decoded : raw;
  } catch {
    return raw;
  }
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
    full_name: repairTextEncoding(user.full_name),
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
    ? ["overview", "analysis", "alerts", "history", "admin"]
    : ["overview", "analysis", "history"];
  return allowed.includes(safeRoute) ? safeRoute : "overview";
}

function normalizeUserRecord(user, roleFallback = "operator") {
  const role = user.role || roleFallback;
  return {
    ...user,
    full_name: repairTextEncoding(user.full_name),
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
      full_name: "João Fonseca",
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
    metric_rules: {
      production: { red_max: 70, amber_max: 100 },
      effectiveness: { red_max: 70, amber_max: 90 },
      quality: { red_max: 70, amber_max: 90 },
    },
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

function normalizeD1DailyMetricRow(row) {
  return {
    id: Number(row.id),
    user_id: Number(row.user_id),
    metric_date: String(row.metric_date || "").trim(),
    production: toInt(row.production),
    production_0800: toInt(row.production_0800),
    production_nuvidio: toInt(row.production_nuvidio),
    calls_0800_approved: toInt(row.calls_0800_approved),
    calls_0800_rejected: toInt(row.calls_0800_rejected),
    calls_0800_pending: toInt(row.calls_0800_pending),
    calls_0800_no_action: toInt(row.calls_0800_no_action),
    calls_nuvidio_approved: toInt(row.calls_nuvidio_approved),
    calls_nuvidio_rejected: toInt(row.calls_nuvidio_rejected),
    calls_nuvidio_no_action: toInt(row.calls_nuvidio_no_action),
    calls_nuvidio_empty: toInt(row.calls_nuvidio_empty),
    import_source: String(row.import_source || "").trim(),
    created_at: row.created_at || nowIso(),
    updated_at: row.updated_at || nowIso(),
  };
}

function normalizeD1QualityScoreRow(row) {
  const normalizeMonitoria = (value) => (value === null || value === undefined || value === "" ? null : toFloat(value));
  return {
    id: Number(row.id),
    user_id: Number(row.user_id),
    reference_month: String(row.reference_month || "").trim(),
    score: toFloat(row.score),
    monitoria_1: normalizeMonitoria(row.monitoria_1),
    monitoria_2: normalizeMonitoria(row.monitoria_2),
    monitoria_3: normalizeMonitoria(row.monitoria_3),
    monitoria_4: normalizeMonitoria(row.monitoria_4),
    notes: String(row.notes || "").trim(),
    created_at: row.created_at || nowIso(),
    updated_at: row.updated_at || nowIso(),
  };
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

async function loadDailyMetricsFromD1(connection) {
  const result = await connection.prepare(`
    SELECT
      id,
      user_id,
      metric_date,
      production,
      production_0800,
      production_nuvidio,
      calls_0800_approved,
      calls_0800_rejected,
      calls_0800_pending,
      calls_0800_no_action,
      calls_nuvidio_approved,
      calls_nuvidio_rejected,
      calls_nuvidio_no_action,
      calls_nuvidio_empty,
      import_source,
      created_at,
      updated_at
    FROM daily_metrics
    ORDER BY metric_date, id
  `).all();
  return (result?.results || []).map(normalizeD1DailyMetricRow);
}

async function loadQualityScoresFromD1(connection) {
  const result = await connection.prepare(`
    SELECT
      id,
      user_id,
      reference_month,
      score,
      monitoria_1,
      monitoria_2,
      monitoria_3,
      monitoria_4,
      notes,
      created_at,
      updated_at
    FROM quality_scores
    ORDER BY reference_month, id
  `).all();
  return (result?.results || []).map(normalizeD1QualityScoreRow);
}

async function persistUserRecordToD1(connection, user) {
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

async function persistDailyMetricRecordToD1(connection, metric) {
  await connection.prepare(`
    INSERT INTO daily_metrics (
      id,
      user_id,
      metric_date,
      production,
      production_0800,
      production_nuvidio,
      calls_0800_approved,
      calls_0800_rejected,
      calls_0800_pending,
      calls_0800_no_action,
      calls_nuvidio_approved,
      calls_nuvidio_rejected,
      calls_nuvidio_no_action,
      calls_nuvidio_empty,
      import_source,
      created_at,
      updated_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      user_id = excluded.user_id,
      metric_date = excluded.metric_date,
      production = excluded.production,
      production_0800 = excluded.production_0800,
      production_nuvidio = excluded.production_nuvidio,
      calls_0800_approved = excluded.calls_0800_approved,
      calls_0800_rejected = excluded.calls_0800_rejected,
      calls_0800_pending = excluded.calls_0800_pending,
      calls_0800_no_action = excluded.calls_0800_no_action,
      calls_nuvidio_approved = excluded.calls_nuvidio_approved,
      calls_nuvidio_rejected = excluded.calls_nuvidio_rejected,
      calls_nuvidio_no_action = excluded.calls_nuvidio_no_action,
      calls_nuvidio_empty = excluded.calls_nuvidio_empty,
      import_source = excluded.import_source,
      created_at = excluded.created_at,
      updated_at = excluded.updated_at
  `).bind(
    Number(metric.id),
    Number(metric.user_id),
    String(metric.metric_date || "").trim(),
    toInt(metric.production),
    toInt(metric.production_0800),
    toInt(metric.production_nuvidio),
    toInt(metric.calls_0800_approved),
    toInt(metric.calls_0800_rejected),
    toInt(metric.calls_0800_pending),
    toInt(metric.calls_0800_no_action),
    toInt(metric.calls_nuvidio_approved),
    toInt(metric.calls_nuvidio_rejected),
    toInt(metric.calls_nuvidio_no_action),
    toInt(metric.calls_nuvidio_empty),
    String(metric.import_source || "").trim(),
    metric.created_at || nowIso(),
    metric.updated_at || nowIso(),
  ).run();
}

async function persistQualityScoreRecordToD1(connection, score) {
  await connection.prepare(`
    INSERT INTO quality_scores (
      id,
      user_id,
      reference_month,
      score,
      monitoria_1,
      monitoria_2,
      monitoria_3,
      monitoria_4,
      notes,
      created_at,
      updated_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      user_id = excluded.user_id,
      reference_month = excluded.reference_month,
      score = excluded.score,
      monitoria_1 = excluded.monitoria_1,
      monitoria_2 = excluded.monitoria_2,
      monitoria_3 = excluded.monitoria_3,
      monitoria_4 = excluded.monitoria_4,
      notes = excluded.notes,
      created_at = excluded.created_at,
      updated_at = excluded.updated_at
  `).bind(
    Number(score.id),
    Number(score.user_id),
    String(score.reference_month || "").trim(),
    toFloat(score.score),
    score.monitoria_1 === null || score.monitoria_1 === undefined || score.monitoria_1 === "" ? null : toFloat(score.monitoria_1),
    score.monitoria_2 === null || score.monitoria_2 === undefined || score.monitoria_2 === "" ? null : toFloat(score.monitoria_2),
    score.monitoria_3 === null || score.monitoria_3 === undefined || score.monitoria_3 === "" ? null : toFloat(score.monitoria_3),
    score.monitoria_4 === null || score.monitoria_4 === undefined || score.monitoria_4 === "" ? null : toFloat(score.monitoria_4),
    String(score.notes || "").trim(),
    score.created_at || nowIso(),
    score.updated_at || nowIso(),
  ).run();
}

async function persistUsersToD1(connection, users) {
  for (const user of users) {
    await persistUserRecordToD1(connection, user);
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

async function persistDailyMetricsToD1(connection, dailyMetrics) {
  for (const metric of dailyMetrics) {
    await persistDailyMetricRecordToD1(connection, metric);
  }
}

async function persistQualityScoresToD1(connection, qualityScores) {
  for (const score of qualityScores) {
    await persistQualityScoreRecordToD1(connection, score);
  }
}

async function deleteDailyMetricRecordFromD1(connection, metricId) {
  await connection.prepare("DELETE FROM daily_metrics WHERE id = ?").bind(Number(metricId)).run();
}

async function deleteUserDataFromD1(connection, userId) {
  await connection.prepare("DELETE FROM quality_scores WHERE user_id = ?").bind(Number(userId)).run();
  await connection.prepare("DELETE FROM daily_metrics WHERE user_id = ?").bind(Number(userId)).run();
  await connection.prepare("DELETE FROM users WHERE id = ?").bind(Number(userId)).run();
}

function buildPersistableMetaState(db) {
  return {
    counters: db.counters || seedState().counters,
    importLogs: db.importLogs || [],
    r2ProcessedKeys: db.r2ProcessedKeys || [],
    appSettings: db.appSettings || seedState().appSettings,
  };
}

async function persistMetaStateToD1(connection, db) {
  await connection.prepare(`
    INSERT INTO app_state (id, data, updated_at)
    VALUES (?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      data = excluded.data,
      updated_at = excluded.updated_at
  `).bind(D1_STATE_ID, JSON.stringify(buildPersistableMetaState(db)), nowIso()).run();
}

async function ensureD1Schema(connection) {
  for (const statement of D1_SCHEMA_STATEMENTS) {
    await connection.prepare(statement).run();
  }
  const userTableInfo = await connection.prepare("PRAGMA table_info(users)").all();
  const userColumns = new Set((userTableInfo?.results || []).map((row) => String(row.name)));
  if (!userColumns.has("platform_0800_id")) {
    await connection.exec("ALTER TABLE users ADD COLUMN platform_0800_id TEXT");
  }
  if (!userColumns.has("nuvidio_id")) {
    await connection.exec("ALTER TABLE users ADD COLUMN nuvidio_id TEXT");
  }
  if (!userColumns.has("must_change_password")) {
    await connection.exec("ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 1");
  }
  if (!userColumns.has("password_plain")) {
    await connection.exec("ALTER TABLE users ADD COLUMN password_plain TEXT");
  }
  if (!userColumns.has("is_active")) {
    await connection.exec("ALTER TABLE users ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1");
  }
  if (!userColumns.has("preferred_theme")) {
    await connection.exec("ALTER TABLE users ADD COLUMN preferred_theme TEXT NOT NULL DEFAULT 'dark'");
  }
  if (!userColumns.has("last_route")) {
    await connection.exec("ALTER TABLE users ADD COLUMN last_route TEXT NOT NULL DEFAULT 'overview'");
  }

  const dailyMetricsTableInfo = await connection.prepare("PRAGMA table_info(daily_metrics)").all();
  const dailyMetricsColumns = new Set((dailyMetricsTableInfo?.results || []).map((row) => String(row.name)));
  if (!dailyMetricsColumns.has("production_0800")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN production_0800 INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("production_nuvidio")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN production_nuvidio INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("calls_0800_approved")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN calls_0800_approved INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("calls_0800_rejected")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN calls_0800_rejected INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("calls_0800_pending")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN calls_0800_pending INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("calls_0800_no_action")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN calls_0800_no_action INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("calls_nuvidio_approved")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN calls_nuvidio_approved INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("calls_nuvidio_rejected")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN calls_nuvidio_rejected INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("calls_nuvidio_no_action")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN calls_nuvidio_no_action INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("calls_nuvidio_empty")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN calls_nuvidio_empty INTEGER NOT NULL DEFAULT 0");
  }
  if (!dailyMetricsColumns.has("import_source")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN import_source TEXT");
  }
  if (!dailyMetricsColumns.has("created_at")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN created_at TEXT NOT NULL DEFAULT ''");
  }
  if (!dailyMetricsColumns.has("updated_at")) {
    await connection.exec("ALTER TABLE daily_metrics ADD COLUMN updated_at TEXT NOT NULL DEFAULT ''");
  }

  const qualityScoresTableInfo = await connection.prepare("PRAGMA table_info(quality_scores)").all();
  const qualityScoreColumns = new Set((qualityScoresTableInfo?.results || []).map((row) => String(row.name)));
  if (!qualityScoreColumns.has("monitoria_1")) {
    await connection.exec("ALTER TABLE quality_scores ADD COLUMN monitoria_1 REAL");
  }
  if (!qualityScoreColumns.has("monitoria_2")) {
    await connection.exec("ALTER TABLE quality_scores ADD COLUMN monitoria_2 REAL");
  }
  if (!qualityScoreColumns.has("monitoria_3")) {
    await connection.exec("ALTER TABLE quality_scores ADD COLUMN monitoria_3 REAL");
  }
  if (!qualityScoreColumns.has("monitoria_4")) {
    await connection.exec("ALTER TABLE quality_scores ADD COLUMN monitoria_4 REAL");
  }
  if (!qualityScoreColumns.has("notes")) {
    await connection.exec("ALTER TABLE quality_scores ADD COLUMN notes TEXT");
  }
  if (!qualityScoreColumns.has("created_at")) {
    await connection.exec("ALTER TABLE quality_scores ADD COLUMN created_at TEXT NOT NULL DEFAULT ''");
  }
  if (!qualityScoreColumns.has("updated_at")) {
    await connection.exec("ALTER TABLE quality_scores ADD COLUMN updated_at TEXT NOT NULL DEFAULT ''");
  }
}

function calculateEffectiveness(metric) {
  const actionable =
    toInt(metric.calls_0800_approved) +
    toInt(metric.calls_0800_rejected) +
    toInt(metric.calls_0800_pending) +
    toInt(metric.calls_nuvidio_approved) +
    toInt(metric.calls_nuvidio_rejected);
  const total = actionable + toInt(metric.calls_0800_no_action) + toInt(metric.calls_nuvidio_no_action) + toInt(metric.calls_nuvidio_empty);
  if (!total) return 0;
  return Number(((actionable / total) * 100).toFixed(2));
}

function calculateOperationEffectiveness(metric, operation) {
  if (operation === "0800") {
    const actionable = toInt(metric.calls_0800_approved) + toInt(metric.calls_0800_rejected) + toInt(metric.calls_0800_pending);
    const total = actionable + toInt(metric.calls_0800_no_action);
    return total ? Number(((actionable / total) * 100).toFixed(2)) : 0;
  }
  const actionable = toInt(metric.calls_nuvidio_approved) + toInt(metric.calls_nuvidio_rejected);
  const total = actionable + toInt(metric.calls_nuvidio_no_action) + toInt(metric.calls_nuvidio_empty);
  return total ? Number(((actionable / total) * 100).toFixed(2)) : 0;
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
      metric_rules: {
        production: { red_max: 70, amber_max: 100 },
        effectiveness: { red_max: 70, amber_max: 90 },
        quality: { red_max: 70, amber_max: 90 },
      },
    },
  };
}

function serializeAppSettings(db) {
  const fallback = {
    production: { red_max: 70, amber_max: 100 },
    effectiveness: { red_max: 70, amber_max: 90 },
    quality: { red_max: 70, amber_max: 90 },
  };
  const sourceRules = db?.appSettings?.metric_rules || {};
  const metricRules = {
    production: {
      red_max: Number(sourceRules?.production?.red_max ?? fallback.production.red_max),
      amber_max: Number(sourceRules?.production?.amber_max ?? fallback.production.amber_max),
    },
    effectiveness: {
      red_max: Number(sourceRules?.effectiveness?.red_max ?? fallback.effectiveness.red_max),
      amber_max: Number(sourceRules?.effectiveness?.amber_max ?? fallback.effectiveness.amber_max),
    },
    quality: {
      red_max: Number(sourceRules?.quality?.red_max ?? fallback.quality.red_max),
      amber_max: Number(sourceRules?.quality?.amber_max ?? fallback.quality.amber_max),
    },
  };
  return {
    maintenance_for_operators: Boolean(db?.appSettings?.maintenance_for_operators),
    maintenance_message: String(
      db?.appSettings?.maintenance_message || "Portal em manutenção. Tente novamente em instantes.",
    ),
    metric_rules: metricRules,
  };
}

function isOperatorBlockedByMaintenance(user, db) {
  return Boolean(user && user.role !== "manager" && db?.appSettings?.maintenance_for_operators);
}

function rememberStorage(db) {
  storageCache = db;
  storageCacheLoadedAt = Date.now();
  return db;
}

async function ensureStorage(env = {}) {
  if (env?.DB) {
    if (storageCache && (Date.now() - storageCacheLoadedAt) < STORAGE_CACHE_TTL_MS) {
      return storageCache;
    }
    await ensureD1Schema(env.DB);
    const row = await env.DB.prepare("SELECT data FROM app_state WHERE id = ?").bind(D1_STATE_ID).first();
    const seededState = seedState();
    const rawState = row?.data ? JSON.parse(row.data) : {};
    const legacyState = normalizeDbState({
      ...seededState,
      ...rawState,
      users: Array.isArray(rawState?.users) ? rawState.users : seededState.users,
      dailyMetrics: Array.isArray(rawState?.dailyMetrics) ? rawState.dailyMetrics : [],
      qualityScores: Array.isArray(rawState?.qualityScores) ? rawState.qualityScores : [],
    });
    const [loadedUsers, loadedDailyMetrics, loadedQualityScores] = await Promise.all([
      loadUsersFromD1(env.DB),
      loadDailyMetricsFromD1(env.DB),
      loadQualityScoresFromD1(env.DB),
    ]);
    const db = normalizeDbState({
      ...seededState,
      ...legacyState,
      users: loadedUsers.length ? loadedUsers : (legacyState.users?.length ? legacyState.users : seededState.users),
      dailyMetrics: loadedDailyMetrics.length ? loadedDailyMetrics : (legacyState.dailyMetrics || []),
      qualityScores: loadedQualityScores.length ? loadedQualityScores : (legacyState.qualityScores || []),
    });
    if (!loadedUsers.length && !legacyState.users?.length) {
      db.users[0].password_hash = await hashPassword("admin123");
    }
    const repairedPasswords = await ensureDefaultPasswords(db);
    const legacyPayloadHasEmbeddedCollections = Boolean(
      rawState && (
        Object.prototype.hasOwnProperty.call(rawState, "users")
        || Object.prototype.hasOwnProperty.call(rawState, "dailyMetrics")
        || Object.prototype.hasOwnProperty.call(rawState, "qualityScores")
      )
    );
    const shouldPersist =
      !row?.data ||
      !loadedUsers.length ||
      (!loadedDailyMetrics.length && db.dailyMetrics.length > 0) ||
      (!loadedQualityScores.length && db.qualityScores.length > 0) ||
      repairedPasswords ||
      db.users.length !== loadedUsers.length ||
      legacyPayloadHasEmbeddedCollections;
    if (shouldPersist) {
      await persistStorage(db, env);
    }
    return rememberStorage(db);
  }
  if (storageCache) return storageCache;
  if (isLocalNodeRuntime) {
    const { fs, path } = await nodeModules();
    await fs.mkdir(path.dirname(DB_PATH), { recursive: true });
    try {
      const raw = await fs.readFile(DB_PATH, "utf8");
      rememberStorage(normalizeDbState(JSON.parse(raw)));
    } catch {
      const seeded = normalizeDbState(seedState());
      seeded.users[0].password_hash = await hashPassword("admin123");
      rememberStorage(seeded);
      await persistStorage(storageCache, env);
    }
  } else {
    const seeded = normalizeDbState(seedState());
    seeded.users[0].password_hash = await hashPassword("admin123");
    rememberStorage(seeded);
  }
  return storageCache;
}

async function persistStorage(db, env = {}, scope = {}) {
  const options = {
    users: true,
    dailyMetrics: true,
    qualityScores: true,
    meta: true,
    ...scope,
  };
  if (env?.DB) {
    await ensureD1Schema(env.DB);
    if (options.users) await persistUsersToD1(env.DB, db.users || []);
    if (options.dailyMetrics) await persistDailyMetricsToD1(env.DB, db.dailyMetrics || []);
    if (options.qualityScores) await persistQualityScoresToD1(env.DB, db.qualityScores || []);
    if (options.meta) await persistMetaStateToD1(env.DB, db);
    rememberStorage(db);
    return;
  }
  if (!isLocalNodeRuntime || !db) {
    rememberStorage(db);
    return;
  }
  rememberStorage(db);
  const { fs } = await nodeModules();
  await fs.writeFile(DB_PATH, JSON.stringify(db, null, 2), "utf8");
}

async function persistSingleUserChange(db, env, user, includeMeta = false) {
  if (env?.DB) {
    await ensureD1Schema(env.DB);
    await persistUserRecordToD1(env.DB, user);
    if (includeMeta) await persistMetaStateToD1(env.DB, db);
    rememberStorage(db);
    return;
  }
  await persistStorage(db, env, { users: true, dailyMetrics: false, qualityScores: false, meta: includeMeta });
}

async function persistSingleDailyMetricChange(db, env, metric, includeMeta = false) {
  if (env?.DB) {
    await ensureD1Schema(env.DB);
    await persistDailyMetricRecordToD1(env.DB, metric);
    if (includeMeta) await persistMetaStateToD1(env.DB, db);
    rememberStorage(db);
    return;
  }
  await persistStorage(db, env, { users: false, dailyMetrics: true, qualityScores: false, meta: includeMeta });
}

async function persistSingleQualityScoreChange(db, env, score, includeMeta = false) {
  if (env?.DB) {
    await ensureD1Schema(env.DB);
    await persistQualityScoreRecordToD1(env.DB, score);
    if (includeMeta) await persistMetaStateToD1(env.DB, db);
    rememberStorage(db);
    return;
  }
  await persistStorage(db, env, { users: false, dailyMetrics: false, qualityScores: true, meta: includeMeta });
}

async function persistMetaOnly(db, env) {
  if (env?.DB) {
    await ensureD1Schema(env.DB);
    await persistMetaStateToD1(env.DB, db);
    rememberStorage(db);
    return;
  }
  await persistStorage(db, env, { users: false, dailyMetrics: false, qualityScores: false, meta: true });
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
  if (hasFields(row, ["usuario_nuvidio", "nuvidio_aprovadas", "nuvidio_reprovadas", "nuvidio_sem_acao", "nuvidio_vazio", "data"])) return "nuvidio_summary";
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
      calls_nuvidio_empty: 0,
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

function extendImportedPeriod(period, metricDate) {
  if (!metricDate) return period;
  if (!period.start || metricDate < period.start) period.start = metricDate;
  if (!period.end || metricDate > period.end) period.end = metricDate;
  return period;
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
    "nuvidio_vazio",
  ];
  const first = rows[0] || {};
  const missing = required.filter((field) => !(field in first));
  if (missing.length) {
    throw new Error(`Colunas obrigatorias ausentes: ${missing.join(", ")}`);
  }
  let processed = 0;
  let rejected = 0;
  const errors = [];
  const period = { start: "", end: "" };
  rows.forEach((row, index) => {
    try {
      const user = matchUser(users, row);
      if (!user) throw new Error("Operador nao encontrado para os identificadores informados.");
      const metricDate = parseDate(row.data);
      upsertDailyMetric(
        db,
        user.id,
        metricDate,
        {
          production: toInt(row.producao),
          calls_0800_approved: toInt(row["0800_aprovado"]),
          calls_0800_rejected: toInt(row["0800_cancelada"]),
          calls_0800_pending: toInt(row["0800_pendenciada"]),
          calls_0800_no_action: toInt(row["0800_sem_acao"]),
          calls_nuvidio_approved: toInt(row["nuvidio_aprovada"]),
          calls_nuvidio_rejected: toInt(row["nuvidio_reprovada"]),
          calls_nuvidio_no_action: toInt(row["nuvidio_sem_acao"]),
          calls_nuvidio_empty: toInt(row["nuvidio_vazio"]),
        },
        sourceName,
      );
      extendImportedPeriod(period, metricDate);
      processed += 1;
    } catch (error) {
      rejected += 1;
      errors.push({ row: index + 2, error: error.message });
    }
  });
  return { processed, rejected, errors, period };
}

function import0800Rows(db, users, rows, sourceName) {
  const aggregates = new Map();
  const errors = [];
  const period = { start: "", end: "" };
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
      extendImportedPeriod(period, metricDate);
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

  return { processed: aggregates.size, rejected: errors.length, errors, period };
}

function importNuvidioRows(db, users, rows, sourceName) {
  const aggregates = new Map();
  const errors = [];
  const period = { start: "", end: "" };
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
      extendImportedPeriod(period, metricDate);
      const key = `${user.id}::${metricDate}`;
      const current = aggregates.get(key) || {
        userId: user.id,
        metricDate,
        production_nuvidio: 0,
        calls_nuvidio_approved: 0,
        calls_nuvidio_rejected: 0,
        calls_nuvidio_no_action: 0,
        calls_nuvidio_empty: 0,
      };
      current.production_nuvidio += 1;
      const tag = normalizeNuvidioTag(row.tag);
      if (tag === "aprovada" || tag === "aprovado") current.calls_nuvidio_approved += 1;
      else if (tag === "reprovada" || tag === "reprovado") current.calls_nuvidio_rejected += 1;
      else if (!tag || tag === "vazio" || tag === "vazia") current.calls_nuvidio_empty += 1;
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
        calls_nuvidio_empty: aggregate.calls_nuvidio_empty,
      },
      sourceName,
    );
  }

  return { processed: aggregates.size, rejected: errors.length, errors, period };
}

function importNuvidioSummaryRows(db, users, rows, sourceName) {
  let processed = 0;
  let rejected = 0;
  const errors = [];
  const period = { start: "", end: "" };

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
            toInt(row.nuvidio_sem_acao) +
            toInt(row.nuvidio_vazio),
          calls_nuvidio_approved: toInt(row.nuvidio_aprovadas),
          calls_nuvidio_rejected: toInt(row.nuvidio_reprovadas),
          calls_nuvidio_no_action: toInt(row.nuvidio_sem_acao),
          calls_nuvidio_empty: toInt(row.nuvidio_vazio),
        },
        sourceName,
      );
      extendImportedPeriod(period, metricDate);
      processed += 1;
    } catch (error) {
      rejected += 1;
      errors.push({ row: index + 2, error: error.message });
    }
  });

  return { processed, rejected, errors, period };
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
    : "Usuário Nuvidio;Nuvidio Aprovadas;Nuvidio Reprovadas;Nuvidio Sem ação;Nuvidio Vazio;Data";
  const rows = users.map((user) => {
    if (is0800) {
      const id0800 = String(user.platform_0800_id || user.login || "").replaceAll('"', '""');
      return `"${id0800}";0;0;0;0;2026-01-01`;
    }
    const nuvidio = String(user.nuvidio_id || user.login || "").replaceAll('"', '""');
    return `"${nuvidio}";0;0;0;0;2026-01-01`;
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
  else if (normalizedTag.includes("vazi")) values.calls_nuvidio_empty = qty;
  else values.calls_nuvidio_no_action = qty;
  return values;
}

function import0800SummaryRows(db, users, rows, sourceName) {
  let processed = 0;
  let rejected = 0;
  const errors = [];
  const period = { start: "", end: "" };

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
      extendImportedPeriod(period, metricDate);
      processed += 1;
    } catch (error) {
      rejected += 1;
      errors.push({ row: index + 2, error: error.message });
    }
  });

  return { processed, rejected, errors, period };
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
    production: todayRows.reduce((sum, row) => sum + toInt(row.production_0800) + toInt(row.production_nuvidio), 0),
    effectiveness: averageIgnoreZero(todayRows.flatMap((row) => [
      calculateOperationEffectiveness(row, "0800"),
      calculateOperationEffectiveness(row, "nuvidio"),
    ])),
    quality: average(monthRows.map((item) => toFloat(item.score))),
  };
  const trendMap = new Map();
  for (const row of trendRows) {
    const current = trendMap.get(row.metric_date) || { date: row.metric_date, production: 0, effectivenessParts: [] };
    current.production += toInt(row.production_0800) + toInt(row.production_nuvidio);
    current.effectivenessParts.push(calculateOperationEffectiveness(row, "0800"));
    current.effectivenessParts.push(calculateOperationEffectiveness(row, "nuvidio"));
    trendMap.set(row.metric_date, current);
  }
  const trend = [...trendMap.values()].map((item) => ({
    date: item.date,
    production: item.production,
    effectiveness: averageIgnoreZero(item.effectivenessParts),
  }));
  const operators = todayRows.map((row) => ({
    name: repairTextEncoding(db.users.find((entry) => entry.id === row.user_id)?.full_name || "Operador"),
    production: toInt(row.production_0800) + toInt(row.production_nuvidio),
    effectiveness: averageIgnoreZero([
      calculateOperationEffectiveness(row, "0800"),
      calculateOperationEffectiveness(row, "nuvidio"),
    ]),
  }));
  return { cards, trend, operators };
}

function buildDayTop(db, user, url) {
  const date = String(url.searchParams.get("date") || todayIso()).trim();
  const scoped = db.dailyMetrics.filter((row) => row.metric_date === date && (user.role === "manager" || row.user_id === user.id));
  const top = scoped
    .map((row) => {
      const production0800 = toInt(row.production_0800);
      const productionNuvidio = toInt(row.production_nuvidio);
      const production = production0800 + productionNuvidio;
      const eff0800 = calculateOperationEffectiveness(row, "0800");
      const effNuvidio = calculateOperationEffectiveness(row, "nuvidio");
      const effectiveness = averageIgnoreZero([eff0800, effNuvidio]);
      return {
        user_id: row.user_id,
        name: repairTextEncoding(db.users.find((entry) => entry.id === row.user_id)?.full_name || "Operador"),
        production,
        effectiveness,
      };
    })
    .sort((a, b) => b.production - a.production || b.effectiveness - a.effectiveness || a.name.localeCompare(b.name))
    .slice(0, 10);
  return { date, top };
}

function buildMetricTop(db, user, url) {
  const metric = String(url.searchParams.get("metric") || "production").trim().toLowerCase();
  const operation = String(url.searchParams.get("operation") || "all").trim().toLowerCase();
  const date = String(url.searchParams.get("date") || todayIso()).trim();
  const referenceMonth = String(url.searchParams.get("reference_month") || monthRef(date)).trim();
  const qualityField = String(url.searchParams.get("quality_field") || "final").trim().toLowerCase();

  if (metric === "quality") {
    const scopedQuality = db.qualityScores.filter((row) => row.reference_month === referenceMonth && (user.role === "manager" || row.user_id === user.id));
    const resolveQualityValue = (row) => {
      if (qualityField === "m1") return toFloat(row.monitoria_1);
      if (qualityField === "m2") return toFloat(row.monitoria_2);
      if (qualityField === "m3") return toFloat(row.monitoria_3);
      if (qualityField === "m4") return toFloat(row.monitoria_4);
      return toFloat(row.score);
    };
    const top = scopedQuality
      .map((row) => ({
        user_id: row.user_id,
        name: repairTextEncoding(db.users.find((entry) => entry.id === row.user_id)?.full_name || "Operador"),
        value: resolveQualityValue(row),
      }))
      .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name))
      .slice(0, 10);
    return { metric, operation: "all", date: null, reference_month: referenceMonth, top };
  }

  const scopedMetrics = db.dailyMetrics.filter((row) => row.metric_date === date && (user.role === "manager" || row.user_id === user.id));
  const top = scopedMetrics
    .map((row) => {
      let value = 0;
      if (metric === "production") {
        if (operation === "0800") value = toInt(row.production_0800);
        else if (operation === "nuvidio") value = toInt(row.production_nuvidio);
        else value = toInt(row.production_0800) + toInt(row.production_nuvidio);
      } else {
        if (operation === "0800") value = calculateOperationEffectiveness(row, "0800");
        else if (operation === "nuvidio") value = calculateOperationEffectiveness(row, "nuvidio");
        else value = averageIgnoreZero([
          calculateOperationEffectiveness(row, "0800"),
          calculateOperationEffectiveness(row, "nuvidio"),
        ]);
      }
      return {
        user_id: row.user_id,
        name: repairTextEncoding(db.users.find((entry) => entry.id === row.user_id)?.full_name || "Operador"),
        value,
      };
    })
    .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name))
    .slice(0, 10);
  return { metric, operation, date, reference_month: null, top };
}

function buildAnalysis(db, user, url) {
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const users = db.users.filter((entry) => entry.is_active && (user.role === "manager" || entry.id === user.id));
  const ranking = users.map((entry) => {
    const metrics = db.dailyMetrics.filter((row) => row.user_id === entry.id && row.metric_date >= start && row.metric_date <= end);
    const productionMetrics = metrics.filter((row) => !isSaturdayIsoDate(row.metric_date));
    const quality = db.qualityScores.find((score) => score.user_id === entry.id && score.reference_month === monthRef(end));
    return {
      user_id: entry.id,
      name: repairTextEncoding(entry.full_name),
      avg_production: productionMetrics.length ? Number((productionMetrics.reduce((sum, row) => sum + toInt(row.production), 0) / productionMetrics.length).toFixed(2)) : 0,
      total_production: metrics.reduce((sum, row) => sum + toInt(row.production), 0),
      effectiveness: average(metrics.map(calculateEffectiveness)),
      quality: toFloat(quality?.score),
      active_days: metrics.length,
    };
  }).sort((a, b) => b.total_production - a.total_production || a.name.localeCompare(b.name));
  return { ranking, period: { start, end } };
}

function buildBootstrap(db, user, url) {
  return {
    app_settings: serializeAppSettings(db),
    users: user.role === "manager" ? db.users.map(serializeUser) : [],
    overview: buildOverview(db, user, url),
    analysis: buildAnalysis(db, user, url),
    history: buildHistory(db, user, url),
  };
}

function buildAlerts(db, user, url) {
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const requestedRawUserId = String(url.searchParams.get("user_id") || "all").trim().toLowerCase();
  const includeAllUsers = requestedRawUserId === "all";
  const requestedUserId = Number(requestedRawUserId || 0);
  const users = db.users.filter((entry) => entry.is_active && entry.role === "operator")
    .filter((entry) => includeAllUsers || entry.id === requestedUserId);
  const metricRules = serializeAppSettings(db).metric_rules;
  const qualityByUser = new Map();
  for (const score of (db.qualityScores || []).slice().sort((a, b) => b.reference_month.localeCompare(a.reference_month) || b.updated_at.localeCompare(a.updated_at))) {
    if (!qualityByUser.has(score.user_id)) qualityByUser.set(score.user_id, score);
  }
  const alerts = [];
  const scoredOperators = [];

  function clampScore(value) {
    return Math.max(0, Math.min(10, Number(value || 0)));
  }

  function toneFromScore(score) {
    if (score >= 7) return "red";
    if (score >= 4) return "amber";
    return "green";
  }

  function labelFromScore(score) {
    if (score >= 8.5) return "Crítico";
    if (score >= 7) return "Alto";
    if (score >= 4) return "Atenção";
    return "Baixo";
  }

  for (const entry of users) {
    const metrics = db.dailyMetrics.filter((row) => row.user_id === entry.id && row.metric_date >= start && row.metric_date <= end);
    const productionMetrics = metrics.filter((row) => !isSaturdayIsoDate(row.metric_date));
    const avgProduction = averageIgnoreZero(productionMetrics.map((row) => toInt(row.production_0800) + toInt(row.production_nuvidio)));
    const effectiveness = averageIgnoreZero(metrics.flatMap((row) => [
      calculateOperationEffectiveness(row, "0800"),
      calculateOperationEffectiveness(row, "nuvidio"),
    ]));
    const totalNoAction = metrics.reduce((sum, row) => sum + toInt(row.calls_0800_no_action) + toInt(row.calls_nuvidio_no_action) + toInt(row.calls_nuvidio_empty), 0);
    const totalActionable = metrics.reduce((sum, row) => (
      sum
      + toInt(row.calls_0800_approved)
      + toInt(row.calls_0800_rejected)
      + toInt(row.calls_0800_pending)
      + toInt(row.calls_nuvidio_approved)
      + toInt(row.calls_nuvidio_rejected)
    ), 0);
    const totalContacts = totalActionable + totalNoAction;
    const noActionShare = totalContacts ? Number(((totalNoAction / totalContacts) * 100).toFixed(2)) : 0;
    const latestQuality = qualityByUser.get(entry.id) || null;
    const qualityScore = latestQuality ? toFloat(latestQuality.score) : 0;
    const reasons = [];
    const scoreParts = [];

    if (totalContacts > 0) {
      let noActionRisk = 0;
      if (noActionShare >= 25) noActionRisk = 10;
      else if (noActionShare >= 20) noActionRisk = 8.5;
      else if (noActionShare >= 15) noActionRisk = 6.5;
      else if (noActionShare >= 10) noActionRisk = 4.5;
      else if (noActionShare >= 5) noActionRisk = 2.5;

      if (noActionRisk > 0 && avgProduction >= metricRules.production.amber_max) noActionRisk = clampScore(noActionRisk + 1);
      else if (noActionRisk > 0 && avgProduction >= metricRules.production.red_max) noActionRisk = clampScore(noActionRisk + 0.5);

      scoreParts.push(noActionRisk);
      if (noActionRisk >= 7) {
        reasons.push({
          tone: "red",
          text: `Sem ação / vazio muito alto para o volume lançado (${Math.round(noActionShare)}%).`,
        });
      } else if (noActionRisk >= 4) {
        reasons.push({
          tone: "amber",
          text: `Sem ação / vazio acima do ideal (${Math.round(noActionShare)}%).`,
        });
      }

      let effectivenessRisk = 0;
      if (effectiveness <= metricRules.effectiveness.red_max) effectivenessRisk = 10;
      else if (effectiveness <= metricRules.effectiveness.amber_max) effectivenessRisk = 6.5;
      else if (effectiveness < 95) effectivenessRisk = 3;

      scoreParts.push(effectivenessRisk);
      if (effectivenessRisk >= 7) {
        reasons.push({
          tone: "red",
          text: `Efetividade crítica (${Math.round(effectiveness)}%).`,
        });
      } else if (effectivenessRisk >= 4) {
        reasons.push({
          tone: "amber",
          text: `Efetividade abaixo do ideal (${Math.round(effectiveness)}%).`,
        });
      }
    }

    if (latestQuality) {
      let qualityRisk = 0;
      if (qualityScore <= metricRules.quality.red_max) qualityRisk = 10;
      else if (qualityScore <= metricRules.quality.amber_max) qualityRisk = 6.5;
      else if (qualityScore < 95) qualityRisk = 3;

      scoreParts.push(qualityRisk);
      if (qualityRisk >= 7) {
        reasons.push({
          tone: "red",
          text: `Qualidade crítica (${Math.round(qualityScore)}).`,
        });
      } else if (qualityRisk >= 4) {
        reasons.push({
          tone: "amber",
          text: `Qualidade em atenção (${Math.round(qualityScore)}).`,
        });
      }
    }

    const availableScores = scoreParts.filter((value) => Number.isFinite(value));
    if (!availableScores.length) continue;
    const alertScore = Number((availableScores.reduce((sum, value) => sum + value, 0) / availableScores.length).toFixed(1));
    scoredOperators.push(alertScore);
    if (alertScore <= 0) continue;

    alerts.push({
      user_id: entry.id,
      name: repairTextEncoding(entry.full_name),
      login: entry.login,
      alert_score: alertScore,
      alert_tone: toneFromScore(alertScore),
      alert_label: labelFromScore(alertScore),
      avg_production: Number(avgProduction.toFixed(2)),
      effectiveness: Number(effectiveness.toFixed(2)),
      quality: qualityScore,
      no_action_share: noActionShare,
      total_contacts: totalContacts,
      active_days: metrics.length,
      latest_quality_month: latestQuality?.reference_month || "",
      reasons,
    });
  }

  alerts.sort((a, b) =>
    b.alert_score - a.alert_score
    || b.no_action_share - a.no_action_share
    || a.quality - b.quality
    || b.avg_production - a.avg_production
    || a.name.localeCompare(b.name)
  );

  return {
    period: { start, end },
    summary: {
      monitored: scoredOperators.length,
      total: alerts.length,
      average_score: scoredOperators.length
        ? Number((scoredOperators.reduce((sum, value) => sum + value, 0) / scoredOperators.length).toFixed(1))
        : 0,
      max_score: alerts.length ? alerts[0].alert_score : 0,
    },
    alerts,
  };
}

function buildHistory(db, user, url) {
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const startMonth = String(start).slice(0, 7);
  const endMonth = String(end).slice(0, 7);
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
      if (String(item.reference_month || "") < startMonth || String(item.reference_month || "") > endMonth) return false;
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

function averageIgnoreZero(values) {
  const filtered = values.map((value) => Number(value || 0)).filter((value) => Number.isFinite(value) && value > 0);
  if (!filtered.length) return 0;
  return Number((filtered.reduce((sum, value) => sum + value, 0) / filtered.length).toFixed(2));
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
      await persistSingleUserChange(db, env, user, false);
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
    await persistSingleUserChange(db, env, auth.user, false);
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
    await persistSingleUserChange(db, env, auth.user, false);
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
    if (payload.metric_rules && typeof payload.metric_rules === "object") {
      const toNumber = (value, fallbackValue) => {
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : fallbackValue;
      };
      const current = serializeAppSettings(db).metric_rules;
      db.appSettings.metric_rules = {
        production: {
          red_max: toNumber(payload.metric_rules?.production?.red_max, current.production.red_max),
          amber_max: toNumber(payload.metric_rules?.production?.amber_max, current.production.amber_max),
        },
        effectiveness: {
          red_max: toNumber(payload.metric_rules?.effectiveness?.red_max, current.effectiveness.red_max),
          amber_max: toNumber(payload.metric_rules?.effectiveness?.amber_max, current.effectiveness.amber_max),
        },
        quality: {
          red_max: toNumber(payload.metric_rules?.quality?.red_max, current.quality.red_max),
          amber_max: toNumber(payload.metric_rules?.quality?.amber_max, current.quality.amber_max),
        },
      };
    }
    await persistMetaOnly(db, env);
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

  if (url.pathname === "/api/dashboard/day-top" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    if (isOperatorBlockedByMaintenance(auth.user, db)) {
      return jsonResponse({ error: "Portal em manutenção para operadores." }, 503);
    }
    return jsonResponse(buildDayTop(db, auth.user, url));
  }

  if (url.pathname === "/api/dashboard/top-metric" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    if (isOperatorBlockedByMaintenance(auth.user, db)) {
      return jsonResponse({ error: "Portal em manutenção para operadores." }, 503);
    }
    return jsonResponse(buildMetricTop(db, auth.user, url));
  }

  if (url.pathname === "/api/analysis" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    if (isOperatorBlockedByMaintenance(auth.user, db)) {
      return jsonResponse({ error: "Portal em manutenção para operadores." }, 503);
    }
    return jsonResponse(buildAnalysis(db, auth.user, url));
  }

  if (url.pathname === "/api/alerts" && request.method === "GET") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse(buildAlerts(db, auth.user, url));
  }

  if (url.pathname === "/api/bootstrap" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    if (isOperatorBlockedByMaintenance(auth.user, db)) {
      return jsonResponse({ error: "Portal em manutenção para operadores." }, 503);
    }
    return jsonResponse(buildBootstrap(db, auth.user, url));
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
    await persistSingleUserChange(db, env, user, true);
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
    await persistSingleUserChange(db, env, user, false);
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
    if (env?.DB) {
      await ensureD1Schema(env.DB);
      await deleteUserDataFromD1(env.DB, userId);
      await persistMetaStateToD1(env.DB, db);
      rememberStorage(db);
    } else {
      await persistStorage(db, env);
    }
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
    const monitoria1 = toOptionalMonitoria(payload.monitoria_1);
    const monitoria2 = toOptionalMonitoria(payload.monitoria_2);
    const monitoria3 = toOptionalMonitoria(payload.monitoria_3);
    const monitoria4 = toOptionalMonitoria(payload.monitoria_4);
    const monitorias = [monitoria1, monitoria2, monitoria3, monitoria4]
      .filter((value) => value !== null && Number.isFinite(value) && value >= 0 && value <= 100);
    const resolvedScore = monitorias.length
      ? Number((monitorias.reduce((sum, value) => sum + value, 0) / monitorias.length).toFixed(2))
      : toFloat(payload.score);

    let score = db.qualityScores.find((item) => item.user_id === userId && item.reference_month === payload.reference_month);
    if (!score) {
      score = {
        id: nextId(db, "qualityScores"),
        user_id: userId,
        reference_month: payload.reference_month,
        score: resolvedScore,
        monitoria_1: monitoria1,
        monitoria_2: monitoria2,
        monitoria_3: monitoria3,
        monitoria_4: monitoria4,
        notes: String(payload.notes || "").trim(),
        created_at: nowIso(),
        updated_at: nowIso(),
      };
      db.qualityScores.push(score);
    } else {
      score.score = resolvedScore;
      score.monitoria_1 = monitoria1;
      score.monitoria_2 = monitoria2;
      score.monitoria_3 = monitoria3;
      score.monitoria_4 = monitoria4;
      score.notes = String(payload.notes || "").trim();
      score.updated_at = nowIso();
    }
    await persistSingleQualityScoreChange(db, env, score, true);
    return jsonResponse({ quality: score });
  }

  if (url.pathname.startsWith("/api/admin/quality/") && request.method === "PUT") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const qualityId = Number(url.pathname.split("/").pop());
    const score = db.qualityScores.find((entry) => entry.id === qualityId);
    if (!score) return jsonResponse({ error: "Registro de qualidade nao encontrado" }, 404);
    const payload = await request.json();
    const monitoria1 = payload.monitoria_1 === "" ? null : toOptionalMonitoria(payload.monitoria_1 ?? score.monitoria_1);
    const monitoria2 = payload.monitoria_2 === "" ? null : toOptionalMonitoria(payload.monitoria_2 ?? score.monitoria_2);
    const monitoria3 = payload.monitoria_3 === "" ? null : toOptionalMonitoria(payload.monitoria_3 ?? score.monitoria_3);
    const monitoria4 = payload.monitoria_4 === "" ? null : toOptionalMonitoria(payload.monitoria_4 ?? score.monitoria_4);
    const monitorias = [monitoria1, monitoria2, monitoria3, monitoria4]
      .filter((value) => value !== null && Number.isFinite(value) && value >= 0 && value <= 100);
    const resolvedScore = monitorias.length
      ? Number((monitorias.reduce((sum, value) => sum + Number(value || 0), 0) / monitorias.length).toFixed(2))
      : 0;
    score.monitoria_1 = monitoria1;
    score.monitoria_2 = monitoria2;
    score.monitoria_3 = monitoria3;
    score.monitoria_4 = monitoria4;
    score.score = resolvedScore;
    score.notes = String(payload.notes ?? score.notes ?? "").trim();
    score.updated_at = nowIso();
    await persistSingleQualityScoreChange(db, env, score, true);
    return jsonResponse({ quality: score });
  }

  if (url.pathname.startsWith("/api/admin/quality/") && request.method === "DELETE") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const qualityId = Number(url.pathname.split("/").pop());
    const index = db.qualityScores.findIndex((entry) => entry.id === qualityId);
    if (index === -1) return jsonResponse({ error: "Registro de qualidade nao encontrado" }, 404);
    db.qualityScores.splice(index, 1);
    if (env?.DB) {
      await ensureD1Schema(env.DB);
      await env.DB.prepare("DELETE FROM quality_scores WHERE id = ?").bind(qualityId).run();
      rememberStorage(db);
    } else {
      await persistStorage(db, env, { users: false, dailyMetrics: false, qualityScores: true, meta: false });
    }
    return jsonResponse({ ok: true });
  }

  if (url.pathname === "/api/admin/quality/template" && request.method === "GET") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    if (!isLocalNodeRuntime) {
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

    if (!isLocalNodeRuntime && lowerName.endsWith(".xlsx")) {
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

    await persistStorage(db, env, { users: false, dailyMetrics: false, qualityScores: true, meta: true });
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
    const empty = Math.max(0, toInt(payload.empty));
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
          production_nuvidio: approved + rejected + noAction + empty,
          calls_nuvidio_approved: approved,
          calls_nuvidio_rejected: rejected,
          calls_nuvidio_no_action: noAction,
          calls_nuvidio_empty: empty,
        };
    upsertDailyMetric(db, user.id, metricDate, values, "manual_tag");
    const metric = db.dailyMetrics.find((entry) => entry.user_id === user.id && entry.metric_date === metricDate);
    await persistSingleDailyMetricChange(db, env, metric, true);
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
    metric.calls_nuvidio_empty = toInt(payload.calls_nuvidio_empty ?? metric.calls_nuvidio_empty);
    metric.production_0800 = toInt(metric.calls_0800_approved) + toInt(metric.calls_0800_rejected) + toInt(metric.calls_0800_pending) + toInt(metric.calls_0800_no_action);
    metric.production_nuvidio = toInt(metric.calls_nuvidio_approved) + toInt(metric.calls_nuvidio_rejected) + toInt(metric.calls_nuvidio_no_action) + toInt(metric.calls_nuvidio_empty);
    metric.production = toInt(metric.production_0800) + toInt(metric.production_nuvidio);
    metric.updated_at = nowIso();
    await persistSingleDailyMetricChange(db, env, metric, false);
    return jsonResponse({ metric: { ...metric, effectiveness: calculateEffectiveness(metric) } });
  }

  if (url.pathname.startsWith("/api/admin/daily-metrics/") && request.method === "DELETE") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const metricId = Number(url.pathname.split("/").pop());
    const index = db.dailyMetrics.findIndex((entry) => entry.id === metricId);
    if (index === -1) return jsonResponse({ error: "Registro nao encontrado" }, 404);
    db.dailyMetrics.splice(index, 1);
    if (env?.DB) {
      await ensureD1Schema(env.DB);
      await deleteDailyMetricRecordFromD1(env.DB, metricId);
      rememberStorage(db);
    } else {
      await persistStorage(db, env, { users: false, dailyMetrics: true, qualityScores: false, meta: false });
    }
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

    await persistStorage(db, env, { users: false, dailyMetrics: true, qualityScores: false, meta: true });
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
    const { processed, rejected, errors, period } = outcome;
    registerImportLog(db, auth.user.id, file.name, processed, rejected);
    await persistStorage(db, env, { users: false, dailyMetrics: true, qualityScores: false, meta: true });
    return jsonResponse({ processed, rejected, period, errors: errors.slice(0, 20) });
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




