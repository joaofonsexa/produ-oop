const CORS_HEADERS = {
  "access-control-allow-origin": "*",
  "access-control-allow-methods": "GET,POST,OPTIONS",
  "access-control-allow-headers": "content-type"
};

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...CORS_HEADERS,
      "content-type": "application/json; charset=utf-8"
    }
  });
}

function decodeBase64ToBytes(value) {
  const binary = atob(String(value || ""));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function safeFileName(name) {
  return String(name || "import.xlsx").replace(/[^a-zA-Z0-9._-]/g, "_");
}

function getAuthBase(env) {
  return String(env.AUTH_API_BASE || env.CENTRAL_AUTH_BASE || "").trim().replace(/\/+$/, "");
}

function getSsoConfig(env, requestUrl) {
  const url = new URL(requestUrl);
  const expectedAudience = String(env.SSO_EXPECTED_AUDIENCE || url.origin).trim();
  const expectedIssuer = String(env.SSO_EXPECTED_ISSUER || getAuthBase(env)).trim();
  const sharedSecret = String(env.SSO_SHARED_SECRET || "").trim();
  return { expectedAudience, expectedIssuer, sharedSecret };
}

function base64UrlToBytes(value) {
  const normalized = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
  const padding = normalized.length % 4 === 0 ? "" : "=".repeat(4 - (normalized.length % 4));
  const binary = atob(normalized + padding);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function base64UrlDecodeJson(value) {
  const bytes = base64UrlToBytes(value);
  const text = new TextDecoder().decode(bytes);
  return JSON.parse(text);
}

function normalizeLookupKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

async function importHmacKey(sharedSecret) {
  const keyData = new TextEncoder().encode(sharedSecret);
  return crypto.subtle.importKey("raw", keyData, { name: "HMAC", hash: "SHA-256" }, false, ["verify"]);
}

async function verifyJwtHs256(token, sharedSecret) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3) {
    throw new Error("Formato de token invalido.");
  }

  const [headerB64, payloadB64, signatureB64] = parts;
  const header = base64UrlDecodeJson(headerB64);
  const payload = base64UrlDecodeJson(payloadB64);

  if (String(header?.alg || "") !== "HS256") {
    throw new Error("Algoritmo do token nao suportado.");
  }

  const key = await importHmacKey(sharedSecret);
  const data = new TextEncoder().encode(`${headerB64}.${payloadB64}`);
  const signature = base64UrlToBytes(signatureB64);
  const valid = await crypto.subtle.verify("HMAC", key, signature, data);
  if (!valid) {
    throw new Error("Assinatura do token invalida.");
  }

  return payload;
}

function normalizeDateKey(value) {
  const raw = String(value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  return "";
}

async function ensureResultsTable(db) {
  await db.prepare(
    `CREATE TABLE IF NOT EXISTS operator_results_daily (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      user_name TEXT NOT NULL DEFAULT '',
      username TEXT NOT NULL DEFAULT '',
      username_0800 TEXT NOT NULL DEFAULT '',
      username_nuvidio TEXT NOT NULL DEFAULT '',
      result_date TEXT NOT NULL,
      funnel_0800_approved REAL NOT NULL DEFAULT 0,
      funnel_0800_cancelled REAL NOT NULL DEFAULT 0,
      funnel_0800_pending REAL NOT NULL DEFAULT 0,
      funnel_0800_no_action REAL NOT NULL DEFAULT 0,
      funnel_nuvidio_approved REAL NOT NULL DEFAULT 0,
      funnel_nuvidio_reproved REAL NOT NULL DEFAULT 0,
      funnel_nuvidio_no_action REAL NOT NULL DEFAULT 0,
      production_0800 REAL NOT NULL DEFAULT 0,
      production_nuvidio REAL NOT NULL DEFAULT 0,
      production_total REAL NOT NULL DEFAULT 0,
      effectiveness_0800 REAL NOT NULL DEFAULT 0,
      effectiveness_nuvidio REAL NOT NULL DEFAULT 0,
      effectiveness REAL NOT NULL DEFAULT 0,
      quality_score REAL NOT NULL DEFAULT 0,
      updated_by_id TEXT NOT NULL DEFAULT '',
      updated_by_name TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(user_id, result_date)
    )`
  ).run();

  const migrations = [
    "ALTER TABLE operator_results_daily ADD COLUMN username_0800 TEXT NOT NULL DEFAULT ''",
    "ALTER TABLE operator_results_daily ADD COLUMN username_nuvidio TEXT NOT NULL DEFAULT ''",
    "ALTER TABLE operator_results_daily ADD COLUMN funnel_0800_approved REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN funnel_0800_cancelled REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN funnel_0800_pending REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN funnel_0800_no_action REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN funnel_nuvidio_approved REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN funnel_nuvidio_reproved REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN funnel_nuvidio_no_action REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN production_0800 REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN production_nuvidio REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN effectiveness_0800 REAL NOT NULL DEFAULT 0",
    "ALTER TABLE operator_results_daily ADD COLUMN effectiveness_nuvidio REAL NOT NULL DEFAULT 0"
  ];

  for (const sql of migrations) {
    try {
      await db.prepare(sql).run();
    } catch (error) {
      const message = String(error?.message || "");
      if (!message.includes("duplicate column name")) throw error;
    }
  }
}

async function ensureSsoReplayTable(db) {
  await db.prepare(
    `CREATE TABLE IF NOT EXISTS sso_consumed_tokens (
      jti TEXT PRIMARY KEY,
      expires_at INTEGER NOT NULL,
      consumed_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`
  ).run();
}

async function ensureSystemSettingsTable(db) {
  await db.prepare(
    `CREATE TABLE IF NOT EXISTS system_settings (
      setting_key TEXT PRIMARY KEY,
      value_text TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_by_id TEXT NOT NULL DEFAULT '',
      updated_by_name TEXT NOT NULL DEFAULT ''
    )`
  ).run();
}

async function ensureSheetSyncTables(db) {
  await db.prepare(
    `CREATE TABLE IF NOT EXISTS sheet_sync_daily (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      user_name TEXT NOT NULL DEFAULT '',
      username TEXT NOT NULL DEFAULT '',
      username_0800 TEXT NOT NULL DEFAULT '',
      username_nuvidio TEXT NOT NULL DEFAULT '',
      result_date TEXT NOT NULL,
      funnel_0800_approved REAL NOT NULL DEFAULT 0,
      funnel_0800_cancelled REAL NOT NULL DEFAULT 0,
      funnel_0800_pending REAL NOT NULL DEFAULT 0,
      funnel_0800_no_action REAL NOT NULL DEFAULT 0,
      funnel_nuvidio_approved REAL NOT NULL DEFAULT 0,
      funnel_nuvidio_reproved REAL NOT NULL DEFAULT 0,
      funnel_nuvidio_no_action REAL NOT NULL DEFAULT 0,
      production_0800 REAL NOT NULL DEFAULT 0,
      production_nuvidio REAL NOT NULL DEFAULT 0,
      production_total REAL NOT NULL DEFAULT 0,
      effectiveness_0800 REAL NOT NULL DEFAULT 0,
      effectiveness_nuvidio REAL NOT NULL DEFAULT 0,
      effectiveness REAL NOT NULL DEFAULT 0,
      quality_score REAL NOT NULL DEFAULT 0,
      source_origin TEXT NOT NULL DEFAULT '',
      updated_by_id TEXT NOT NULL DEFAULT '',
      updated_by_name TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(user_id, result_date)
    )`
  ).run();

  await db.prepare(
    `CREATE TABLE IF NOT EXISTS sheet_sync_events (
      id TEXT PRIMARY KEY,
      source_origin TEXT NOT NULL DEFAULT '',
      platform TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT '',
      user_id TEXT NOT NULL DEFAULT '',
      user_name TEXT NOT NULL DEFAULT '',
      username TEXT NOT NULL DEFAULT '',
      username_0800 TEXT NOT NULL DEFAULT '',
      username_nuvidio TEXT NOT NULL DEFAULT '',
      event_date TEXT NOT NULL DEFAULT '',
      quantity REAL NOT NULL DEFAULT 0,
      provider TEXT NOT NULL DEFAULT '',
      department TEXT NOT NULL DEFAULT '',
      subtag TEXT NOT NULL DEFAULT '',
      tma_seconds REAL NOT NULL DEFAULT 0,
      wait_seconds REAL NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`
  ).run();
}

async function ensureOperatorAccessLinksTable(db) {
  await db.prepare(
    `CREATE TABLE IF NOT EXISTS operator_access_links (
      user_id TEXT PRIMARY KEY,
      user_name TEXT NOT NULL DEFAULT '',
      username TEXT NOT NULL DEFAULT '',
      username_0800 TEXT NOT NULL DEFAULT '',
      username_nuvidio TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_by_id TEXT NOT NULL DEFAULT '',
      updated_by_name TEXT NOT NULL DEFAULT ''
    )`
  ).run();
}

async function ensureLocalUsersTable(db) {
  await db.prepare(
    `CREATE TABLE IF NOT EXISTS local_portal_users (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      username TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT 'operador',
      access_level TEXT NOT NULL DEFAULT '',
      username_0800 TEXT NOT NULL DEFAULT '',
      username_nuvidio TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_by_id TEXT NOT NULL DEFAULT '',
      updated_by_name TEXT NOT NULL DEFAULT ''
    )`
  ).run();
}

async function sha256Hex(value) {
  const bytes = new TextEncoder().encode(String(value || ""));
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(digest)].map((item) => item.toString(16).padStart(2, "0")).join("");
}

function sanitizeLocalUser(row) {
  return {
    id: String(row?.id || ""),
    name: String(row?.name || ""),
    username: String(row?.username || ""),
    username0800: String(row?.username_0800 || ""),
    usernameNuvidio: String(row?.username_nuvidio || ""),
    role: String(row?.role || "operador"),
    accessLevel: String(row?.access_level || "")
  };
}

function normalizeSpreadsheetSourceSettings(settings) {
  return {
    nuvidioUrl: String(settings?.nuvidioUrl || "").trim(),
    url0800General: String(settings?.url0800General || "").trim(),
    url0800Approved: String(settings?.url0800Approved || "").trim(),
    url0800Cancelled: String(settings?.url0800Cancelled || "").trim(),
    url0800Pending: String(settings?.url0800Pending || "").trim(),
    url0800NoAction: String(settings?.url0800NoAction || "").trim(),
    autoSyncOnOpen: Boolean(settings?.autoSyncOnOpen),
    lastSyncedAt: String(settings?.lastSyncedAt || "").trim(),
    lastSyncSummary: settings?.lastSyncSummary && typeof settings.lastSyncSummary === "object"
      ? settings.lastSyncSummary
      : null,
    updatedById: String(settings?.updatedById || "").trim(),
    updatedByName: String(settings?.updatedByName || "").trim()
  };
}

async function readJsonSystemSetting(db, key, fallback) {
  await ensureSystemSettingsTable(db);
  const row = await db
    .prepare("SELECT value_text FROM system_settings WHERE setting_key = ? LIMIT 1")
    .bind(key)
    .first();
  if (!row?.value_text) return fallback;
  try {
    return JSON.parse(String(row.value_text || ""));
  } catch {
    return fallback;
  }
}

async function writeJsonSystemSetting(db, key, value, actor = {}) {
  await ensureSystemSettingsTable(db);
  const updatedAt = new Date().toISOString();
  await db.prepare(
    `INSERT INTO system_settings (setting_key, value_text, updated_at, updated_by_id, updated_by_name)
     VALUES (?, ?, ?, ?, ?)
     ON CONFLICT(setting_key) DO UPDATE SET
       value_text = excluded.value_text,
       updated_at = excluded.updated_at,
       updated_by_id = excluded.updated_by_id,
       updated_by_name = excluded.updated_by_name`
  ).bind(
    key,
    JSON.stringify(value || {}),
    updatedAt,
    String(actor?.updatedById || actor?.actorId || "").trim(),
    String(actor?.updatedByName || actor?.actorName || "").trim()
  ).run();
}

async function readSpreadsheetSourceSettings(db) {
  const parsed = await readJsonSystemSetting(db, "spreadsheet_sources", {});
  return normalizeSpreadsheetSourceSettings(parsed);
}

async function saveSpreadsheetSourceSettings(db, payload = {}) {
  const current = await readSpreadsheetSourceSettings(db);
  const next = normalizeSpreadsheetSourceSettings({
    ...current,
    ...payload,
    updatedById: String(payload?.updatedById || current.updatedById || "").trim(),
    updatedByName: String(payload?.updatedByName || current.updatedByName || "").trim()
  });
  await writeJsonSystemSetting(db, "spreadsheet_sources", next, {
    updatedById: next.updatedById,
    updatedByName: next.updatedByName
  });
  return next;
}

function normalizeStatusKey(value) {
  const text = normalizeLookupKey(value);
  if (!text) return "";
  if (text.includes("aprov")) return "approved";
  if (text.includes("reprov")) return "reproved";
  if (text.includes("cancel")) return "cancelled";
  if (text.includes("pend")) return "pending";
  if (text.includes("sem acao") || text.includes("semacao") || text.includes("sem_acao")) return "noAction";
  return "";
}

function toTitleCaseLabel(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  return text
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function parseNumberLoose(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  let text = String(value || "").trim();
  if (!text) return 0;
  text = text.replace(/\s+/g, "");
  const hasComma = text.includes(",");
  const hasDot = text.includes(".");
  if (hasComma && hasDot) {
    text = text.replace(/\./g, "").replace(",", ".");
  } else if (hasComma) {
    text = text.replace(",", ".");
  }
  const numeric = Number(text.replace(/[^0-9.-]/g, ""));
  return Number.isFinite(numeric) ? numeric : 0;
}

function normalizeDateValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const ddmmyyyy = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (ddmmyyyy) {
    return `${ddmmyyyy[3]}-${String(ddmmyyyy[2]).padStart(2, "0")}-${String(ddmmyyyy[1]).padStart(2, "0")}`;
  }
  const ddmmyy = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2})$/);
  if (ddmmyy) {
    const year = Number(ddmmyy[3]) >= 70 ? 1900 + Number(ddmmyy[3]) : 2000 + Number(ddmmyy[3]);
    return `${year}-${String(ddmmyy[2]).padStart(2, "0")}-${String(ddmmyy[1]).padStart(2, "0")}`;
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return "";
  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  const day = String(parsed.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function buildGoogleCsvUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  try {
    const url = new URL(raw);
    if (url.hostname.includes("docs.google.com") && url.pathname.includes("/spreadsheets/d/")) {
      const pathParts = url.pathname.split("/");
      const sheetId = pathParts[pathParts.indexOf("d") + 1] || "";
      const gid = url.searchParams.get("gid") || "0";
      if (sheetId) {
        return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
      }
    }
  } catch {}
  return raw;
}

function parseCsvText(text) {
  const rows = [];
  let current = [];
  let field = "";
  let inQuotes = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === "\"") {
      if (inQuotes && next === "\"") {
        field += "\"";
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (!inQuotes && (char === ";" || char === ",")) {
      current.push(field);
      field = "";
      continue;
    }
    if (!inQuotes && (char === "\n" || char === "\r")) {
      if (char === "\r" && next === "\n") index += 1;
      current.push(field);
      rows.push(current);
      current = [];
      field = "";
      continue;
    }
    field += char;
  }
  if (field.length || current.length) {
    current.push(field);
    rows.push(current);
  }

  const scored = [
    { delimiter: ";", score: rows[0]?.filter((item) => String(item).includes(";")).length || 0 },
    { delimiter: ",", score: rows[0]?.filter((item) => String(item).includes(",")).length || 0 }
  ];
  if (rows.length <= 1) return rows.map((row) => row.map((item) => String(item || "").trim()));

  const firstRaw = String(text.split(/\r?\n/)[0] || "");
  const delimiter = (firstRaw.match(/;/g) || []).length >= (firstRaw.match(/,/g) || []).length ? ";" : ",";
  const normalizedRows = [];
  for (const line of String(text || "").split(/\r?\n/)) {
    if (!line.trim()) continue;
    const fields = [];
    let token = "";
    let quoted = false;
    for (let index = 0; index < line.length; index += 1) {
      const char = line[index];
      const next = line[index + 1];
      if (char === "\"") {
        if (quoted && next === "\"") {
          token += "\"";
          index += 1;
        } else {
          quoted = !quoted;
        }
        continue;
      }
      if (!quoted && char === delimiter) {
        fields.push(token.trim());
        token = "";
        continue;
      }
      token += char;
    }
    fields.push(token.trim());
    normalizedRows.push(fields);
  }
  return normalizedRows;
}

async function fetchCsvRows(url) {
  const normalizedUrl = buildGoogleCsvUrl(url);
  if (!normalizedUrl) return [];
  const response = await fetch(normalizedUrl, {
    method: "GET",
    headers: { accept: "text/csv,text/plain,application/octet-stream,*/*" }
  });
  if (!response.ok) {
    throw new Error(`Nao foi possivel ler a planilha remota: ${normalizedUrl}`);
  }
  const text = await response.text();
  return parseCsvText(text);
}

function findHeaderIndex(header, aliases) {
  const normalizedAliases = (aliases || []).map((alias) => normalizeLookupKey(alias));
  for (let index = 0; index < header.length; index += 1) {
    const column = normalizeLookupKey(header[index]);
    if (!column) continue;
    if (normalizedAliases.some((alias) => column === alias || column.includes(alias) || alias.includes(column))) {
      return index;
    }
  }
  return -1;
}

function calculateProduction0800(values = {}) {
  return Number(values.approved || 0) + Number(values.cancelled || 0) + Number(values.pending || 0) + Number(values.noAction || 0);
}

function calculateProductionNuvidio(values = {}) {
  return Number(values.approved || 0) + Number(values.reproved || 0) + Number(values.noAction || 0);
}

function calculateEffectiveness0800(values = {}) {
  const approved = Number(values.approved || 0);
  const cancelled = Number(values.cancelled || 0);
  const pending = Number(values.pending || 0);
  const noAction = Number(values.noAction || 0);
  const total = approved + cancelled + pending + noAction;
  if (total <= 0) return 0;
  return ((approved + cancelled + pending) / total) * 100;
}

function calculateEffectivenessNuvidio(values = {}) {
  const approved = Number(values.approved || 0);
  const reproved = Number(values.reproved || 0);
  const noAction = Number(values.noAction || 0);
  const total = approved + reproved + noAction;
  if (total <= 0) return 0;
  return ((approved + reproved) / total) * 100;
}

function averagePlatformEffectiveness(values = {}) {
  const items = [Number(values.effectiveness0800), Number(values.effectivenessNuvidio)].filter(Number.isFinite);
  if (!items.length) return 0;
  return items.reduce((sum, value) => sum + value, 0) / items.length;
}

function buildOperatorLookup(operators) {
  const byId = new Map();
  const byName = new Map();
  const byUsername = new Map();
  const byUsername0800 = new Map();
  const byUsernameNuvidio = new Map();
  for (const operator of operators || []) {
    const safe = sanitizeUser(operator);
    if (!safe.id) continue;
    byId.set(safe.id, safe);
    const nameKey = normalizeLookupKey(safe.name);
    const usernameKey = normalizeLookupKey(safe.username);
    const username0800Key = normalizeLookupKey(safe.username0800);
    const usernameNuvidioKey = normalizeLookupKey(safe.usernameNuvidio);
    if (nameKey) byName.set(nameKey, safe);
    if (usernameKey) byUsername.set(usernameKey, safe);
    if (username0800Key) byUsername0800.set(username0800Key, safe);
    if (usernameNuvidioKey) byUsernameNuvidio.set(usernameNuvidioKey, safe);
    const emailLocalPart = safe.usernameNuvidio.split("@")[0] || "";
    const emailLocalKey = normalizeLookupKey(emailLocalPart);
    if (emailLocalKey && !byUsernameNuvidio.has(emailLocalKey)) byUsernameNuvidio.set(emailLocalKey, safe);
  }
  return { byId, byName, byUsername, byUsername0800, byUsernameNuvidio };
}

function resolveOperatorCandidate(lookup, options = {}) {
  const values = [
    options.userId,
    options.usernameNuvidio,
    options.username0800,
    options.username,
    options.name
  ].map((item) => normalizeLookupKey(item)).filter(Boolean);
  for (const key of values) {
    if (lookup.byId.has(key)) return lookup.byId.get(key);
    if (lookup.byUsernameNuvidio.has(key)) return lookup.byUsernameNuvidio.get(key);
    if (lookup.byUsername0800.has(key)) return lookup.byUsername0800.get(key);
    if (lookup.byUsername.has(key)) return lookup.byUsername.get(key);
    if (lookup.byName.has(key)) return lookup.byName.get(key);
  }
  return null;
}

function createDailyAccumulator(operator) {
  return {
    userId: String(operator?.id || ""),
    userName: String(operator?.name || ""),
    username: String(operator?.username || ""),
    username0800: String(operator?.username0800 || ""),
    usernameNuvidio: String(operator?.usernameNuvidio || ""),
    funnel0800Approved: 0,
    funnel0800Cancelled: 0,
    funnel0800Pending: 0,
    funnel0800NoAction: 0,
    funnelNuvidioApproved: 0,
    funnelNuvidioReproved: 0,
    funnelNuvidioNoAction: 0
  };
}

function ensureDailyAccumulator(store, operator, date) {
  const key = `${String(operator?.id || "")}::${date}`;
  let item = store.get(key);
  if (!item) {
    item = createDailyAccumulator(operator);
    item.date = date;
    store.set(key, item);
  }
  return item;
}

function finalizeDailyAccumulator(item) {
  const production0800 = calculateProduction0800({
    approved: item.funnel0800Approved,
    cancelled: item.funnel0800Cancelled,
    pending: item.funnel0800Pending,
    noAction: item.funnel0800NoAction
  });
  const productionNuvidio = calculateProductionNuvidio({
    approved: item.funnelNuvidioApproved,
    reproved: item.funnelNuvidioReproved,
    noAction: item.funnelNuvidioNoAction
  });
  const effectiveness0800 = calculateEffectiveness0800({
    approved: item.funnel0800Approved,
    cancelled: item.funnel0800Cancelled,
    pending: item.funnel0800Pending,
    noAction: item.funnel0800NoAction
  });
  const effectivenessNuvidio = calculateEffectivenessNuvidio({
    approved: item.funnelNuvidioApproved,
    reproved: item.funnelNuvidioReproved,
    noAction: item.funnelNuvidioNoAction
  });
  return {
    ...item,
    production0800,
    productionNuvidio,
    productionTotal: production0800 + productionNuvidio,
    effectiveness0800,
    effectivenessNuvidio,
    effectiveness: averagePlatformEffectiveness({ effectiveness0800, effectivenessNuvidio }),
    qualityScore: 0
  };
}

async function persistSheetSync(db, payload = {}) {
  await ensureSheetSyncTables(db);
  await db.prepare("DELETE FROM sheet_sync_daily").run();
  await db.prepare("DELETE FROM sheet_sync_events").run();

  const actorId = String(payload?.actorId || "").trim();
  const actorName = String(payload?.actorName || "").trim();
  const dailyRows = Array.isArray(payload?.dailyRows) ? payload.dailyRows : [];
  const events = Array.isArray(payload?.events) ? payload.events : [];

  for (const row of dailyRows) {
    await db.prepare(
      `INSERT INTO sheet_sync_daily (
        id, user_id, user_name, username, username_0800, username_nuvidio, result_date,
        funnel_0800_approved, funnel_0800_cancelled, funnel_0800_pending, funnel_0800_no_action,
        funnel_nuvidio_approved, funnel_nuvidio_reproved, funnel_nuvidio_no_action,
        production_0800, production_nuvidio, production_total,
        effectiveness_0800, effectiveness_nuvidio, effectiveness, quality_score,
        source_origin, updated_by_id, updated_by_name, updated_at, created_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))`
    ).bind(
      `${row.userId}::${row.date}`,
      row.userId,
      row.userName,
      row.username,
      row.username0800,
      row.usernameNuvidio,
      row.date,
      row.funnel0800Approved,
      row.funnel0800Cancelled,
      row.funnel0800Pending,
      row.funnel0800NoAction,
      row.funnelNuvidioApproved,
      row.funnelNuvidioReproved,
      row.funnelNuvidioNoAction,
      row.production0800,
      row.productionNuvidio,
      row.productionTotal,
      row.effectiveness0800,
      row.effectivenessNuvidio,
      row.effectiveness,
      row.qualityScore || 0,
      String(row.sourceOrigin || "spreadsheet"),
      actorId,
      actorName
    ).run();
  }

  for (const event of events) {
    await db.prepare(
      `INSERT INTO sheet_sync_events (
        id, source_origin, platform, status, user_id, user_name, username, username_0800, username_nuvidio,
        event_date, quantity, provider, department, subtag, tma_seconds, wait_seconds, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))`
    ).bind(
      event.id,
      event.sourceOrigin || "",
      event.platform || "",
      event.status || "",
      event.userId || "",
      event.userName || "",
      event.username || "",
      event.username0800 || "",
      event.usernameNuvidio || "",
      event.date || "",
      Number(event.quantity || 0),
      event.provider || "",
      event.department || "",
      event.subtag || "",
      Number(event.tmaSeconds || 0),
      Number(event.waitSeconds || 0)
    ).run();
  }
}

async function readSheetSyncDailyRows(db) {
  await ensureSheetSyncTables(db);
  const rows = await db.prepare("SELECT * FROM sheet_sync_daily ORDER BY user_id ASC, result_date ASC, updated_at ASC").all();
  return rows.results || [];
}

async function readSheetSyncEvents(db) {
  await ensureSheetSyncTables(db);
  const rows = await db.prepare("SELECT * FROM sheet_sync_events ORDER BY event_date ASC, platform ASC, status ASC").all();
  return rows.results || [];
}

function mergeRowsBySourcePreference(manualRows, sourceRows) {
  const byKey = new Map();
  for (const row of sourceRows || []) {
    byKey.set(`${row.user_id}::${row.result_date}`, { source: row, manual: null });
  }
  for (const row of manualRows || []) {
    const key = `${row.user_id}::${row.result_date}`;
    const current = byKey.get(key) || { source: null, manual: null };
    current.manual = row;
    byKey.set(key, current);
  }

  const merged = [];
  for (const pair of byKey.values()) {
    if (pair.source && pair.manual) {
      merged.push({
        ...pair.source,
        quality_score: Number(pair.manual.quality_score || pair.source.quality_score || 0),
        updated_by_id: pair.manual.updated_by_id || pair.source.updated_by_id || "",
        updated_by_name: pair.manual.updated_by_name || pair.source.updated_by_name || "",
        updated_at: pair.manual.updated_at || pair.source.updated_at || ""
      });
      continue;
    }
    merged.push(pair.source || pair.manual);
  }
  merged.sort((a, b) => `${a.user_id}::${a.result_date}`.localeCompare(`${b.user_id}::${b.result_date}`));
  return merged;
}

async function syncSpreadsheetSources(env, actor = {}) {
  const settings = await readSpreadsheetSourceSettings(env.DB);
  const operators = await resolveOperators(env);
  const lookup = buildOperatorLookup(operators);
  const dailyStore = new Map();
  const events = [];
  let unmatchedOperators = 0;

  if (settings.nuvidioUrl) {
    const rows = await fetchCsvRows(settings.nuvidioUrl);
    const header = rows[0] || [];
    const idxTag = findHeaderIndex(header, ["tag"]);
    const idxSubtag = findHeaderIndex(header, ["subtag"]);
    const idxAtendente = findHeaderIndex(header, ["atendente"]);
    const idxAtendenteEmail = findHeaderIndex(header, ["email do atendente", "email atendente"]);
    const idxPrestadora = findHeaderIndex(header, ["prestadora"]);
    const idxDepartamento = findHeaderIndex(header, ["nome do departamento", "departamento"]);
    const idxDate = findHeaderIndex(header, ["data abreviada", "data de entrada", "data"]);
    const idxTma = findHeaderIndex(header, ["tma", "duracao em segundos"]);
    const idxWait = findHeaderIndex(header, ["espera em segundos"]);

    for (let index = 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      const date = normalizeDateValue(row[idxDate]);
      const status = normalizeStatusKey(row[idxTag]);
      if (!date || !status) continue;
      const operator = resolveOperatorCandidate(lookup, {
        usernameNuvidio: row[idxAtendenteEmail],
        name: row[idxAtendente],
        username: row[idxAtendenteEmail]
      });
      if (!operator) {
        unmatchedOperators += 1;
        continue;
      }
      const daily = ensureDailyAccumulator(dailyStore, operator, date);
      if (status === "approved") daily.funnelNuvidioApproved += 1;
      if (status === "reproved") daily.funnelNuvidioReproved += 1;
      if (status === "noAction") daily.funnelNuvidioNoAction += 1;

      events.push({
        id: `nuvidio-${index}-${crypto.randomUUID()}`,
        sourceOrigin: "nuvidio",
        platform: "nuvidio",
        status,
        userId: operator.id,
        userName: operator.name,
        username: operator.username,
        username0800: operator.username0800,
        usernameNuvidio: operator.usernameNuvidio,
        date,
        quantity: 1,
        provider: toTitleCaseLabel(row[idxPrestadora]),
        department: toTitleCaseLabel(row[idxDepartamento]),
        subtag: String(row[idxSubtag] || "").trim(),
        tmaSeconds: parseNumberLoose(row[idxTma]),
        waitSeconds: parseNumberLoose(row[idxWait])
      });
    }
  }

  if (settings.url0800General) {
    const rows = await fetchCsvRows(settings.url0800General);
    const header = rows[0] || [];
    const idxName = findHeaderIndex(header, ["nome do operador", "nome operador", "operador", "nome"]);
    const idxUsername = findHeaderIndex(header, ["usuario", "username", "login"]);
    const idxUsername0800 = findHeaderIndex(header, ["usuario 0800", "username 0800", "login 0800"]);
    const idxDate = findHeaderIndex(header, ["data", "dia", "data resultado"]);
    const idxStatus = findHeaderIndex(header, ["status", "resultado", "situacao"]);
    const idxQuantity = findHeaderIndex(header, ["quantidade", "qtd", "qtde", "volume", "producao"]);
    const idxApproved = findHeaderIndex(header, ["aprovadas", "aprovada"]);
    const idxCancelled = findHeaderIndex(header, ["canceladas", "cancelada"]);
    const idxPending = findHeaderIndex(header, ["pendenciadas", "pendenciada"]);
    const idxNoAction = findHeaderIndex(header, ["sem acao", "sem ação"]);

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const operator = resolveOperatorCandidate(lookup, {
        username0800: row[idxUsername0800 >= 0 ? idxUsername0800 : idxUsername],
        username: row[idxUsername],
        name: row[idxName]
      });
      if (!operator) {
        if (String(row[idxName >= 0 ? idxName : 0] || "").trim()) unmatchedOperators += 1;
        continue;
      }
      const date = normalizeDateValue(row[idxDate]);
      if (!date) continue;
      const daily = ensureDailyAccumulator(dailyStore, operator, date);

      if (idxStatus >= 0 && idxQuantity >= 0) {
        const status = normalizeStatusKey(row[idxStatus]);
        const quantity = parseNumberLoose(row[idxQuantity]);
        if (!status || !quantity) continue;
        if (status === "approved") daily.funnel0800Approved += quantity;
        if (status === "cancelled") daily.funnel0800Cancelled += quantity;
        if (status === "pending") daily.funnel0800Pending += quantity;
        if (status === "noAction") daily.funnel0800NoAction += quantity;
        events.push({
          id: `0800-general-${rowIndex}-${crypto.randomUUID()}`,
          sourceOrigin: "0800-general",
          platform: "0800",
          status,
          userId: operator.id,
          userName: operator.name,
          username: operator.username,
          username0800: operator.username0800,
          usernameNuvidio: operator.usernameNuvidio,
          date,
          quantity,
          provider: "",
          department: "",
          subtag: ""
        });
        continue;
      }

      const approved = idxApproved >= 0 ? parseNumberLoose(row[idxApproved]) : 0;
      const cancelled = idxCancelled >= 0 ? parseNumberLoose(row[idxCancelled]) : 0;
      const pending = idxPending >= 0 ? parseNumberLoose(row[idxPending]) : 0;
      const noAction = idxNoAction >= 0 ? parseNumberLoose(row[idxNoAction]) : 0;
      if (!(approved || cancelled || pending || noAction)) continue;
      daily.funnel0800Approved += approved;
      daily.funnel0800Cancelled += cancelled;
      daily.funnel0800Pending += pending;
      daily.funnel0800NoAction += noAction;

      const eventMap = [
        ["approved", approved],
        ["cancelled", cancelled],
        ["pending", pending],
        ["noAction", noAction]
      ];
      for (const [status, quantity] of eventMap) {
        if (!quantity) continue;
        events.push({
          id: `0800-general-${status}-${rowIndex}-${crypto.randomUUID()}`,
          sourceOrigin: "0800-general",
          platform: "0800",
          status,
          userId: operator.id,
          userName: operator.name,
          username: operator.username,
          username0800: operator.username0800,
          usernameNuvidio: operator.usernameNuvidio,
          date,
          quantity,
          provider: "",
          department: "",
          subtag: ""
        });
      }
    }
  }

  const matrixSources = [
    { url: settings.url0800Approved, status: "approved", sourceOrigin: "0800-approved" },
    { url: settings.url0800Cancelled, status: "cancelled", sourceOrigin: "0800-cancelled" },
    { url: settings.url0800Pending, status: "pending", sourceOrigin: "0800-pending" },
    { url: settings.url0800NoAction, status: "noAction", sourceOrigin: "0800-no-action" }
  ].filter((item) => item.url);

  for (const source of matrixSources) {
    const rows = await fetchCsvRows(source.url);
    const header = rows[0] || [];
    const dateColumns = header.map((value, index) => ({ index, date: normalizeDateValue(value) })).filter((item) => item.index > 0 && item.date);
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const operator = resolveOperatorCandidate(lookup, {
        username0800: row[0],
        username: row[0],
        name: row[0]
      });
      if (!operator) {
        if (String(row[0] || "").trim()) unmatchedOperators += 1;
        continue;
      }
      for (const dateColumn of dateColumns) {
        const quantity = parseNumberLoose(row[dateColumn.index]);
        if (!quantity) continue;
        const daily = ensureDailyAccumulator(dailyStore, operator, dateColumn.date);
        if (source.status === "approved") daily.funnel0800Approved += quantity;
        if (source.status === "cancelled") daily.funnel0800Cancelled += quantity;
        if (source.status === "pending") daily.funnel0800Pending += quantity;
        if (source.status === "noAction") daily.funnel0800NoAction += quantity;
        events.push({
          id: `${source.sourceOrigin}-${rowIndex}-${dateColumn.date}-${crypto.randomUUID()}`,
          sourceOrigin: source.sourceOrigin,
          platform: "0800",
          status: source.status,
          userId: operator.id,
          userName: operator.name,
          username: operator.username,
          username0800: operator.username0800,
          usernameNuvidio: operator.usernameNuvidio,
          date: dateColumn.date,
          quantity,
          provider: "",
          department: "",
          subtag: ""
        });
      }
    }
  }

  const dailyRows = [...dailyStore.values()].map((item) => ({
    ...finalizeDailyAccumulator(item),
    sourceOrigin: "spreadsheet"
  }));

  await persistSheetSync(env.DB, {
    dailyRows,
    events,
    actorId: String(actor?.actorId || "").trim(),
    actorName: String(actor?.actorName || "").trim()
  });

  const nextSettings = normalizeSpreadsheetSourceSettings({
    ...settings,
    lastSyncedAt: new Date().toISOString(),
    lastSyncSummary: {
      importedDailyRows: dailyRows.length,
      importedEvents: events.length,
      unmatchedOperators
    },
    updatedById: String(actor?.actorId || settings.updatedById || "").trim(),
    updatedByName: String(actor?.actorName || settings.updatedByName || "").trim()
  });
  await writeJsonSystemSetting(env.DB, "spreadsheet_sources", nextSettings, {
    updatedById: nextSettings.updatedById,
    updatedByName: nextSettings.updatedByName
  });
  return nextSettings;
}

async function buildSourceInsights(db) {
  const settings = await readSpreadsheetSourceSettings(db);
  const dailyRows = await readSheetSyncDailyRows(db);
  const events = await readSheetSyncEvents(db);
  const uniqueOperators = new Set(dailyRows.map((row) => String(row.user_id || "")).filter(Boolean));
  const uniqueDates = new Set(dailyRows.map((row) => String(row.result_date || "")).filter(Boolean));

  const totals = dailyRows.reduce((acc, row) => {
    acc.production0800 += Number(row.production_0800 || 0);
    acc.productionNuvidio += Number(row.production_nuvidio || 0);
    acc.funnel0800Approved += Number(row.funnel_0800_approved || 0);
    acc.funnel0800Cancelled += Number(row.funnel_0800_cancelled || 0);
    acc.funnel0800Pending += Number(row.funnel_0800_pending || 0);
    acc.funnel0800NoAction += Number(row.funnel_0800_no_action || 0);
    acc.funnelNuvidioApproved += Number(row.funnel_nuvidio_approved || 0);
    acc.funnelNuvidioReproved += Number(row.funnel_nuvidio_reproved || 0);
    acc.funnelNuvidioNoAction += Number(row.funnel_nuvidio_no_action || 0);
    return acc;
  }, {
    production0800: 0,
    productionNuvidio: 0,
    funnel0800Approved: 0,
    funnel0800Cancelled: 0,
    funnel0800Pending: 0,
    funnel0800NoAction: 0,
    funnelNuvidioApproved: 0,
    funnelNuvidioReproved: 0,
    funnelNuvidioNoAction: 0
  });

  const byDepartment = new Map();
  const bySubtag = new Map();
  const byProvider = new Map();
  const byOperator = new Map();
  const by0800Status = new Map();
  let nuvidioTmaSum = 0;
  let nuvidioWaitSum = 0;
  let nuvidioTimedEvents = 0;

  for (const event of events) {
    const quantity = Number(event.quantity || 0);
    const departmentKey = String(event.department || "").trim();
    const subtagKey = String(event.subtag || "").trim();
    const providerKey = String(event.provider || "").trim();
    const operatorKey = String(event.user_id || "").trim();

    if (event.platform === "nuvidio") {
      if (departmentKey) byDepartment.set(departmentKey, (byDepartment.get(departmentKey) || 0) + quantity);
      if (subtagKey) bySubtag.set(subtagKey, (bySubtag.get(subtagKey) || 0) + quantity);
      if (providerKey) byProvider.set(providerKey, (byProvider.get(providerKey) || 0) + quantity);
      if (Number(event.tma_seconds || 0) > 0) {
        nuvidioTmaSum += Number(event.tma_seconds || 0);
        nuvidioWaitSum += Number(event.wait_seconds || 0);
        nuvidioTimedEvents += 1;
      }
    }

    if (event.platform === "0800") {
      by0800Status.set(String(event.status || ""), (by0800Status.get(String(event.status || "")) || 0) + quantity);
    }

    if (operatorKey) {
      const previous = byOperator.get(operatorKey) || {
        userId: operatorKey,
        userName: String(event.user_name || ""),
        production0800: 0,
        productionNuvidio: 0
      };
      if (event.platform === "0800") previous.production0800 += quantity;
      if (event.platform === "nuvidio") previous.productionNuvidio += quantity;
      byOperator.set(operatorKey, previous);
    }
  }

  const effectiveness0800 = calculateEffectiveness0800({
    approved: totals.funnel0800Approved,
    cancelled: totals.funnel0800Cancelled,
    pending: totals.funnel0800Pending,
    noAction: totals.funnel0800NoAction
  });
  const effectivenessNuvidio = calculateEffectivenessNuvidio({
    approved: totals.funnelNuvidioApproved,
    reproved: totals.funnelNuvidioReproved,
    noAction: totals.funnelNuvidioNoAction
  });

  return {
    lastSyncedAt: settings.lastSyncedAt || "",
    summary: {
      operators: uniqueOperators.size,
      days: uniqueDates.size,
      production0800: totals.production0800,
      productionNuvidio: totals.productionNuvidio,
      effectiveness0800,
      effectivenessNuvidio,
      avgNuvidioTmaSeconds: nuvidioTimedEvents ? nuvidioTmaSum / nuvidioTimedEvents : 0,
      avgNuvidioWaitSeconds: nuvidioTimedEvents ? nuvidioWaitSum / nuvidioTimedEvents : 0
    },
    departments: [...byDepartment.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value).slice(0, 6),
    subtags: [...bySubtag.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value).slice(0, 8),
    providers: [...byProvider.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value).slice(0, 6),
    status0800: [...by0800Status.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value),
    topOperators: [...byOperator.values()].map((item) => ({
      ...item,
      productionTotal: Number(item.production0800 || 0) + Number(item.productionNuvidio || 0)
    })).sort((a, b) => b.productionTotal - a.productionTotal).slice(0, 8)
  };
}

async function readLocalUsers(db) {
  await ensureLocalUsersTable(db);
  const rows = await db.prepare("SELECT * FROM local_portal_users ORDER BY name ASC, username ASC").all();
  return (rows.results || []).map(sanitizeLocalUser);
}

async function saveLocalUser(db, payload = {}) {
  await ensureLocalUsersTable(db);
  const id = String(payload?.id || "").trim() || `local-${crypto.randomUUID()}`;
  const name = String(payload?.name || "").trim();
  const username = String(payload?.username || "").trim();
  const role = String(payload?.role || "operador").trim() || "operador";
  const password = String(payload?.password || "");
  const username0800 = String(payload?.username0800 || "").trim();
  const usernameNuvidio = String(payload?.usernameNuvidio || "").trim();
  const updatedById = String(payload?.updatedById || "").trim();
  const updatedByName = String(payload?.updatedByName || "").trim();

  if (!name || !username) {
    return { ok: false, error: "Nome e login do portal sao obrigatorios." };
  }

  const existing = await db.prepare("SELECT * FROM local_portal_users WHERE id = ? LIMIT 1").bind(id).first();
  let passwordHash = String(existing?.password_hash || "");
  if (password) {
    passwordHash = await sha256Hex(password);
  }
  if (!passwordHash) {
    return { ok: false, error: "Informe uma senha para o usuario." };
  }

  await db.prepare(
    `INSERT INTO local_portal_users (
      id, name, username, password_hash, role, access_level, username_0800, username_nuvidio,
      created_at, updated_at, updated_by_id, updated_by_name
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      name = excluded.name,
      username = excluded.username,
      password_hash = excluded.password_hash,
      role = excluded.role,
      access_level = excluded.access_level,
      username_0800 = excluded.username_0800,
      username_nuvidio = excluded.username_nuvidio,
      updated_at = excluded.updated_at,
      updated_by_id = excluded.updated_by_id,
      updated_by_name = excluded.updated_by_name`
  )
    .bind(
      id,
      name,
      username,
      passwordHash,
      role,
      "",
      username0800,
      usernameNuvidio,
      new Date().toISOString(),
      updatedById,
      updatedByName
    )
    .run();

  return { ok: true, id };
}

async function verifyLocalUserLogin(db, username, password) {
  await ensureLocalUsersTable(db);
  const row = await db.prepare("SELECT * FROM local_portal_users WHERE username = ? LIMIT 1").bind(username).first();
  if (!row) {
    throw new Error("Usuario ou senha invalidos.");
  }
  const passwordHash = await sha256Hex(password);
  if (String(row.password_hash || "") !== passwordHash) {
    throw new Error("Usuario ou senha invalidos.");
  }
  return sanitizeLocalUser(row);
}

async function readOperatorAccessLinks(db) {
  await ensureOperatorAccessLinksTable(db);
  const rows = await db.prepare("SELECT * FROM operator_access_links").all();
  const byUserId = new Map();
  const byName = new Map();
  const byUsername = new Map();
  for (const row of rows.results || []) {
    const userId = String(row.user_id || "").trim();
    const entry = {
      userId,
      userName: String(row.user_name || ""),
      username: String(row.username || ""),
      username0800: String(row.username_0800 || ""),
      usernameNuvidio: String(row.username_nuvidio || "")
    };
    if (userId) byUserId.set(userId, entry);
    const nameKey = normalizeLookupKey(entry.userName);
    if (nameKey) byName.set(nameKey, entry);
    const usernameKey = normalizeLookupKey(entry.username);
    if (usernameKey) byUsername.set(usernameKey, entry);
  }
  return { byUserId, byName, byUsername };
}

async function saveOperatorAccessLink(db, payload = {}) {
  await ensureOperatorAccessLinksTable(db);
  const userId = String(payload?.userId || "").trim();
  if (!userId) {
    return { ok: false, error: "userId obrigatorio para salvar acessos do operador." };
  }

  const userName = String(payload?.userName || "").trim();
  const username = String(payload?.username || "").trim();
  const username0800 = String(payload?.username0800 || "").trim();
  const usernameNuvidio = String(payload?.usernameNuvidio || "").trim();
  const updatedById = String(payload?.updatedById || "").trim();
  const updatedByName = String(payload?.updatedByName || "").trim();
  const updatedAt = new Date().toISOString();

  await db.prepare(
    `INSERT INTO operator_access_links (
      user_id, user_name, username, username_0800, username_nuvidio, updated_at, updated_by_id, updated_by_name
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(user_id) DO UPDATE SET
      user_name = excluded.user_name,
      username = excluded.username,
      username_0800 = excluded.username_0800,
      username_nuvidio = excluded.username_nuvidio,
      updated_at = excluded.updated_at,
      updated_by_id = excluded.updated_by_id,
      updated_by_name = excluded.updated_by_name`
  )
    .bind(userId, userName, username, username0800, usernameNuvidio, updatedAt, updatedById, updatedByName)
    .run();

  return { ok: true, userId };
}

function normalizeMaintenanceStatus(raw = {}) {
  return {
    enabled: Boolean(raw?.enabled),
    message: String(raw?.message || "O portal esta temporariamente em manutencao. Tente novamente em alguns minutos."),
    updatedAt: String(raw?.updatedAt || ""),
    updatedById: String(raw?.updatedById || ""),
    updatedByName: String(raw?.updatedByName || "")
  };
}

async function readSystemMaintenanceStatus(db) {
  await ensureSystemSettingsTable(db);
  const row = await db
    .prepare("SELECT value_text FROM system_settings WHERE setting_key = ? LIMIT 1")
    .bind("maintenance_mode")
    .first();
  if (!row?.value_text) {
    return normalizeMaintenanceStatus({ enabled: false });
  }

  try {
    const parsed = JSON.parse(String(row.value_text || "{}"));
    return normalizeMaintenanceStatus(parsed);
  } catch {
    return normalizeMaintenanceStatus({ enabled: false });
  }
}

async function updateSystemMaintenanceStatus(db, payload = {}) {
  await ensureSystemSettingsTable(db);
  const nextStatus = normalizeMaintenanceStatus({
    enabled: Boolean(payload?.enabled),
    message: String(payload?.message || "").trim(),
    updatedAt: new Date().toISOString(),
    updatedById: String(payload?.updatedById || "").trim(),
    updatedByName: String(payload?.updatedByName || "").trim()
  });

  await db
    .prepare(
      `INSERT INTO system_settings (setting_key, value_text, updated_at, updated_by_id, updated_by_name)
       VALUES (?, ?, ?, ?, ?)
       ON CONFLICT(setting_key) DO UPDATE SET
         value_text = excluded.value_text,
         updated_at = excluded.updated_at,
         updated_by_id = excluded.updated_by_id,
         updated_by_name = excluded.updated_by_name`
    )
    .bind(
      "maintenance_mode",
      JSON.stringify(nextStatus),
      nextStatus.updatedAt,
      nextStatus.updatedById,
      nextStatus.updatedByName
    )
    .run();

  return nextStatus;
}

async function markSsoTokenAsConsumed(db, jti, expSeconds) {
  await ensureSsoReplayTable(db);
  const nowSeconds = Math.floor(Date.now() / 1000);
  await db.prepare("DELETE FROM sso_consumed_tokens WHERE expires_at < ?").bind(nowSeconds).run();
  await db
    .prepare("INSERT INTO sso_consumed_tokens (jti, expires_at, consumed_at) VALUES (?, ?, datetime('now'))")
    .bind(jti, expSeconds)
    .run();
}

function buildRecord(userId, rows) {
  const entries = (rows || []).map((row) => ({
    date: row.result_date,
    funnel0800Approved: Number(row.funnel_0800_approved || 0),
    funnel0800Cancelled: Number(row.funnel_0800_cancelled || 0),
    funnel0800Pending: Number(row.funnel_0800_pending || 0),
    funnel0800NoAction: Number(row.funnel_0800_no_action || 0),
    funnelNuvidioApproved: Number(row.funnel_nuvidio_approved || 0),
    funnelNuvidioReproved: Number(row.funnel_nuvidio_reproved || 0),
    funnelNuvidioNoAction: Number(row.funnel_nuvidio_no_action || 0),
    production0800: Number(row.production_0800 || 0),
    productionNuvidio: Number(row.production_nuvidio || 0),
    productionTotal: Number(row.production_total || 0),
    effectiveness0800: Number(row.effectiveness_0800 || 0),
    effectivenessNuvidio: Number(row.effectiveness_nuvidio || 0),
    effectiveness: Number(row.effectiveness || 0),
    qualityScore: Number(row.quality_score || 0),
    updatedById: row.updated_by_id || "",
    updatedByName: row.updated_by_name || "",
    updatedAt: row.updated_at || row.created_at || ""
  }));

  if (!entries.length) return null;
  entries.sort((a, b) => String(a.date).localeCompare(String(b.date)));
  const latest = entries[entries.length - 1];
  const productionAverage = entries.reduce((sum, entry) => sum + Number(entry.productionTotal || 0), 0) / entries.length;
  const firstRow = rows[0] || {};

  return {
    userId,
    userName: firstRow.user_name || "",
    username: firstRow.username || "",
    username0800: firstRow.username_0800 || "",
    usernameNuvidio: firstRow.username_nuvidio || "",
    entries,
    daysCount: entries.length,
    productionAverage,
    productionTotal: latest.productionTotal,
    effectiveness: latest.effectiveness,
    qualityScore: latest.qualityScore,
    updatedAt: latest.updatedAt,
    updatedById: latest.updatedById,
    updatedByName: latest.updatedByName
  };
}

async function fetchCentral(env, pathname, options = {}) {
  const authBase = getAuthBase(env);
  if (!authBase) {
    throw new Error("AUTH_API_BASE nao configurado.");
  }

  const response = await fetch(`${authBase}${pathname}`, options);
  const payload = await response.json().catch(() => null);
  if (!response.ok || payload?.ok === false) {
    throw new Error(payload?.error || "Falha ao consultar a Central do Operador.");
  }
  return payload;
}

function sanitizeUser(user) {
  return {
    id: String(user?.id || ""),
    name: String(user?.name || ""),
    username: String(user?.username || ""),
    username0800: String(
      user?.username0800 ||
      user?.username_0800 ||
      user?.login0800 ||
      user?.login_0800 ||
      ""
    ),
    usernameNuvidio: String(
      user?.usernameNuvidio ||
      user?.username_nuvidio ||
      user?.loginNuvidio ||
      user?.login_nuvidio ||
      ""
    ),
    role: String(user?.role || "operador"),
    accessLevel: String(user?.accessLevel || "")
  };
}

async function resolveOperators(env) {
  let users = [];
  try {
    const payload = await fetchCentral(env, "/api/users", { method: "GET" });
    users = Array.isArray(payload?.users) ? payload.users : [];
  } catch {}
  const savedAccessLinks = await readOperatorAccessLinks(env.DB);
  return users
    .map(sanitizeUser)
    .map((user) => {
      const saved =
        savedAccessLinks.byUserId.get(String(user.id || "")) ||
        savedAccessLinks.byUsername.get(normalizeLookupKey(user.username)) ||
        savedAccessLinks.byName.get(normalizeLookupKey(user.name));
      if (!saved) return user;
      return {
        ...user,
        username0800: saved.username0800 || user.username0800 || "",
        usernameNuvidio: saved.usernameNuvidio || user.usernameNuvidio || ""
      };
    })
    .filter((user) => user.id && user.role !== "gestor")
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt-BR"));
}

async function readRecord(db, userId) {
  await ensureResultsTable(db);
  await ensureSheetSyncTables(db);
  const manualRows = await db
    .prepare("SELECT * FROM operator_results_daily WHERE user_id = ? ORDER BY result_date ASC, updated_at ASC")
    .bind(userId)
    .all();
  const sourceRows = await db
    .prepare("SELECT * FROM sheet_sync_daily WHERE user_id = ? ORDER BY result_date ASC, updated_at ASC")
    .bind(userId)
    .all();
  return buildRecord(userId, mergeRowsBySourcePreference(manualRows.results || [], sourceRows.results || []));
}

async function upsertOperatorResult(db, payload) {
  const userId = String(payload?.userId || "").trim();
  const userName = String(payload?.userName || "").trim();
  const username = String(payload?.username || "").trim();
  const username0800 = String(payload?.username0800 || "").trim();
  const usernameNuvidio = String(payload?.usernameNuvidio || "").trim();
  const date = normalizeDateKey(payload?.date);
  const funnel0800Approved = Number(payload?.funnel0800Approved);
  const funnel0800Cancelled = Number(payload?.funnel0800Cancelled);
  const funnel0800Pending = Number(payload?.funnel0800Pending);
  const funnel0800NoAction = Number(payload?.funnel0800NoAction);
  const funnelNuvidioApproved = Number(payload?.funnelNuvidioApproved);
  const funnelNuvidioReproved = Number(payload?.funnelNuvidioReproved);
  const funnelNuvidioNoAction = Number(payload?.funnelNuvidioNoAction);
  const production0800 = Number(payload?.production0800);
  const productionNuvidio = Number(payload?.productionNuvidio);
  const productionTotal = Number(payload?.productionTotal);
  const effectiveness0800 = Number(payload?.effectiveness0800);
  const effectivenessNuvidio = Number(payload?.effectivenessNuvidio);
  const effectiveness = Number(payload?.effectiveness);
  const qualityScore = Number(payload?.qualityScore);
  const updatedById = String(payload?.updatedById || "").trim();
  const updatedByName = String(payload?.updatedByName || "").trim();

  if (
    !userId ||
    !date ||
    !Number.isFinite(funnel0800Approved) ||
    !Number.isFinite(funnel0800Cancelled) ||
    !Number.isFinite(funnel0800Pending) ||
    !Number.isFinite(funnel0800NoAction) ||
    !Number.isFinite(funnelNuvidioApproved) ||
    !Number.isFinite(funnelNuvidioReproved) ||
    !Number.isFinite(funnelNuvidioNoAction) ||
    !Number.isFinite(production0800) ||
    !Number.isFinite(productionNuvidio) ||
    !Number.isFinite(productionTotal) ||
    !Number.isFinite(effectiveness0800) ||
    !Number.isFinite(effectivenessNuvidio) ||
    !Number.isFinite(effectiveness) ||
    !Number.isFinite(qualityScore)
  ) {
    return { ok: false, error: "Payload invalido para resultado do operador." };
  }

  await ensureResultsTable(db);
  await db.prepare(
    `INSERT INTO operator_results_daily (
      id, user_id, user_name, username, username_0800, username_nuvidio, result_date,
      funnel_0800_approved, funnel_0800_cancelled, funnel_0800_pending, funnel_0800_no_action,
      funnel_nuvidio_approved, funnel_nuvidio_reproved, funnel_nuvidio_no_action,
      production_0800, production_nuvidio, production_total,
      effectiveness_0800, effectiveness_nuvidio, effectiveness, quality_score,
      updated_by_id, updated_by_name, updated_at, created_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(user_id, result_date) DO UPDATE SET
      user_name = excluded.user_name,
      username = excluded.username,
      username_0800 = excluded.username_0800,
      username_nuvidio = excluded.username_nuvidio,
      funnel_0800_approved = excluded.funnel_0800_approved,
      funnel_0800_cancelled = excluded.funnel_0800_cancelled,
      funnel_0800_pending = excluded.funnel_0800_pending,
      funnel_0800_no_action = excluded.funnel_0800_no_action,
      funnel_nuvidio_approved = excluded.funnel_nuvidio_approved,
      funnel_nuvidio_reproved = excluded.funnel_nuvidio_reproved,
      funnel_nuvidio_no_action = excluded.funnel_nuvidio_no_action,
      production_0800 = excluded.production_0800,
      production_nuvidio = excluded.production_nuvidio,
      production_total = excluded.production_total,
      effectiveness_0800 = excluded.effectiveness_0800,
      effectiveness_nuvidio = excluded.effectiveness_nuvidio,
      effectiveness = excluded.effectiveness,
      quality_score = excluded.quality_score,
      updated_by_id = excluded.updated_by_id,
      updated_by_name = excluded.updated_by_name,
      updated_at = excluded.updated_at`
  )
    .bind(
      `${userId}__${date}`,
      userId,
      userName,
      username,
      username0800,
      usernameNuvidio,
      date,
      funnel0800Approved,
      funnel0800Cancelled,
      funnel0800Pending,
      funnel0800NoAction,
      funnelNuvidioApproved,
      funnelNuvidioReproved,
      funnelNuvidioNoAction,
      production0800,
      productionNuvidio,
      productionTotal,
      effectiveness0800,
      effectivenessNuvidio,
      effectiveness,
      qualityScore,
      updatedById,
      updatedByName,
      new Date().toISOString()
    )
    .run();

  return { ok: true, userId };
}

async function deleteOperatorResult(db, payload) {
  const userId = String(payload?.userId || "").trim();
  const date = normalizeDateKey(payload?.date);
  if (!userId || !date) {
    return { ok: false, error: "Informe userId e date para excluir o lancamento." };
  }

  await ensureResultsTable(db);
  const result = await db
    .prepare("DELETE FROM operator_results_daily WHERE user_id = ? AND result_date = ?")
    .bind(userId, date)
    .run();

  const deleted = Number(result?.meta?.changes || 0);
  if (!deleted) {
    return { ok: false, error: "Nenhum lancamento encontrado para exclusao nessa data." };
  }

  return { ok: true, userId, date, deleted };
}

async function deleteAllOperatorResults(db) {
  await ensureResultsTable(db);
  const result = await db.prepare("DELETE FROM operator_results_daily").run();
  const deleted = Number(result?.meta?.changes || 0);
  return { ok: true, deleted };
}

async function readAllRecords(db) {
  await ensureResultsTable(db);
  await ensureSheetSyncTables(db);
  const manualRows = await db
    .prepare("SELECT * FROM operator_results_daily ORDER BY user_id ASC, result_date ASC, updated_at ASC")
    .all();
  const sourceRows = await db
    .prepare("SELECT * FROM sheet_sync_daily ORDER BY user_id ASC, result_date ASC, updated_at ASC")
    .all();
  const rows = mergeRowsBySourcePreference(manualRows.results || [], sourceRows.results || []);

  const byUser = new Map();
  for (const row of rows || []) {
    const userId = String(row.user_id || "").trim();
    if (!userId) continue;
    const list = byUser.get(userId) || [];
    list.push(row);
    byUser.set(userId, list);
  }

  const records = [];
  for (const [userId, entries] of byUser.entries()) {
    const record = buildRecord(userId, entries);
    if (record) records.push(record);
  }
  return records;
}

async function verifySsoTokenAndBuildUser(env, requestUrl, token, db) {
  const { sharedSecret, expectedIssuer, expectedAudience } = getSsoConfig(env, requestUrl);
  if (!sharedSecret) {
    throw new Error("SSO_SHARED_SECRET nao configurado.");
  }

  const payload = await verifyJwtHs256(token, sharedSecret);
  const nowSeconds = Math.floor(Date.now() / 1000);
  const exp = Number(payload?.exp);
  const iat = Number(payload?.iat);
  const iss = String(payload?.iss || "");
  const aud = String(payload?.aud || "");
  const jti = String(payload?.jti || "");

  if (!Number.isFinite(exp) || exp <= nowSeconds) {
    throw new Error("Token expirado.");
  }
  if (!Number.isFinite(iat) || iat > nowSeconds + 30) {
    throw new Error("Token com iat invalido.");
  }
  if (!jti) {
    throw new Error("Token sem jti.");
  }
  if (expectedIssuer && iss !== expectedIssuer) {
    throw new Error("Issuer invalido.");
  }
  if (expectedAudience && aud !== expectedAudience) {
    throw new Error("Audience invalida.");
  }

  try {
    await markSsoTokenAsConsumed(db, jti, exp);
  } catch (error) {
    const detail = String(error?.message || "");
    if (detail.includes("UNIQUE")) {
      throw new Error("Token SSO ja utilizado.");
    }
    throw error;
  }

  const id = String(payload?.id || payload?.sub || "").trim();
  const name = String(payload?.name || "").trim();
  const username = String(payload?.username || "").trim();
  const role = String(payload?.role || "operador").trim();
  const accessLevel = String(payload?.accessLevel || payload?.access_level || "").trim();

  if (!id || !name || !username) {
    throw new Error("Token sem dados obrigatorios do usuario.");
  }

  return {
    user: { id, name, username, role, accessLevel },
    payload
  };
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return jsonResponse({ ok: true }, 204);
    }

    if (!url.pathname.startsWith("/api/")) {
      return env.ASSETS.fetch(request);
    }

    if (url.pathname === "/api/health") {
      const ssoConfig = getSsoConfig(env, request.url);
      return jsonResponse({
        ok: true,
        service: "portal-resultados-operador",
        hasDB: Boolean(env.DB),
        hasR2: Boolean(env.RESULTS_BUCKET || env.IMPORTS_BUCKET),
        hasImportsBucket: Boolean(env.IMPORTS_BUCKET),
        hasResultsBucket: Boolean(env.RESULTS_BUCKET),
        hasImportR2: Boolean(env.IMPORTS_BUCKET),
        authBaseConfigured: Boolean(getAuthBase(env)),
        ssoSecretConfigured: Boolean(ssoConfig.sharedSecret),
        ssoAudience: ssoConfig.expectedAudience || null,
        ssoIssuer: ssoConfig.expectedIssuer || null
      });
    }

    if (!env.DB) {
      return jsonResponse({ ok: false, error: "D1 binding DB nao configurado." }, 500);
    }

    try {
      if (url.pathname === "/api/login" && request.method === "POST") {
        const body = await request.json();
        const username = String(body?.username || "").trim();
        const password = String(body?.password || "");
        if (!username || !password) {
          return jsonResponse({ ok: false, error: "Usuario e senha obrigatorios." }, 400);
        }

        try {
          const payload = await fetchCentral(env, "/api/login", {
            method: "POST",
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ username, password })
          });
          return jsonResponse({ ok: true, user: sanitizeUser(payload.user) });
        } catch (centralError) {
          try {
            const user = await verifyLocalUserLogin(env.DB, username, password);
            return jsonResponse({ ok: true, user });
          } catch {
            return jsonResponse({ ok: false, error: "Usuario ou senha invalidos." }, 401);
          }
        }
      }

      if (url.pathname === "/api/sso/consume" && request.method === "POST") {
        const body = await request.json();
        const token = String(body?.token || "").trim();
        if (!token) {
          return jsonResponse({ ok: false, error: "Token SSO obrigatorio." }, 400);
        }

        const result = await verifySsoTokenAndBuildUser(env, request.url, token, env.DB);
        return jsonResponse({
          ok: true,
          user: sanitizeUser(result.user),
          exp: Number(result.payload.exp || 0)
        });
      }

      if (url.pathname === "/api/operators" && request.method === "GET") {
        const operators = await resolveOperators(env);
        return jsonResponse({ ok: true, operators });
      }

      if (url.pathname === "/api/users/local" && request.method === "GET") {
        const users = await readLocalUsers(env.DB);
        return jsonResponse({ ok: true, users });
      }

      if (url.pathname === "/api/users/local" && request.method === "POST") {
        const body = await request.json().catch(() => ({}));
        const result = await saveLocalUser(env.DB, body || {});
        if (!result.ok) {
          return jsonResponse({ ok: false, error: result.error || "Nao foi possivel salvar o usuario." }, 400);
        }
        const users = await readLocalUsers(env.DB);
        return jsonResponse({ ok: true, users, id: result.id });
      }

      if (url.pathname === "/api/operators/access" && request.method === "POST") {
        const body = await request.json().catch(() => ({}));
        const result = await saveOperatorAccessLink(env.DB, body || {});
        if (!result.ok) {
          return jsonResponse({ ok: false, error: result.error || "Nao foi possivel salvar os acessos do operador." }, 400);
        }
        const operators = await resolveOperators(env);
        return jsonResponse({ ok: true, operators, userId: result.userId });
      }

      if (url.pathname === "/api/system-status" && request.method === "GET") {
        const status = await readSystemMaintenanceStatus(env.DB);
        return jsonResponse({ ok: true, status });
      }

      if (url.pathname === "/api/system-maintenance" && request.method === "POST") {
        const body = await request.json().catch(() => ({}));
        const actorRole = String(body?.actorRole || "").trim().toLowerCase();
        if (actorRole !== "gestor") {
          return jsonResponse({ ok: false, error: "Somente gestor pode alterar o modo manutencao." }, 403);
        }
        const status = await updateSystemMaintenanceStatus(env.DB, body || {});
        return jsonResponse({ ok: true, status });
      }

      if (url.pathname === "/api/source-settings" && request.method === "GET") {
        const settings = await readSpreadsheetSourceSettings(env.DB);
        return jsonResponse({ ok: true, settings });
      }

      if (url.pathname === "/api/source-settings" && request.method === "POST") {
        const body = await request.json().catch(() => ({}));
        const settings = await saveSpreadsheetSourceSettings(env.DB, body || {});
        return jsonResponse({ ok: true, settings });
      }

      if (url.pathname === "/api/source-sync" && request.method === "POST") {
        const body = await request.json().catch(() => ({}));
        const settings = await syncSpreadsheetSources(env, body || {});
        const insights = await buildSourceInsights(env.DB);
        return jsonResponse({ ok: true, settings, insights });
      }

      if (url.pathname === "/api/source-insights" && request.method === "GET") {
        const insights = await buildSourceInsights(env.DB);
        return jsonResponse({ ok: true, insights });
      }

      if (url.pathname === "/api/results" && request.method === "GET") {
        const userId = String(url.searchParams.get("userId") || "").trim();
        if (!userId) {
          return jsonResponse({ ok: false, error: "userId obrigatorio." }, 400);
        }

        const record = await readRecord(env.DB, userId);
        return jsonResponse({ ok: true, record });
      }

      if (url.pathname === "/api/results/all" && request.method === "GET") {
        const records = await readAllRecords(env.DB);
        return jsonResponse({ ok: true, records });
      }

      if (url.pathname === "/api/operator-results" && request.method === "POST") {
        const body = await request.json();
        const result = await upsertOperatorResult(env.DB, body || {});
        if (!result.ok) {
          return jsonResponse({ ok: false, error: result.error || "Payload invalido para resultado do operador." }, 400);
        }
        const record = await readRecord(env.DB, result.userId);
        return jsonResponse({ ok: true, record });
      }

      if (url.pathname === "/api/operator-results/delete" && request.method === "POST") {
        const body = await request.json();
        const result = await deleteOperatorResult(env.DB, body || {});
        if (!result.ok) {
          return jsonResponse({ ok: false, error: result.error || "Nao foi possivel excluir o lancamento." }, 400);
        }
        const record = await readRecord(env.DB, result.userId);
        return jsonResponse({ ok: true, deleted: result.deleted, record });
      }

      if (url.pathname === "/api/operator-results/delete-all" && request.method === "POST") {
        const body = await request.json().catch(() => ({}));
        if (!body?.confirm) {
          return jsonResponse({ ok: false, error: "Confirme a exclusao em lote para continuar." }, 400);
        }
        const result = await deleteAllOperatorResults(env.DB);
        return jsonResponse({ ok: true, deleted: result.deleted });
      }

      if (url.pathname === "/api/operator-results/bulk" && request.method === "POST") {
        const body = await request.json();
        const items = Array.isArray(body?.items) ? body.items : [];
        if (!items.length) {
          return jsonResponse({ ok: false, error: "Envie ao menos um item para importacao em lote." }, 400);
        }

        let imported = 0;
        const errors = [];
        for (let index = 0; index < items.length; index += 1) {
          const result = await upsertOperatorResult(env.DB, items[index]);
          if (result.ok) {
            imported += 1;
          } else {
            errors.push({ index, error: result.error || "Linha invalida" });
          }
        }

        return jsonResponse({
          ok: true,
          imported,
          failed: errors.length,
          errors
        });
      }

      if (url.pathname === "/api/import/remove-by-sheet" && request.method === "POST") {
        const body = await request.json();
        const items = Array.isArray(body?.items) ? body.items : [];
        if (!items.length) {
          return jsonResponse({ ok: false, error: "Envie ao menos um item para remocao em lote." }, 400);
        }

        let removed = 0;
        const errors = [];
        for (let index = 0; index < items.length; index += 1) {
          const result = await deleteOperatorResult(env.DB, items[index]);
          if (result.ok) {
            removed += Number(result.deleted || 1);
          } else {
            errors.push({ index, error: result.error || "Linha invalida ou nao encontrada" });
          }
        }

        return jsonResponse({
          ok: true,
          removed,
          failed: errors.length,
          errors
        });
      }

      if (url.pathname === "/api/import/upload-and-process" && request.method === "POST") {
        if (!env.IMPORTS_BUCKET) {
          return jsonResponse({ ok: false, error: "Binding R2 IMPORTS_BUCKET nao configurado." }, 500);
        }

        const body = await request.json();
        const items = Array.isArray(body?.items) ? body.items : [];
        const fileBase64 = String(body?.fileBase64 || "");
        const fileName = safeFileName(body?.fileName || "import.xlsx");
        const mimeType = String(body?.mimeType || "application/octet-stream");
        if (!items.length) {
          return jsonResponse({ ok: false, error: "Nenhum item valido enviado para importacao." }, 400);
        }
        if (!fileBase64) {
          return jsonResponse({ ok: false, error: "Arquivo da planilha nao enviado." }, 400);
        }

        const importKey = `imports/${Date.now()}-${crypto.randomUUID()}-${fileName}`;
        let imported = 0;
        const errors = [];
        let uploadStored = false;
        let removedFromR2 = false;

        try {
          const bytes = decodeBase64ToBytes(fileBase64);
          await env.IMPORTS_BUCKET.put(importKey, bytes, {
            httpMetadata: { contentType: mimeType }
          });
          uploadStored = true;

          for (let index = 0; index < items.length; index += 1) {
            const result = await upsertOperatorResult(env.DB, items[index]);
            if (result.ok) {
              imported += 1;
            } else {
              errors.push({ index, error: result.error || "Linha invalida" });
            }
          }
        } finally {
          if (uploadStored) {
            await env.IMPORTS_BUCKET.delete(importKey);
            removedFromR2 = true;
          }
        }

        return jsonResponse({
          ok: true,
          imported,
          failed: errors.length,
          errors,
          importKey,
          removedFromR2
        });
      }

      return jsonResponse({ ok: false, error: "Rota nao encontrada." }, 404);
    } catch (error) {
      return jsonResponse(
        { ok: false, error: "Falha interna na API.", detail: String(error?.message || error) },
        500
      );
    }
  }
};
