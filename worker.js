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

function sanitizeOperationType(value) {
  return String(value || "").trim().toLowerCase() === "0800" ? "0800" : "nuvidio";
}

function sanitizeR2Kind(value) {
  const kind = String(value || "").trim().toLowerCase();
  if (kind === "parsed") return "parsed";
  return "source";
}

function buildR2BaseKey(operationType, r2Kind, fileName) {
  const safeName = safeFileName(fileName || (r2Kind === "parsed" ? "base.parsed.csv" : "base.xlsx"));
  if (r2Kind === "parsed") {
    return `bases/${operationType}/parsed/${Date.now()}-${crypto.randomUUID()}-${safeName}`;
  }
  return `bases/${operationType}/source/${Date.now()}-${crypto.randomUUID()}-${safeName}`;
}

function getAuthBase(env) {
  return String(env.AUTH_API_BASE || env.CENTRAL_AUTH_BASE || "").trim().replace(/\/+$/, "");
}

function normalizeLooseKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]/g, "")
    .trim()
    .toLowerCase();
}

function collectStringFields(source, bucket = []) {
  if (!source || typeof source !== "object") return bucket;
  for (const [key, value] of Object.entries(source)) {
    if (typeof value === "string" && value.trim()) {
      bucket.push({ key: normalizeLooseKey(key), value: value.trim() });
      continue;
    }
    if (value && typeof value === "object" && !Array.isArray(value)) {
      collectStringFields(value, bucket);
    }
  }
  return bucket;
}

function findUserFieldByHeuristics(source, matcher) {
  const entries = collectStringFields(source, []);
  const found = entries.find((entry) => matcher(entry.key, entry.value));
  return found?.value || "";
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

function normalizeOperationType(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "0800") return "0800";
  if (raw === "nuvidio") return "nuvidio";
  return "nuvidio";
}

async function migrateResultsTableToOperationGranularity(db) {
  const schema = await db
    .prepare("SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'operator_results_daily' LIMIT 1")
    .first();
  const tableSql = String(schema?.sql || "");
  const usesOldUnique = tableSql.includes("UNIQUE(user_id, result_date)") && !tableSql.includes("UNIQUE(user_id, result_date, operation_type)");
  if (!usesOldUnique) return;

  await db.prepare("DROP TABLE IF EXISTS operator_results_daily_v2").run();
  await db.prepare(
    `CREATE TABLE operator_results_daily_v2 (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      user_name TEXT NOT NULL DEFAULT '',
      username TEXT NOT NULL DEFAULT '',
      result_date TEXT NOT NULL,
      operation_type TEXT NOT NULL DEFAULT '',
      approved_count REAL NOT NULL DEFAULT 0,
      reproved_count REAL NOT NULL DEFAULT 0,
      pending_count REAL NOT NULL DEFAULT 0,
      no_action_count REAL NOT NULL DEFAULT 0,
      production_total REAL NOT NULL DEFAULT 0,
      effectiveness REAL NOT NULL DEFAULT 0,
      quality_score REAL NOT NULL DEFAULT 0,
      updated_by_id TEXT NOT NULL DEFAULT '',
      updated_by_name TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(user_id, result_date, operation_type)
    )`
  ).run();

  await db.prepare(
    `INSERT INTO operator_results_daily_v2 (
      id, user_id, user_name, username, result_date, operation_type,
      approved_count, reproved_count, pending_count, no_action_count,
      production_total, effectiveness, quality_score,
      updated_by_id, updated_by_name, updated_at, created_at
    )
    SELECT
      (user_id || '__' || result_date || '__' || lower(CASE WHEN trim(operation_type) = '' THEN 'nuvidio' ELSE operation_type END)) AS id,
      user_id,
      user_name,
      username,
      result_date,
      lower(CASE WHEN trim(operation_type) = '' THEN 'nuvidio' ELSE operation_type END) AS operation_type,
      COALESCE(approved_count, 0),
      COALESCE(reproved_count, 0),
      COALESCE(pending_count, 0),
      COALESCE(no_action_count, 0),
      COALESCE(production_total, 0),
      COALESCE(effectiveness, 0),
      COALESCE(quality_score, 0),
      COALESCE(updated_by_id, ''),
      COALESCE(updated_by_name, ''),
      COALESCE(updated_at, datetime('now')),
      COALESCE(created_at, datetime('now'))
    FROM operator_results_daily`
  ).run();

  await db.prepare("DROP TABLE operator_results_daily").run();
  await db.prepare("ALTER TABLE operator_results_daily_v2 RENAME TO operator_results_daily").run();
}

async function ensureResultsTable(db) {
  await db.prepare(
    `CREATE TABLE IF NOT EXISTS operator_results_daily (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      user_name TEXT NOT NULL DEFAULT '',
      username TEXT NOT NULL DEFAULT '',
      result_date TEXT NOT NULL,
      operation_type TEXT NOT NULL DEFAULT '',
      approved_count REAL NOT NULL DEFAULT 0,
      reproved_count REAL NOT NULL DEFAULT 0,
      pending_count REAL NOT NULL DEFAULT 0,
      no_action_count REAL NOT NULL DEFAULT 0,
      production_total REAL NOT NULL DEFAULT 0,
      effectiveness REAL NOT NULL DEFAULT 0,
      quality_score REAL NOT NULL DEFAULT 0,
      updated_by_id TEXT NOT NULL DEFAULT '',
      updated_by_name TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(user_id, result_date, operation_type)
    )`
  ).run();
  try {
    await db.prepare("ALTER TABLE operator_results_daily ADD COLUMN operation_type TEXT NOT NULL DEFAULT ''").run();
  } catch {}
  try {
    await db.prepare("ALTER TABLE operator_results_daily ADD COLUMN approved_count REAL NOT NULL DEFAULT 0").run();
  } catch {}
  try {
    await db.prepare("ALTER TABLE operator_results_daily ADD COLUMN reproved_count REAL NOT NULL DEFAULT 0").run();
  } catch {}
  try {
    await db.prepare("ALTER TABLE operator_results_daily ADD COLUMN pending_count REAL NOT NULL DEFAULT 0").run();
  } catch {}
  try {
    await db.prepare("ALTER TABLE operator_results_daily ADD COLUMN no_action_count REAL NOT NULL DEFAULT 0").run();
  } catch {}
  await migrateResultsTableToOperationGranularity(db);
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

async function ensureR2InsightsSnapshotTable(db) {
  await db.prepare(
    `CREATE TABLE IF NOT EXISTS r2_insights_snapshot (
      snapshot_key TEXT PRIMARY KEY,
      views_json TEXT NOT NULL DEFAULT '{}',
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`
  ).run();
}

async function readR2InsightsSnapshot(db) {
  await ensureR2InsightsSnapshotTable(db);
  const row = await db
    .prepare("SELECT views_json, updated_at FROM r2_insights_snapshot WHERE snapshot_key = ? LIMIT 1")
    .bind("latest")
    .first();
  if (!row?.views_json) return null;
  try {
    const views = JSON.parse(String(row.views_json || "{}"));
    if (!views || typeof views !== "object") return null;
    return {
      views,
      updatedAt: String(row.updated_at || "")
    };
  } catch {
    return null;
  }
}

async function writeR2InsightsSnapshot(db, views) {
  await ensureR2InsightsSnapshotTable(db);
  const payload = views && typeof views === "object" ? views : {};
  const now = new Date().toISOString();
  await db
    .prepare(
      `INSERT INTO r2_insights_snapshot (snapshot_key, views_json, updated_at)
       VALUES (?, ?, ?)
       ON CONFLICT(snapshot_key) DO UPDATE SET
         views_json = excluded.views_json,
         updated_at = excluded.updated_at`
    )
    .bind("latest", JSON.stringify(payload), now)
    .run();
  return { ok: true, updatedAt: now };
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
    operationType: String(row.operation_type || ""),
    approvedCount: Number(row.approved_count || 0),
    reprovedCount: Number(row.reproved_count || 0),
    pendingCount: Number(row.pending_count || 0),
    noActionCount: Number(row.no_action_count || 0),
    productionTotal: Number(row.production_total || 0),
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
  let nuvidioUsername = String(
    user?.nuvidioUsername ??
    user?.usuarioNuvidio ??
    user?.usernameNuvidio ??
    user?.nuvidio_user ??
    user?.nuvidio_login ??
    user?.usuario_nuvidio ??
    ""
  ).trim();
  let line0800Username = String(
    user?.line0800Username ??
    user?.usuario0800 ??
    user?.username0800 ??
    user?.["0800Username"] ??
    user?.["0800_user"] ??
    user?.["0800_login"] ??
    user?.usuario_0800 ??
    ""
  ).trim();

  if (!nuvidioUsername) {
    nuvidioUsername = findUserFieldByHeuristics(user, (key) => (
      key.includes("nuvidio") && (key.includes("usuario") || key.includes("username") || key.includes("login") || key.includes("user"))
    ));
  }
  if (!line0800Username) {
    line0800Username = findUserFieldByHeuristics(user, (key) => (
      key.includes("0800") && (key.includes("usuario") || key.includes("username") || key.includes("login") || key.includes("user"))
    ));
  }
  return {
    id: String(user?.id || ""),
    name: String(user?.name || ""),
    username: String(user?.username || ""),
    nuvidioUsername,
    line0800Username,
    role: String(user?.role || "operador"),
    accessLevel: String(user?.accessLevel || "")
  };
}

async function resolveOperators(env) {
  const payload = await fetchCentral(env, "/api/users", { method: "GET" });
  const users = Array.isArray(payload?.users) ? payload.users : [];
  return users
    .map(sanitizeUser)
    .filter((user) => user.id && user.role !== "gestor")
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt-BR"));
}

async function readRecord(db, userId) {
  await ensureResultsTable(db);
  const rows = await db
    .prepare("SELECT * FROM operator_results_daily WHERE user_id = ? ORDER BY result_date ASC, updated_at ASC")
    .bind(userId)
    .all();
  return buildRecord(userId, rows.results || []);
}

async function upsertOperatorResult(db, payload) {
  const userId = String(payload?.userId || "").trim();
  const userName = String(payload?.userName || "").trim();
  const username = String(payload?.username || "").trim();
  const date = normalizeDateKey(payload?.date);
  const operationType = normalizeOperationType(payload?.operationType);
  const approvedCount = Number(payload?.approvedCount ?? 0);
  const reprovedCount = Number(payload?.reprovedCount ?? 0);
  const pendingCount = Number(payload?.pendingCount ?? 0);
  const noActionCount = Number(payload?.noActionCount ?? 0);
  const productionTotal = Number(payload?.productionTotal);
  const effectiveness = Number(payload?.effectiveness);
  const qualityScore = Number(payload?.qualityScore);
  const updatedById = String(payload?.updatedById || "").trim();
  const updatedByName = String(payload?.updatedByName || "").trim();

  if (
    !userId ||
    !date ||
    !Number.isFinite(productionTotal) ||
    !Number.isFinite(effectiveness) ||
    !Number.isFinite(qualityScore) ||
    !Number.isFinite(approvedCount) ||
    !Number.isFinite(reprovedCount) ||
    !Number.isFinite(pendingCount) ||
    !Number.isFinite(noActionCount)
  ) {
    return { ok: false, error: "Payload invalido para resultado do operador." };
  }

  await ensureResultsTable(db);
  await db.prepare(
    `INSERT INTO operator_results_daily (
      id, user_id, user_name, username, result_date, operation_type,
      approved_count, reproved_count, pending_count, no_action_count,
      production_total, effectiveness, quality_score,
      updated_by_id, updated_by_name, updated_at, created_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(user_id, result_date, operation_type) DO UPDATE SET
      user_name = excluded.user_name,
      username = excluded.username,
      approved_count = excluded.approved_count,
      reproved_count = excluded.reproved_count,
      pending_count = excluded.pending_count,
      no_action_count = excluded.no_action_count,
      production_total = excluded.production_total,
      effectiveness = excluded.effectiveness,
      quality_score = excluded.quality_score,
      updated_by_id = excluded.updated_by_id,
      updated_by_name = excluded.updated_by_name,
      updated_at = excluded.updated_at`
  )
    .bind(
      `${userId}__${date}__${operationType}`,
      userId,
      userName,
      username,
      date,
      operationType,
      approvedCount,
      reprovedCount,
      pendingCount,
      noActionCount,
      productionTotal,
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
  const operationTypeRaw = String(payload?.operationType || "").trim();
  const operationType = operationTypeRaw ? normalizeOperationType(operationTypeRaw) : "";
  if (!userId || !date) {
    return { ok: false, error: "Informe userId e date para excluir o lancamento." };
  }

  await ensureResultsTable(db);
  const result = operationType
    ? await db
      .prepare("DELETE FROM operator_results_daily WHERE user_id = ? AND result_date = ? AND operation_type = ?")
      .bind(userId, date, operationType)
      .run()
    : await db
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
  const rows = await db
    .prepare("SELECT * FROM operator_results_daily ORDER BY user_id ASC, result_date ASC, updated_at ASC")
    .all();

  const byUser = new Map();
  for (const row of rows.results || []) {
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

function normalizeCsvText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function parseCsvSemicolon(content) {
  const source = String(content || "").replace(/^\uFEFF/, "");
  const lines = source.split(/\r?\n/).filter((line) => line.trim().length);
  if (!lines.length) return [];

  const parseLine = (line) => {
    const out = [];
    let current = "";
    let quoted = false;
    for (let i = 0; i < line.length; i += 1) {
      const char = line[i];
      if (char === "\"") {
        const next = line[i + 1];
        if (quoted && next === "\"") {
          current += "\"";
          i += 1;
          continue;
        }
        quoted = !quoted;
        continue;
      }
      if (char === ";" && !quoted) {
        out.push(current);
        current = "";
        continue;
      }
      current += char;
    }
    out.push(current);
    return out.map((part) => String(part || "").trim());
  };

  const header = parseLine(lines[0]);
  const rows = [];
  for (let i = 1; i < lines.length; i += 1) {
    const cells = parseLine(lines[i]);
    if (!cells.length) continue;
    const row = {};
    for (let col = 0; col < header.length; col += 1) {
      row[header[col]] = String(cells[col] || "").trim();
    }
    rows.push(row);
  }
  return rows;
}

function getRowValueByHeader(row, aliases = []) {
  const keys = Object.keys(row || {});
  for (const alias of aliases) {
    const normAlias = normalizeCsvText(alias);
    const found = keys.find((key) => {
      const normKey = normalizeCsvText(key);
      if (!normKey || !normAlias) return false;
      if (normKey === normAlias) return true;
      if (normKey.includes(normAlias) || normAlias.includes(normKey)) return true;
      return false;
    });
    if (found) return String(row[found] || "").trim();
  }
  return "";
}

function resolveOperatorNameFromRow(row, aliases = []) {
  const explicit = getRowValueByHeader(row, aliases);
  if (explicit) return explicit;
  const keys = Object.keys(row || {});
  const guessedKey = keys.find((key) => {
    const norm = normalizeCsvText(key);
    return (
      norm.includes("atendente") ||
      norm.includes("analista") ||
      norm.includes("operador") ||
      norm.includes("funcionario") ||
      norm.includes("usuario") ||
      norm.includes("login") ||
      norm.includes("emaildoatendente")
    );
  });
  if (!guessedKey) return "";
  const raw = String(row[guessedKey] || "").trim();
  if (!raw) return "";
  const emailMatch = raw.match(/^[^@\s]+@[^@\s]+\.[^@\s]+$/);
  if (emailMatch) return raw.split("@")[0];
  return raw;
}

function parseNumberLoose(value) {
  const raw = String(value || "").trim();
  if (!raw) return NaN;
  const normalized = raw.replace(/\./g, "").replace(",", ".");
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : NaN;
}

function parseDurationSeconds(value) {
  const raw = String(value || "").trim();
  if (!raw) return NaN;
  if (/^\d+(\.\d+)?$/.test(raw)) return Number(raw);
  const match = raw.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (!match) return NaN;
  const [, h, m, s] = match;
  return (Number(h) * 3600) + (Number(m) * 60) + Number(s);
}

function normalizeStatusBucket(value) {
  const text = normalizeCsvText(value);
  if (!text) return "semAcao";
  if (text.includes("aprovad")) return "aprovadas";
  if (text.includes("reprovad")) return "reprovadas";
  if (text.includes("pendenc")) return "pendenciadas";
  if (text.includes("sem acao") || text.includes("semacao")) return "semAcao";
  return "semAcao";
}

function pushTopCounter(map, key) {
  const normalized = String(key || "").trim();
  if (!normalized) return;
  map.set(normalized, Number(map.get(normalized) || 0) + 1);
}

function mapToTopList(counterMap, maxItems = 5) {
  return [...counterMap.entries()]
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, maxItems);
}

function aggregateNuvidio(rows) {
  const statuses = { aprovadas: 0, reprovadas: 0, semAcao: 0 };
  const subtags = new Map();
  const atendentes = new Map();
  let total = 0;
  let waitSum = 0;
  let waitCount = 0;
  let tmaSum = 0;
  let tmaCount = 0;

  for (const row of rows) {
    total += 1;
    const status = normalizeStatusBucket(getRowValueByHeader(row, ["Tag", "Motivo"]));
    if (status === "aprovadas") statuses.aprovadas += 1;
    else if (status === "reprovadas") statuses.reprovadas += 1;
    else statuses.semAcao += 1;

    const wait = parseNumberLoose(getRowValueByHeader(row, ["Espera em segundos"]));
    if (Number.isFinite(wait)) {
      waitSum += wait;
      waitCount += 1;
    }

    const tmaRaw = getRowValueByHeader(row, ["Duracao em segundos", "TMA"]);
    const tma = parseDurationSeconds(tmaRaw);
    if (Number.isFinite(tma)) {
      tmaSum += tma;
      tmaCount += 1;
    }

    pushTopCounter(subtags, getRowValueByHeader(row, ["Subtag"]));
    const atendente = resolveOperatorNameFromRow(row, [
      "Atendente",
      "Analista",
      "Usuario de Abertura da Ocorrencia",
      "Usuário de Abertura da Ocorrência",
      "Funcionario",
      "Funcionário",
      "Operador",
      "Usuario",
      "Usuário"
    ]);
    if (atendente) {
      atendentes.set(atendente, Number(atendentes.get(atendente) || 0) + 1);
    }
  }

  const producaoPorOperador = [...atendentes.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  return {
    totalAtendimentos: total,
    statuses,
    avgWaitSeconds: waitCount ? (waitSum / waitCount) : 0,
    avgTmaSeconds: tmaCount ? (tmaSum / tmaCount) : 0,
    topSubtags: mapToTopList(subtags, 5),
    topAtendentes: producaoPorOperador.slice(0, 5),
    producaoPorOperador
  };
}

function aggregate0800(rows) {
  const statuses = { aprovadas: 0, reprovadas: 0, pendenciadas: 0, semAcao: 0 };
  const subMotivos = new Map();
  const analistas = new Map();
  let total = 0;
  let daysSum = 0;
  let daysCount = 0;
  let fcrYes = 0;
  let fcrTotal = 0;

  for (const row of rows) {
    total += 1;
    const status = normalizeStatusBucket(getRowValueByHeader(row, ["Motivo", "Tag"]));
    if (status === "aprovadas") statuses.aprovadas += 1;
    else if (status === "reprovadas") statuses.reprovadas += 1;
    else if (status === "pendenciadas") statuses.pendenciadas += 1;
    else statuses.semAcao += 1;

    const days = parseNumberLoose(getRowValueByHeader(row, ["Dias Para Resolucao", "Dias Para Resolução"]));
    if (Number.isFinite(days)) {
      daysSum += days;
      daysCount += 1;
    }

    const fcr = normalizeCsvText(getRowValueByHeader(row, ["FCR (Sim / Nao)", "FCR (Sim / Não)", "FCR"]));
    if (fcr) {
      fcrTotal += 1;
      if (fcr.startsWith("sim")) fcrYes += 1;
    }

    pushTopCounter(subMotivos, getRowValueByHeader(row, ["Sub-Motivo", "Sub Motivo"]));
    pushTopCounter(analistas, resolveOperatorNameFromRow(row, [
      "Analista",
      "Usuario de Abertura da Ocorrencia",
      "Usuário de Abertura da Ocorrência",
      "Funcionario",
      "Funcionário",
      "Operador",
      "Usuario",
      "Usuário"
    ]));
  }

  const producaoPorOperador = [...analistas.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  return {
    totalOcorrencias: total,
    statuses,
    avgResolutionDays: daysCount ? (daysSum / daysCount) : 0,
    fcrSimRate: fcrTotal ? ((fcrYes * 100) / fcrTotal) : 0,
    topSubMotivos: mapToTopList(subMotivos, 5),
    topAnalistas: producaoPorOperador.slice(0, 5),
    producaoPorOperador
  };
}

async function listAllR2Objects(bucket) {
  const files = [];
  let cursor = undefined;
  let truncated = true;
  while (truncated) {
    const page = await bucket.list({ limit: 1000, cursor });
    files.push(...(page.objects || []));
    truncated = Boolean(page.truncated);
    cursor = page.cursor;
  }
  return files;
}

async function resolveR2SourceFiles(bucket) {
  const objects = await listAllR2Objects(bucket);
  const pickLatest = (matcher) => {
    return objects
      .filter((obj) => matcher(normalizeCsvText(obj.key || "")))
      .sort((a, b) => new Date(String(b.uploaded || 0)).getTime() - new Date(String(a.uploaded || 0)).getTime())[0] || null;
  };

  const line0800 = pickLatest((key) => (
    key.includes("bases/0800/parsed/") ||
    key.includes("bases/0800/") ||
    key.includes("detalhes do protocolo com ocorrencias") ||
    key.includes("detalhesdoprotocolo")
  ));
  const nuvidio = pickLatest((key) => (
    key.includes("bases/nuvidio/parsed/") ||
    key.includes("bases/nuvidio/") ||
    key.includes("todos-atendimentos") ||
    key.includes("todos atendimentos")
  ));
  return { line0800, nuvidio };
}

async function readCsvFromR2Object(bucket, objectMeta) {
  if (!objectMeta?.key) return [];
  const object = await bucket.get(objectMeta.key);
  if (!object) return [];
  const content = await object.text();
  return parseCsvSemicolon(content);
}

async function buildR2Insights(env) {
  const bucket = env.IMPORTS_BUCKET || env.RESULTS_BUCKET;
  if (!bucket) {
    return { views: null, sources: null };
  }

  const files = await resolveR2SourceFiles(bucket);
  const nuvidioRows = await readCsvFromR2Object(bucket, files.nuvidio);
  const line0800Rows = await readCsvFromR2Object(bucket, files.line0800);

  return {
    views: {
      nuvidio: aggregateNuvidio(nuvidioRows),
      line0800: aggregate0800(line0800Rows)
    },
    sources: {
      nuvidio: files.nuvidio?.key || "",
      line0800: files.line0800?.key || ""
    }
  };
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

        const payload = await fetchCentral(env, "/api/login", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ username, password })
        });

        return jsonResponse({ ok: true, user: sanitizeUser(payload.user) });
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

      if (url.pathname === "/api/r2-base-upload/multipart/init" && request.method === "POST") {
        const bucket = env.IMPORTS_BUCKET || env.RESULTS_BUCKET;
        if (!bucket) {
          return jsonResponse({ ok: false, error: "Binding R2 nao configurado." }, 500);
        }

        const body = await request.json().catch(() => ({}));
        const operationType = sanitizeOperationType(body?.operationType || "nuvidio");
        const r2Kind = sanitizeR2Kind(body?.kind || "source");
        const fileName = String(body?.fileName || "").trim() || (r2Kind === "parsed" ? "base.parsed.csv" : "base.xlsx");
        const contentType = String(body?.contentType || (r2Kind === "parsed" ? "text/csv" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")).trim();
        const key = buildR2BaseKey(operationType, r2Kind, fileName);

        const upload = await bucket.createMultipartUpload(key, {
          httpMetadata: { contentType },
          customMetadata: {
            source: "portal",
            operationType,
            kind: r2Kind
          }
        });

        return jsonResponse({
          ok: true,
          key,
          uploadId: upload.uploadId,
          operationType,
          kind: r2Kind
        });
      }

      if (url.pathname === "/api/r2-base-upload/multipart/part" && request.method === "POST") {
        const bucket = env.IMPORTS_BUCKET || env.RESULTS_BUCKET;
        if (!bucket) {
          return jsonResponse({ ok: false, error: "Binding R2 nao configurado." }, 500);
        }
        const key = String(url.searchParams.get("key") || "").trim();
        const uploadId = String(url.searchParams.get("uploadId") || "").trim();
        const partNumber = Number(url.searchParams.get("partNumber"));
        if (!key || !uploadId || !Number.isInteger(partNumber) || partNumber < 1) {
          return jsonResponse({ ok: false, error: "Informe key, uploadId e partNumber validos." }, 400);
        }
        if (!request.body) {
          return jsonResponse({ ok: false, error: "Parte vazia." }, 400);
        }

        const upload = bucket.resumeMultipartUpload(key, uploadId);
        const part = await upload.uploadPart(partNumber, request.body);
        return jsonResponse({
          ok: true,
          key,
          uploadId,
          partNumber,
          etag: String(part?.etag || "")
        });
      }

      if (url.pathname === "/api/r2-base-upload/multipart/complete" && request.method === "POST") {
        const bucket = env.IMPORTS_BUCKET || env.RESULTS_BUCKET;
        if (!bucket) {
          return jsonResponse({ ok: false, error: "Binding R2 nao configurado." }, 500);
        }
        const body = await request.json().catch(() => ({}));
        const key = String(body?.key || "").trim();
        const uploadId = String(body?.uploadId || "").trim();
        const partsRaw = Array.isArray(body?.parts) ? body.parts : [];
        const parts = partsRaw
          .map((item) => ({
            partNumber: Number(item?.partNumber),
            etag: String(item?.etag || "")
          }))
          .filter((item) => Number.isInteger(item.partNumber) && item.partNumber > 0 && item.etag);
        if (!key || !uploadId || !parts.length) {
          return jsonResponse({ ok: false, error: "Informe key, uploadId e parts para concluir upload multipart." }, 400);
        }

        const upload = bucket.resumeMultipartUpload(key, uploadId);
        await upload.complete(parts);
        return jsonResponse({ ok: true, key, uploadId, parts: parts.length });
      }

      if (url.pathname === "/api/r2-base-upload" && request.method === "POST") {
        const bucket = env.IMPORTS_BUCKET || env.RESULTS_BUCKET;
        if (!bucket) {
          return jsonResponse({ ok: false, error: "Binding R2 nao configurado." }, 500);
        }

        const operationType = sanitizeOperationType(
          request.headers.get("x-operation-type") || url.searchParams.get("operationType") || "nuvidio"
        );
        const r2Kind = sanitizeR2Kind(
          request.headers.get("x-r2-kind") || url.searchParams.get("kind") || "source"
        );
        const fileNameRaw = request.headers.get("x-file-name") || url.searchParams.get("fileName") || "base.csv";
        const fileName = safeFileName(fileNameRaw || "base.csv");
        const contentType = String(request.headers.get("content-type") || "text/csv").trim() || "text/csv";
        const key = buildR2BaseKey(operationType, r2Kind, fileName);

        if (!request.body) {
          return jsonResponse({ ok: false, error: "Arquivo nao enviado no corpo da requisicao." }, 400);
        }

        await bucket.put(key, request.body, {
          httpMetadata: { contentType },
          customMetadata: {
            source: "portal",
            operationType,
            kind: r2Kind
          }
        });

        return jsonResponse({
          ok: true,
          key,
          operationType,
          kind: r2Kind
        });
      }

      if (url.pathname === "/api/r2-insights" && request.method === "GET") {
        const snapshot = await readR2InsightsSnapshot(env.DB);
        if (snapshot?.views) {
          return jsonResponse({
            ok: true,
            views: snapshot.views,
            sources: { mode: "snapshot" },
            updatedAt: snapshot.updatedAt
          });
        }
        const payload = await buildR2Insights(env);
        if (payload?.views) {
          await writeR2InsightsSnapshot(env.DB, payload.views);
        }
        return jsonResponse({ ok: true, views: payload.views, sources: payload.sources });
      }

      if (url.pathname === "/api/r2-insights/rebuild" && request.method === "POST") {
        const payload = await buildR2Insights(env);
        if (payload?.views) {
          await writeR2InsightsSnapshot(env.DB, payload.views);
        }
        return jsonResponse({ ok: true, views: payload.views, sources: payload.sources });
      }

      if (url.pathname === "/api/r2-insights/snapshot" && request.method === "POST") {
        const body = await request.json().catch(() => ({}));
        const views = body?.views;
        if (!views || typeof views !== "object") {
          return jsonResponse({ ok: false, error: "Envie views validas para atualizar o snapshot." }, 400);
        }
        const result = await writeR2InsightsSnapshot(env.DB, views);
        return jsonResponse({ ok: true, updatedAt: result.updatedAt });
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
