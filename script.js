const state = {
  user: null,
  appSettings: {
    maintenance_for_operators: false,
    maintenance_message: "Portal em manutenção. Tente novamente em instantes.",
    metric_rules: {
      production: { red_max: 70, amber_max: 100 },
      effectiveness: { red_max: 70, amber_max: 90 },
      quality: { red_max: 70, amber_max: 90 },
    },
    alert_rules: {
      production_nuvidio: { critical_min: 70 },
      production_0800: { critical_min: 70 },
      effectiveness_0800: { critical_min: 70 },
      effectiveness_nuvidio: { critical_min: 70 },
      quality: { critical_min: 70 },
    },
  },
  dayTopCache: {},
  analysisTopCache: {},
  route: "overview",
  overview: null,
  analysis: null,
  alerts: null,
  history: null,
  reportDataset: null,
  users: [],
  flash: null,
  forcePasswordChange: false,
  theme: "dark",
  filters: {
    today: new Date().toISOString().slice(0, 10),
    start: "",
    end: "",
    historyUserId: "",
    historyQuery: "",
    analysisUserId: "all",
    usersQuery: "",
    reportsType: "consolidado",
    reportsView: "detalhada",
    reportsSector: "all",
    detailedPage: 1,
    operation: "all",
  },
  reportSorts: {
    consolidado: { column: "date", direction: "desc" },
    operacional: { column: "date", direction: "desc" },
    qualidade: { column: "reference_month", direction: "desc" },
    ofensores: { column: "alert_score", direction: "asc" },
  },
};

const DEFAULT_PASSWORD_HINT = "Trocar@01";
const app = document.getElementById("app");
const reportDatasetCache = {
  key: "",
  value: null,
};
const bootLoader = document.createElement("div");
bootLoader.className = "boot-loader";
bootLoader.innerHTML = `
  <div class="boot-loader-card">
    <div class="boot-loader-spinner"></div>
    <strong id="boot-loader-message">Carregando portal...</strong>
  </div>
`;
document.body.appendChild(bootLoader);
let maintenanceWatcher = null;
let maintenanceWatcherInFlight = false;
let maintenanceVisibilityHandlerBound = false;

function setBootLoaderMessage(message) {
  const label = document.getElementById("boot-loader-message");
  if (label) label.textContent = message || "Carregando portal...";
}

function showBootLoader() {
  bootLoader.classList.add("visible");
}

function hideBootLoader() {
  bootLoader.classList.remove("visible");
}

function stopMaintenanceWatcher() {
  if (maintenanceWatcher) {
    clearInterval(maintenanceWatcher);
    maintenanceWatcher = null;
  }
  maintenanceWatcherInFlight = false;
}

async function checkMaintenanceNow() {
  if (!state.user || isManager() || maintenanceWatcherInFlight) return;
  maintenanceWatcherInFlight = true;
  try {
    const auth = await api("/api/auth/me");
    if (auth.app_settings) {
      const wasActive = Boolean(state.appSettings?.maintenance_for_operators);
      state.appSettings = auth.app_settings;
      const isActive = Boolean(state.appSettings.maintenance_for_operators);
      if (!wasActive && isActive) {
        setFlash("error", state.appSettings.maintenance_message || "Portal em manutenção para operadores.");
        render();
      }
    }
  } catch {
    // silencioso para nao poluir a UX
  } finally {
    maintenanceWatcherInFlight = false;
  }
}

function startMaintenanceWatcher() {
  stopMaintenanceWatcher();
  if (!state.user || isManager()) return;
  maintenanceWatcher = setInterval(() => {
    checkMaintenanceNow();
  }, 5000);
  checkMaintenanceNow();
  if (!maintenanceVisibilityHandlerBound) {
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) checkMaintenanceNow();
    });
    window.addEventListener("focus", () => {
      checkMaintenanceNow();
    });
    maintenanceVisibilityHandlerBound = true;
  }
}

function brandLogoSrc() {
  return window.__brandLogo || "/logos_KR-02.png?v=20260429";
}

function applyTheme() {
  document.body.classList.toggle("theme-contrast", state.theme === "contrast");
}

function setFlash(type, message, details = []) {
  state.flash = {
    type,
    message,
    details: Array.isArray(details) ? details.filter(Boolean) : [],
  };
  render();
}

function clearFlash() {
  state.flash = null;
}

function refreshDashboardInBackground(successMessage = "") {
  if (successMessage) setFlash("success", successMessage);
  loadBootstrap()
    .then(async () => {
      if (state.route === "alerts") {
        await loadAlerts();
      }
      render();
    })
    .catch((error) => setFlash("error", error.message || "Falha ao atualizar os dados."));
}

function setButtonProcessing(button, processing, processingText = "Processando...") {
  if (!button) return () => {};
  const originalText = button.textContent;
  if (processing) {
    button.disabled = true;
    button.classList.add("is-loading");
    if (processingText) button.textContent = processingText;
  } else {
    button.disabled = false;
    button.classList.remove("is-loading");
    button.textContent = originalText;
  }
  return () => {
    button.disabled = false;
    button.classList.remove("is-loading");
    button.textContent = originalText;
  };
}

function enhancePasswordFields(root = document) {
  root.querySelectorAll('input[type="password"]').forEach((input) => {
    if (input.dataset.passwordEnhanced === "true") return;
    const parent = input.parentElement;
    if (!parent) return;
    input.dataset.passwordEnhanced = "true";
    const wrapper = document.createElement("div");
    wrapper.className = "password-input-wrap";
    parent.insertBefore(wrapper, input);
    wrapper.appendChild(input);
    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "password-toggle-btn";
    toggle.setAttribute("aria-label", "Mostrar senha");
    toggle.innerHTML = `
      <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <path d="M2.2 12c2.1-3.9 5.7-6.1 9.8-6.1s7.7 2.2 9.8 6.1c-2.1 3.9-5.7 6.1-9.8 6.1S4.3 15.9 2.2 12Z"></path>
        <circle cx="12" cy="12" r="3.2"></circle>
      </svg>`;
    toggle.addEventListener("click", () => {
      const visible = input.type === "text";
      input.type = visible ? "password" : "text";
      toggle.classList.toggle("is-active", !visible);
      toggle.setAttribute("aria-label", visible ? "Mostrar senha" : "Ocultar senha");
      input.focus({ preventScroll: true });
      const length = input.value.length;
      input.setSelectionRange(length, length);
    });
    wrapper.appendChild(toggle);
  });
}

async function api(path, options = {}) {
  const headers = { ...(options.headers || {}) };
  if (!(options.body instanceof FormData) && !headers["Content-Type"]) headers["Content-Type"] = "application/json";
  const response = await fetch(path, { headers, credentials: "same-origin", ...options });
  const contentType = response.headers.get("content-type") || "";
  const data = contentType.includes("application/json")
    ? await response.json()
    : { error: await response.text() };
  if (!response.ok) {
    const message = [data.error, data.details].filter(Boolean).join(": ");
    throw new Error(message || "Erro inesperado");
  }
  return data;
}

async function downloadFile(url, fallbackName) {
  const response = await fetch(url, { credentials: "same-origin" });
  const contentType = response.headers.get("content-type") || "";
  if (!response.ok) {
    const data = contentType.includes("application/json")
      ? await response.json()
      : { error: await response.text() };
    const message = [data.error, data.details].filter(Boolean).join(": ");
    throw new Error(message || "Falha ao baixar o arquivo.");
  }
  const blob = await response.blob();
  const disposition = response.headers.get("content-disposition") || "";
  const match = disposition.match(/filename="?([^"]+)"?/i);
  const fileName = match?.[1] || fallbackName || "arquivo";
  const fileUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = fileUrl;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(fileUrl), 1000);
}

function esc(value) {
  return repairTextEncoding(String(value ?? ""))
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
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

function number(value) {
  return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 1 }).format(Number(value || 0));
}

function integer(value) {
  return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 }).format(Number(value || 0));
}

function percent(value) {
  return `${Math.round(Number(value || 0))}%`;
}

function metricRules() {
  return state.appSettings?.metric_rules || {
    production: { red_max: 70, amber_max: 100 },
    effectiveness: { red_max: 70, amber_max: 90 },
    quality: { red_max: 70, amber_max: 90 },
  };
}

function alertRules() {
  return state.appSettings?.alert_rules || {
    production_nuvidio: { critical_min: 70 },
    production_0800: { critical_min: 70 },
    effectiveness_0800: { critical_min: 70 },
    effectiveness_nuvidio: { critical_min: 70 },
    quality: { critical_min: 70 },
  };
}

function metricTone(value, metricType) {
  const numeric = Number(value || 0);
  const rules = metricRules()[metricType] || { red_max: 70, amber_max: 90 };
  const redMax = Number(rules.red_max);
  const amberMax = Number(rules.amber_max);
  if (numeric <= redMax) return "red";
  if (numeric <= amberMax) return "amber";
  return "green";
}

function initials(name) {
  return repairTextEncoding(String(name || "KR")).split(" ").filter(Boolean).slice(0, 2).map((item) => item[0]).join("").toUpperCase();
}

function normalizeUserPayload(user) {
  if (!user) return user;
  return {
    ...user,
    full_name: repairTextEncoding(user.full_name),
  };
}

function normalizeUsersPayload(users) {
  return Array.isArray(users) ? users.map(normalizeUserPayload) : [];
}

function average(values, options = {}) {
  const ignoreZero = Boolean(options.ignoreZero);
  const numericValues = values
    .map((value) => Number(value || 0))
    .filter((value) => Number.isFinite(value) && (!ignoreZero || value !== 0));
  if (!numericValues.length) return 0;
  return numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length;
}

function isSaturday(dateValue) {
  const raw = String(dateValue || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return false;
  return new Date(`${raw}T00:00:00`).getDay() === 6;
}

function isManager() {
  return state.user?.role === "manager";
}

function normalizeRoute(route, role = state.user?.role) {
  const value = String(route || "").trim();
  const allowed = role === "manager"
    ? ["overview", "analysis", "alerts", "history", "reports", "detailed", "admin"]
    : ["overview", "analysis", "history"];
  return allowed.includes(value) ? value : "overview";
}

function applyUserPreferences() {
  if (!state.user) return;
  state.theme = state.user.preferred_theme === "contrast" ? "contrast" : "dark";
  state.route = normalizeRoute(state.user.last_route, state.user.role);
}

async function saveUserPreferences(partial) {
  if (!state.user) return;
  const response = await api("/api/auth/preferences", {
    method: "PATCH",
    body: JSON.stringify(partial),
  });
  state.user = normalizeUserPayload(response.user);
}

function enforceOperatorScope() {
  if (!state.user || isManager()) return;
  state.filters.historyUserId = String(state.user.id);
  state.filters.analysisUserId = String(state.user.id);
  state.filters.historyQuery = state.user.full_name;
}

function sortUsersByName(users = []) {
  return [...users].sort((a, b) =>
    repairTextEncoding(String(a?.full_name || "")).localeCompare(
      repairTextEncoding(String(b?.full_name || "")),
      "pt-BR",
      { sensitivity: "base" },
    ),
  );
}

function getOperatorUsers() {
  return sortUsersByName(state.users.filter((user) => user.role === "operator" && user.is_active));
}

function getAnyOperatorUsers() {
  return sortUsersByName(state.users.filter((user) => user.role === "operator"));
}

function getFilteredUsers() {
  const query = repairTextEncoding(String(state.filters.usersQuery || ""))
    .trim()
    .toLocaleLowerCase("pt-BR");
  const sortedUsers = sortUsersByName(state.users);
  if (!query) return sortedUsers;
  return sortedUsers.filter((user) => {
    const fullName = repairTextEncoding(String(user.full_name || "")).toLocaleLowerCase("pt-BR");
    const login = repairTextEncoding(String(user.login || "")).toLocaleLowerCase("pt-BR");
    const role = user.role === "manager" ? "gestor" : "operador";
    return fullName.includes(query) || login.includes(query) || role.includes(query);
  });
}

function applyManagementUserFilter() {
  const query = repairTextEncoding(String(state.filters.usersQuery || ""))
    .trim()
    .toLocaleLowerCase("pt-BR");
  const rows = [...document.querySelectorAll("[data-management-user-search]")];
  const emptyState = document.getElementById("management-users-empty");
  if (!rows.length) return;
  let visible = 0;
  rows.forEach((row) => {
    const haystack = String(row.dataset.managementUserSearch || "").toLocaleLowerCase("pt-BR");
    const matches = !query || haystack.includes(query);
    row.hidden = !matches;
    if (matches) visible += 1;
  });
  if (emptyState) emptyState.hidden = visible > 0;
}

function ensureManagerUserFilters() {
  if (!isManager()) return;
  const operators = getOperatorUsers();
  if (state.filters.analysisUserId !== "all" && !operators.some((user) => String(user.id) === String(state.filters.analysisUserId))) {
    state.filters.analysisUserId = "all";
  }
  if (state.filters.historyUserId !== "all" && !operators.some((user) => String(user.id) === String(state.filters.historyUserId))) {
    state.filters.historyUserId = "all";
  }
  if (state.filters.historyUserId === "all") {
    state.filters.historyQuery = "";
  } else if (!state.filters.historyQuery || !operators.some((user) => user.full_name === state.filters.historyQuery)) {
    state.filters.historyQuery = getUserLabelById(state.filters.historyUserId) || "";
  }
}

function getUserLabelById(userId) {
  return getAnyOperatorUsers().find((user) => String(user.id) === String(userId))?.full_name || "";
}

function resolveHistoryUserId(query) {
  const normalized = String(query || "").trim().toLowerCase();
  if (!normalized) return "";
  if (normalized === "todos os operadores") return "all";
  const exact = getOperatorUsers().find((user) => user.full_name.trim().toLowerCase() === normalized);
  if (exact) return String(exact.id);
  const partial = getOperatorUsers().find((user) => user.full_name.trim().toLowerCase().includes(normalized));
  return partial ? String(partial.id) : "";
}

function shortDate(value) {
  const raw = String(value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    const [year, month, day] = raw.split("-");
    return `${day}/${month}`;
  }
  return raw;
}

function formatDateBr(value) {
  const raw = String(value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    const [year, month, day] = raw.split("-");
    return `${day}/${month}/${year}`;
  }
  return raw || "--";
}

function formatDateTimeBr(value) {
  const raw = String(value || "").trim();
  if (!raw) return "--";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return raw;
  return new Intl.DateTimeFormat("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(date);
}

function formatMonthLabel(value) {
  const raw = String(value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(raw)) return raw;
  const [year, month] = raw.split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);
  const label = new Intl.DateTimeFormat("pt-BR", {
    month: "long",
    year: "numeric",
  }).format(date);
  return label.charAt(0).toUpperCase() + label.slice(1);
}

function formatQualityScopeLabel(value) {
  const scope = String(value || "all").trim().toLowerCase();
  if (scope === "0800") return "0800";
  if (scope === "nuvidio") return "Nuvidio";
  return "Geral";
}

function normalizeQualityScope(value) {
  const scope = String(value || "all").trim().toLowerCase();
  if (scope === "0800") return "0800";
  if (scope === "nuvidio") return "nuvidio";
  return "all";
}

function monthOptions(selectedValue = "") {
  const months = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
  ];
  return months
    .map((label, index) => {
      const value = String(index + 1).padStart(2, "0");
      return `<option value="${value}" ${value === selectedValue ? "selected" : ""}>${label}</option>`;
    })
    .join("");
}

function yearOptions(selectedYear = "") {
  const currentYear = new Date().getFullYear();
  const years = [];
  for (let year = currentYear - 2; year <= currentYear + 2; year += 1) {
    years.push(`<option value="${year}" ${String(year) === String(selectedYear) ? "selected" : ""}>${year}</option>`);
  }
  return years.join("");
}

function percentageDelta(previous, current) {
  const base = Number(previous || 0);
  const value = Number(current || 0);
  if (!base) return value ? 100 : 0;
  return ((value - base) / base) * 100;
}

function trendLabel(value) {
  const text = `${value > 0 ? "+" : ""}${value.toFixed(1)}%`;
  return `<span class="pill ${value >= 0 ? "green" : "red"}">${text}</span>`;
}

function combinedScore(item) {
  return Number(item.avg_production || 0) * 0.35 + Number(item.effectiveness || 0) * 0.4 + Number(item.quality || 0) * 6;
}

function findQualityForDate(metricDate) {
  const month = String(metricDate).slice(0, 7);
  return state.history?.quality?.find((item) => item.reference_month === month)?.score || 0;
}

function calcOperationEffectiveness(row, operation) {
  if (operation === "0800") {
    const actionable = Number(row.calls_0800_approved) + Number(row.calls_0800_rejected) + Number(row.calls_0800_pending);
    const total = actionable + Number(row.calls_0800_no_action);
    return total ? (actionable / total) * 100 : 0;
  }
  const actionable = Number(row.calls_nuvidio_approved) + Number(row.calls_nuvidio_rejected);
  const total = actionable + Number(row.calls_nuvidio_no_action) + Number(row.calls_nuvidio_empty || 0);
  return total ? (actionable / total) * 100 : 0;
}

function getScopedHistory() {
  const rows = state.history?.history || [];
  const qualityRows = state.history?.quality || [];
  const operation = state.filters.operation;
  const qualityByMonth = new Map(qualityRows.map((item) => [item.reference_month, item.score || 0]));
  const metricRows = rows.flatMap((row) => {
    const list = [];
    const qualityScore = qualityByMonth.get(String(row.metric_date || "").slice(0, 7)) || 0;
    if (operation === "all" || operation === "0800") {
      list.push({
        entryType: "metric",
        metricId: row.id,
        userId: row.user_id,
        date: row.metric_date,
        dateLabel: formatDateBr(row.metric_date),
        operation: "0800",
        production: Number(row.production_0800 || 0),
        production_0800: Number(row.production_0800 || 0),
        production_nuvidio: Number(row.production_nuvidio || 0),
        effectiveness: calcOperationEffectiveness(row, "0800"),
        quality: qualityScore,
        updatedAt: row.updated_at,
        calls_approved: Number(row.calls_0800_approved || 0),
        calls_rejected: Number(row.calls_0800_rejected || 0),
        calls_pending: Number(row.calls_0800_pending || 0),
        calls_no_action: Number(row.calls_0800_no_action || 0),
      });
    }
    if (operation === "all" || operation === "nuvidio") {
      list.push({
        entryType: "metric",
        metricId: row.id,
        userId: row.user_id,
        date: row.metric_date,
        dateLabel: formatDateBr(row.metric_date),
        operation: "Nuvidio",
        production: Number(row.production_nuvidio || 0),
        production_0800: Number(row.production_0800 || 0),
        production_nuvidio: Number(row.production_nuvidio || 0),
        effectiveness: calcOperationEffectiveness(row, "nuvidio"),
        quality: qualityScore,
        updatedAt: row.updated_at,
        calls_approved: Number(row.calls_nuvidio_approved || 0),
        calls_rejected: Number(row.calls_nuvidio_rejected || 0),
        calls_pending: 0,
        calls_no_action: Number(row.calls_nuvidio_no_action || 0),
        calls_empty: Number(row.calls_nuvidio_empty || 0),
      });
    }
    return list;
  });
  const historyQualityRows = qualityRows.map((row) => ({
    entryType: "quality",
    metricId: row.id,
    userId: row.user_id,
    date: `${row.reference_month}-01`,
    dateLabel: formatMonthLabel(row.reference_month),
    operation: `Qualidade · ${formatQualityScopeLabel(row.quality_scope)}`,
    production: null,
    effectiveness: null,
    quality: Number(row.score || 0),
    scoreType: row.score_type || "monitorias",
    qualityScope: row.quality_scope || "all",
    updatedAt: row.updated_at,
    referenceMonth: row.reference_month,
    monitoria_1: row.monitoria_1,
    monitoria_2: row.monitoria_2,
    monitoria_3: row.monitoria_3,
    monitoria_4: row.monitoria_4,
    m1_entered: Boolean(row.m1_entered),
    m2_entered: Boolean(row.m2_entered),
    m3_entered: Boolean(row.m3_entered),
    m4_entered: Boolean(row.m4_entered),
    notes: row.notes || "",
  }));
  return [...metricRows, ...historyQualityRows]
    .sort((a, b) => String(b.date).localeCompare(String(a.date)) || String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
}

function buildOverviewModel() {
  const trend = state.overview?.trend || [];
  const productionTrend = trend.filter((item) => !isSaturday(item.date));
  const ranking = state.analysis?.ranking || [];
  const quality = state.history?.quality || [];
  const historyRows = state.history?.history || [];
  const latest = trend[trend.length - 1];
  const previous = trend[trend.length - 2];
  const byDate0800 = new Map();
  const byDateNuvidio = new Map();
  historyRows.forEach((row) => {
    const date = String(row.metric_date || "").trim();
    if (!date) return;
    const current0800 = byDate0800.get(date) || { date, production: 0, effectivenessParts: [] };
    const currentNuvidio = byDateNuvidio.get(date) || { date, production: 0, effectivenessParts: [] };
    current0800.production += Number(row.production_0800 || 0);
    currentNuvidio.production += Number(row.production_nuvidio || 0);
    current0800.effectivenessParts.push(calcOperationEffectiveness(row, "0800"));
    currentNuvidio.effectivenessParts.push(calcOperationEffectiveness(row, "nuvidio"));
    byDate0800.set(date, current0800);
    byDateNuvidio.set(date, currentNuvidio);
  });
  const trend0800 = [...byDate0800.values()]
    .map((item) => ({
      date: item.date,
      production: item.production,
      effectiveness: average(item.effectivenessParts, { ignoreZero: true }),
    }))
    .sort((a, b) => a.date.localeCompare(b.date));
  const trendNuvidio = [...byDateNuvidio.values()]
    .map((item) => ({
      date: item.date,
      production: item.production,
      effectiveness: average(item.effectivenessParts, { ignoreZero: true }),
    }))
    .sort((a, b) => a.date.localeCompare(b.date));
  const latestMetricDate = [...new Set([...trend0800.map((item) => item.date), ...trendNuvidio.map((item) => item.date)])]
    .sort((a, b) => a.localeCompare(b))
    .pop() || "";
  const latest0800 = trend0800.find((item) => item.date === latestMetricDate);
  const latestNuvidio = trendNuvidio.find((item) => item.date === latestMetricDate);
  const latestEffectivenessValues = [
    Number(latest0800?.effectiveness || 0),
    Number(latestNuvidio?.effectiveness || 0),
  ];
  const latestProduction = Number(latest0800?.production || 0) + Number(latestNuvidio?.production || 0);
  return {
    totalAttended: trend.reduce((sum, item) => sum + Number(item.production || 0), 0),
    avgProduction: average(productionTrend.map((item) => item.production), { ignoreZero: true }),
    avgEffectiveness: latestMetricDate
      ? average(latestEffectivenessValues, { ignoreZero: true })
      : average(trend.map((item) => item.effectiveness), { ignoreZero: true }),
    avgQuality: average(quality.map((item) => item.score)),
    latest,
    latestDate: latestMetricDate || latest?.date || latest?.metric_date || "",
    latestProduction,
    latestEffectiveness: average(latestEffectivenessValues, { ignoreZero: true }),
    daysTracked: trend.length,
    prodDelta: latest && previous ? percentageDelta(previous.production, latest.production) : 0,
    effDelta: latest && previous ? percentageDelta(previous.effectiveness, latest.effectiveness) : 0,
    ranking: [...ranking].sort((a, b) => combinedScore(b) - combinedScore(a)).slice(0, 5),
  };
}

function buildAnalysisModel() {
  const ranking = [...(state.analysis?.ranking || [])];
  const selectedUser = state.filters.analysisUserId;
  const filteredRanking = selectedUser === "all" ? ranking : ranking.filter((item) => String(item.user_id) === String(selectedUser));
  const trend = state.overview?.trend || [];
  const historyRows = state.history?.history || [];
  const qualityRows = (state.history?.quality || []).slice().sort((a, b) => a.reference_month.localeCompare(b.reference_month));
  const qualityByMonthMap = new Map();
  const hasValue = (value) => value !== null && value !== undefined && String(value).trim() !== "";
  qualityRows.forEach((item) => {
    const scope = normalizeQualityScope(item.quality_scope || item.qualityScope || "all");
    const key = `${String(item.reference_month || "")}|${scope}`;
    const bucket = qualityByMonthMap.get(key) || {
      reference_month: String(item.reference_month || ""),
      quality_scope: scope,
      scoreValues: [],
      generalValues: [],
      m1Values: [],
      m2Values: [],
      m3Values: [],
      m4Values: [],
      rows: [],
    };
    if (hasValue(item.score)) {
      bucket.scoreValues.push(Number(item.score));
      if (String(item.score_type || "monitorias") === "general") bucket.generalValues.push(Number(item.score));
    }
    if (hasValue(item.monitoria_1)) bucket.m1Values.push(Number(item.monitoria_1));
    if (hasValue(item.monitoria_2)) bucket.m2Values.push(Number(item.monitoria_2));
    if (hasValue(item.monitoria_3)) bucket.m3Values.push(Number(item.monitoria_3));
      if (hasValue(item.monitoria_4)) bucket.m4Values.push(Number(item.monitoria_4));
      bucket.rows.push(item);
    qualityByMonthMap.set(key, bucket);
  });
  const qualityMonths = [...qualityByMonthMap.values()].map((bucket) => {
    const avg = (arr) => arr.length ? (arr.reduce((s, v) => s + Number(v || 0), 0) / arr.length) : 0;
    const m1 = avg(bucket.m1Values);
    const m2 = avg(bucket.m2Values);
    const m3 = avg(bucket.m3Values);
    const m4 = avg(bucket.m4Values);
    const launchedMonitorias = [];
    if (bucket.m1Values.length) launchedMonitorias.push(m1);
    if (bucket.m2Values.length) launchedMonitorias.push(m2);
    if (bucket.m3Values.length) launchedMonitorias.push(m3);
    if (bucket.m4Values.length) launchedMonitorias.push(m4);
    const final = launchedMonitorias.length
      ? (launchedMonitorias.reduce((sum, value) => sum + value, 0) / launchedMonitorias.length)
      : 0;
    const general = avg(bucket.generalValues);
    const hasGeneralScore = bucket.generalValues.length > 0;
    return {
      reference_month: bucket.reference_month,
      quality_scope: bucket.quality_scope,
      monitoria_1: m1,
      monitoria_2: m2,
      monitoria_3: m3,
      monitoria_4: m4,
      has_general_score: hasGeneralScore,
      general_score: general,
      score: hasGeneralScore ? general : final,
      final_score: final,
      rows: bucket.rows,
    };
  }).sort((a, b) =>
    a.reference_month.localeCompare(b.reference_month)
    || String(a.quality_scope || "").localeCompare(String(b.quality_scope || ""))
  );
  const status0800 = historyRows.reduce((acc, row) => {
    acc.approved += Number(row.calls_0800_approved || 0);
    acc.pending += Number(row.calls_0800_pending || 0);
    acc.rejected += Number(row.calls_0800_rejected || 0);
    acc.noAction += Number(row.calls_0800_no_action || 0);
    return acc;
  }, { approved: 0, pending: 0, rejected: 0, noAction: 0 });
  const statusNuvidio = historyRows.reduce((acc, row) => {
    acc.approved += Number(row.calls_nuvidio_approved || 0);
    acc.rejected += Number(row.calls_nuvidio_rejected || 0);
    acc.noAction += Number(row.calls_nuvidio_no_action || 0);
    acc.empty += Number(row.calls_nuvidio_empty || 0);
    return acc;
  }, { approved: 0, pending: 0, rejected: 0, noAction: 0, empty: 0 });
  const status = historyRows.reduce((acc, row) => {
    if (state.filters.operation === "all" || state.filters.operation === "0800") {
      acc.approved += Number(row.calls_0800_approved || 0);
      acc.pending += Number(row.calls_0800_pending || 0);
      acc.rejected += Number(row.calls_0800_rejected || 0);
      acc.noAction += Number(row.calls_0800_no_action || 0);
    }
    if (state.filters.operation === "all" || state.filters.operation === "nuvidio") {
      acc.approved += Number(row.calls_nuvidio_approved || 0);
      acc.rejected += Number(row.calls_nuvidio_rejected || 0);
      acc.noAction += Number(row.calls_nuvidio_no_action || 0);
      acc.empty += Number(row.calls_nuvidio_empty || 0);
    }
    return acc;
  }, { approved: 0, pending: 0, rejected: 0, noAction: 0, empty: 0 });
  const statusTotal = status.approved + status.pending + status.rejected + status.noAction + status.empty;
  const statusBreakdown = [
    { key: "approved", label: "Aprovado", value: status.approved, tone: "green" },
    { key: "rejected", label: "Reprovado", value: status.rejected, tone: "red" },
    { key: "pending", label: "Pendenciado", value: status.pending, tone: "amber" },
    { key: "noAction", label: "Sem ação", value: status.noAction, tone: "blue" },
    { key: "empty", label: "Vazio", value: status.empty, tone: "muted" },
  ].map((item) => ({
    ...item,
    share: statusTotal ? (item.value / statusTotal) * 100 : 0,
  }));
  const buildStatusBreakdown = (bucket, visibleKeys = ["approved", "rejected", "pending", "noAction", "empty"]) => {
    const breakdownBase = [
      { key: "approved", label: "Aprovado", value: bucket.approved, tone: "green" },
      { key: "rejected", label: "Reprovado", value: bucket.rejected, tone: "red" },
      { key: "pending", label: "Pendenciado", value: bucket.pending, tone: "amber" },
      { key: "noAction", label: "Sem ação", value: bucket.noAction, tone: "blue" },
      { key: "empty", label: "Vazio", value: bucket.empty || 0, tone: "muted" },
    ].filter((item) => visibleKeys.includes(item.key));
    const total = breakdownBase.reduce((sum, item) => sum + Number(item.value || 0), 0);
    const breakdown = breakdownBase.map((item) => ({
      ...item,
      share: total ? (item.value / total) * 100 : 0,
    }));
    return { total, breakdown };
  };
  const tags0800 = buildStatusBreakdown(status0800, ["approved", "rejected", "pending", "noAction"]);
  const tagsNuvidio = buildStatusBreakdown(statusNuvidio, ["approved", "rejected", "noAction", "empty"]);
  const byDate0800 = new Map();
  const byDateNuvidio = new Map();
  historyRows.forEach((row) => {
    const date = String(row.metric_date || "").trim();
    if (!date) return;
    const current0800 = byDate0800.get(date) || { date, production: 0, effectivenessParts: [] };
    const currentNuvidio = byDateNuvidio.get(date) || { date, production: 0, effectivenessParts: [] };
    current0800.production += Number(row.production_0800 || 0);
    currentNuvidio.production += Number(row.production_nuvidio || 0);
    current0800.effectivenessParts.push(calcOperationEffectiveness(row, "0800"));
    currentNuvidio.effectivenessParts.push(calcOperationEffectiveness(row, "nuvidio"));
    byDate0800.set(date, current0800);
    byDateNuvidio.set(date, currentNuvidio);
  });
  const trend0800 = [...byDate0800.values()]
    .map((item) => ({
      date: item.date,
      production: item.production,
      effectiveness: average(item.effectivenessParts, { ignoreZero: true }),
    }))
    .sort((a, b) => a.date.localeCompare(b.date));
  const trendNuvidio = [...byDateNuvidio.values()]
    .map((item) => ({
      date: item.date,
      production: item.production,
      effectiveness: average(item.effectivenessParts, { ignoreZero: true }),
    }))
    .sort((a, b) => a.date.localeCompare(b.date));
  const latestTrendDate = [...new Set([...trend0800.map((item) => item.date), ...trendNuvidio.map((item) => item.date)])]
    .sort((a, b) => a.localeCompare(b))
    .pop() || "";
  const latest0800 = trend0800.find((item) => item.date === latestTrendDate);
  const latestNuvidio = trendNuvidio.find((item) => item.date === latestTrendDate);
  const averageEffectivenessValues = [];
  if (state.filters.operation === "all" || state.filters.operation === "0800") {
    averageEffectivenessValues.push(...trend0800.map((item) => Number(item.effectiveness || 0)));
  }
  if (state.filters.operation === "all" || state.filters.operation === "nuvidio") {
    averageEffectivenessValues.push(...trendNuvidio.map((item) => Number(item.effectiveness || 0)));
  }
  return {
    trend,
    trend0800,
    trendNuvidio,
    qualityMonths,
    filteredRanking,
    status,
    statusTotal,
    statusBreakdown,
    tags0800,
    tagsNuvidio,
    summary: {
      production: average(filteredRanking.map((item) => item.avg_production), { ignoreZero: true }),
      effectiveness: average(averageEffectivenessValues, { ignoreZero: true }),
    },
  };
}

function renderOfflineHint() {
  app.innerHTML = `
    <section class="login-screen">
      <div class="login-card">
        <div class="login-copy">
          <div class="brand">
            <div class="brand-logo-wrap">
              <img class="brand-logo" src="${brandLogoSrc()}" alt="KR Consulting" onerror="this.style.display='none';this.nextElementSibling.style.display='grid';">
              <div class="brand-mark" style="display:none;">KR</div>
            </div>
            <div class="brand-copy">
              <span class="eyebrow">Performance operacional</span>
              <h1>PORTAL DE RESULTADOS</h1>
              <p>Abra via worker local.</p>
            </div>
          </div>
        </div>
        <div class="login-form">
          <span class="eyebrow">Inicializacao</span>
          <h2>Como testar</h2>
          <div class="info-box">& "C:\\Users\\joao.fonseca.KRCONSULTORIA\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\bin\\node.exe" worker.js</div>
        </div>
      </div>
    </section>
  `;
}

async function boot() {
  setBootLoaderMessage("Carregando portal...");
  showBootLoader();
  applyTheme();
  if (window.location.protocol === "file:") {
    renderOfflineHint();
    hideBootLoader();
    return;
  }
  try {
    const auth = await api("/api/auth/me");
    state.user = normalizeUserPayload(auth.user);
    if (auth.app_settings) state.appSettings = auth.app_settings;
    if (state.user) {
      applyUserPreferences();
      state.filters.historyUserId = isManager() ? "all" : String(state.user.id);
      if (!isManager()) state.filters.analysisUserId = String(state.user.id);
      enforceOperatorScope();
      if (!(state.user.role !== "manager" && state.appSettings.maintenance_for_operators)) {
        await loadAll();
      }
      state.filters.historyQuery = isManager() ? "" : (getUserLabelById(state.filters.historyUserId) || state.user.full_name);
      state.forcePasswordChange = Boolean(state.user.must_change_password);
    }
  } catch {
    state.user = null;
  }
  render();
  hideBootLoader();
}

async function loadAll() {
  if (isManager() && !state.users.length) {
    await loadUsers();
  }
  await loadBootstrap();
  if (state.route === "reports" || state.route === "detailed") {
    await loadReportsData();
  }
  if (state.route === "alerts" || ((state.route === "detailed" || state.route === "reports") && state.filters.reportsType === "ofensores")) {
    await loadAlerts();
    if (state.route === "detailed") {
      state.reportDataset = buildReportDatasetModel();
    }
  }
}

async function loadOverview() {
  state.overview = await api(`/api/dashboard/overview?date=${state.filters.today}&start=${state.filters.start}&end=${state.filters.end}`);
}

async function loadAnalysis() {
  state.analysis = await api(`/api/analysis?start=${state.filters.start}&end=${state.filters.end}`);
}

async function loadHistory() {
  const userId = isManager() && (state.filters.analysisUserId === "all" || state.filters.historyUserId === "all")
    ? "all"
    : (isManager() ? state.filters.historyUserId || state.user.id : state.user.id);
  state.history = await api(`/api/history?user_id=${userId}&start=${state.filters.start}&end=${state.filters.end}`);
  state.reportDataset = null;
  invalidateReportDatasetCache();
}

async function loadBootstrap() {
  const userId = isManager() && (state.filters.analysisUserId === "all" || state.filters.historyUserId === "all")
    ? "all"
    : (isManager() ? state.filters.historyUserId || state.user.id : state.user.id);
  const data = await api(`/api/bootstrap?date=${state.filters.today}&start=${state.filters.start}&end=${state.filters.end}&user_id=${userId}`);
  state.overview = data.overview;
  state.analysis = data.analysis;
  state.history = data.history;
  state.reportDataset = null;
  invalidateReportDatasetCache();
  if (data.app_settings) state.appSettings = data.app_settings;
  if (isManager() && Array.isArray(data.users) && data.users.length) {
    state.users = normalizeUsersPayload(data.users);
    ensureManagerUserFilters();
    state.filters.historyQuery = getUserLabelById(state.filters.historyUserId) || state.filters.historyQuery;
  }
}

async function loadAlerts() {
  const userId = isManager() ? (state.filters.analysisUserId || "all") : String(state.user?.id || "");
  state.alerts = await api(`/api/alerts?start=${state.filters.start}&end=${state.filters.end}&user_id=${encodeURIComponent(userId)}`);
  state.reportDataset = null;
  invalidateReportDatasetCache();
}

async function loadReportsData() {
  await loadHistory();
  state.reportDataset = buildReportDatasetModel();
}

async function loadUsers() {
  const response = await api("/api/admin/users");
  state.users = normalizeUsersPayload(response.users);
  if (isManager()) {
    ensureManagerUserFilters();
    state.filters.historyQuery = getUserLabelById(state.filters.historyUserId) || state.filters.historyQuery;
  }
}

function navMeta() {
  return {
    overview: "Visão geral",
    analysis: "Análises",
    ...(isManager() ? { alerts: "Ofensores" } : {}),
    history: "Histórico",
    reports: "Relatórios",
    detailed: "Detalhada",
    admin: "Gestão",
  };
}

function render() {
  applyTheme();
  if (!state.user) {
    stopMaintenanceWatcher();
    app.innerHTML = loginTemplate();
    bindLogin();
    return;
  }
  if (!isManager() && state.appSettings.maintenance_for_operators) {
    startMaintenanceWatcher();
    app.innerHTML = maintenanceTemplate();
    bindMaintenance();
    return;
  }
  startMaintenanceWatcher();
  enforceOperatorScope();
  app.innerHTML = shellTemplate();
  bindShellEvents();
}

function maintenanceTemplate() {
  return `
    <section class="login-screen">
      <div class="login-card">
        <div class="login-copy">
          <div class="brand">
            <div class="brand-logo-wrap">
              <img class="brand-logo" src="${brandLogoSrc()}" alt="KR Consulting" onerror="this.style.display='none';this.nextElementSibling.style.display='grid';">
              <div class="brand-mark" style="display:none;">KR</div>
            </div>
            <div class="brand-copy">
              <span class="eyebrow">Manutenção</span>
              <h1>PORTAL DE RESULTADOS</h1>
              <p>${esc(state.appSettings.maintenance_message || "Portal em manutenção. Tente novamente em instantes.")}</p>
            </div>
          </div>
        </div>
        <div class="login-form">
          <span class="eyebrow">Acesso temporariamente indisponível</span>
          <h2>Em manutenção</h2>
          <div class="info-box">O acesso para operadores está temporariamente pausado.</div>
          <button class="btn" id="maintenance-logout">Sair</button>
        </div>
      </div>
    </section>
  `;
}

function themeToggleIcon() {
  if (state.theme === "contrast") {
    return `
      <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <g fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8">
          <circle cx="12" cy="12" r="3.6"/>
          <path d="M12 3.75v2.1M12 18.15v2.1M3.75 12h2.1M18.15 12h2.1M6.15 6.15l1.48 1.48M16.37 16.37l1.48 1.48M17.85 6.15l-1.48 1.48M7.63 16.37l-1.48 1.48"/>
        </g>
      </svg>
    `;
  }
  return `
    <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
      <path fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.9" d="M15.25 4.65A7.95 7.95 0 1 0 19.35 18a6.9 6.9 0 1 1-4.1-13.35Z"/>
    </svg>
  `;
}

function bindMaintenance() {
  const button = document.getElementById("maintenance-logout");
  if (!button) return;
  button.addEventListener("click", async () => {
    try {
      await api("/api/auth/logout", { method: "POST" });
    } finally {
      state.user = null;
      render();
    }
  });
}

function flashTemplate() {
  if (!state.flash) return "";
  const details = Array.isArray(state.flash.details) ? state.flash.details : [];
  return `
    <div class="notice toast ${esc(state.flash.type)} ${details.length ? "has-details" : ""}">
      <div class="toast-message">${esc(state.flash.message)}</div>
      ${details.length ? `
        <details class="toast-details">
          <summary>Ver falhas (${details.length})</summary>
          <div class="toast-details-list">
            ${details.map((item) => `<div class="toast-details-item">${esc(item)}</div>`).join("")}
          </div>
        </details>
      ` : ""}
    </div>
  `;
}

function shellTemplate() {
  const titles = {
    overview: { title: isManager() ? "Visão da operação" : "Performance do operador", desc: "" },
    analysis: { title: "Análises", desc: "" },
    alerts: { title: "Ofensores", desc: "" },
    history: { title: "Histórico", desc: "" },
    reports: { title: "Relatórios", desc: "" },
    detailed: { title: "Detalhada", desc: "" },
    admin: { title: "Gestão", desc: "" },
  };
  const current = titles[state.route];
  const operatorUsers = getOperatorUsers();
  const selectedLabel = isManager()
    ? (state.filters.analysisUserId === "all" ? "Todos os operadores" : operatorUsers.find((user) => String(user.id) === String(state.filters.analysisUserId))?.full_name || "Operador")
    : state.user.full_name;
  return `
    <div class="shell">
      <aside class="sidebar">
        <div class="brand-box">
          <div class="brand">
            <div class="brand-logo-wrap">
              <img class="brand-logo" src="${brandLogoSrc()}" alt="KR Consulting" onerror="this.style.display='none';this.nextElementSibling.style.display='grid';">
              <div class="brand-mark" style="display:none;">KR</div>
            </div>
            <div class="brand-copy">
              <h1>PORTAL DE RESULTADOS</h1>
              <p>Performance operacional</p>
            </div>
          </div>
        </div>
        <nav class="nav">
          ${Object.entries(navMeta()).filter(([key]) => (key !== "admin" && key !== "reports" && key !== "detailed") || isManager()).map(([key, label]) => `
            <button class="${state.route === key ? "active" : ""}" data-route="${key}">${label}</button>
          `).join("")}
        </nav>
      </aside>
      <main class="main">
        <div class="topbar">
          <div class="topbar-left">
            <div class="page-copy">
              <h2>${current.title}</h2>
              ${current.desc ? `<p>${current.desc}</p>` : ""}
            </div>
          </div>
          <div class="topbar-actions">
            <div class="select-wrap">
              <select id="global-user-select" ${isManager() ? "" : 'disabled aria-disabled="true"'}>
                ${isManager() ? `
                  <option value="all" ${state.filters.analysisUserId === "all" ? "selected" : ""}>Todos os operadores</option>
                  ${operatorUsers.map((user) => `<option value="${user.id}" ${String(user.id) === String(state.filters.analysisUserId) ? "selected" : ""}>${esc(user.full_name)}</option>`).join("")}
                ` : `<option value="${state.user.id}">${esc(selectedLabel)}</option>`}
              </select>
            </div>
            <button class="btn" data-action="refresh-all">Atualizar</button>
            <button class="toggle theme-toggle" data-action="toggle-theme" title="${state.theme === "contrast" ? "Ativar tema escuro" : "Ativar tema claro"}" aria-label="${state.theme === "contrast" ? "Ativar tema escuro" : "Ativar tema claro"}">
              ${themeToggleIcon()}
            </button>
            <div class="profile-menu profile-menu-topbar">
              <button class="topbar-user menu-trigger topbar-user-trigger" type="button" id="profile-menu-trigger" aria-haspopup="true" aria-expanded="false">
                <div class="avatar">${initials(state.user.full_name)}</div>
                <div>
                  <strong>${esc(state.user.full_name)}</strong>
                  <span>${isManager() ? "Gestor" : "Operador"}</span>
                </div>
              </button>
              <div class="menu-popover" id="profile-menu-popover" hidden>
                <button class="menu-item" type="button" id="open-password-modal">Redefinir senha</button>
                <button class="menu-item danger" type="button" data-action="logout">Sair</button>
              </div>
            </div>
          </div>
        </div>
        ${renderPage()}
      </main>
      <div class="modal-backdrop" id="password-modal" hidden>
        <div class="modal-card">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Segurança</span>
              <h3>${state.forcePasswordChange ? "Primeiro acesso" : "Redefinir senha"}</h3>
            </div>
            ${state.forcePasswordChange ? "" : `<button class="icon-close" type="button" id="close-password-modal" aria-label="Fechar">×</button>`}
          </div>
          <form id="password-form" class="section compact-form">
            ${state.forcePasswordChange ? `<div class="info-box">Para continuar, defina uma nova senha.</div>` : `<label>Senha atual<input name="current_password" type="password" required></label>`}
            <label>Nova senha<input name="new_password" type="password" minlength="4" required></label>
            <label>Confirmar nova senha<input name="confirm_password" type="password" minlength="4" required></label>
            <div class="action-grid">
              ${state.forcePasswordChange ? "" : `<button class="btn-secondary" type="button" id="cancel-password-modal">Cancelar</button>`}
              <button class="btn" type="submit">${state.forcePasswordChange ? "Salvar e continuar" : "Salvar senha"}</button>
            </div>
          </form>
        </div>
      </div>
      <div class="modal-backdrop" id="user-modal" hidden>
        <div class="modal-card">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Usuário</span>
              <h3>Editar cadastro</h3>
            </div>
            <button class="icon-close" type="button" id="close-user-modal" aria-label="Fechar">×</button>
          </div>
          <form id="user-edit-form" class="section compact-form">
            <input type="hidden" name="user_id" id="edit-user-id">
            <div class="form-grid">
              <label>Nome completo<input name="full_name" id="edit-full-name" required></label>
              <label>Login<input name="login" id="edit-login" required></label>
              <label>Perfil
                <select name="role" id="edit-role" required>
                  <option value="operator">Operador</option>
                  <option value="manager">Gestor</option>
                </select>
              </label>
              <label>Status
                <select name="is_active" id="edit-is-active" required>
                  <option value="true">Ativo</option>
                  <option value="false">Inativo</option>
                </select>
              </label>
              <label>ID 0800<input name="platform_0800_id" id="edit-platform-0800-id"></label>
              <label>ID Nuvidio<input name="nuvidio_id" id="edit-nuvidio-id"></label>
            </div>
            <label>Nova senha (opcional)<input name="password" id="edit-password" type="password"></label>
            <div class="info-box">Se você informar uma nova senha, o usuário será obrigado a trocá-la no próximo login.</div>
            <div class="action-grid">
              <button class="btn-secondary" type="button" id="cancel-user-modal">Cancelar</button>
              <button class="btn" type="submit">Salvar alterações</button>
            </div>
          </form>
        </div>
      </div>
      <div class="modal-backdrop" id="history-edit-modal" hidden>
        <div class="modal-card">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Histórico</span>
              <h3>Editar lançamento</h3>
            </div>
            <button class="icon-close" type="button" id="close-history-edit-modal" aria-label="Fechar">×</button>
          </div>
          <form id="history-edit-form" class="section compact-form">
            <input type="hidden" id="history-edit-type" name="entry_type">
            <input type="hidden" id="history-edit-metric-id" name="metric_id">
            <input type="hidden" id="history-edit-operation" name="operation">
            <div class="form-grid">
              <label>Operador<input id="history-edit-operator" readonly></label>
              <label>Data<input id="history-edit-date" readonly></label>
              <label>Operação<input id="history-edit-operation-label" readonly></label>
            </div>
            <div class="form-grid history-edit-metric-fields">
              <label>Aprovado<input type="number" min="0" step="1" id="history-edit-approved" name="approved"></label>
              <label>Reprovado<input type="number" min="0" step="1" id="history-edit-rejected" name="rejected"></label>
              <label class="history-edit-pending">Pendenciado<input type="number" min="0" step="1" id="history-edit-pending" name="pending"></label>
              <label>Sem ação<input type="number" min="0" step="1" id="history-edit-no-action" name="no_action"></label>
              <label class="history-edit-empty">Vazio<input type="number" min="0" step="1" id="history-edit-empty" name="empty"></label>
            </div>
            <div class="form-grid history-edit-quality-fields" hidden>
              <label>Esteira
                <select id="history-edit-quality-scope" name="quality_scope">
                  <option value="0800">0800</option>
                  <option value="nuvidio">Nuvidio</option>
                </select>
              </label>
              <label>Nota bruta do mês<input type="number" min="0" max="100" step="0.01" id="history-edit-score" name="score" required></label>
              <label>Observações<input id="history-edit-notes" name="notes"></label>
            </div>
            <div class="action-grid">
              <button class="btn-secondary" type="button" id="cancel-history-edit-modal">Cancelar</button>
              <button class="btn" type="submit">Salvar lançamento</button>
            </div>
          </form>
        </div>
      </div>
      <div class="modal-backdrop" id="history-view-modal" hidden>
        <div class="modal-card">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Histórico</span>
              <h3>Qualidade lançada</h3>
            </div>
            <button class="icon-close" type="button" id="close-history-view-modal" aria-label="Fechar">×</button>
          </div>
          <div class="section compact-form" id="history-view-content"></div>
          <div class="action-grid">
            <button class="btn-secondary" type="button" id="cancel-history-view-modal">Fechar</button>
          </div>
        </div>
      </div>
      <div class="modal-backdrop" id="history-delete-day-modal" hidden>
        <div class="modal-card">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Histórico</span>
              <h3>Excluir registros do dia</h3>
            </div>
            <button class="icon-close" type="button" id="close-history-delete-day-modal" aria-label="Fechar">×</button>
          </div>
          <form id="history-delete-day-form" class="section compact-form">
            <div class="info-box" id="history-delete-day-info">Essa ação remove os registros de todos os operadores na data e setor selecionados.</div>
            <label id="history-delete-day-date-label">Data com registros
              <select id="history-delete-day-date" name="date" required>
                <option value="">Selecione um dia</option>
                ${historyAvailableDates().map((date) => `<option value="${esc(date)}">${esc(formatDateBr(date))}</option>`).join("")}
              </select>
            </label>
            <label id="history-delete-quality-month-label" hidden>Mês com qualidade lançada
              <select id="history-delete-quality-month" name="reference_month">
                <option value="">Selecione um mês</option>
                ${historyAvailableQualityMonths().map((month) => `<option value="${esc(month)}">${esc(formatMonthLabel(month))}</option>`).join("")}
              </select>
            </label>
            <label>Setor
              <select name="operation" id="history-delete-day-operation" required>
                <option value="all">0800 + Nuvidio</option>
                <option value="0800">0800</option>
                <option value="nuvidio">Nuvidio</option>
                <option value="quality-all">Qualidade · 0800 + Nuvidio</option>
                <option value="quality-0800">Qualidade · 0800</option>
                <option value="quality-nuvidio">Qualidade · Nuvidio</option>
              </select>
            </label>
            <div class="action-grid">
              <button class="btn-secondary" type="button" id="cancel-history-delete-day-modal">Cancelar</button>
              <button class="btn" type="submit">Excluir registros</button>
            </div>
          </form>
        </div>
      </div>
      ${flashTemplate()}
    </div>
  `;
}

function loginTemplate() {
  return `
    <section class="login-screen">
      <div class="login-card">
        <div class="login-copy">
          <div class="brand">
            <div class="brand-logo-wrap">
              <img class="brand-logo" src="${brandLogoSrc()}" alt="KR Consulting" onerror="this.style.display='none';this.nextElementSibling.style.display='grid';">
              <div class="brand-mark" style="display:none;">KR</div>
            </div>
            <div class="brand-copy">
              <span class="eyebrow">Performance operacional</span>
              <h1>PORTAL DE RESULTADOS</h1>
              <p>Acesso ao portal.</p>
            </div>
          </div>
        </div>
        <form class="login-form" id="login-form">
          <span class="eyebrow">Login</span>
          <h2>Entrar</h2>
          <label>Login<input name="login" required></label>
          <label>Senha<input name="password" type="password" required></label>
          <button class="btn" type="submit">Acessar</button>
          <button class="btn-secondary" type="button" id="open-self-reset-modal">Esqueci minha senha</button>
        </form>
      </div>
      <div class="modal-backdrop" id="self-reset-modal" hidden>
        <div class="modal-card">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Acesso</span>
              <h3>Redefinir minha senha</h3>
            </div>
            <button class="icon-close" type="button" id="close-self-reset-modal" aria-label="Fechar">×</button>
          </div>
          <form id="self-reset-form" class="section compact-form">
            <div class="info-box">Confirme seus dados cadastrados para definir uma nova senha.</div>
            <label>Login<input name="login" required></label>
            <label>Nome completo<input name="full_name" required></label>
            <label>ID 0800 (opcional)<input name="platform_0800_id"></label>
            <label>Qual seu usuário da Nuvidio?<input name="nuvidio_id"></label>
            <label>Nova senha<input name="new_password" type="password" minlength="4" required></label>
            <label>Confirmar nova senha<input name="confirm_password" type="password" minlength="4" required></label>
            <div class="action-grid">
              <button class="btn-secondary" type="button" id="cancel-self-reset-modal">Cancelar</button>
              <button class="btn" type="submit">Salvar nova senha</button>
            </div>
          </form>
        </div>
      </div>
      ${flashTemplate()}
    </section>
  `;
}

function renderPage() {
  if (state.route === "analysis") return analysisTemplate();
  if (state.route === "alerts") return alertsTemplate();
  if (state.route === "history") return historyTemplate();
  if (state.route === "reports") return reportsTemplate();
  if (state.route === "detailed") return detailedTemplate();
  if (state.route === "admin") return adminTemplate();
  return overviewTemplate();
}

function overviewTemplate() {
  const model = buildOverviewModel();
  const trendRows = state.overview?.trend || [];
  const maxTrendProduction = Math.max(...trendRows.map((row) => Number(row.production || 0)), 1);
  const latestDate = formatDateBr(model.latestDate || "--");
  const latestProduction = model.latestDate ? integer(model.latestProduction) : "--";
  const latestEffectiveness = model.latestDate ? percent(model.latestEffectiveness) : "--%";
  const cards = [
    { label: "Total atendido", value: integer(model.totalAttended) },
    { label: "Média de produção", value: integer(model.avgProduction) || "--" },
    { label: "Efetividade média", value: percent(model.avgEffectiveness) },
    { label: "Qualidade média", value: model.avgQuality ? number(model.avgQuality) : "--" },
  ];
  return `
    <section class="section">
      <div class="hero-grid">
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Painel principal</span>
              <h3>${isManager() ? "Resumo operacional" : "Resumo individual"}</h3>
            </div>
          </div>
          <div class="mini-grid">
            <div class="mini-card">
              <span class="muted">Última data</span>
              <div class="metric-value">${latestDate}</div>
            </div>
            <div class="mini-card">
              <span class="muted">Dias lançados</span>
              <div class="metric-value">${integer(model.daysTracked)}</div>
            </div>
          </div>
        </article>
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Última atualização</span>
              <h3>${model.latest ? "Base atualizada" : "Aguardando lançamentos"}</h3>
            </div>
          </div>
          <div class="mini-grid">
            <div class="mini-card">
              <span class="muted">Produção</span>
              <div class="metric-value">${latestProduction}</div>
            </div>
            <div class="mini-card">
              <span class="muted">Efetividade</span>
              <div class="metric-value">${latestEffectiveness}</div>
            </div>
          </div>
        </article>
      </div>

      <article class="panel">
        <div class="panel-head">
          <div>
            <span class="eyebrow">Resumo rápido</span>
            <h3>Indicadores principais</h3>
          </div>
        </div>
        <div class="kpi-grid">
          ${cards.map((card) => `
            <div class="metric-card">
              <span class="muted">${card.label}</span>
              <div class="metric-value">${card.value}</div>
            </div>
          `).join("")}
        </div>
      </article>

      <article class="panel">
        <div class="panel-head">
          <div>
            <span class="eyebrow">Evolução visual</span>
            <h3>Tendência recente</h3>
          </div>
        </div>
        <div class="chart">
          ${trendRows.length ? trendRows.map((item) => `
            <div class="chart-col trend-point" data-trend-date="${esc(item.date)}">
              <span class="chart-value">${integer(item.production)}</span>
              <div class="column ${metricTone(item.production, "production")}" style="height:${Math.max(14, (Number(item.production || 0) / maxTrendProduction) * 110)}px;"></div>
              <small>${shortDate(item.date)}</small>
            </div>
          `).join("") : `<div class="empty">Sem tendência disponível.</div>`}
        </div>
        ${isManager() ? `<div class="trend-tooltip" id="trend-tooltip" hidden></div>` : ""}
      </article>
    </section>
  `;
}

function analysisTemplate() {
  const model = buildAnalysisModel();
  const trend = model.trend;
  const maxProduction = Math.max(...trend.map((item) => Number(item.production || 0)), 1);
  const maxEffectiveness = Math.max(...trend.map((item) => Number(item.effectiveness || 0)), 1);
  const qualityMonths = model.qualityMonths;
  const maxQuality = Math.max(...qualityMonths.flatMap((item) => [
    Number(item.monitoria_1 || 0),
    Number(item.monitoria_2 || 0),
    Number(item.monitoria_3 || 0),
    Number(item.monitoria_4 || 0),
    Number(item.general_score || 0),
    Number(item.final_score || item.score || 0),
  ]), 10);
  const trend0800 = model.trend0800 || [];
  const trendNuvidio = model.trendNuvidio || [];
  const maxProd0800 = Math.max(...trend0800.map((item) => Number(item.production || 0)), 1);
  const maxEff0800 = Math.max(...trend0800.map((item) => Number(item.effectiveness || 0)), 1);
  const maxProdNuvidio = Math.max(...trendNuvidio.map((item) => Number(item.production || 0)), 1);
  const maxEffNuvidio = Math.max(...trendNuvidio.map((item) => Number(item.effectiveness || 0)), 1);
  const showAllOperatorsTop = isManager() && String(state.filters.analysisUserId) === "all";
  return `
    <section class="section">
      <article class="panel analysis-filter-panel">
        <div class="analysis-filter-bar">
          <div class="analysis-filter-copy">
            <span class="eyebrow">Filtros</span>
            <h3>Recorte</h3>
          </div>
          <label class="analysis-inline-field">Início<input type="date" id="start-filter" value="${state.filters.start}"></label>
          <label class="analysis-inline-field">Fim<input type="date" id="end-filter" value="${state.filters.end}"></label>
          <div class="analysis-filter-actions">
            <button class="btn" data-action="refresh-analysis">Aplicar</button>
            <button class="btn-secondary" data-action="reset-analysis">Limpar</button>
          </div>
        </div>
      </article>
      <div class="analysis-dashboard">
          <div class="analysis-grid-split">
            <article class="panel analysis-panel-primary">
              <div class="panel-head"><div><span class="eyebrow">0800</span><h3>Produção diária</h3></div></div>
              <div class="chart">
                ${trend0800.length ? trend0800.map((item) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="production" data-operation="0800" data-date="${esc(item.date)}"` : ""}><span class="chart-value">${integer(item.production)}</span><div class="column ${metricTone(item.production, "production")}" style="height:${Math.max(12, (item.production / maxProd0800) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
            <article class="panel analysis-panel-secondary">
              <div class="panel-head"><div><span class="eyebrow">0800</span><h3>Efetividade diária</h3></div></div>
              <div class="chart">
                ${trend0800.length ? trend0800.map((item) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="effectiveness" data-operation="0800" data-date="${esc(item.date)}"` : ""}><span class="chart-value">${percent(item.effectiveness)}</span><div class="column ${metricTone(item.effectiveness, "effectiveness")}" style="height:${Math.max(12, (item.effectiveness / maxEff0800) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
          </div>

          <div class="analysis-grid-split">
            <article class="panel analysis-panel-primary">
              <div class="panel-head"><div><span class="eyebrow">Nuvidio</span><h3>Produção diária</h3></div></div>
              <div class="chart">
                ${trendNuvidio.length ? trendNuvidio.map((item) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="production" data-operation="nuvidio" data-date="${esc(item.date)}"` : ""}><span class="chart-value">${integer(item.production)}</span><div class="column ${metricTone(item.production, "production")}" style="height:${Math.max(12, (item.production / maxProdNuvidio) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
            <article class="panel analysis-panel-secondary">
              <div class="panel-head"><div><span class="eyebrow">Nuvidio</span><h3>Efetividade diária</h3></div></div>
              <div class="chart">
                ${trendNuvidio.length ? trendNuvidio.map((item) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="effectiveness" data-operation="nuvidio" data-date="${esc(item.date)}"` : ""}><span class="chart-value">${percent(item.effectiveness)}</span><div class="column ${metricTone(item.effectiveness, "effectiveness")}" style="height:${Math.max(12, (item.effectiveness / maxEffNuvidio) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
          </div>

          <div class="analysis-grid-split">
            <article class="panel analysis-panel-primary">
              <div class="panel-head"><div><span class="eyebrow">Qualidade</span><h3>Meses lançados</h3></div></div>
              <div class="chart">
                ${qualityMonths.length ? qualityMonths.map((monthItem) => `
                  <div class="chart-col quality-summary-point" data-quality-summary="true" data-reference-month="${esc(monthItem.reference_month)}" data-quality-scope="${esc(monthItem.quality_scope || "all")}">
                    <span class="chart-value">${number(Number(monthItem.score || 0))}</span>
                    <div class="column ${metricTone(Number(monthItem.score || 0), "quality")}" style="height:${Math.max(12, (Number(monthItem.score || 0) / maxQuality) * 110)}px;"></div>
                    <small>${esc(formatQualityScopeLabel(monthItem.quality_scope || "all"))}</small>
                    <small>${esc(formatMonthLabel(monthItem.reference_month).split(" de ")[0])}</small>
                  </div>
                `).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
            <article class="panel analysis-panel-secondary">
              <div class="panel-head"><div><span class="eyebrow">Resumo</span><h3>Consolidado</h3></div></div>
              <div class="mini-grid">
                <div class="mini-card"><span class="muted">Produção média</span><div class="metric-value">${integer(model.summary.production)}</div></div>
                <div class="mini-card"><span class="muted">Efetividade média</span><div class="metric-value">${percent(model.summary.effectiveness)}</div></div>
              </div>
            </article>
          </div>

          <div class="analysis-grid-split analysis-grid-equal">
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Tags · 0800</span><h3>Divisão por tags</h3></div></div>
              <div class="tag-split">
                ${model.tags0800.breakdown.some((item) => item.value > 0) ? model.tags0800.breakdown.map((item) => `
                  <div class="tag-split-row">
                    <div class="tag-split-head">
                      <span>${item.label}</span>
                      <strong>${integer(item.value)} (${Math.round(item.share)}%)</strong>
                    </div>
                    <div class="tag-split-track">
                      <div class="tag-split-fill ${item.tone}" style="width:${Math.max(item.share, item.value ? 2 : 0)}%"></div>
                    </div>
                  </div>
                `).join("") : `<div class="empty">Sem dados de tags no período.</div>`}
              </div>
            </article>
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Tags · Nuvidio</span><h3>Divisão por tags</h3></div></div>
              <div class="tag-split">
                ${model.tagsNuvidio.breakdown.some((item) => item.value > 0) ? model.tagsNuvidio.breakdown.map((item) => `
                  <div class="tag-split-row">
                    <div class="tag-split-head">
                      <span>${item.label}</span>
                      <strong>${integer(item.value)} (${Math.round(item.share)}%)</strong>
                    </div>
                    <div class="tag-split-track">
                      <div class="tag-split-fill ${item.tone}" style="width:${Math.max(item.share, item.value ? 2 : 0)}%"></div>
                    </div>
                  </div>
                `).join("") : `<div class="empty">Sem dados de tags no período.</div>`}
              </div>
            </article>
          </div>
      </div>
      <div class="trend-tooltip" id="analysis-metric-tooltip" hidden></div>
    </section>
  `;
}

function alertsTemplate() {
  if (!state.alerts) {
    return `
      <section class="section">
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Ofensores</span>
              <h3>Leitura de risco operacional</h3>
            </div>
          </div>
          <div class="empty">Carregando alertas...</div>
        </article>
      </section>
    `;
  }

  const summary = state.alerts.summary || { monitored: 0, total: 0, average_score: 0, max_score: 0 };
  const alerts = state.alerts.alerts || [];
  const selectedName = isManager() && state.filters.analysisUserId !== "all"
    ? (getOperatorUsers().find((user) => String(user.id) === String(state.filters.analysisUserId))?.full_name || "Operador")
    : (isManager() ? "Todos os operadores" : state.user.full_name);
  const selfScope = !isManager() || state.alerts.scope === "self";

  return `
    <section class="section">
      <div class="hero-grid">
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Monitoramento</span>
              <h3>${selfScope ? "Seus principais ofensores" : "Operadores em alerta"}</h3>
            </div>
          </div>
          <div class="history-tools alerts-tools">
            <label>Início<input type="date" id="alerts-start-filter" value="${state.filters.start}"></label>
            <label>Fim<input type="date" id="alerts-end-filter" value="${state.filters.end}"></label>
            <button class="btn" data-action="refresh-alerts">Aplicar</button>
            <button class="btn-secondary" data-action="reset-alerts">Limpar</button>
          </div>
        </article>
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Resumo</span>
              <h3>${selfScope ? "Seu resultado atual" : "Média da operação"}</h3>
            </div>
          </div>
          <div class="mini-grid">
            <div class="mini-card"><span class="muted">Monitorados</span><div class="metric-value">${integer(summary.monitored)}</div></div>
            <div class="mini-card"><span class="muted">Em alerta</span><div class="metric-value">${integer(summary.total)}</div></div>
            <div class="mini-card"><span class="muted">Nota média</span><div class="metric-value">${number(summary.average_score)}</div></div>
            <div class="mini-card"><span class="muted">Menor nota</span><div class="metric-value">${number(summary.max_score)}</div></div>
          </div>
        </article>
      </div>

      <article class="panel">
          <div class="panel-head">
          <div>
            <span class="eyebrow">Prioridade</span>
            <h3>${selfScope ? "Itens que mais prejudicam seu resultado" : "Pontos de atenção por operador"}</h3>
          </div>
        </div>
        ${alerts.length ? `
          <div class="alert-grid">
            ${alerts.map((item) => {
              return `
                <article class="alert-card">
                  <div class="alert-card-head">
                    <div>
                      <h4>${esc(item.name)}</h4>
                      <p>${esc(item.login)}</p>
                    </div>
                    <span class="pill ${item.alert_tone}">Nota ${number(item.alert_score)} · ${esc(item.alert_label)}</span>
                  </div>
                  <div class="alert-metrics">
                    <div class="mini-card">
                      <span class="muted">Nota do operador</span>
                      <div class="metric-value">${number(item.alert_score)}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Produção média 0800</span>
                      <div class="metric-value">${integer(item.avg_production_0800)}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Efetividade 0800</span>
                      <div class="metric-value">${percent(item.effectiveness_0800)}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Produção média Nuvidio</span>
                      <div class="metric-value">${integer(item.avg_production_nuvidio)}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Efetividade Nuvidio</span>
                      <div class="metric-value">${percent(item.effectiveness_nuvidio)}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Qualidade</span>
                      <div class="metric-value">${item.quality ? number(item.quality) : "--"}</div>
                    </div>
                  </div>
                  <div class="alert-reasons">
                    ${item.reasons.map((reason) => `<div class="alert-reason ${reason.tone}">${esc(reason.text)}</div>`).join("")}
                  </div>
                  <div class="alert-footer">
                    <span>Dias ativos: <strong>${integer(item.active_days)}</strong></span>
                    <span>Última qualidade: <strong>${item.latest_quality_month ? esc(formatMonthLabel(item.latest_quality_month)) : "Sem registro"}</strong></span>
                  </div>
                </article>
              `;
            }).join("")}
          </div>
        ` : `<div class="empty">${selfScope ? "Nenhum item crítico encontrado no seu resultado neste período." : "Nenhum operador em alerta no recorte atual."}</div>`}
      </article>
    </section>
  `;
}

function reportsTemplate() {
  const reportType = state.filters.reportsType;
  const reportView = state.filters.reportsView;
  const model = state.reportDataset;
  if (!model) {
    return `
      <section class="section">
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Relatórios</span>
              <h3>Preparando dados</h3>
            </div>
          </div>
          <div class="empty">Carregando base para exportação...</div>
        </article>
      </section>
    `;
  }
  const selectedName = model.selectedName;
  const reportRows = model.reportRows;
  const qualityRows = model.qualityRows;
  const sectorLabel = model.sectorLabel;
  const reportTypeLabel = model.reportTypeLabel;
  const reportViewLabel = model.reportViewLabel;
  const consolidatedCount = model.consolidatedRows.length;
  const offendersCount = model.offendersRows.length;
  const exportColumns = {
    consolidado: ["Data", "Operador", "Setor", "Produção", "Efetividade"],
    operacional: ["Data", "Operador", "Setor", "Produção", "Efetividade"],
    qualidade: ["Mês", "Operador", "Esteira", "Nota", "Observações"],
    ofensores: ["Operador", "Nota", "Prod. 0800", "Efet. 0800", "Prod. Nuvidio", "Efet. Nuvidio", "Qualidade"],
  }[reportType] || ["Data", "Operador", "Setor", "Produção", "Efetividade"];
  const previewBody = `
    <div class="mini-grid report-preview-cards">
      <div class="mini-card"><span class="muted">Linhas operacionais</span><div class="metric-value">${integer(reportRows.length)}</div></div>
      <div class="mini-card"><span class="muted">Linhas consolidadas</span><div class="metric-value">${integer(consolidatedCount)}</div></div>
      <div class="mini-card"><span class="muted">Linhas de qualidade</span><div class="metric-value">${integer(qualityRows.length)}</div></div>
      <div class="mini-card"><span class="muted">Ofensores no recorte</span><div class="metric-value">${integer(offendersCount)}</div></div>
    </div>
    <div class="report-preview-columns">
      <span class="eyebrow">Colunas do arquivo</span>
      <div class="report-preview-column-list">
        ${exportColumns.map((column) => `<span class="report-column-pill">${esc(column)}</span>`).join("")}
      </div>
    </div>
    <div class="info-box">
      A exportação sai completa. Para consulta pesada e comparação linha a linha, use a aba <strong>Detalhada</strong>.
    </div>
  `;
  return `
    <section class="section">
      <div class="hero-grid">
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Exportação</span>
              <h3>Relatórios gerenciais</h3>
            </div>
          </div>
          <div class="history-tools reports-tools">
            <label>Início<input type="date" id="reports-start-filter" value="${state.filters.start}"></label>
            <label>Fim<input type="date" id="reports-end-filter" value="${state.filters.end}"></label>
            <label>Tipo
              <select id="reports-type-filter">
                <option value="consolidado" ${state.filters.reportsType === "consolidado" ? "selected" : ""}>Consolidado</option>
                <option value="operacional" ${state.filters.reportsType === "operacional" ? "selected" : ""}>Operacional</option>
                <option value="qualidade" ${state.filters.reportsType === "qualidade" ? "selected" : ""}>Qualidade</option>
                <option value="ofensores" ${state.filters.reportsType === "ofensores" ? "selected" : ""}>Ofensores</option>
              </select>
            </label>
            <label>Visão
              <select id="reports-view-filter">
                <option value="detalhada" ${state.filters.reportsView === "detalhada" ? "selected" : ""}>Detalhada</option>
                <option value="sintetica" ${state.filters.reportsView === "sintetica" ? "selected" : ""}>Sintética</option>
              </select>
            </label>
            <label>Setor
              <select id="reports-sector-filter">
                <option value="all" ${state.filters.reportsSector === "all" ? "selected" : ""}>0800 + Nuvidio</option>
                <option value="0800" ${state.filters.reportsSector === "0800" ? "selected" : ""}>0800</option>
                <option value="nuvidio" ${state.filters.reportsSector === "nuvidio" ? "selected" : ""}>Nuvidio</option>
              </select>
            </label>
            <button class="btn" data-action="refresh-reports">Aplicar</button>
            <button class="btn-secondary" data-action="reset-reports">Limpar</button>
          </div>
          <div class="mini-grid">
            <div class="mini-card"><span class="muted">Escopo</span><div class="metric-value">${esc(selectedName)}</div></div>
            <div class="mini-card"><span class="muted">Tipo</span><div class="metric-value">${esc(reportTypeLabel)}</div></div>
            <div class="mini-card"><span class="muted">Visão</span><div class="metric-value">${esc(reportViewLabel)}</div></div>
            <div class="mini-card"><span class="muted">Setor</span><div class="metric-value">${esc(sectorLabel)}</div></div>
            <div class="mini-card"><span class="muted">Base operacional</span><div class="metric-value">${integer(reportRows.length)}</div></div>
            <div class="mini-card"><span class="muted">Qualidade</span><div class="metric-value">${integer(qualityRows.length)}</div></div>
            <div class="mini-card"><span class="muted">Período</span><div class="metric-value">${esc(formatDateBr(state.filters.start) || "Todos")} - ${esc(formatDateBr(state.filters.end) || "Todos")}</div></div>
          </div>
        </article>
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Arquivos</span>
              <h3>Exportar relatório</h3>
            </div>
          </div>
          <div class="info-box">Os relatórios usam o operador selecionado no topo, o período, o tipo, a visão e o setor definidos nesta aba.</div>
          <div class="action-grid reports-actions">
            <button class="btn" type="button" data-action="export-report-excel">Exportar Excel</button>
            <button class="btn-secondary" type="button" data-action="export-report-pdf">Exportar PDF</button>
          </div>
        </article>
      </div>
      <article class="panel">
        <div class="panel-head">
          <div>
            <span class="eyebrow">Conteúdo</span>
            <h3>O que será enviado no relatório</h3>
          </div>
        </div>
        <div class="mini-grid">
          <div class="mini-card"><span class="muted">Consolidado</span><div class="helper">Resumo por operador e operação com produção média, efetividade e volume.</div></div>
          <div class="mini-card"><span class="muted">Operacional</span><div class="helper">Base linha a linha por data, operador, setor, produção e efetividade.</div></div>
          <div class="mini-card"><span class="muted">Qualidade</span><div class="helper">Mês, operador, esteira, nota bruta e observações.</div></div>
          <div class="mini-card"><span class="muted">Ofensores</span><div class="helper">Nota, produção, efetividade, qualidade e principais alertas do período.</div></div>
        </div>
      </article>
      <article class="panel report-preview-panel">
        <div class="panel-head">
          <div>
            <span class="eyebrow">Pré-visualização</span>
            <h3>Como o relatório será exportado</h3>
          </div>
        </div>
        <div class="report-preview-shell">
          <div class="report-preview-header">
            <div class="report-preview-brand">
              <div class="report-preview-badge">KR</div>
              <div>
                <strong>PORTAL DE RESULTADOS</strong>
                <span>Performance operacional</span>
              </div>
            </div>
            <div class="report-preview-meta">
              <strong>${esc(reportTypeLabel)}</strong>
              <span>${esc(reportViewLabel)} · ${esc(sectorLabel)}</span>
            </div>
          </div>
          <div class="report-preview-summary">
            <div><span>Escopo</span><strong>${esc(selectedName)}</strong></div>
            <div><span>Período</span><strong>${esc(formatDateBr(state.filters.start) || "Todos")} - ${esc(formatDateBr(state.filters.end) || "Todos")}</strong></div>
          </div>
          ${previewBody}
        </div>
      </article>
    </section>
  `;
}

function detailedTemplate() {
  const DETAILED_PAGE_SIZE = 200;
  const reportType = state.filters.reportsType;
  const reportView = state.filters.reportsView;
  const model = state.reportDataset;
  if (!model) {
    return `
      <section class="section">
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Consulta</span>
              <h3>Detalhada</h3>
            </div>
          </div>
          <div class="empty">Carregando consulta detalhada...</div>
        </article>
      </section>
    `;
  }
  const selectedName = model.selectedName;
  const typeLabel = model.reportTypeLabel;
  const sectorLabel = model.sectorLabel;
  let tableHead = "";
  let tableBody = "";
  let totalRows = 0;
  let allRows = [];

  if (reportType === "qualidade") {
    allRows = model.qualityPreparedRows;
    totalRows = allRows.length;
    tableHead = `
      <tr>
        <th>${reportSortHeader("qualidade", "reference_month", "Mês")}</th>
        <th>${reportSortHeader("qualidade", "operator", "Operador")}</th>
        <th>${reportSortHeader("qualidade", "quality_scope", "Esteira")}</th>
        <th>${reportSortHeader("qualidade", "gross_score", "Nota")}</th>
        <th>Observações</th>
      </tr>`;
  } else if (reportType === "ofensores") {
    allRows = model.offendersRows;
    totalRows = allRows.length;
    tableHead = `
      <tr>
        <th>${reportSortHeader("ofensores", "name", "Operador")}</th>
        <th>${reportSortHeader("ofensores", "alert_score", "Nota")}</th>
        <th>${reportSortHeader("ofensores", "avg_production_0800", "Prod. 0800")}</th>
        <th>${reportSortHeader("ofensores", "effectiveness_0800", "Efet. 0800")}</th>
        <th>${reportSortHeader("ofensores", "avg_production_nuvidio", "Prod. Nuvidio")}</th>
        <th>${reportSortHeader("ofensores", "effectiveness_nuvidio", "Efet. Nuvidio")}</th>
        <th>${reportSortHeader("ofensores", "quality", "Qualidade")}</th>
        <th>Principais ofensores</th>
      </tr>`;
  } else if (reportType === "consolidado") {
    allRows = model.consolidatedRows;
    totalRows = allRows.length;
    tableHead = `
      <tr>
        <th>${reportSortHeader("consolidado", "date", "Data")}</th>
        <th>${reportSortHeader("consolidado", "operator", "Operador")}</th>
        <th>${reportSortHeader("consolidado", "operation", "Setor")}</th>
        <th>${reportSortHeader("consolidado", "production", "Produção")}</th>
        <th>${reportSortHeader("consolidado", "effectiveness", "Efetividade")}</th>
      </tr>`;
  } else {
    allRows = model.operationalRows;
    totalRows = allRows.length;
    tableHead = `
      <tr>
        <th>${reportSortHeader("operacional", "date", "Data")}</th>
        <th>${reportSortHeader("operacional", "operator", "Operador")}</th>
        <th>${reportSortHeader("operacional", "operation", "Setor")}</th>
        <th>${reportSortHeader("operacional", "production", "Produção")}</th>
        <th>${reportSortHeader("operacional", "effectiveness", "Efetividade")}</th>
      </tr>`;
  }

  const totalPages = Math.max(1, Math.ceil(totalRows / DETAILED_PAGE_SIZE));
  const currentPage = Math.min(Math.max(1, Number(state.filters.detailedPage || 1)), totalPages);
  state.filters.detailedPage = currentPage;
  const startIndex = (currentPage - 1) * DETAILED_PAGE_SIZE;
  const endIndex = startIndex + DETAILED_PAGE_SIZE;
  const visibleRows = allRows.slice(startIndex, endIndex);
  const visibleFrom = totalRows ? startIndex + 1 : 0;
  const visibleTo = totalRows ? Math.min(endIndex, totalRows) : 0;

  if (reportType === "qualidade") {
    tableBody = visibleRows.length ? visibleRows.map((row) => `
      <tr>
        <td>${esc(formatMonthLabel(row.reference_month))}</td>
        <td>${esc(row.operator)}</td>
        <td>${esc(formatQualityScopeLabel(row.quality_scope || "all"))}</td>
        <td>${esc(number(row.gross_score ?? row.score ?? 0))}</td>
        <td>${esc(row.notes || "—")}</td>
      </tr>`).join("") : `<tr><td colspan="5">Sem dados para visualização.</td></tr>`;
  } else if (reportType === "ofensores") {
    tableBody = visibleRows.length ? visibleRows.map((row) => `
      <tr>
        <td>${esc(row.name)}</td>
        <td>${esc(number(row.alert_score))}</td>
        <td>${esc(integer(row.avg_production_0800))}</td>
        <td>${esc(percent(row.effectiveness_0800))}</td>
        <td>${esc(integer(row.avg_production_nuvidio))}</td>
        <td>${esc(percent(row.effectiveness_nuvidio))}</td>
        <td>${esc(row.quality ? number(row.quality) : "--")}</td>
        <td>${esc((row.reasons || []).map((reason) => reason.text).join(" | ") || "Sem ofensores destacados")}</td>
      </tr>`).join("") : `<tr><td colspan="8">Sem dados para visualização.</td></tr>`;
  } else if (reportType === "consolidado") {
    tableBody = visibleRows.length ? visibleRows.map((row) => `
      <tr>
        <td>${esc(formatDateBr(row.date))}</td>
        <td>${esc(row.operator)}</td>
        <td>${esc(row.operation)}</td>
        <td>${esc(integer(row.production))}</td>
        <td>${esc(percent(row.effectiveness || 0))}</td>
      </tr>`).join("") : `<tr><td colspan="5">Sem dados para visualização.</td></tr>`;
  } else {
    tableBody = visibleRows.length ? visibleRows.map((row) => `
      <tr>
        <td>${esc(formatDateBr(row.date))}</td>
        <td>${esc(row.operator)}</td>
        <td>${esc(row.operation)}</td>
        <td>${esc(integer(row.production || 0))}</td>
        <td>${esc(percent(row.effectiveness || 0))}</td>
      </tr>`).join("") : `<tr><td colspan="5">Sem dados para visualização.</td></tr>`;
  }

  return `
    <section class="section detailed-dashboard">
      <article class="panel analysis-filter-panel detailed-filter-panel">
        <div class="analysis-filter-bar detailed-filter-bar">
          <div class="analysis-filter-copy">
            <span class="eyebrow">Consulta</span>
            <h3>Detalhada</h3>
          </div>
          <label class="analysis-inline-field">Início
            <input type="date" id="detailed-start-filter" value="${state.filters.start}">
          </label>
          <label class="analysis-inline-field">Fim
            <input type="date" id="detailed-end-filter" value="${state.filters.end}">
          </label>
          <label class="analysis-inline-field">Tipo
            <select id="detailed-type-filter">
              <option value="consolidado" ${reportType === "consolidado" ? "selected" : ""}>Consolidado</option>
              <option value="operacional" ${reportType === "operacional" ? "selected" : ""}>Operacional</option>
              <option value="qualidade" ${reportType === "qualidade" ? "selected" : ""}>Qualidade</option>
              <option value="ofensores" ${reportType === "ofensores" ? "selected" : ""}>Ofensores</option>
            </select>
          </label>
          <label class="analysis-inline-field">Setor
            <select id="detailed-sector-filter">
              <option value="all" ${state.filters.reportsSector === "all" ? "selected" : ""}>0800 + Nuvidio</option>
              <option value="0800" ${state.filters.reportsSector === "0800" ? "selected" : ""}>0800</option>
              <option value="nuvidio" ${state.filters.reportsSector === "nuvidio" ? "selected" : ""}>Nuvidio</option>
            </select>
          </label>
          <div class="analysis-filter-actions">
            <button class="btn" data-action="refresh-detailed">Aplicar</button>
            <button class="btn-secondary" data-action="reset-detailed">Limpar</button>
          </div>
        </div>
      </article>
      <article class="panel detailed-table-panel">
        <div class="panel-head">
          <div>
            <span class="eyebrow">${esc(typeLabel)}</span>
            <h3>Comparativo detalhado</h3>
          </div>
          <div class="detailed-table-meta">
            <span>${esc(selectedName)}</span>
            <strong>${integer(totalRows)} linha(s)</strong>
            <span>Exibindo ${integer(visibleFrom)}-${integer(visibleTo)}</span>
            <span>${esc(sectorLabel)}</span>
            <span>${esc(reportView === "sintetica" ? "Sintética" : "Detalhada")}</span>
          </div>
        </div>
        <div class="table-wrap detailed-table-wrap">
          <table class="report-preview-table detailed-table detailed-table--${esc(reportType)}">
            <thead>${tableHead}</thead>
            <tbody>${tableBody}</tbody>
          </table>
        </div>
        ${totalPages > 1 ? `
          <div class="detailed-pagination">
            <button class="btn-secondary" type="button" data-action="detailed-prev-page" ${currentPage <= 1 ? "disabled" : ""}>Anterior</button>
            <span>Página ${integer(currentPage)} de ${integer(totalPages)}</span>
            <button class="btn-secondary" type="button" data-action="detailed-next-page" ${currentPage >= totalPages ? "disabled" : ""}>Próxima</button>
          </div>
        ` : ""}
      </article>
    </section>
  `;
}

function historyAvailableDates() {
  const dates = (state.history?.history || [])
    .map((row) => String(row.metric_date || row.date || "").trim())
    .filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(value));
  return [...new Set(dates)].sort((a, b) => b.localeCompare(a));
}

function historyAvailableQualityMonths() {
  const months = (state.history?.quality || [])
    .map((row) => String(row.reference_month || "").trim())
    .filter((value) => /^\d{4}-\d{2}$/.test(value));
  return [...new Set(months)].sort((a, b) => b.localeCompare(a));
}

function compareReportValues(left, right, direction = "asc") {
  const multiplier = direction === "desc" ? -1 : 1;
  const a = left ?? "";
  const b = right ?? "";
  const aNumber = typeof a === "number" ? a : Number(a);
  const bNumber = typeof b === "number" ? b : Number(b);
  const aIsNumber = Number.isFinite(aNumber) && String(a).trim() !== "";
  const bIsNumber = Number.isFinite(bNumber) && String(b).trim() !== "";
  if (aIsNumber && bIsNumber) {
    return (aNumber - bNumber) * multiplier;
  }
  return String(a).localeCompare(String(b), "pt-BR", { sensitivity: "base", numeric: true }) * multiplier;
}

function getReportSort(scope) {
  return state.reportSorts?.[scope] || { column: "", direction: "asc" };
}

function sortReportRows(scope, rows) {
  const list = Array.isArray(rows) ? [...rows] : [];
  const sort = getReportSort(scope);
  const column = String(sort.column || "").trim();
  const direction = sort.direction === "desc" ? "desc" : "asc";
  if (!column) return list;
  return list.sort((a, b) => compareReportValues(a?.[column], b?.[column], direction));
}

function reportSortHeader(scope, column, label) {
  const current = getReportSort(scope);
  const active = current.column === column;
  const marker = active ? (current.direction === "desc" ? "↓" : "↑") : "↕";
  return `<button class="report-sort-btn ${active ? "is-active" : ""}" type="button" data-report-sort-scope="${esc(scope)}" data-report-sort-column="${esc(column)}">${esc(label)} <span>${marker}</span></button>`;
}

function invalidateReportDatasetCache() {
  reportDatasetCache.key = "";
  reportDatasetCache.value = null;
}

function buildReportDatasetModel() {
  const key = JSON.stringify({
    route: state.route,
    userId: state.filters.analysisUserId,
    sector: state.filters.reportsSector,
    type: state.filters.reportsType,
    view: state.filters.reportsView,
    start: state.filters.start,
    end: state.filters.end,
    historyCount: state.history?.history?.length || 0,
    qualityCount: state.history?.quality?.length || 0,
    alertsCount: state.alerts?.alerts?.length || 0,
    sortConsolidado: state.reportSorts?.consolidado,
    sortOperacional: state.reportSorts?.operacional,
    sortQualidade: state.reportSorts?.qualidade,
    sortOfensores: state.reportSorts?.ofensores,
  });
  if (reportDatasetCache.key === key && reportDatasetCache.value) {
    return reportDatasetCache.value;
  }

  const reportType = state.filters.reportsType;
  const reportView = state.filters.reportsView;
  const selectedName = state.filters.analysisUserId === "all"
    ? "Todos os operadores"
    : (getAnyOperatorUsers().find((user) => String(user.id) === String(state.filters.analysisUserId))?.full_name || "Operador");
  const reportRows = getScopedHistory().filter((row) => (
    state.filters.reportsSector === "all"
    || String(row.operation || "").toLowerCase() === String(state.filters.reportsSector || "").toLowerCase()
  ));
  const qualityRows = state.history?.quality || [];
  const sectorLabel = state.filters.reportsSector === "0800"
    ? "0800"
    : state.filters.reportsSector === "nuvidio"
      ? "Nuvidio"
      : "0800 + Nuvidio";
  const reportTypeLabel = {
    consolidado: "Consolidado",
    operacional: "Operacional",
    qualidade: "Qualidade",
    ofensores: "Ofensores",
  }[reportType] || "Consolidado";
  const reportViewLabel = reportView === "sintetica" ? "Sintética" : "Detalhada";
  const operationalRows = reportRows.map((row) => ({
    ...row,
    operator: row.operator || getUserLabelById(row.userId) || "Operador",
  }));
  const preparedQualityRows = qualityRows.map((row) => ({
    ...row,
    score_type: row.score_type || "monitorias",
    quality_scope: row.quality_scope || "all",
    gross_score: String(row.score_type || "monitorias") === "general" ? Number(row.score || 0) : null,
    operator: row.operator || getUserLabelById(row.user_id) || "Operador",
  }));
  const consolidatedMap = new Map();
  operationalRows.forEach((row) => {
    const operatorName = row.operator || "Operador";
    const keyPart = `${row.date}::${operatorName}::${row.operation}`;
    const current = consolidatedMap.get(keyPart) || {
      date: row.date,
      operator: operatorName,
      operation: row.operation,
      totalProduction: 0,
      effectivenessParts: [],
    };
    current.totalProduction += Number(row.production || 0);
    if (Number(row.effectiveness || 0) > 0) current.effectivenessParts.push(Number(row.effectiveness || 0));
    consolidatedMap.set(keyPart, current);
  });
  const consolidatedRows = [...consolidatedMap.values()].map((row) => {
    const avgEffectiveness = row.effectivenessParts.length
      ? row.effectivenessParts.reduce((sum, value) => sum + value, 0) / row.effectivenessParts.length
      : 0;
    return {
      date: row.date,
      operator: row.operator,
      operation: row.operation,
      production: row.totalProduction,
      effectiveness: avgEffectiveness,
    };
  });

  const model = {
    selectedName,
    sectorLabel,
    reportTypeLabel,
    reportViewLabel,
    reportRows,
    qualityRows,
    operationalRows: sortReportRows("operacional", operationalRows),
    qualityPreparedRows: sortReportRows("qualidade", preparedQualityRows),
    offendersRows: sortReportRows("ofensores", [...(state.alerts?.alerts || [])]),
    consolidatedRows: sortReportRows("consolidado", consolidatedRows),
  };

  reportDatasetCache.key = key;
  reportDatasetCache.value = model;
  return model;
}

function historyTemplate() {
  const rows = getScopedHistory();
  const currentHistoryUser = getUserLabelById(state.filters.historyUserId);
  const showActionsColumn = isManager() || rows.some((row) => row.entryType === "quality");
  const historyInputValue = isManager()
    ? (state.filters.historyUserId === "all" && !String(state.filters.historyQuery || "").trim()
      ? "Todos os operadores"
      : (state.filters.historyQuery || currentHistoryUser))
    : (state.filters.historyQuery || currentHistoryUser);
  const showOperatorColumn = isManager() && (state.filters.historyUserId === "all" || !String(state.filters.historyQuery || "").trim());
  const userNameById = new Map((state.users || []).map((user) => [String(user.id), user.full_name]));
  return `
    <section class="section">
      <article class="table-card">
        <div class="history-header">
          <div>
            <span class="eyebrow">Histórico</span>
            <h3>Resultados por data</h3>
          </div>
          <div class="filter-actions history-tools">
            ${isManager() ? `
              <label>Operador
                <input id="history-user-search" list="history-user-options" value="${esc(historyInputValue)}" placeholder="Pesquisar operador">
                <datalist id="history-user-options">
                  ${getOperatorUsers().map((user) => `<option value="${esc(user.full_name)}"></option>`).join("")}
                </datalist>
              </label>
              <button class="btn-secondary" type="button" id="open-history-delete-day-modal">Excluir dia</button>
            ` : ""}
            <button class="btn" data-action="refresh-history">Atualizar</button>
          </div>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                ${showOperatorColumn ? "<th>Operador</th>" : ""}
                <th>Data</th>
                <th>Operação</th>
                <th>Produção</th>
                <th>Efetividade</th>
                <th>Qualidade</th>
                <th>Atualizado</th>
                ${showActionsColumn ? "<th>Ações</th>" : ""}
              </tr>
            </thead>
            <tbody>
              ${rows.length ? rows.map((row) => `
                <tr>
                  ${showOperatorColumn ? `<td>${esc(userNameById.get(String(row.userId)) || "—")}</td>` : ""}
                  <td>${esc(row.dateLabel || formatDateBr(row.date))}</td>
                  <td><span class="pill ${row.operation === "0800" ? "amber" : row.operation === "Qualidade" ? "green" : "blue"}">${esc(row.operation)}</span></td>
                  <td>${row.production === null ? "—" : integer(row.production)}</td>
                  <td>${row.effectiveness === null ? "—" : percent(row.effectiveness)}</td>
                  <td>${number(row.quality)}</td>
                  <td>${esc(formatDateTimeBr(row.updatedAt))}</td>
                  ${showActionsColumn ? `
                    <td>
                      <div class="row-actions">
                        ${row.entryType === "quality" ? `<button class="btn-secondary btn-small" type="button" data-history-view="${row.metricId}" data-history-type="${row.entryType}">Ver</button>` : ""}
                        ${isManager() ? `<button class="btn-secondary btn-small" type="button" data-history-edit="${row.metricId}" data-history-operation="${row.operation}" data-history-type="${row.entryType}">Editar</button>` : ""}
                        ${isManager() ? `<button class="btn-secondary btn-small danger-outline" type="button" data-history-delete="${row.metricId}" data-history-operation="${row.operation}" data-history-type="${row.entryType}" data-history-user-id="${row.userId ?? ""}" data-history-reference-month="${row.referenceMonth ?? ""}" data-history-quality-scope="${row.qualityScope ?? ""}">Remover</button>` : ""}
                      </div>
                    </td>
                  ` : ""}
                </tr>
              `).join("") : `<tr><td colspan="${6 + (showOperatorColumn ? 1 : 0) + (showActionsColumn ? 1 : 0)}"><div class="empty">Sem resultados.</div></td></tr>`}
            </tbody>
          </table>
        </div>
      </article>
    </section>
  `;
}

function adminTemplate() {
  const currentReferenceMonth = new Date().toISOString().slice(0, 7);
  const [defaultYear, defaultMonth] = currentReferenceMonth.split("-");
  const operatorUsers = getOperatorUsers();
  const managementUsers = getFilteredUsers();
  const hasAnyUsers = state.users.length > 0;
  return `
    <section class="section">
      <div class="hero-grid">
        <article class="panel">
          <div class="panel-head"><div><span class="eyebrow">Operação</span><h3>Modo manutenção</h3></div></div>
          <form id="admin-maintenance-form" class="section">
            <label class="switch-row">
              <input type="checkbox" id="maintenance-toggle" ${state.appSettings.maintenance_for_operators ? "checked" : ""}>
              <span>Ativar manutenção para operadores</span>
            </label>
            <label>Mensagem para operadores
              <input id="maintenance-message" value="${esc(state.appSettings.maintenance_message || "")}" maxlength="220">
            </label>
            <div class="mini-grid">
              <div class="mini-card">
                <span class="muted">Produção</span>
                <div class="form-grid">
                  <label>Vermelho até <input type="number" id="metric-production-red" value="${Number(state.appSettings?.metric_rules?.production?.red_max ?? 70)}"></label>
                  <label>Laranja até <input type="number" id="metric-production-amber" value="${Number(state.appSettings?.metric_rules?.production?.amber_max ?? 100)}"></label>
                </div>
              </div>
              <div class="mini-card">
                <span class="muted">Efetividade</span>
                <div class="form-grid">
                  <label>Vermelho até <input type="number" id="metric-effectiveness-red" value="${Number(state.appSettings?.metric_rules?.effectiveness?.red_max ?? 70)}"></label>
                  <label>Laranja até <input type="number" id="metric-effectiveness-amber" value="${Number(state.appSettings?.metric_rules?.effectiveness?.amber_max ?? 90)}"></label>
                </div>
              </div>
              <div class="mini-card">
                <span class="muted">Qualidade</span>
                <div class="form-grid">
                  <label>Vermelho até <input type="number" id="metric-quality-red" value="${Number(state.appSettings?.metric_rules?.quality?.red_max ?? 70)}"></label>
                  <label>Laranja até <input type="number" id="metric-quality-amber" value="${Number(state.appSettings?.metric_rules?.quality?.amber_max ?? 90)}"></label>
                </div>
              </div>
            </div>
            <div class="mini-grid">
              <div class="mini-card">
                <span class="muted">Ofensores · crítico até</span>
                <div class="form-grid">
                  <label>Produção Nuvidio <input type="number" id="alert-production-nuvidio" value="${Number(alertRules().production_nuvidio.critical_min ?? 70)}"></label>
                  <label>Produção 0800 <input type="number" id="alert-production-0800" value="${Number(alertRules().production_0800.critical_min ?? 70)}"></label>
                  <label>Efetividade 0800 <input type="number" id="alert-effectiveness-0800" value="${Number(alertRules().effectiveness_0800.critical_min ?? 70)}"></label>
                  <label>Efetividade Nuvidio <input type="number" id="alert-effectiveness-nuvidio" value="${Number(alertRules().effectiveness_nuvidio.critical_min ?? 70)}"></label>
                  <label>Qualidade <input type="number" id="alert-quality" value="${Number(alertRules().quality.critical_min ?? 70)}"></label>
                </div>
              </div>
            </div>
            <div class="action-grid">
              <button class="btn" type="submit">Salvar metas</button>
            </div>
          </form>
        </article>
        <article class="panel">
          <div class="management-header">
            <div>
              <span class="eyebrow">Gestão</span>
              <h3>Base principal</h3>
            </div>
          </div>
          <div class="panel-head"><div><span class="eyebrow">Base operacional</span><h3>Carga por planilha</h3></div></div>
          <form id="base-import-form" class="section">
            <div class="info-box">Baixe o modelo, preencha e anexe o arquivo CSV para importar a base.</div>
            <label>Modelo da planilha
              <select name="model" id="base-model-select">
                <option value="nuvidio">Nuvidio</option>
                <option value="0800">0800</option>
              </select>
            </label>
            <label>Arquivo base (CSV)<input type="file" name="file" accept=".csv" required></label>
            <div class="action-grid">
              <button class="btn-secondary" type="button" id="download-base-template">Baixar modelo</button>
              <button class="btn" type="submit">Anexar arquivo</button>
            </div>
          </form>
        </article>
        <article class="panel">
          <div class="panel-head"><div><span class="eyebrow">Lançamento manual</span><h3>Campos de tag</h3></div></div>
          <form id="manual-tag-form" class="section">
            <div class="form-grid">
              <label>Operador
                <select name="user_id" required>
                  <option value="">Selecione</option>
                  ${operatorUsers.map((user) => `<option value="${user.id}">${esc(user.full_name)}</option>`).join("")}
                </select>
              </label>
              <label>Data<input type="date" name="date" value="${state.filters.today}" required></label>
              <label>Operação
                <select name="operation" required>
                  <option value="0800">0800</option>
                  <option value="nuvidio">Nuvidio</option>
                </select>
              </label>
              <label>Aprovado<input type="number" min="0" step="1" name="approved" value="0" required></label>
              <label>Reprovado<input type="number" min="0" step="1" name="rejected" value="0" required></label>
              <label>Sem ação<input type="number" min="0" step="1" name="no_action" value="0" required></label>
              <label class="pending-only">Pendenciado<input type="number" min="0" step="1" name="pending" value="0"></label>
              <label class="empty-only">Vazio<input type="number" min="0" step="1" name="empty" value="0"></label>
            </div>
            <div class="action-grid">
              <button class="btn" type="submit">Salvar lançamento</button>
            </div>
          </form>
        </article>
      </div>
      <div class="hero-grid">
        <article class="panel">
          <div class="panel-head"><div><span class="eyebrow">Qualidade</span><h3>Referência mensal</h3></div></div>
          <form id="quality-upload-form" class="section">
            <input type="hidden" name="reference_month" id="quality-reference-month" value="${currentReferenceMonth}">
            <input type="hidden" name="quality_mode" value="general">
            <div class="form-grid">
              <label>Esteira
                <select name="quality_scope" id="quality-upload-scope-select">
                  <option value="0800">0800</option>
                  <option value="nuvidio">Nuvidio</option>
                </select>
              </label>
              <label>Mês
                <select id="quality-reference-month-select">
                  ${monthOptions(defaultMonth)}
                </select>
              </label>
              <label>Ano
                <select id="quality-reference-year-select">
                  ${yearOptions(defaultYear)}
                </select>
              </label>
            </div>
            <label>Arquivo<input type="file" name="file" accept=".xlsx,.csv" required></label>
            <div class="action-grid">
              <button class="btn-secondary" type="button" id="download-quality-template">Baixar modelo</button>
              <button class="btn" type="submit">Importar qualidade</button>
            </div>
          </form>
        </article>
        <article class="panel">
          <div class="panel-head"><div><span class="eyebrow">Qualidade</span><h3>Lançamento manual</h3></div></div>
          <form id="quality-manual-form" class="section">
            <div class="form-grid">
              <label>Operador
                <select name="user_id" required>
                  <option value="">Selecione</option>
                  ${operatorUsers.map((user) => `<option value="${user.id}">${esc(user.full_name)}</option>`).join("")}
                </select>
              </label>
              <label>Esteira
                <select name="quality_scope" required>
                  <option value="0800">0800</option>
                  <option value="nuvidio">Nuvidio</option>
                </select>
              </label>
              <label>Mês
                <select id="quality-manual-month-select" required>
                  ${monthOptions(defaultMonth)}
                </select>
              </label>
              <label>Ano
                <select id="quality-manual-year-select" required>
                  ${yearOptions(defaultYear)}
                </select>
              </label>
              <label>Nota bruta do mês<input type="number" min="0" max="100" step="0.01" name="score" required></label>
              <label>Observações<input name="notes" maxlength="220" placeholder="Opcional"></label>
            </div>
            <input type="hidden" name="reference_month" id="quality-manual-reference-month" value="${currentReferenceMonth}">
            <div class="action-grid">
              <button class="btn" type="submit">Salvar qualidade</button>
            </div>
          </form>
        </article>
      </div>
      <div class="hero-grid">
        <article class="panel">
          <div class="panel-head"><div><span class="eyebrow">Usuários</span><h3>Novo usuário</h3></div></div>
          <form id="admin-user-form" class="section">
            <div class="form-grid">
              <label>Nome completo<input name="full_name" required></label>
              <label>Login<input name="login" required></label>
              <label>Perfil
                <select name="role" required>
                  <option value="operator">Operador</option>
                  <option value="manager">Gestor</option>
                </select>
              </label>
              <label>ID 0800<input name="platform_0800_id"></label>
              <label>ID Nuvidio<input name="nuvidio_id"></label>
            </div>
            <div class="info-box">Senha inicial automática: <strong>Trocar@01</strong>. No primeiro login, o usuário será obrigado a definir uma nova senha.</div>
            <div class="action-grid">
              <button class="btn" type="submit">Cadastrar usuário</button>
            </div>
          </form>
        </article>
        <article class="panel">
          <div class="panel-head"><div><span class="eyebrow">Usuários</span><h3>Cadastrados</h3></div></div>
          <div class="management-user-search">
            <label>Pesquisar usuário
              <input id="management-user-search" value="${esc(state.filters.usersQuery)}" placeholder="Nome, login ou perfil">
            </label>
          </div>
          <div class="list">
            ${managementUsers.length ? managementUsers.map((user) => `
              <div class="list-row" data-management-user-search="${esc(`${user.full_name} ${user.login} ${user.role === "manager" ? "gestor" : "operador"}`)}">
                <div class="list-row-copy">
                  <strong>${esc(user.full_name)}</strong>
                  <span>${esc(user.login)} · ${user.role === "manager" ? "Gestor" : "Operador"}</span>
                </div>
                <div class="list-row-actions">
                  <button class="pill ${user.is_active ? "green" : "red"} user-toggle-btn" type="button" data-user-toggle="${user.id}">
                    ${user.is_active ? "Ativo" : "Inativo"}
                  </button>
                  <button class="btn-secondary btn-small" type="button" data-user-edit="${user.id}">Editar</button>
                  <button class="btn-secondary btn-small" type="button" data-user-reset="${user.id}">Resetar senha</button>
                  <button class="btn-secondary btn-small danger-outline" type="button" data-user-delete="${user.id}">Apagar</button>
                </div>
              </div>
            `).join("") : ""}
            <div class="empty" id="management-users-empty" ${managementUsers.length ? "hidden" : ""}>${hasAnyUsers ? "Nenhum usuário encontrado." : "Nenhum usuário cadastrado."}</div>
          </div>
        </article>
      </div>
    </section>
  `;
}
function bindLogin() {
  enhancePasswordFields();
  const loginForm = document.getElementById("login-form");
  const selfResetModal = document.getElementById("self-reset-modal");
  const selfResetForm = document.getElementById("self-reset-form");
  const closeSelfResetModal = () => {
    if (!selfResetModal) return;
    selfResetModal.hidden = true;
    if (selfResetForm) selfResetForm.reset();
  };

  document.getElementById("open-self-reset-modal")?.addEventListener("click", () => {
    if (selfResetModal) selfResetModal.hidden = false;
  });
  document.querySelectorAll("#close-self-reset-modal, #cancel-self-reset-modal").forEach((button) => {
    button.addEventListener("click", closeSelfResetModal);
  });
  if (selfResetModal) {
    selfResetModal.addEventListener("click", (event) => {
      if (event.target === selfResetModal) closeSelfResetModal();
    });
  }
  if (selfResetForm) {
    selfResetForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const submitButton = event.currentTarget.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Salvando...");
      try {
        await api("/api/auth/reset-self", {
          method: "POST",
          body: JSON.stringify(Object.fromEntries(new FormData(event.currentTarget).entries())),
        });
        closeSelfResetModal();
        setFlash("success", "Senha redefinida com sucesso. Faça login com a nova senha.");
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  loginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const submitButton = event.currentTarget.querySelector('button[type="submit"]');
    const originalButtonText = submitButton?.textContent || "Acessar";
    try {
      clearFlash();
      if (submitButton) {
        submitButton.disabled = true;
        submitButton.textContent = "Verificando credenciais...";
      }
      const response = await api("/api/auth/login", {
        method: "POST",
        body: JSON.stringify({ login: form.get("login"), password: form.get("password") }),
      });
      state.user = normalizeUserPayload(response.user);
      if (response.app_settings) state.appSettings = response.app_settings;
      applyUserPreferences();
      state.forcePasswordChange = Boolean(state.user.must_change_password);
      state.filters.historyUserId = isManager() ? "all" : String(state.user.id);
      state.filters.analysisUserId = isManager() ? "all" : String(state.user.id);
      enforceOperatorScope();
      render();
      if (!(state.user.role !== "manager" && state.appSettings.maintenance_for_operators)) {
        loadAll()
          .then(() => {
            state.filters.historyQuery = isManager() ? "" : (getUserLabelById(state.filters.historyUserId) || state.user.full_name);
            render();
          })
          .catch((error) => {
            setFlash("error", error.message || "Não foi possível carregar os dados iniciais.");
          });
      }
    } catch (error) {
      setFlash("error", error.message);
    } finally {
      if (submitButton) {
        submitButton.disabled = false;
        submitButton.textContent = originalButtonText;
      }
    }
  });
}

function bindShellEvents() {
  enhancePasswordFields();
  const profileMenuTrigger = document.getElementById("profile-menu-trigger");
  const profileMenuPopover = document.getElementById("profile-menu-popover");
  const passwordModal = document.getElementById("password-modal");
  const userModal = document.getElementById("user-modal");
  const historyEditModal = document.getElementById("history-edit-modal");
  const historyViewModal = document.getElementById("history-view-modal");
  const historyDeleteDayModal = document.getElementById("history-delete-day-modal");
  const trendTooltip = document.getElementById("trend-tooltip");
  const analysisMetricTooltip = document.getElementById("analysis-metric-tooltip");

  const closeProfileMenu = () => {
    if (!profileMenuPopover || !profileMenuTrigger) return;
    profileMenuPopover.hidden = true;
    profileMenuTrigger.setAttribute("aria-expanded", "false");
  };

  const openPasswordModal = () => {
    if (!passwordModal) return;
    passwordModal.hidden = false;
    closeProfileMenu();
  };

  const closePasswordModal = () => {
    if (state.forcePasswordChange) return;
    if (!passwordModal) return;
    passwordModal.hidden = true;
    const form = document.getElementById("password-form");
    if (form) form.reset();
  };

  const closeUserModal = () => {
    if (!userModal) return;
    userModal.hidden = true;
    const form = document.getElementById("user-edit-form");
    if (form) form.reset();
  };

  const closeHistoryEditModal = () => {
    if (!historyEditModal) return;
    historyEditModal.hidden = true;
    const form = document.getElementById("history-edit-form");
    if (form) form.reset();
  };

  const closeHistoryViewModal = () => {
    if (!historyViewModal) return;
    historyViewModal.hidden = true;
    const content = document.getElementById("history-view-content");
    if (content) content.innerHTML = "";
  };

  const closeHistoryDeleteDayModal = () => {
    if (!historyDeleteDayModal) return;
    historyDeleteDayModal.hidden = true;
    const form = document.getElementById("history-delete-day-form");
    if (form) form.reset();
  };

  const hideMetricPopovers = () => {
    if (trendTooltip) trendTooltip.hidden = true;
    if (analysisMetricTooltip) analysisMetricTooltip.hidden = true;
  };

  const attachPopoverClose = (tooltip) => {
    tooltip?.querySelector("[data-close-popover]")?.addEventListener("click", (event) => {
      event.stopPropagation();
      tooltip.hidden = true;
    });
  };

  const openUserModal = (userId) => {
    const user = state.users.find((item) => String(item.id) === String(userId));
    if (!user || !userModal) return;
    const setValue = (id, value) => {
      const input = document.getElementById(id);
      if (input) input.value = value ?? "";
    };
    setValue("edit-user-id", user.id);
    setValue("edit-full-name", user.full_name);
    setValue("edit-login", user.login);
    setValue("edit-role", user.role);
    setValue("edit-is-active", String(Boolean(user.is_active)));
    setValue("edit-platform-0800-id", user.platform_0800_id || "");
    setValue("edit-nuvidio-id", user.nuvidio_id || "");
    setValue("edit-password", "");
    userModal.hidden = false;
  };

  if (profileMenuTrigger && profileMenuPopover) {
    profileMenuTrigger.addEventListener("click", (event) => {
      event.stopPropagation();
      const isHidden = profileMenuPopover.hidden;
      profileMenuPopover.hidden = !isHidden;
      profileMenuTrigger.setAttribute("aria-expanded", String(isHidden));
    });

    document.addEventListener("click", (event) => {
      const target = event.target;
      const clickedInsideProfileMenu = target instanceof Element && target.closest(".profile-menu");
      if (!profileMenuPopover.hidden && !clickedInsideProfileMenu) {
        closeProfileMenu();
      }
    });
  }

  const openPasswordButton = document.getElementById("open-password-modal");
  if (openPasswordButton) {
    openPasswordButton.addEventListener("click", openPasswordModal);
  }

  document.querySelectorAll("#close-password-modal, #cancel-password-modal").forEach((button) => {
    button.addEventListener("click", closePasswordModal);
  });

  document.querySelectorAll("#close-user-modal, #cancel-user-modal").forEach((button) => {
    button.addEventListener("click", closeUserModal);
  });

  document.querySelectorAll("#close-history-edit-modal, #cancel-history-edit-modal").forEach((button) => {
    button.addEventListener("click", closeHistoryEditModal);
  });

  document.querySelectorAll("#close-history-view-modal, #cancel-history-view-modal").forEach((button) => {
    button.addEventListener("click", closeHistoryViewModal);
  });

  document.querySelectorAll("#close-history-delete-day-modal, #cancel-history-delete-day-modal").forEach((button) => {
    button.addEventListener("click", closeHistoryDeleteDayModal);
  });

  if (passwordModal) {
    passwordModal.addEventListener("click", (event) => {
      if (!state.forcePasswordChange && event.target === passwordModal) closePasswordModal();
    });

    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && !passwordModal.hidden && !state.forcePasswordChange) {
        closePasswordModal();
      }
      if (event.key === "Escape" && userModal && !userModal.hidden) {
        closeUserModal();
      }
      if (event.key === "Escape" && historyDeleteDayModal && !historyDeleteDayModal.hidden) {
        closeHistoryDeleteDayModal();
      }
    });
  }

  if (userModal) {
    userModal.addEventListener("click", (event) => {
      if (event.target === userModal) closeUserModal();
    });
  }
  if (historyEditModal) {
    historyEditModal.addEventListener("click", (event) => {
      if (event.target === historyEditModal) closeHistoryEditModal();
    });
  }
  if (historyDeleteDayModal) {
    historyDeleteDayModal.addEventListener("click", (event) => {
      if (event.target === historyDeleteDayModal) closeHistoryDeleteDayModal();
    });
  }

  document.getElementById("open-history-delete-day-modal")?.addEventListener("click", () => {
    if (!historyDeleteDayModal) return;
    historyDeleteDayModal.hidden = false;
  });

  const placeTooltipCentered = (tooltip) => {
    if (!tooltip) return;
    tooltip.hidden = false;
    tooltip.style.visibility = "hidden";
    const tooltipWidth = tooltip.offsetWidth || 280;
    const tooltipHeight = tooltip.offsetHeight || 120;
    const resolvedLeft = Math.max(16, (window.innerWidth - tooltipWidth) / 2);
    const resolvedTop = Math.max(16, (window.innerHeight - tooltipHeight) / 2);
    tooltip.style.left = `${resolvedLeft}px`;
    tooltip.style.top = `${resolvedTop}px`;
    tooltip.style.visibility = "visible";
  };

  if (trendTooltip && isManager()) {
    const trendPoints = document.querySelectorAll(".trend-point[data-trend-date]");
    const showTrendTooltip = async (point) => {
      const date = point?.dataset?.trendDate;
      if (!date) return;
      const key = String(date);
      if (!state.dayTopCache[key]) {
        state.dayTopCache[key] = { loading: true, top: [] };
        try {
          const response = await api(`/api/dashboard/day-top?date=${encodeURIComponent(key)}`);
          state.dayTopCache[key] = { loading: false, top: response.top || [] };
        } catch {
          state.dayTopCache[key] = { loading: false, top: [] };
        }
      }
      const bucket = state.dayTopCache[key];
      const lines = bucket.loading
        ? `<div class="trend-tooltip-line trend-tooltip-line-single">Carregando...</div>`
        : (bucket.top.length
          ? bucket.top.map((item, idx) => `
            <div class="trend-tooltip-rank">
              <div class="trend-rank-order">${idx + 1}</div>
              <div class="trend-rank-copy">
                <strong>${esc(item.name)}</strong>
                <span>${integer(item.production)} · ${percent(item.effectiveness)}</span>
              </div>
            </div>
          `).join("")
          : `<div class="trend-tooltip-line trend-tooltip-line-single">Sem dados para ${esc(formatDateBr(key))}.</div>`);
      trendTooltip.innerHTML = `
        <div class="trend-tooltip-head">
          <div class="trend-tooltip-title">Top 10 · ${esc(formatDateBr(key))}</div>
          <button class="trend-tooltip-close" type="button" data-close-popover aria-label="Fechar">×</button>
        </div>
        ${lines}
      `;
      attachPopoverClose(trendTooltip);
      placeTooltipCentered(trendTooltip);
    };

    trendPoints.forEach((point) => {
      point.addEventListener("click", async (event) => {
        event.stopPropagation();
        const alreadyVisible = !trendTooltip.hidden && trendTooltip.dataset.anchorKey === String(point.dataset.trendDate || "");
        hideMetricPopovers();
        if (alreadyVisible) return;
        trendTooltip.dataset.anchorKey = String(point.dataset.trendDate || "");
        await showTrendTooltip(point);
      });
    });
  }

  if (analysisMetricTooltip) {
    const qualitySummaryPoints = document.querySelectorAll(".quality-summary-point[data-reference-month]");
    const showQualitySummaryTooltip = (referenceMonth, qualityScope) => {
        const scopeKey = normalizeQualityScope(qualityScope || "all");
        const monthRows = (state.history?.quality || []).filter((item) =>
          String(item.reference_month || "") === String(referenceMonth || "")
          && normalizeQualityScope(item.quality_scope || item.qualityScope || "all") === scopeKey
        );
        const average = (values) => values.length ? values.reduce((sum, value) => sum + Number(value || 0), 0) / values.length : null;
        const finalScore = average(monthRows.map((item) => Number(item.score || 0)).filter((value) => Number.isFinite(value)));
        const userMap = new Map((state.users || []).map((user) => [String(user.id), repairTextEncoding(user.full_name)]));
        const useRanking = isManager() && String(state.filters.analysisUserId) === "all";
        const qualityDescription = scopeKey === "nuvidio"
          ? "Nota média refente a monitorias dos atendimentos na fila Nuvidio auditadas pela equipe de qualidade."
          : "Nota média refente a monitorias dos atendimentos na fila do 0800 auditadas pela equipe de qualidade.";
        let lines = `<div class="trend-tooltip-line trend-tooltip-line-single">Sem lançamentos neste mês.</div>`;
        if (useRanking) {
          const rankedEntries = monthRows.flatMap((item) => {
            const operatorName = userMap.get(String(item.user_id || "")) || "Operador";
            return [{
              name: operatorName,
              label: "Bruto",
              value: Number(item.quality || item.score || 0),
              detail: `Mês - ${number(item.quality || item.score || 0)}`,
            }];
          })
            .filter((item) => Number.isFinite(item.value))
            .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name, "pt-BR", { sensitivity: "base" }))
            .slice(0, 10);
          lines = rankedEntries.length
            ? rankedEntries.map((item, index) => `
                <div class="trend-tooltip-rank">
                  <div class="trend-rank-order">${index + 1}</div>
                  <div class="trend-rank-copy">
                    <strong>${esc(item.name)}</strong>
                    <span>${esc(item.detail)}</span>
                  </div>
                </div>
              `).join("")
            : lines;
        } else {
          lines = "";
        }
        analysisMetricTooltip.innerHTML = `
          <div class="trend-tooltip-head">
            <div class="trend-tooltip-title">${useRanking ? "Top 10 Qualidade" : "Qualidade"} · ${esc(formatQualityScopeLabel(scopeKey))} · ${esc(formatMonthLabel(referenceMonth || ""))}</div>
            <button class="trend-tooltip-close" type="button" data-close-popover aria-label="Fechar">×</button>
          </div>
          ${finalScore !== null ? `<div class="trend-tooltip-line"><span>Nota média do mês</span><strong>${number(finalScore)}</strong></div>` : ""}
          ${!useRanking ? `<div class="trend-tooltip-line trend-tooltip-line-single">${esc(qualityDescription)}</div>` : ""}
          ${lines}
        `;
      attachPopoverClose(analysisMetricTooltip);
      placeTooltipCentered(analysisMetricTooltip);
    };

      qualitySummaryPoints.forEach((point) => {
        point.addEventListener("click", (event) => {
          event.stopPropagation();
          const key = `quality-summary|${String(point.dataset.referenceMonth || "")}|${String(point.dataset.qualityScope || "all")}`;
          const alreadyVisible = !analysisMetricTooltip.hidden && analysisMetricTooltip.dataset.anchorKey === key;
          hideMetricPopovers();
          if (alreadyVisible) return;
          analysisMetricTooltip.dataset.anchorKey = key;
          showQualitySummaryTooltip(point.dataset.referenceMonth || "", point.dataset.qualityScope || "all");
        });
      });
  }

  if (analysisMetricTooltip && isManager() && String(state.filters.analysisUserId) === "all") {
    const analysisPoints = document.querySelectorAll(".analysis-point[data-metric]");
    const getAnalysisCacheKey = (dataset) => [
      dataset.metric || "",
      dataset.operation || "",
      dataset.date || "",
      dataset.referenceMonth || "",
      dataset.qualityField || "",
    ].join("|");
    const formatMetricValue = (metric, value) => {
      if (metric === "production") return integer(value);
      if (metric === "effectiveness") return percent(value);
      return number(value);
    };
    const metricLabel = (metric, operation, qualityField) => {
      if (metric === "production") return `Top 10 Produção · ${operation === "0800" ? "0800" : "Nuvidio"}`;
      if (metric === "effectiveness") return `Top 10 Efetividade · ${operation === "0800" ? "0800" : "Nuvidio"}`;
      const map = { m1: "M1", m2: "M2", m3: "M3", m4: "M4", final: "Final" };
      return `Top 10 Qualidade · ${map[qualityField] || "Final"}`;
    };
    const metricWhenLabel = (dataset) => {
      if (dataset.metric === "quality") return formatMonthLabel(dataset.referenceMonth || "");
      return formatDateBr(dataset.date || "");
    };
    const showAnalysisTooltip = async (point) => {
      const { dataset } = point;
      const key = getAnalysisCacheKey(dataset);
      if (!key) return;
      if (!state.analysisTopCache[key]) {
        state.analysisTopCache[key] = { loading: true, top: [] };
        const params = new URLSearchParams();
        params.set("metric", dataset.metric || "");
        if (dataset.operation) params.set("operation", dataset.operation);
        if (dataset.date) params.set("date", dataset.date);
        if (dataset.referenceMonth) params.set("reference_month", dataset.referenceMonth);
        if (dataset.qualityField) params.set("quality_field", dataset.qualityField);
        try {
          const response = await api(`/api/dashboard/top-metric?${params.toString()}`);
          state.analysisTopCache[key] = { loading: false, top: response.top || [] };
        } catch {
          state.analysisTopCache[key] = { loading: false, top: [] };
        }
      }
      const bucket = state.analysisTopCache[key];
      const lines = bucket.loading
        ? `<div class="trend-tooltip-line trend-tooltip-line-single">Carregando...</div>`
        : (bucket.top.length
          ? bucket.top.map((item, idx) => `
            <div class="trend-tooltip-rank">
              <div class="trend-rank-order">${idx + 1}</div>
              <div class="trend-rank-copy">
                <strong>${esc(item.name)}</strong>
                <span>${formatMetricValue(dataset.metric, item.value)}</span>
              </div>
            </div>
          `).join("")
          : `<div class="trend-tooltip-line trend-tooltip-line-single">Sem dados para ${esc(metricWhenLabel(dataset))}.</div>`);
      analysisMetricTooltip.innerHTML = `
        <div class="trend-tooltip-head">
          <div class="trend-tooltip-title">${esc(metricLabel(dataset.metric, dataset.operation, dataset.qualityField))}</div>
          <button class="trend-tooltip-close" type="button" data-close-popover aria-label="Fechar">×</button>
        </div>
        <div class="trend-tooltip-line trend-tooltip-line-single"><span>${esc(metricWhenLabel(dataset))}</span></div>
        ${lines}
      `;
      attachPopoverClose(analysisMetricTooltip);
      placeTooltipCentered(analysisMetricTooltip);
    };
    analysisPoints.forEach((point) => {
      point.addEventListener("click", async (event) => {
        event.stopPropagation();
        const key = getAnalysisCacheKey(point.dataset);
        const alreadyVisible = !analysisMetricTooltip.hidden && analysisMetricTooltip.dataset.anchorKey === key;
        hideMetricPopovers();
        if (alreadyVisible) return;
        analysisMetricTooltip.dataset.anchorKey = key;
        await showAnalysisTooltip(point);
      });
    });
  }

  document.addEventListener("click", hideMetricPopovers);

  if (state.forcePasswordChange && passwordModal) {
    passwordModal.hidden = false;
  }

  document.querySelectorAll("[data-route]").forEach((button) => {
    button.addEventListener("click", async () => {
      clearFlash();
      state.route = button.dataset.route;
      if (state.route === "alerts") state.alerts = null;
      if (state.route === "reports" || state.route === "detailed") state.reportDataset = null;
      render();
      try {
        if (state.route === "alerts") {
          await loadAlerts();
          render();
        } else if (state.route === "reports" || state.route === "detailed") {
          await loadReportsData();
          if (state.route === "detailed" && state.filters.reportsType === "ofensores") {
            await loadAlerts();
            state.reportDataset = buildReportDatasetModel();
          }
          render();
        }
        await saveUserPreferences({ last_route: state.route });
      } catch (error) {
        setFlash("error", error.message || "Não foi possível salvar a aba atual.");
      }
    });
  });

  document.querySelectorAll("#start-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.start = event.target.value; }));
  document.querySelectorAll("#end-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.end = event.target.value; }));
  document.querySelectorAll("#alerts-start-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.start = event.target.value; }));
  document.querySelectorAll("#alerts-end-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.end = event.target.value; }));
  document.querySelectorAll("#reports-start-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.start = event.target.value; }));
  document.querySelectorAll("#reports-end-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.end = event.target.value; }));
  document.querySelectorAll("#reports-type-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.reportsType = event.target.value; }));
  document.querySelectorAll("#reports-view-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.reportsView = event.target.value; }));
  document.querySelectorAll("#reports-sector-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.reportsSector = event.target.value; }));
  document.querySelectorAll("#detailed-start-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.start = event.target.value; state.filters.detailedPage = 1; }));
  document.querySelectorAll("#detailed-end-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.end = event.target.value; state.filters.detailedPage = 1; }));
  document.querySelectorAll("#detailed-type-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.reportsType = event.target.value; state.filters.detailedPage = 1; }));
  document.querySelectorAll("#detailed-sector-filter").forEach((input) => input.addEventListener("change", (event) => { state.filters.reportsSector = event.target.value; state.filters.detailedPage = 1; }));
  document.querySelectorAll("[data-report-sort-scope][data-report-sort-column]").forEach((button) => {
    button.addEventListener("click", () => {
      const scope = String(button.dataset.reportSortScope || "").trim();
      const column = String(button.dataset.reportSortColumn || "").trim();
      if (!scope || !column) return;
      const current = getReportSort(scope);
      const direction = current.column === column && current.direction === "asc" ? "desc" : "asc";
      state.reportSorts = {
        ...state.reportSorts,
        [scope]: { column, direction },
      };
      state.filters.detailedPage = 1;
      invalidateReportDatasetCache();
      state.reportDataset = buildReportDatasetModel();
      render();
    });
  });

  const historyUserSearch = document.getElementById("history-user-search");
  if (historyUserSearch) {
    historyUserSearch.addEventListener("input", (event) => {
      state.filters.historyQuery = event.target.value;
    });
  }

  const managementUserSearch = document.getElementById("management-user-search");
  if (managementUserSearch) {
    applyManagementUserFilter();
    managementUserSearch.addEventListener("input", (event) => {
      state.filters.usersQuery = event.target.value;
      applyManagementUserFilter();
    });
  }

  const globalUserSelect = document.getElementById("global-user-select");
  if (globalUserSelect && isManager()) {
    globalUserSelect.addEventListener("change", async (event) => {
      state.filters.analysisUserId = event.target.value;
      state.filters.historyUserId = event.target.value === "all" ? "all" : event.target.value;
      state.filters.historyQuery = event.target.value === "all" ? "" : (getUserLabelById(state.filters.historyUserId) || "");
      if (state.route === "alerts") {
        state.alerts = null;
        render();
        await loadAlerts();
      } else if (state.route === "detailed" && state.filters.reportsType === "ofensores") {
        state.alerts = null;
        render();
        await loadAlerts();
      } else {
        await loadHistory();
      }
      render();
    });
  }

  const actionMap = {
    "logout": async () => {
      await api("/api/auth/logout", { method: "POST" });
      state.user = null;
      state.route = "overview";
      state.theme = "dark";
      setFlash("success", "Sessão encerrada.");
    },
    "refresh-all": async () => {
      await loadAll();
      setFlash("success", "Atualizado.");
    },
    "refresh-analysis": async () => {
      await loadBootstrap();
      setFlash("success", "Análises atualizadas.");
    },
    "refresh-history": async () => {
      if (isManager()) {
        if (!String(state.filters.historyQuery || "").trim()) {
          state.filters.historyUserId = "all";
          state.filters.analysisUserId = "all";
        } else {
          const resolvedUserId = resolveHistoryUserId(state.filters.historyQuery);
          if (!resolvedUserId) {
            throw new Error("Selecione um operador válido para pesquisar.");
          }
          state.filters.historyUserId = resolvedUserId;
          state.filters.historyQuery = getUserLabelById(resolvedUserId);
        }
      }
      if (state.route === "alerts") {
        await loadAlerts();
        setFlash("success", "Alertas atualizados.");
      } else {
        await loadHistory();
        setFlash("success", "Histórico atualizado.");
      }
    },
    "refresh-alerts": async () => {
      await loadAlerts();
      setFlash("success", "Ofensores atualizados.");
    },
    "reset-alerts": async () => {
      state.filters.start = "";
      state.filters.end = "";
      await loadAlerts();
      render();
      setFlash("success", "Período redefinido.");
    },
    "refresh-reports": async () => {
      await loadReportsData();
      setFlash("success", "Relatórios atualizados.");
    },
    "refresh-detailed": async () => {
      state.filters.detailedPage = 1;
      await loadReportsData();
      render();
      setFlash("success", "Consulta detalhada atualizada.");
    },
    "reset-reports": async () => {
      state.filters.start = "";
      state.filters.end = "";
      state.filters.reportsType = "consolidado";
      state.filters.reportsView = "detalhada";
      state.filters.reportsSector = "all";
      await loadReportsData();
      render();
      setFlash("success", "Período redefinido.");
    },
    "reset-detailed": async () => {
      state.filters.start = "";
      state.filters.end = "";
      state.filters.reportsType = "consolidado";
      state.filters.reportsSector = "all";
      state.filters.detailedPage = 1;
      await loadReportsData();
      render();
      setFlash("success", "Filtros redefinidos.");
    },
    "detailed-prev-page": async () => {
      state.filters.detailedPage = Math.max(1, Number(state.filters.detailedPage || 1) - 1);
      render();
    },
    "detailed-next-page": async () => {
      state.filters.detailedPage = Number(state.filters.detailedPage || 1) + 1;
      render();
    },
    "export-report-excel": async () => {
      const userId = state.filters.analysisUserId || "all";
      const params = new URLSearchParams({
        format: "excel",
        start: state.filters.start || "",
        end: state.filters.end || "",
        user_id: userId,
        type: state.filters.reportsType || "consolidado",
        view: state.filters.reportsView || "detalhada",
        sector: state.filters.reportsSector || "all",
      });
      await downloadFile(`/api/reports/export?${params.toString()}`, "relatorio-operacional.xls");
    },
    "export-report-pdf": async () => {
      const userId = state.filters.analysisUserId || "all";
      const params = new URLSearchParams({
        format: "pdf",
        start: state.filters.start || "",
        end: state.filters.end || "",
        user_id: userId,
        type: state.filters.reportsType || "consolidado",
        view: state.filters.reportsView || "detalhada",
        sector: state.filters.reportsSector || "all",
      });
      await downloadFile(`/api/reports/export?${params.toString()}`, "relatorio-operacional.pdf");
    },
    "reset-analysis": async () => {
      state.filters.start = "";
      state.filters.end = "";
      state.filters.operation = "all";
      state.filters.analysisUserId = isManager() ? "all" : String(state.user.id);
      await loadBootstrap();
      setFlash("success", "Filtros redefinidos.");
    },
    "toggle-theme": async () => {
      state.theme = state.theme === "dark" ? "contrast" : "dark";
      render();
      try {
        await saveUserPreferences({ preferred_theme: state.theme });
      } catch (error) {
        setFlash("error", error.message || "Não foi possível salvar o tema.");
      }
    },
  };

  document.querySelectorAll("[data-action]").forEach((button) => {
    const action = actionMap[button.dataset.action];
    if (!action) return;
    button.addEventListener("click", async () => {
      const processingLabel = (
        button.dataset.action === "export-report-excel" || button.dataset.action === "export-report-pdf"
      ) ? "Gerando..." : "Processando...";
      const restoreButton = setButtonProcessing(button, true, processingLabel);
      try {
        await action();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  });

  const qualityUploadForm = document.getElementById("quality-upload-form");
  if (qualityUploadForm) {
    const monthInput = document.getElementById("quality-reference-month");
    const monthSelect = document.getElementById("quality-reference-month-select");
    const yearSelect = document.getElementById("quality-reference-year-select");
    const uploadModeSelect = document.getElementById("quality-upload-mode-select");
    const syncReferenceMonth = () => {
      if (!monthInput || !monthSelect || !yearSelect) return;
      monthInput.value = `${yearSelect.value}-${monthSelect.value}`;
    };
    syncReferenceMonth();
    if (monthSelect) monthSelect.addEventListener("change", syncReferenceMonth);
    if (yearSelect) yearSelect.addEventListener("change", syncReferenceMonth);
    qualityUploadForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const submitButton = qualityUploadForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Importando...");
      try {
        const response = await fetch("/api/admin/quality/import", { method: "POST", body: form, credentials: "same-origin" });
        const data = await response.json();
        if (!response.ok) throw new Error(data.details ? `${data.error || "Erro ao importar qualidade"}: ${data.details}` : (data.error || "Erro ao importar qualidade"));
        await loadBootstrap();
        setFlash("success", `Qualidade importada com ${data.processed} operadores.`);
        render();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  const qualityManualForm = document.getElementById("quality-manual-form");
  if (qualityManualForm) {
    const monthSelect = document.getElementById("quality-manual-month-select");
    const yearSelect = document.getElementById("quality-manual-year-select");
    const monthInput = document.getElementById("quality-manual-reference-month");
    const syncManualReference = () => {
      if (!monthSelect || !yearSelect || !monthInput) return;
      monthInput.value = `${yearSelect.value}-${monthSelect.value}`;
    };
    syncManualReference();
    if (monthSelect) monthSelect.addEventListener("change", syncManualReference);
    if (yearSelect) yearSelect.addEventListener("change", syncManualReference);

    qualityManualForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const submitButton = qualityManualForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Salvando...");
      try {
        const payload = Object.fromEntries(new FormData(event.currentTarget).entries());
        const generalScore = Number(String(payload.score || "").replace(",", "."));
        if (!Number.isFinite(generalScore) || generalScore < 0 || generalScore > 100) {
          throw new Error("Informe a nota bruta do mês entre 0 e 100.");
        }
        payload.score_type = "general";
        payload.score = generalScore;
        payload.monitoria_1 = "";
        payload.monitoria_2 = "";
        payload.monitoria_3 = "";
        payload.monitoria_4 = "";
        await api("/api/admin/quality", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        refreshDashboardInBackground("Qualidade salva com sucesso.");
        qualityManualForm.reset();
        syncManualReference();
        syncManualQualityMode();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  const baseImportForm = document.getElementById("base-import-form");
  if (baseImportForm) {
    baseImportForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const submitButton = baseImportForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Importando...");
      try {
        const response = await fetch("/api/admin/import", { method: "POST", body: form, credentials: "same-origin" });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "Erro ao importar base.");
        if (data.period?.start && data.period?.end) {
          state.filters.start = state.filters.start ? (state.filters.start < data.period.start ? state.filters.start : data.period.start) : data.period.start;
          state.filters.end = state.filters.end ? (state.filters.end > data.period.end ? state.filters.end : data.period.end) : data.period.end;
          state.filters.today = state.filters.end;
        }
        if (isManager()) {
          state.filters.analysisUserId = "all";
          state.filters.historyUserId = "all";
          state.filters.historyQuery = "";
          state.filters.operation = "all";
        }
        state.route = "history";
        await loadBootstrap();
        setFlash(
          "success",
          `Base importada: ${data.processed} registro(s), ${data.rejected} rejeição(ões).`,
          (data.errors || []).map((entry) => `Linha ${entry.row}: ${entry.error}`),
        );
        render();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  const downloadBaseTemplate = document.getElementById("download-base-template");
  if (downloadBaseTemplate) {
    downloadBaseTemplate.addEventListener("click", async () => {
      const modelSelect = document.getElementById("base-model-select");
      const model = modelSelect ? modelSelect.value : "nuvidio";
      const restoreButton = setButtonProcessing(downloadBaseTemplate, true, "Baixando...");
      try {
        await downloadFile(`/api/admin/import/template?model=${encodeURIComponent(model)}`, `modelo-base-${model}.csv`);
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  const manualTagForm = document.getElementById("manual-tag-form");
  if (manualTagForm) {
    const operationSelect = manualTagForm.querySelector('select[name="operation"]');
    const pendingField = manualTagForm.querySelector(".pending-only");
    const emptyField = manualTagForm.querySelector(".empty-only");
    const syncPendingVisibility = () => {
      if (!operationSelect) return;
      pendingField.style.display = operationSelect.value === "0800" ? "grid" : "none";
      if (emptyField) emptyField.style.display = operationSelect.value === "nuvidio" ? "grid" : "none";
    };
    syncPendingVisibility();
    if (operationSelect) operationSelect.addEventListener("change", syncPendingVisibility);

    manualTagForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const payload = Object.fromEntries(new FormData(event.currentTarget).entries());
      const submitButton = manualTagForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Salvando...");
      try {
        await api("/api/admin/manual-tag", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        refreshDashboardInBackground("Lançamento manual salvo.");
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  const adminUserForm = document.getElementById("admin-user-form");
  if (adminUserForm) {
    adminUserForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const payload = Object.fromEntries(form.entries());
      const submitButton = adminUserForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Salvando...");
      try {
        const data = await api("/api/admin/users", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        await loadUsers();
        setFlash("success", `Usuário cadastrado com senha inicial ${data.default_password}.`);
        render();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  const userEditForm = document.getElementById("user-edit-form");
  if (userEditForm) {
    userEditForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const userId = form.get("user_id");
      const payload = {
        full_name: form.get("full_name"),
        login: form.get("login"),
        role: form.get("role"),
        is_active: String(form.get("is_active")) === "true",
        platform_0800_id: form.get("platform_0800_id"),
        nuvidio_id: form.get("nuvidio_id"),
        password: form.get("password"),
      };
      const submitButton = userEditForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Salvando...");
      try {
        const data = await api(`/api/admin/users/${userId}`, {
          method: "PUT",
          body: JSON.stringify(payload),
        });
        await loadUsers();
        if (String(state.user?.id) === String(data.user.id)) {
          state.user = normalizeUserPayload({ ...state.user, ...data.user });
        }
        closeUserModal();
        setFlash("success", "Usuário atualizado com sucesso.");
        render();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  document.querySelectorAll("[data-user-edit]").forEach((button) => {
    button.addEventListener("click", () => openUserModal(button.dataset.userEdit));
  });

  document.querySelectorAll("[data-user-toggle]").forEach((button) => {
    button.addEventListener("click", async () => {
      const user = state.users.find((item) => String(item.id) === String(button.dataset.userToggle));
      if (!user) return;
      const restoreButton = setButtonProcessing(button, true, user.is_active ? "Desativando..." : "Ativando...");
      try {
        const data = await api(`/api/admin/users/${user.id}`, {
          method: "PUT",
          body: JSON.stringify({ is_active: !user.is_active }),
        });
        await loadUsers();
        if (String(state.user?.id) === String(data.user.id)) {
          state.user = normalizeUserPayload({ ...state.user, ...data.user });
        }
        setFlash("success", `Usuário ${data.user.is_active ? "ativado" : "desativado"} com sucesso.`);
        render();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  });

  document.querySelectorAll("[data-user-delete]").forEach((button) => {
    button.addEventListener("click", async () => {
      const user = state.users.find((item) => String(item.id) === String(button.dataset.userDelete));
      if (!user) return;
      if (!window.confirm(`Apagar o usuário ${user.full_name}?`)) return;
      const restoreButton = setButtonProcessing(button, true, "Apagando...");
      try {
        await api(`/api/admin/users/${user.id}`, { method: "DELETE" });
        await Promise.all([loadUsers(), loadOverview(), loadAnalysis(), loadHistory()]);
        setFlash("success", "Usuário apagado com sucesso.");
        render();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  });

  document.querySelectorAll("[data-user-reset]").forEach((button) => {
    button.addEventListener("click", async () => {
      const user = state.users.find((item) => String(item.id) === String(button.dataset.userReset));
      if (!user) return;
      if (!window.confirm(`Resetar a senha de ${user.full_name} para ${DEFAULT_PASSWORD_HINT}?`)) return;
      const restoreButton = setButtonProcessing(button, true, "Resetando...");
      try {
        const data = await api(`/api/admin/users/${user.id}/reset-password`, {
          method: "POST",
        });
        await loadUsers();
        setFlash("success", `Senha resetada para ${data.default_password}. O usuário deverá trocar no próximo login.`);
        render();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  });

  document.querySelectorAll("[data-history-edit]").forEach((button) => {
    button.addEventListener("click", () => {
      const metricId = Number(button.dataset.historyEdit);
      const operation = String(button.dataset.historyOperation || "");
      const entryType = String(button.dataset.historyType || "metric");
      const historyRow = getScopedHistory().find((item) => Number(item.metricId) === metricId && String(item.entryType) === entryType);
      if (!historyRow) {
        setFlash("error", "Registro não encontrado para edição.");
        return;
      }
      const setField = (id, value) => {
        const input = document.getElementById(id);
        if (input) input.value = value ?? "";
      };
      const operatorName = state.users.find((user) => String(user.id) === String(historyRow.userId))?.full_name || "Operador";
      const metricFields = document.querySelector(".history-edit-metric-fields");
      const qualityFields = document.querySelector(".history-edit-quality-fields");
      const pendingRow = document.querySelector(".history-edit-pending");
      const emptyRow = document.querySelector(".history-edit-empty");
      const submitButton = document.querySelector("#history-edit-form button[type='submit']");
      setField("history-edit-metric-id", metricId);
      setField("history-edit-type", entryType);
      setField("history-edit-operation", operation);
      setField("history-edit-operator", operatorName);
      setField("history-edit-date", historyRow.dateLabel || formatDateBr(historyRow.date));
      setField("history-edit-operation-label", operation);
      if (entryType === "quality") {
        if (metricFields) metricFields.hidden = true;
        if (qualityFields) qualityFields.hidden = false;
        if (submitButton) submitButton.textContent = "Salvar qualidade";
        setField("history-edit-quality-scope", historyRow.qualityScope === "nuvidio" ? "nuvidio" : "0800");
        setField("history-edit-score", historyRow.quality ?? "");
        setField("history-edit-notes", historyRow.notes ?? "");
      } else {
        const is0800 = operation === "0800";
        if (metricFields) metricFields.hidden = false;
        if (qualityFields) qualityFields.hidden = true;
        if (submitButton) submitButton.textContent = "Salvar lançamento";
        setField("history-edit-approved", Number(historyRow.calls_approved || 0));
        setField("history-edit-rejected", Number(historyRow.calls_rejected || 0));
        setField("history-edit-pending", Number(historyRow.calls_pending || 0));
        setField("history-edit-no-action", Number(historyRow.calls_no_action || 0));
        setField("history-edit-empty", Number(historyRow.calls_empty || 0));
        if (pendingRow) pendingRow.style.display = is0800 ? "grid" : "none";
        if (emptyRow) emptyRow.style.display = is0800 ? "none" : "grid";
      }
      if (historyEditModal) historyEditModal.hidden = false;
    });
  });

  document.querySelectorAll("[data-history-view]").forEach((button) => {
    button.addEventListener("click", () => {
      const metricId = Number(button.dataset.historyView);
      const historyRow = getScopedHistory().find((item) => Number(item.metricId) === metricId && String(item.entryType) === "quality");
      if (!historyRow || !historyViewModal) {
        setFlash("error", "Registro de qualidade não encontrado.");
        return;
      }
      const operatorName = getUserLabelById(historyRow.userId) || state.user?.full_name || "Operador";
      const launchedMonitorias = [
        historyRow.m1_entered ? ["M1", historyRow.monitoria_1] : null,
        historyRow.m2_entered ? ["M2", historyRow.monitoria_2] : null,
        historyRow.m3_entered ? ["M3", historyRow.monitoria_3] : null,
        historyRow.m4_entered ? ["M4", historyRow.monitoria_4] : null,
      ].filter(Boolean);
      const content = document.getElementById("history-view-content");
      if (content) {
        content.innerHTML = `
          <div class="form-grid">
            <label>Operador<input value="${esc(operatorName)}" readonly></label>
            <label>Mês<input value="${esc(historyRow.dateLabel || formatMonthLabel(historyRow.referenceMonth || ""))}" readonly></label>
            <label>Esteira<input value="${esc(formatQualityScopeLabel(historyRow.qualityScope || "all"))}" readonly></label>
            <label>Nota<input value="${esc(number(historyRow.quality || 0))}" readonly></label>
          </div>
          ${launchedMonitorias.length ? `
            <div class="info-box">Monitorias lançadas para este mês.</div>
            <div class="mini-grid">
              ${launchedMonitorias.map(([label, value]) => `<div class="mini-card"><span class="muted">${esc(label)}</span><div class="metric-value">${esc(number(value || 0))}</div></div>`).join("")}
            </div>
          ` : `<div class="info-box">Lançamento bruto do mês.</div>`}
          <label>Observações<input value="${esc(historyRow.notes || "")}" readonly></label>
        `;
      }
      historyViewModal.hidden = false;
    });
  });

  const historyEditForm = document.getElementById("history-edit-form");
  if (historyEditForm) {
    historyEditForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const metricId = Number(form.get("metric_id"));
      const operation = String(form.get("operation") || "");
      const entryType = String(form.get("entry_type") || "metric");
      const historyRow = getScopedHistory().find((item) => Number(item.metricId) === metricId && String(item.entryType) === entryType);
      if (!historyRow) {
        setFlash("error", "Registro não encontrado para edição.");
        return;
      }
      const submitButton = historyEditForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Salvando...");
      try {
        if (entryType === "quality") {
          const payload = {
            score_type: "general",
            quality_scope: form.get("quality_scope") || "0800",
            score: form.get("score"),
            monitoria_1: "",
            monitoria_2: "",
            monitoria_3: "",
            monitoria_4: "",
            notes: form.get("notes"),
          };
          await api(`/api/admin/quality/${metricId}`, {
            method: "PUT",
            body: JSON.stringify(payload),
          });
        } else {
          const is0800 = operation === "0800";
          const approvedN = Math.max(0, Number(form.get("approved") || 0));
          const rejectedN = Math.max(0, Number(form.get("rejected") || 0));
          const pendingN = Math.max(0, Number(form.get("pending") || 0));
          const noActionN = Math.max(0, Number(form.get("no_action") || 0));
          const emptyN = Math.max(0, Number(form.get("empty") || 0));
          const payload = is0800
            ? {
                production: Math.max(0, Number(historyRow.production_nuvidio || 0)) + approvedN + rejectedN + pendingN + noActionN,
                calls_0800_approved: approvedN,
                calls_0800_rejected: rejectedN,
                calls_0800_pending: pendingN,
                calls_0800_no_action: noActionN,
              }
            : {
                production: Math.max(0, Number(historyRow.production_0800 || 0)) + approvedN + rejectedN + noActionN + emptyN,
                calls_nuvidio_approved: approvedN,
                calls_nuvidio_rejected: rejectedN,
                calls_nuvidio_no_action: noActionN,
                calls_nuvidio_empty: emptyN,
              };
          await api(`/api/admin/daily-metrics/${metricId}`, {
            method: "PUT",
            body: JSON.stringify(payload),
          });
        }
        closeHistoryEditModal();
        refreshDashboardInBackground(entryType === "quality" ? "Qualidade atualizada." : `Registro ${operation} atualizado.`);
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  document.querySelectorAll("[data-history-delete]").forEach((button) => {
    button.addEventListener("click", async () => {
      const metricId = Number(button.dataset.historyDelete);
      const entryType = String(button.dataset.historyType || "metric");
      if (!window.confirm(entryType === "quality" ? "Remover este registro de qualidade?" : "Remover este registro diário?")) return;
      const restoreButton = setButtonProcessing(button, true, "Removendo...");
      try {
        if (entryType === "quality") {
          const params = new URLSearchParams();
          const fallbackUserId = String(button.dataset.historyUserId || "").trim();
          const fallbackReferenceMonth = String(button.dataset.historyReferenceMonth || "").trim();
          const fallbackScope = String(button.dataset.historyQualityScope || "").trim();
          if (fallbackUserId) params.set("user_id", fallbackUserId);
          if (fallbackReferenceMonth) params.set("reference_month", fallbackReferenceMonth);
          if (fallbackScope) params.set("scope", fallbackScope);
          const suffix = params.toString() ? `?${params.toString()}` : "";
          await api(`/api/admin/quality/${metricId}${suffix}`, { method: "DELETE" });
        } else {
          await api(`/api/admin/daily-metrics/${metricId}`, { method: "DELETE" });
        }
        refreshDashboardInBackground(entryType === "quality" ? "Qualidade removida com sucesso." : "Registro removido com sucesso.");
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  });

  const historyDeleteDayForm = document.getElementById("history-delete-day-form");
  if (historyDeleteDayForm) {
    const operationSelect = document.getElementById("history-delete-day-operation");
    const dateLabel = document.getElementById("history-delete-day-date-label");
    const dateSelect = document.getElementById("history-delete-day-date");
    const qualityMonthLabel = document.getElementById("history-delete-quality-month-label");
    const qualityMonthSelect = document.getElementById("history-delete-quality-month");
    const infoBox = document.getElementById("history-delete-day-info");
    const syncBulkDeleteMode = () => {
      const operation = String(operationSelect?.value || "all").trim().toLowerCase();
      const isQuality = operation.startsWith("quality-");
      if (dateLabel) dateLabel.hidden = isQuality;
      if (qualityMonthLabel) qualityMonthLabel.hidden = !isQuality;
      if (dateSelect) dateSelect.required = !isQuality;
      if (qualityMonthSelect) qualityMonthSelect.required = isQuality;
      if (infoBox) {
        infoBox.textContent = isQuality
          ? "Essa ação remove os lançamentos de qualidade de todos os operadores no mês e esteira selecionados."
          : "Essa ação remove os registros de todos os operadores na data e setor selecionados.";
      }
    };
    operationSelect?.addEventListener("change", syncBulkDeleteMode);
    syncBulkDeleteMode();

    historyDeleteDayForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const operation = String(form.get("operation") || "all").trim().toLowerCase();
      const submitButton = historyDeleteDayForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Excluindo...");
      try {
        if (operation.startsWith("quality-")) {
          const referenceMonth = String(form.get("reference_month") || "").trim();
          if (!referenceMonth) {
            setFlash("error", "Selecione um mês para excluir a qualidade.");
            return;
          }
          const scope = operation === "quality-0800" ? "0800" : operation === "quality-nuvidio" ? "nuvidio" : "all";
          const scopeLabel = scope === "0800" ? "0800" : scope === "nuvidio" ? "Nuvidio" : "0800 + Nuvidio";
          if (!window.confirm(`Excluir toda a qualidade de ${scopeLabel} do mês ${formatMonthLabel(referenceMonth)} para todos os operadores?`)) return;
          await api(`/api/admin/quality/by-month?reference_month=${encodeURIComponent(referenceMonth)}&scope=${encodeURIComponent(scope)}`, {
            method: "DELETE",
          });
        } else {
          const metricDate = String(form.get("date") || "").trim();
          if (!metricDate) {
            setFlash("error", "Selecione um dia para excluir.");
            return;
          }
          const operationLabel = operation === "0800" ? "0800" : operation === "nuvidio" ? "Nuvidio" : "0800 + Nuvidio";
          if (!window.confirm(`Excluir todos os registros de ${operationLabel} do dia ${formatDateBr(metricDate)} para todos os operadores?`)) return;
          await api(`/api/admin/daily-metrics/by-day?date=${encodeURIComponent(metricDate)}&operation=${encodeURIComponent(operation)}`, {
            method: "DELETE",
          });
        }
        closeHistoryDeleteDayModal();
        refreshDashboardInBackground(operation.startsWith("quality-") ? "Qualidade em massa removida com sucesso." : "Registros do dia removidos com sucesso.");
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  const downloadQualityTemplate = document.getElementById("download-quality-template");
  if (downloadQualityTemplate) {
    downloadQualityTemplate.addEventListener("click", () => {
      window.location.href = "/api/admin/quality/template?mode=general";
    });
  }

  const maintenanceForm = document.getElementById("admin-maintenance-form");
  if (maintenanceForm) {
    const maintenanceToggle = document.getElementById("maintenance-toggle");
    const maintenanceMessage = document.getElementById("maintenance-message");
    const metricProductionRed = document.getElementById("metric-production-red");
    const metricProductionAmber = document.getElementById("metric-production-amber");
    const metricEffectivenessRed = document.getElementById("metric-effectiveness-red");
    const metricEffectivenessAmber = document.getElementById("metric-effectiveness-amber");
    const metricQualityRed = document.getElementById("metric-quality-red");
    const metricQualityAmber = document.getElementById("metric-quality-amber");
    const alertProductionNuvidio = document.getElementById("alert-production-nuvidio");
    const alertProduction0800 = document.getElementById("alert-production-0800");
    const alertEffectiveness0800 = document.getElementById("alert-effectiveness-0800");
    const alertEffectivenessNuvidio = document.getElementById("alert-effectiveness-nuvidio");
    const alertQuality = document.getElementById("alert-quality");
    const buildMetricRulesPayload = () => ({
      production: {
        red_max: Number(metricProductionRed?.value || 70),
        amber_max: Number(metricProductionAmber?.value || 100),
      },
      effectiveness: {
        red_max: Number(metricEffectivenessRed?.value || 70),
        amber_max: Number(metricEffectivenessAmber?.value || 90),
      },
        quality: {
          red_max: Number(metricQualityRed?.value || 70),
          amber_max: Number(metricQualityAmber?.value || 90),
        },
      });
    const buildAlertRulesPayload = () => ({
      production_nuvidio: {
        critical_min: Number(alertProductionNuvidio?.value || 70),
      },
      production_0800: {
        critical_min: Number(alertProduction0800?.value || 70),
      },
      effectiveness_0800: {
        critical_min: Number(alertEffectiveness0800?.value || 70),
      },
      effectiveness_nuvidio: {
        critical_min: Number(alertEffectivenessNuvidio?.value || 70),
      },
      quality: {
        critical_min: Number(alertQuality?.value || 70),
      },
    });

    if (maintenanceToggle) {
      maintenanceToggle.addEventListener("change", async () => {
        maintenanceToggle.disabled = true;
        try {
          const response = await api("/api/admin/settings", {
            method: "PATCH",
            body: JSON.stringify({
              maintenance_for_operators: Boolean(maintenanceToggle.checked),
            }),
          });
          state.appSettings = response.app_settings || state.appSettings;
          setFlash("success", maintenanceToggle.checked ? "Manutenção ativada." : "Manutenção desativada.");
          render();
        } catch (error) {
          maintenanceToggle.checked = !maintenanceToggle.checked;
          setFlash("error", error.message);
        } finally {
          maintenanceToggle.disabled = false;
        }
      });
    }

    maintenanceForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const submitButton = maintenanceForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Salvando...");
      try {
        const response = await api("/api/admin/settings", {
          method: "PATCH",
          body: JSON.stringify({
            maintenance_message: String(maintenanceMessage?.value || "").trim(),
            metric_rules: buildMetricRulesPayload(),
            alert_rules: buildAlertRulesPayload(),
          }),
        });
        state.appSettings = response.app_settings || state.appSettings;
        setFlash("success", "Metas salvas com sucesso.");
        render();
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  }

  const passwordForm = document.getElementById("password-form");
  if (passwordForm) {
    passwordForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const payload = Object.fromEntries(form.entries());
      try {
        await api("/api/auth/password", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        state.user = { ...state.user, must_change_password: false };
        state.forcePasswordChange = false;
        if (passwordModal) passwordModal.hidden = true;
        setFlash("success", "Senha atualizada com sucesso.");
      } catch (error) {
        setFlash("error", error.message);
      }
    });
  }
}

boot();

