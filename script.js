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
  },
  dayTopCache: {},
  analysisTopCache: {},
  route: "overview",
  overview: null,
  analysis: null,
  alerts: null,
  history: null,
  users: [],
  flash: null,
  forcePasswordChange: false,
  theme: "dark",
  filters: {
    today: new Date().toISOString().slice(0, 10),
    start: new Date(Date.now() - 29 * 86400000).toISOString().slice(0, 10),
    end: new Date().toISOString().slice(0, 10),
    historyUserId: "",
    historyQuery: "",
    analysisUserId: "all",
    operation: "all",
  },
};

const app = document.getElementById("app");
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

function setFlash(type, message) {
  state.flash = { type, message };
  render();
}

function clearFlash() {
  state.flash = null;
}

function refreshDashboardInBackground(successMessage = "") {
  if (successMessage) setFlash("success", successMessage);
  loadBootstrap()
    .then(async () => {
      if (state.route === "alerts" && isManager()) {
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
    ? ["overview", "analysis", "alerts", "history", "admin"]
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

function getOperatorUsers() {
  return state.users.filter((user) => user.role === "operator" && user.is_active);
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
  return getOperatorUsers().find((user) => String(user.id) === String(userId))?.full_name || "";
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
    operation: "Qualidade",
    production: null,
    effectiveness: null,
    quality: Number(row.score || 0),
    updatedAt: row.updated_at,
    referenceMonth: row.reference_month,
    monitoria_1: row.monitoria_1,
    monitoria_2: row.monitoria_2,
    monitoria_3: row.monitoria_3,
    monitoria_4: row.monitoria_4,
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
  const latest = trend[trend.length - 1];
  const previous = trend[trend.length - 2];
  return {
    totalAttended: trend.reduce((sum, item) => sum + Number(item.production || 0), 0),
    avgProduction: average(productionTrend.map((item) => item.production), { ignoreZero: true }),
    avgEffectiveness: average(trend.map((item) => item.effectiveness), { ignoreZero: true }),
    avgQuality: average(quality.map((item) => item.score)),
    latest,
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
  const qualityRows = (state.history?.quality || []).slice().sort((a, b) => a.reference_month.localeCompare(b.reference_month));
  const qualityByMonthMap = new Map();
  const hasValue = (value) => value !== null && value !== undefined && String(value).trim() !== "";
  qualityRows.forEach((item) => {
    const key = String(item.reference_month || "");
    const bucket = qualityByMonthMap.get(key) || {
      reference_month: key,
      scoreValues: [],
      m1Values: [],
      m2Values: [],
      m3Values: [],
      m4Values: [],
    };
    if (hasValue(item.score)) bucket.scoreValues.push(Number(item.score));
    if (hasValue(item.monitoria_1)) bucket.m1Values.push(Number(item.monitoria_1));
    if (hasValue(item.monitoria_2)) bucket.m2Values.push(Number(item.monitoria_2));
    if (hasValue(item.monitoria_3)) bucket.m3Values.push(Number(item.monitoria_3));
    if (hasValue(item.monitoria_4)) bucket.m4Values.push(Number(item.monitoria_4));
    qualityByMonthMap.set(key, bucket);
  });
  const qualityMonths = [...qualityByMonthMap.values()].map((bucket) => {
    const avg = (arr) => arr.length ? (arr.reduce((s, v) => s + Number(v || 0), 0) / arr.length) : 0;
    const m1 = avg(bucket.m1Values);
    const m2 = avg(bucket.m2Values);
    const m3 = avg(bucket.m3Values);
    const m4 = avg(bucket.m4Values);
    const final = (m1 + m2 + m3 + m4) / 4;
    return {
      reference_month: bucket.reference_month,
      monitoria_1: m1,
      monitoria_2: m2,
      monitoria_3: m3,
      monitoria_4: m4,
      score: avg(bucket.scoreValues),
      final_score: final,
    };
  }).sort((a, b) => a.reference_month.localeCompare(b.reference_month));
  const status0800 = (state.history?.history || []).reduce((acc, row) => {
    acc.approved += Number(row.calls_0800_approved || 0);
    acc.pending += Number(row.calls_0800_pending || 0);
    acc.rejected += Number(row.calls_0800_rejected || 0);
    acc.noAction += Number(row.calls_0800_no_action || 0);
    return acc;
  }, { approved: 0, pending: 0, rejected: 0, noAction: 0 });
  const statusNuvidio = (state.history?.history || []).reduce((acc, row) => {
    acc.approved += Number(row.calls_nuvidio_approved || 0);
    acc.rejected += Number(row.calls_nuvidio_rejected || 0);
    acc.noAction += Number(row.calls_nuvidio_no_action || 0);
    acc.empty += Number(row.calls_nuvidio_empty || 0);
    return acc;
  }, { approved: 0, pending: 0, rejected: 0, noAction: 0, empty: 0 });
  const status = (state.history?.history || []).reduce((acc, row) => {
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
  const buildStatusBreakdown = (bucket) => {
    const total = bucket.approved + bucket.pending + bucket.rejected + bucket.noAction + (bucket.empty || 0);
    const breakdown = [
      { key: "approved", label: "Aprovado", value: bucket.approved, tone: "green" },
      { key: "rejected", label: "Reprovado", value: bucket.rejected, tone: "red" },
      { key: "pending", label: "Pendenciado", value: bucket.pending, tone: "amber" },
      { key: "noAction", label: "Sem ação", value: bucket.noAction, tone: "blue" },
      { key: "empty", label: "Vazio", value: bucket.empty || 0, tone: "muted" },
    ].map((item) => ({
      ...item,
      share: total ? (item.value / total) * 100 : 0,
    }));
    return { total, breakdown };
  };
  const tags0800 = buildStatusBreakdown(status0800);
  const tagsNuvidio = buildStatusBreakdown(statusNuvidio);
  return {
    trend,
    qualityMonths,
    filteredRanking,
    status,
    statusTotal,
    statusBreakdown,
    tags0800,
    tagsNuvidio,
    summary: {
      production: average(filteredRanking.map((item) => item.avg_production), { ignoreZero: true }),
      effectiveness: average(filteredRanking.map((item) => item.effectiveness), { ignoreZero: true }),
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
  if (isManager() && state.route === "alerts") {
    await loadAlerts();
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
}

async function loadBootstrap() {
  const userId = isManager() && (state.filters.analysisUserId === "all" || state.filters.historyUserId === "all")
    ? "all"
    : (isManager() ? state.filters.historyUserId || state.user.id : state.user.id);
  const data = await api(`/api/bootstrap?date=${state.filters.today}&start=${state.filters.start}&end=${state.filters.end}&user_id=${userId}`);
  state.overview = data.overview;
  state.analysis = data.analysis;
  state.history = data.history;
  if (data.app_settings) state.appSettings = data.app_settings;
  if (isManager() && Array.isArray(data.users) && data.users.length) {
    state.users = normalizeUsersPayload(data.users);
    ensureManagerUserFilters();
    state.filters.historyQuery = getUserLabelById(state.filters.historyUserId) || state.filters.historyQuery;
  }
}

async function loadAlerts() {
  if (!isManager()) {
    state.alerts = null;
    return;
  }
  const userId = state.filters.analysisUserId || "all";
  state.alerts = await api(`/api/alerts?start=${state.filters.start}&end=${state.filters.end}&user_id=${encodeURIComponent(userId)}`);
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
    alerts: "Alertas",
    history: "Histórico",
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
  return `<div class="notice toast ${esc(state.flash.type)}">${esc(state.flash.message)}</div>`;
}

function shellTemplate() {
  const titles = {
    overview: { title: isManager() ? "Visão da operação" : "Performance do operador", desc: "" },
    analysis: { title: "Análises", desc: "" },
    alerts: { title: "Alertas", desc: "" },
    history: { title: "Histórico", desc: "" },
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
          ${Object.entries(navMeta()).filter(([key]) => (key !== "admin" && key !== "alerts") || isManager()).map(([key, label]) => `
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
            <button class="toggle" data-action="toggle-theme">${state.theme === "contrast" ? "◐" : "◑"}</button>
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
              <label>Monitoria 1<input type="number" min="0" max="100" step="0.01" id="history-edit-monitoria-1" name="monitoria_1"></label>
              <label>Monitoria 2<input type="number" min="0" max="100" step="0.01" id="history-edit-monitoria-2" name="monitoria_2"></label>
              <label>Monitoria 3<input type="number" min="0" max="100" step="0.01" id="history-edit-monitoria-3" name="monitoria_3"></label>
              <label>Monitoria 4<input type="number" min="0" max="100" step="0.01" id="history-edit-monitoria-4" name="monitoria_4"></label>
              <label>Observações<input id="history-edit-notes" name="notes"></label>
            </div>
            <div class="action-grid">
              <button class="btn-secondary" type="button" id="cancel-history-edit-modal">Cancelar</button>
              <button class="btn" type="submit">Salvar lançamento</button>
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
        </form>
      </div>
      ${flashTemplate()}
    </section>
  `;
}

function renderPage() {
  if (state.route === "analysis") return analysisTemplate();
  if (state.route === "alerts") return alertsTemplate();
  if (state.route === "history") return historyTemplate();
  if (state.route === "admin") return adminTemplate();
  return overviewTemplate();
}

function overviewTemplate() {
  const model = buildOverviewModel();
  const trendRows = state.overview?.trend || [];
  const maxTrendProduction = Math.max(...trendRows.map((row) => Number(row.production || 0)), 1);
  const latestDate = formatDateBr(model.latest?.date || model.latest?.metric_date || "--");
  const latestProduction = model.latest ? integer(model.latest.production) : "--";
  const latestEffectiveness = model.latest ? percent(model.latest.effectiveness) : "--%";
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
    Number(item.final_score || item.score || 0),
  ]), 10);
  const historyRows = state.history?.history || [];
  const byDate0800 = new Map();
  const byDateNuvidio = new Map();
  historyRows.forEach((row) => {
    const date = row.metric_date;
    if (!date) return;
    byDate0800.set(date, {
      date,
      production: Number(row.production_0800 || 0),
      effectiveness: calcOperationEffectiveness(row, "0800"),
    });
    byDateNuvidio.set(date, {
      date,
      production: Number(row.production_nuvidio || 0),
      effectiveness: calcOperationEffectiveness(row, "nuvidio"),
    });
  });
  const trend0800 = [...byDate0800.values()].sort((a, b) => a.date.localeCompare(b.date));
  const trendNuvidio = [...byDateNuvidio.values()].sort((a, b) => a.date.localeCompare(b.date));
  const maxProd0800 = Math.max(...trend0800.map((item) => Number(item.production || 0)), 1);
  const maxEff0800 = Math.max(...trend0800.map((item) => Number(item.effectiveness || 0)), 1);
  const maxProdNuvidio = Math.max(...trendNuvidio.map((item) => Number(item.production || 0)), 1);
  const maxEffNuvidio = Math.max(...trendNuvidio.map((item) => Number(item.effectiveness || 0)), 1);
  const showAllOperatorsTop = isManager() && String(state.filters.analysisUserId) === "all";
  return `
    <section class="section">
      <div class="two-col">
        <aside class="filter-card">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Filtros</span>
              <h3>Recorte</h3>
            </div>
          </div>
          <label>Início<input type="date" id="start-filter" value="${state.filters.start}"></label>
          <label>Fim<input type="date" id="end-filter" value="${state.filters.end}"></label>
          <div class="filter-actions">
            <button class="btn" data-action="refresh-analysis">Aplicar</button>
            <button class="btn-secondary" data-action="reset-analysis">Limpar</button>
          </div>
        </aside>
        <div class="section">
          <div class="hero-grid">
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">0800</span><h3>Produção diária</h3></div></div>
              <div class="chart">
                ${trend0800.length ? trend0800.map((item) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="production" data-operation="0800" data-date="${esc(item.date)}"` : ""}><span class="chart-value">${integer(item.production)}</span><div class="column ${metricTone(item.production, "production")}" style="height:${Math.max(12, (item.production / maxProd0800) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">0800</span><h3>Efetividade diária</h3></div></div>
              <div class="chart">
                ${trend0800.length ? trend0800.map((item) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="effectiveness" data-operation="0800" data-date="${esc(item.date)}"` : ""}><span class="chart-value">${percent(item.effectiveness)}</span><div class="column ${metricTone(item.effectiveness, "effectiveness")}" style="height:${Math.max(12, (item.effectiveness / maxEff0800) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
          </div>

          <div class="hero-grid">
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Nuvidio</span><h3>Produção diária</h3></div></div>
              <div class="chart">
                ${trendNuvidio.length ? trendNuvidio.map((item) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="production" data-operation="nuvidio" data-date="${esc(item.date)}"` : ""}><span class="chart-value">${integer(item.production)}</span><div class="column ${metricTone(item.production, "production")}" style="height:${Math.max(12, (item.production / maxProdNuvidio) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Nuvidio</span><h3>Efetividade diária</h3></div></div>
              <div class="chart">
                ${trendNuvidio.length ? trendNuvidio.map((item) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="effectiveness" data-operation="nuvidio" data-date="${esc(item.date)}"` : ""}><span class="chart-value">${percent(item.effectiveness)}</span><div class="column ${metricTone(item.effectiveness, "effectiveness")}" style="height:${Math.max(12, (item.effectiveness / maxEffNuvidio) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
          </div>

          <div class="hero-grid">
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Qualidade</span><h3>Monitorias + média final</h3></div></div>
              <div class="quality-months-wrap">
                ${qualityMonths.length ? qualityMonths.map((monthItem) => `
                  <div class="quality-month-block">
                    <strong class="quality-month-title">${esc(formatMonthLabel(monthItem.reference_month).split(" de ")[0])}</strong>
                    <div class="chart">
                      ${[
                        { label: "M1", value: Number(monthItem.monitoria_1 || 0), field: "m1" },
                        { label: "M2", value: Number(monthItem.monitoria_2 || 0), field: "m2" },
                        { label: "M3", value: Number(monthItem.monitoria_3 || 0), field: "m3" },
                        { label: "M4", value: Number(monthItem.monitoria_4 || 0), field: "m4" },
                        { label: "Final", value: Number(monthItem.final_score || monthItem.score || 0), field: "final" },
                      ].map((entry) => `<div class="chart-col ${showAllOperatorsTop ? "analysis-point" : ""}" ${showAllOperatorsTop ? `data-metric="quality" data-operation="all" data-reference-month="${esc(monthItem.reference_month)}" data-quality-field="${entry.field}"` : ""}><span class="chart-value">${number(entry.value)}</span><div class="column ${metricTone(entry.value, "quality")}" style="height:${Math.max(12, (entry.value / maxQuality) * 110)}px;"></div><small>${entry.label}</small></div>`).join("")}
                    </div>
                  </div>
                `).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Resumo</span><h3>Consolidado</h3></div></div>
              <div class="mini-grid">
                <div class="mini-card"><span class="muted">Produção média</span><div class="metric-value">${integer(model.summary.production)}</div></div>
                <div class="mini-card"><span class="muted">Efetividade média</span><div class="metric-value">${percent(model.summary.effectiveness)}</div></div>
              </div>
            </article>
          </div>

          <div class="hero-grid">
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
      </div>
      ${showAllOperatorsTop ? `<div class="trend-tooltip" id="analysis-metric-tooltip" hidden></div>` : ""}
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
              <span class="eyebrow">Alertas</span>
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
    : "Todos os operadores";

  return `
    <section class="section">
      <div class="hero-grid">
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Monitoramento</span>
              <h3>Operadores em alerta</h3>
            </div>
          </div>
          <div class="mini-grid">
            <div class="mini-card">
              <span class="muted">Escopo</span>
              <div class="metric-value">${esc(selectedName)}</div>
            </div>
            <div class="mini-card">
              <span class="muted">Período</span>
              <div class="metric-value">${esc(formatDateBr(state.alerts.period?.start || state.filters.start))} - ${esc(formatDateBr(state.alerts.period?.end || state.filters.end))}</div>
            </div>
          </div>
        </article>
        <article class="panel">
          <div class="panel-head">
            <div>
              <span class="eyebrow">Resumo</span>
              <h3>Média da operação</h3>
            </div>
          </div>
          <div class="mini-grid">
            <div class="mini-card"><span class="muted">Monitorados</span><div class="metric-value">${integer(summary.monitored)}</div></div>
            <div class="mini-card"><span class="muted">Em alerta</span><div class="metric-value">${integer(summary.total)}</div></div>
            <div class="mini-card"><span class="muted">Nota média</span><div class="metric-value">${number(summary.average_score)}</div></div>
            <div class="mini-card"><span class="muted">Maior nota</span><div class="metric-value">${number(summary.max_score)}</div></div>
          </div>
        </article>
      </div>

      <article class="panel">
        <div class="panel-head">
          <div>
            <span class="eyebrow">Prioridade</span>
            <h3>Pontos de atenção por operador</h3>
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
                      <span class="muted">Nota de alerta</span>
                      <div class="metric-value">${number(item.alert_score)}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Produção média</span>
                      <div class="metric-value">${integer(item.avg_production)}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Efetividade</span>
                      <div class="metric-value">${percent(item.effectiveness)}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Qualidade</span>
                      <div class="metric-value">${item.quality ? number(item.quality) : "--"}</div>
                    </div>
                    <div class="mini-card">
                      <span class="muted">Sem ação / vazio</span>
                      <div class="metric-value">${percent(item.no_action_share)}</div>
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
        ` : `<div class="empty">Nenhum operador em alerta no recorte atual.</div>`}
      </article>
    </section>
  `;
}

function historyTemplate() {
  const rows = getScopedHistory();
  const currentHistoryUser = getUserLabelById(state.filters.historyUserId);
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
                ${isManager() ? "<th>Ações</th>" : ""}
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
                  ${isManager() ? `
                    <td>
                      <div class="row-actions">
                        <button class="btn-secondary btn-small" type="button" data-history-edit="${row.metricId}" data-history-operation="${row.operation}" data-history-type="${row.entryType}">Editar</button>
                        <button class="btn-secondary btn-small danger-outline" type="button" data-history-delete="${row.metricId}" data-history-operation="${row.operation}" data-history-type="${row.entryType}">Remover</button>
                      </div>
                    </td>
                  ` : ""}
                </tr>
              `).join("") : `<tr><td colspan="${(isManager() ? 7 : 6) + (showOperatorColumn ? 1 : 0)}"><div class="empty">Sem resultados.</div></td></tr>`}
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
            <div class="action-grid">
              <button class="btn" type="submit">Salvar manutenção</button>
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
            <div class="form-grid">
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
              <button class="btn" type="submit">Importar monitoria</button>
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
              <label>Monitoria 1<input type="number" min="0" max="100" step="0.01" name="monitoria_1"></label>
              <label>Monitoria 2<input type="number" min="0" max="100" step="0.01" name="monitoria_2"></label>
              <label>Monitoria 3<input type="number" min="0" max="100" step="0.01" name="monitoria_3"></label>
              <label>Monitoria 4<input type="number" min="0" max="100" step="0.01" name="monitoria_4"></label>
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
          <div class="list">
            ${state.users.length ? state.users.map((user) => `
              <div class="list-row">
                <div class="list-row-copy">
                  <strong>${esc(user.full_name)}</strong>
                  <span>${esc(user.login)} · ${user.role === "manager" ? "Gestor" : "Operador"}</span>
                </div>
                <div class="list-row-actions">
                  <button class="pill ${user.is_active ? "green" : "red"} user-toggle-btn" type="button" data-user-toggle="${user.id}">
                    ${user.is_active ? "Ativo" : "Inativo"}
                  </button>
                  <button class="btn-secondary btn-small" type="button" data-user-edit="${user.id}">Editar</button>
                  <button class="btn-secondary btn-small danger-outline" type="button" data-user-delete="${user.id}">Apagar</button>
                </div>
              </div>
            `).join("") : `<div class="empty">Nenhum usuário cadastrado.</div>`}
          </div>
        </article>
      </div>
    </section>
  `;
}
function bindLogin() {
  document.getElementById("login-form").addEventListener("submit", async (event) => {
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
  const profileMenuTrigger = document.getElementById("profile-menu-trigger");
  const profileMenuPopover = document.getElementById("profile-menu-popover");
  const passwordModal = document.getElementById("password-modal");
  const userModal = document.getElementById("user-modal");
  const historyEditModal = document.getElementById("history-edit-modal");
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

  const placeTooltipAbovePoint = (tooltip, point) => {
    if (!tooltip || !point) return;
    tooltip.hidden = false;
    tooltip.style.visibility = "hidden";
    const rect = point.getBoundingClientRect();
    const tooltipWidth = tooltip.offsetWidth || 280;
    const tooltipHeight = tooltip.offsetHeight || 120;
    const margin = 12;
    const centeredLeft = rect.left + (rect.width / 2) - (tooltipWidth / 2);
    const resolvedLeft = Math.min(
      window.innerWidth - tooltipWidth - margin,
      Math.max(margin, centeredLeft),
    );
    let resolvedTop = rect.top - tooltipHeight - 10;
    if (resolvedTop < margin) {
      resolvedTop = rect.bottom + 10;
    }
    tooltip.style.left = `${resolvedLeft}px`;
    tooltip.style.top = `${resolvedTop}px`;
    tooltip.style.visibility = "visible";
  };

  if (trendTooltip && isManager()) {
    const trendPoints = document.querySelectorAll(".trend-point[data-trend-date]");
    const showTrendTooltip = async (point, event) => {
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
        ? `<div class="trend-tooltip-line">Carregando...</div>`
        : (bucket.top.length
          ? bucket.top.map((item, idx) => `<div class="trend-tooltip-line"><strong>${idx + 1}. ${esc(item.name)}</strong><span>${integer(item.production)} · ${percent(item.effectiveness)}</span></div>`).join("")
          : `<div class="trend-tooltip-line">Sem dados para ${esc(formatDateBr(key))}.</div>`);
      trendTooltip.innerHTML = `
        <div class="trend-tooltip-title">Top 10 · ${esc(formatDateBr(key))}</div>
        ${lines}
      `;
      placeTooltipAbovePoint(trendTooltip, point);
    };

    trendPoints.forEach((point) => {
      point.addEventListener("mouseenter", (event) => {
        showTrendTooltip(point, event);
      });
      point.addEventListener("mousemove", (event) => {
        showTrendTooltip(point, event);
      });
      point.addEventListener("mouseleave", () => {
        trendTooltip.hidden = true;
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
        ? `<div class="trend-tooltip-line">Carregando...</div>`
        : (bucket.top.length
          ? bucket.top.map((item, idx) => `<div class="trend-tooltip-line"><strong>${idx + 1}. ${esc(item.name)}</strong><span>${formatMetricValue(dataset.metric, item.value)}</span></div>`).join("")
          : `<div class="trend-tooltip-line">Sem dados para ${esc(metricWhenLabel(dataset))}.</div>`);
      analysisMetricTooltip.innerHTML = `
        <div class="trend-tooltip-title">${esc(metricLabel(dataset.metric, dataset.operation, dataset.qualityField))}</div>
        <div class="trend-tooltip-line"><span>${esc(metricWhenLabel(dataset))}</span></div>
        ${lines}
      `;
      placeTooltipAbovePoint(analysisMetricTooltip, point);
    };
    analysisPoints.forEach((point) => {
      point.addEventListener("mouseenter", () => {
        showAnalysisTooltip(point);
      });
      point.addEventListener("mousemove", () => {
        showAnalysisTooltip(point);
      });
      point.addEventListener("mouseleave", () => {
        analysisMetricTooltip.hidden = true;
      });
    });
  }

  if (state.forcePasswordChange && passwordModal) {
    passwordModal.hidden = false;
  }

  document.querySelectorAll("[data-route]").forEach((button) => {
    button.addEventListener("click", async () => {
      clearFlash();
      state.route = button.dataset.route;
      if (state.route === "alerts") state.alerts = null;
      render();
      try {
        if (state.route === "alerts" && isManager()) {
          await loadAlerts();
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

  const historyUserSearch = document.getElementById("history-user-search");
  if (historyUserSearch) {
    historyUserSearch.addEventListener("input", (event) => {
      state.filters.historyQuery = event.target.value;
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
    "reset-analysis": async () => {
      state.filters.start = new Date(Date.now() - 29 * 86400000).toISOString().slice(0, 10);
      state.filters.end = new Date().toISOString().slice(0, 10);
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
      const restoreButton = setButtonProcessing(button, true, "Processando...");
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
      try {
        const response = await fetch("/api/admin/quality/import", { method: "POST", body: form, credentials: "same-origin" });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "Erro ao importar monitoria");
        await loadBootstrap();
        setFlash("success", `Monitoria importada com ${data.processed} operadores.`);
        render();
      } catch (error) {
        setFlash("error", error.message);
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
        const rawMonitorias = [
          payload.monitoria_1,
          payload.monitoria_2,
          payload.monitoria_3,
          payload.monitoria_4,
        ];
        const monitorias = rawMonitorias
          .map((value) => String(value ?? "").trim())
          .filter((value) => value !== "")
          .map((value) => Number(value.replace(",", ".")))
          .filter((value) => Number.isFinite(value) && value >= 0 && value <= 100);
        if (!monitorias.length) {
          throw new Error("Preencha pelo menos uma monitoria (0 a 100).");
        }
        payload.score = monitorias.reduce((sum, value) => sum + value, 0) / monitorias.length;
        await api("/api/admin/quality", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        refreshDashboardInBackground("Qualidade salva com sucesso.");
        qualityManualForm.reset();
        syncManualReference();
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
        setFlash("success", `Base importada: ${data.processed} registro(s), ${data.rejected} rejeição(ões).`);
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
    downloadBaseTemplate.addEventListener("click", () => {
      const modelSelect = document.getElementById("base-model-select");
      const model = modelSelect ? modelSelect.value : "nuvidio";
      window.location.href = `/api/admin/import/template?model=${encodeURIComponent(model)}`;
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
      }
    });
  });

  document.querySelectorAll("[data-user-delete]").forEach((button) => {
    button.addEventListener("click", async () => {
      const user = state.users.find((item) => String(item.id) === String(button.dataset.userDelete));
      if (!user) return;
      if (!window.confirm(`Apagar o usuário ${user.full_name}?`)) return;
      try {
        await api(`/api/admin/users/${user.id}`, { method: "DELETE" });
        await Promise.all([loadUsers(), loadOverview(), loadAnalysis(), loadHistory()]);
        setFlash("success", "Usuário apagado com sucesso.");
        render();
      } catch (error) {
        setFlash("error", error.message);
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
        setField("history-edit-monitoria-1", historyRow.monitoria_1 ?? "");
        setField("history-edit-monitoria-2", historyRow.monitoria_2 ?? "");
        setField("history-edit-monitoria-3", historyRow.monitoria_3 ?? "");
        setField("history-edit-monitoria-4", historyRow.monitoria_4 ?? "");
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
            monitoria_1: form.get("monitoria_1"),
            monitoria_2: form.get("monitoria_2"),
            monitoria_3: form.get("monitoria_3"),
            monitoria_4: form.get("monitoria_4"),
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
        await api(entryType === "quality" ? `/api/admin/quality/${metricId}` : `/api/admin/daily-metrics/${metricId}`, { method: "DELETE" });
        refreshDashboardInBackground(entryType === "quality" ? "Qualidade removida com sucesso." : "Registro removido com sucesso.");
      } catch (error) {
        setFlash("error", error.message);
      } finally {
        restoreButton();
      }
    });
  });

  const downloadQualityTemplate = document.getElementById("download-quality-template");
  if (downloadQualityTemplate) {
    downloadQualityTemplate.addEventListener("click", () => {
      window.location.href = "/api/admin/quality/template";
    });
  }

  const maintenanceForm = document.getElementById("admin-maintenance-form");
  if (maintenanceForm) {
    maintenanceForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const maintenanceToggle = document.getElementById("maintenance-toggle");
      const maintenanceMessage = document.getElementById("maintenance-message");
      const metricProductionRed = document.getElementById("metric-production-red");
      const metricProductionAmber = document.getElementById("metric-production-amber");
      const metricEffectivenessRed = document.getElementById("metric-effectiveness-red");
      const metricEffectivenessAmber = document.getElementById("metric-effectiveness-amber");
      const metricQualityRed = document.getElementById("metric-quality-red");
      const metricQualityAmber = document.getElementById("metric-quality-amber");
      const submitButton = maintenanceForm.querySelector('button[type="submit"]');
      const restoreButton = setButtonProcessing(submitButton, true, "Salvando...");
      try {
        const response = await api("/api/admin/settings", {
          method: "PATCH",
          body: JSON.stringify({
            maintenance_for_operators: Boolean(maintenanceToggle?.checked),
            maintenance_message: String(maintenanceMessage?.value || "").trim(),
            metric_rules: {
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
            },
          }),
        });
        state.appSettings = response.app_settings || state.appSettings;
        setFlash("success", "Modo de manutenção atualizado.");
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

