const state = {
  user: null,
  route: localStorage.getItem("pulse-route") || "overview",
  overview: null,
  analysis: null,
  history: null,
  users: [],
  flash: null,
  forcePasswordChange: false,
  theme: localStorage.getItem("pulse-theme") || "dark",
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
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function number(value) {
  return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 1 }).format(Number(value || 0));
}

function integer(value) {
  return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 }).format(Number(value || 0));
}

function percent(value) {
  return `${Number(value || 0).toFixed(2)}%`;
}

function initials(name) {
  return String(name || "KR").split(" ").filter(Boolean).slice(0, 2).map((item) => item[0]).join("").toUpperCase();
}

function average(values) {
  if (!values.length) return 0;
  return values.reduce((sum, value) => sum + Number(value || 0), 0) / values.length;
}

function isManager() {
  return state.user?.role === "manager";
}

function getUserLabelById(userId) {
  return state.users.find((user) => String(user.id) === String(userId))?.full_name || "";
}

function resolveHistoryUserId(query) {
  const normalized = String(query || "").trim().toLowerCase();
  if (!normalized) return "";
  const exact = state.users.find((user) => user.full_name.trim().toLowerCase() === normalized);
  if (exact) return String(exact.id);
  const partial = state.users.find((user) => user.full_name.trim().toLowerCase().includes(normalized));
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
  const total = actionable + Number(row.calls_nuvidio_no_action);
  return total ? (actionable / total) * 100 : 0;
}

function getScopedHistory() {
  const rows = state.history?.history || [];
  const operation = state.filters.operation;
  return rows.flatMap((row) => {
    const list = [];
    if (operation === "all" || operation === "0800") {
      list.push({
        metricId: row.id,
        date: row.metric_date,
        operation: "0800",
        production: row.production,
        effectiveness: calcOperationEffectiveness(row, "0800"),
        quality: findQualityForDate(row.metric_date),
        updatedAt: row.updated_at,
      });
    }
    if (operation === "all" || operation === "nuvidio") {
      list.push({
        metricId: row.id,
        date: row.metric_date,
        operation: "Nuvidio",
        production: row.production,
        effectiveness: calcOperationEffectiveness(row, "nuvidio"),
        quality: findQualityForDate(row.metric_date),
        updatedAt: row.updated_at,
      });
    }
    return list;
  });
}

function buildOverviewModel() {
  const trend = state.overview?.trend || [];
  const ranking = state.analysis?.ranking || [];
  const quality = state.history?.quality || [];
  const latest = trend[trend.length - 1];
  const previous = trend[trend.length - 2];
  return {
    totalAttended: trend.reduce((sum, item) => sum + Number(item.production || 0), 0),
    avgProduction: average(trend.map((item) => item.production)),
    avgEffectiveness: average(trend.map((item) => item.effectiveness)),
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
  const qualityMonths = (state.history?.quality || []).slice().sort((a, b) => a.reference_month.localeCompare(b.reference_month));
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
    }
    return acc;
  }, { approved: 0, pending: 0, rejected: 0, noAction: 0 });
  return {
    trend,
    qualityMonths,
    filteredRanking,
    status,
    summary: {
      production: average(filteredRanking.map((item) => item.avg_production)),
      effectiveness: average(filteredRanking.map((item) => item.effectiveness)),
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
              <img class="brand-logo" src="/logos_KR-02.png" alt="KR Consulting" onerror="this.style.display='none';this.nextElementSibling.style.display='grid';">
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
  applyTheme();
  if (window.location.protocol === "file:") {
    renderOfflineHint();
    return;
  }
  try {
    const auth = await api("/api/auth/me");
    state.user = auth.user;
    if (state.user) {
      state.filters.historyUserId = String(state.user.id);
      if (!isManager()) state.filters.analysisUserId = String(state.user.id);
      await loadAll();
      state.filters.historyQuery = getUserLabelById(state.filters.historyUserId) || state.user.full_name;
      state.forcePasswordChange = Boolean(state.user.must_change_password);
    }
  } catch {
    state.user = null;
  }
  render();
}

async function loadAll() {
  await Promise.all([
    loadOverview(),
    loadAnalysis(),
    loadHistory(),
    isManager() ? loadUsers() : Promise.resolve(),
  ]);
}

async function loadOverview() {
  state.overview = await api(`/api/dashboard/overview?date=${state.filters.today}&start=${state.filters.start}&end=${state.filters.end}`);
}

async function loadAnalysis() {
  state.analysis = await api(`/api/analysis?start=${state.filters.start}&end=${state.filters.end}`);
}

async function loadHistory() {
  const userId = isManager() ? state.filters.historyUserId || state.user.id : state.user.id;
  state.history = await api(`/api/history?user_id=${userId}&start=${state.filters.start}&end=${state.filters.end}`);
}

async function loadUsers() {
  const response = await api("/api/admin/users");
  state.users = response.users;
  if (isManager()) {
    state.filters.historyQuery = getUserLabelById(state.filters.historyUserId) || state.filters.historyQuery;
  }
}

function navMeta() {
  return {
    overview: "Visão geral",
    analysis: "Análises",
    history: "Histórico",
    admin: "Gestão",
  };
}

function render() {
  applyTheme();
  if (!state.user) {
    app.innerHTML = loginTemplate();
    bindLogin();
    return;
  }
  app.innerHTML = shellTemplate();
  bindShellEvents();
}

function flashTemplate() {
  if (!state.flash) return "";
  return `<div class="notice toast ${esc(state.flash.type)}">${esc(state.flash.message)}</div>`;
}

function shellTemplate() {
  const titles = {
    overview: { title: isManager() ? "Visão da operação" : "Performance do operador", desc: "" },
    analysis: { title: "Análises", desc: "" },
    history: { title: "Histórico", desc: "" },
    admin: { title: "Gestão", desc: "" },
  };
  const current = titles[state.route];
  const selectedLabel = isManager()
    ? (state.filters.analysisUserId === "all" ? "Todos os operadores" : state.users.find((user) => String(user.id) === String(state.filters.analysisUserId))?.full_name || "Operador")
    : state.user.full_name;
  return `
    <div class="shell">
      <aside class="sidebar">
        <div class="brand-box">
          <div class="brand">
            <div class="brand-logo-wrap">
              <img class="brand-logo" src="/logos_KR-02.png" alt="KR Consulting" onerror="this.style.display='none';this.nextElementSibling.style.display='grid';">
              <div class="brand-mark" style="display:none;">KR</div>
            </div>
            <div class="brand-copy">
              <h1>PORTAL DE RESULTADOS</h1>
              <p>Performance operacional</p>
            </div>
          </div>
        </div>
        <nav class="nav">
          ${Object.entries(navMeta()).filter(([key]) => key !== "admin" || isManager()).map(([key, label]) => `
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
              <select id="global-user-select">
                ${isManager() ? `
                  <option value="all" ${state.filters.analysisUserId === "all" ? "selected" : ""}>Todos os operadores</option>
                  ${state.users.map((user) => `<option value="${user.id}" ${String(user.id) === String(state.filters.analysisUserId) ? "selected" : ""}>${esc(user.full_name)}</option>`).join("")}
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
            ${state.forcePasswordChange ? `<div class="info-box">Para continuar, defina uma nova senha para substituir a senha padrão <strong>Trocar@01</strong>.</div>` : `<label>Senha atual<input name="current_password" type="password" required></label>`}
            <label>Nova senha<input name="new_password" type="password" minlength="4" required></label>
            <label>Confirmar nova senha<input name="confirm_password" type="password" minlength="4" required></label>
            <div class="action-grid">
              ${state.forcePasswordChange ? "" : `<button class="btn-secondary" type="button" id="cancel-password-modal">Cancelar</button>`}
              <button class="btn" type="submit">${state.forcePasswordChange ? "Salvar e continuar" : "Salvar senha"}</button>
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
              <img class="brand-logo" src="/logos_KR-02.png" alt="KR Consulting" onerror="this.style.display='none';this.nextElementSibling.style.display='grid';">
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
  if (state.route === "history") return historyTemplate();
  if (state.route === "admin") return adminTemplate();
  return overviewTemplate();
}

function overviewTemplate() {
  const model = buildOverviewModel();
  const latestDate = formatDateBr(model.latest?.date || model.latest?.metric_date || "--");
  const latestProduction = model.latest ? integer(model.latest.production) : "--";
  const latestEffectiveness = model.latest ? percent(model.latest.effectiveness) : "--%";
  const cards = [
    { label: "Total atendido", value: integer(model.totalAttended) },
    { label: "Média de produção", value: number(model.avgProduction) || "--" },
    { label: "Efetividade média", value: model.avgEffectiveness ? percent(model.avgEffectiveness) : "--%" },
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
          ${(state.overview?.trend || []).length ? state.overview.trend.map((item) => `
            <div class="chart-col">
              <div class="column" style="height:${Math.max(14, (Number(item.production || 0) / Math.max(...state.overview.trend.map((row) => Number(row.production || 0)), 1)) * 110)}px;"></div>
              <small>${shortDate(item.date)}</small>
            </div>
          `).join("") : `<div class="empty">Sem tendência disponível.</div>`}
        </div>
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
  const maxQuality = Math.max(...qualityMonths.map((item) => Number(item.score || 0)), 10);
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
          <label>Operação
            <select id="operation-filter">
              <option value="all" ${state.filters.operation === "all" ? "selected" : ""}>Nuvidio + 0800</option>
              <option value="nuvidio" ${state.filters.operation === "nuvidio" ? "selected" : ""}>Nuvidio</option>
              <option value="0800" ${state.filters.operation === "0800" ? "selected" : ""}>0800</option>
            </select>
          </label>
          <div class="filter-actions">
            <button class="btn" data-action="refresh-analysis">Aplicar</button>
            <button class="btn-secondary" data-action="reset-analysis">Limpar</button>
          </div>
        </aside>
        <div class="section">
          <div class="hero-grid">
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Produção</span><h3>Série diária</h3></div></div>
              <div class="chart">
                ${trend.length ? trend.map((item) => `<div class="chart-col"><div class="column" style="height:${Math.max(12, (item.production / maxProduction) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Efetividade</span><h3>Série diária</h3></div></div>
              <div class="chart">
                ${trend.length ? trend.map((item) => `<div class="chart-col"><div class="column green" style="height:${Math.max(12, (item.effectiveness / maxEffectiveness) * 110)}px;"></div><small>${shortDate(item.date)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
          </div>

          <div class="hero-grid">
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Qualidade</span><h3>Mensal</h3></div></div>
              <div class="chart">
                ${qualityMonths.length ? qualityMonths.map((item) => `<div class="chart-col"><div class="column amber" style="height:${Math.max(12, (item.score / maxQuality) * 110)}px;"></div><small>${formatMonthLabel(item.reference_month)}</small></div>`).join("") : `<div class="empty">Sem dados.</div>`}
              </div>
            </article>
            <article class="panel">
              <div class="panel-head"><div><span class="eyebrow">Resumo</span><h3>Consolidado</h3></div></div>
              <div class="mini-grid">
                <div class="mini-card"><span class="muted">Produção média</span><div class="metric-value">${number(model.summary.production)}</div></div>
                <div class="mini-card"><span class="muted">Efetividade média</span><div class="metric-value">${percent(model.summary.effectiveness)}</div></div>
              </div>
            </article>
          </div>
        </div>
      </div>
    </section>
  `;
}

function historyTemplate() {
  const rows = getScopedHistory();
  const currentHistoryUser = getUserLabelById(state.filters.historyUserId);
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
                <input id="history-user-search" list="history-user-options" value="${esc(state.filters.historyQuery || currentHistoryUser)}" placeholder="Pesquisar operador">
                <datalist id="history-user-options">
                  ${state.users.map((user) => `<option value="${esc(user.full_name)}"></option>`).join("")}
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
                <th>Data</th>
                <th>Operação</th>
                <th>Produção</th>
                <th>Efetividade</th>
                <th>Qualidade</th>
                <th>Atualizado</th>
              </tr>
            </thead>
            <tbody>
              ${rows.length ? rows.map((row) => `
                <tr>
                  <td>${esc(formatDateBr(row.date))}</td>
                  <td><span class="pill ${row.operation === "0800" ? "amber" : "blue"}">${esc(row.operation)}</span></td>
                  <td>${integer(row.production)}</td>
                  <td>${percent(row.effectiveness)}</td>
                  <td>${number(row.quality)}</td>
                  <td>${esc(formatDateTimeBr(row.updatedAt))}</td>
                </tr>
              `).join("") : `<tr><td colspan="6"><div class="empty">Sem resultados.</div></td></tr>`}
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
  return `
    <section class="section">
      <div class="hero-grid">
        <article class="panel">
          <div class="management-header">
            <div>
              <span class="eyebrow">Gestão</span>
              <h3>Base principal</h3>
            </div>
          </div>
          <div class="panel-head"><div><span class="eyebrow">Planilha por data</span><h3>Importação em massa</h3></div></div>
          <form id="import-form" class="section">
            <div class="form-grid">
              <label>Operação
                <select name="operation_hint">
                  <option value="nuvidio">Nuvidio</option>
                  <option value="0800">0800</option>
                </select>
              </label>
              <label>Aba (opcional)<input name="sheet_name"></label>
            </div>
            <label>Arquivo<input type="file" name="file" accept=".csv" required></label>
            <div class="action-grid">
              <button class="btn" type="submit">Importar planilha</button>
              <button class="btn-secondary" type="button" data-action="refresh-r2">Atualizar dados (R2)</button>
            </div>
          </form>
        </article>
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
                <div>
                  <strong>${esc(user.full_name)}</strong>
                  <span>${esc(user.login)} · ${user.role === "manager" ? "Gestor" : "Operador"}</span>
                </div>
                <span class="pill ${user.must_change_password ? "amber" : "green"}">${user.must_change_password ? "Troca pendente" : "Ativo"}</span>
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
    try {
      const response = await api("/api/auth/login", {
        method: "POST",
        body: JSON.stringify({ login: form.get("login"), password: form.get("password") }),
      });
      state.user = response.user;
      clearFlash();
      state.forcePasswordChange = Boolean(state.user.must_change_password);
      state.filters.historyUserId = String(state.user.id);
      state.filters.analysisUserId = isManager() ? "all" : String(state.user.id);
      await loadAll();
      state.filters.historyQuery = getUserLabelById(state.filters.historyUserId) || state.user.full_name;
      render();
    } catch (error) {
      setFlash("error", error.message);
    }
  });
}

function bindShellEvents() {
  const profileMenuTrigger = document.getElementById("profile-menu-trigger");
  const profileMenuPopover = document.getElementById("profile-menu-popover");
  const passwordModal = document.getElementById("password-modal");

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

  if (passwordModal) {
    passwordModal.addEventListener("click", (event) => {
      if (!state.forcePasswordChange && event.target === passwordModal) closePasswordModal();
    });

    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && !passwordModal.hidden && !state.forcePasswordChange) {
        closePasswordModal();
      }
    });
  }

  if (state.forcePasswordChange && passwordModal) {
    passwordModal.hidden = false;
  }

  document.querySelectorAll("[data-route]").forEach((button) => {
    button.addEventListener("click", () => {
      clearFlash();
      state.route = button.dataset.route;
      localStorage.setItem("pulse-route", state.route);
      render();
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
      state.filters.historyUserId = event.target.value === "all" ? String(state.user.id) : event.target.value;
      state.filters.historyQuery = getUserLabelById(state.filters.historyUserId) || "";
      if (event.target.value !== "all") await loadHistory();
      render();
    });
  }

  document.querySelectorAll("#operation-filter").forEach((input) => {
    input.addEventListener("change", (event) => {
      state.filters.operation = event.target.value;
      render();
    });
  });

  const actionMap = {
    "logout": async () => {
      await api("/api/auth/logout", { method: "POST" });
      state.user = null;
      state.route = "overview";
      localStorage.setItem("pulse-route", state.route);
      setFlash("success", "Sessão encerrada.");
    },
    "refresh-all": async () => {
      await loadAll();
      setFlash("success", "Atualizado.");
    },
    "refresh-analysis": async () => {
      await Promise.all([loadAnalysis(), loadOverview(), loadHistory()]);
      setFlash("success", "Análises atualizadas.");
    },
    "refresh-history": async () => {
      if (isManager()) {
        const resolvedUserId = resolveHistoryUserId(state.filters.historyQuery);
        if (!resolvedUserId) {
          throw new Error("Selecione um operador válido para pesquisar.");
        }
        state.filters.historyUserId = resolvedUserId;
        state.filters.historyQuery = getUserLabelById(resolvedUserId);
      }
      await loadHistory();
      setFlash("success", "Histórico atualizado.");
    },
    "reset-analysis": async () => {
      state.filters.start = new Date(Date.now() - 29 * 86400000).toISOString().slice(0, 10);
      state.filters.end = new Date().toISOString().slice(0, 10);
      state.filters.operation = "all";
      state.filters.analysisUserId = isManager() ? "all" : String(state.user.id);
      await Promise.all([loadAnalysis(), loadOverview(), loadHistory()]);
      setFlash("success", "Filtros redefinidos.");
    },
    "toggle-theme": async () => {
      state.theme = state.theme === "dark" ? "contrast" : "dark";
      localStorage.setItem("pulse-theme", state.theme);
      render();
    },
    "refresh-r2": async () => setFlash("success", "Atualização do R2 acionada."),
  };

  document.querySelectorAll("[data-action]").forEach((button) => {
    const action = actionMap[button.dataset.action];
    if (!action) return;
    button.addEventListener("click", async () => {
      try {
        await action();
      } catch (error) {
        setFlash("error", error.message);
      }
    });
  });

  const importForm = document.getElementById("import-form");
  if (importForm) {
    importForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      try {
        const response = await fetch("/api/admin/import", { method: "POST", body: form, credentials: "same-origin" });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "Erro ao importar");
        await Promise.all([loadOverview(), loadAnalysis(), loadHistory(), loadUsers()]);
        setFlash("success", `Importação concluída com ${data.processed} linhas.`);
        render();
      } catch (error) {
        setFlash("error", error.message);
      }
    });
  }

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
        await loadHistory();
        setFlash("success", `Monitoria importada com ${data.processed} operadores.`);
        render();
      } catch (error) {
        setFlash("error", error.message);
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

  const downloadQualityTemplate = document.getElementById("download-quality-template");
  if (downloadQualityTemplate) {
    downloadQualityTemplate.addEventListener("click", () => {
      window.location.href = "/api/admin/quality/template";
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
