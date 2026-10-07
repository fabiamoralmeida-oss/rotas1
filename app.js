const PAGE_SIZE = 12;
const AUTOLOAD_PATH = "Planilha%20sem%20t%C3%ADtulo%20(84).xlsx";
const STATUS_COLORS = { close: "#16836f", active: "#4b9cba", planned: "#d6a83c", return_to_station: "#d97861" };
const PROMISE_COLORS = { "On Time": "#16836f", Delay: "#d76b58", Early: "#4b86a8", "Sem Promessa": "#a28b58" };
const PROMISE_ICONS = {
  "On Time": '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="m8 12 2.5 2.5L16 9"></path></svg>',
  Delay: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="M12 7v5l3 2"></path></svg>',
  Early: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="M12 16V8m-4 4 4-4 4 4"></path></svg>',
  "Sem Promessa": '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="5" width="16" height="15" rx="2"></rect><path d="M8 3v4m8-4v4M4 10h16m-12 5h2"></path></svg>',
  default: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="M12 11v5m0-8h.01"></path></svg>',
};
const ADDRESS_COLORS = { residential: "#43856e", business: "#bb7848" };
const ADDRESS_ICONS = {
  residential: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m3 11 9-7 9 7v9H5v-9m5 9v-6h4v6"></path></svg>',
  business: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2"></rect><path d="M8 7h2m4 0h2M8 11h2m4 0h2M10 21v-5h4v5"></path></svg>',
  default: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="M12 11v5m0-8h.01"></path></svg>',
};
const LABELS = { on_route: "Em rota", at_station: "Na base", delivered: "Entregue", delivered_place: "Entregue no local", transferred: "Transferido", problem_solving: "Em tratativa", buyer_absent: "Destinatário ausente", missrouted: "Rota incorreta", unvisited_address: "Endereço não visitado", inaccessible_address: "Acesso indisponível", business_closed: "Estabelecimento fechado", for_return: "Para devolução", soon_deliver: "Próxima entrega", residential: "Residencial", business: "Comercial", close: "Encerrada", active: "Em andamento", planned: "Planejada", return_to_station: "Retorno à base" };
const REQUIRED_COLUMNS = ["Base_Origem", "Destino_XPT_Agencia", "ID_Rota", "Ciclo_Rota", "Status_Rota", "Transportadora", "Tipo_Veiculo", "ID_Pacote", "CIDADE_DESTINO", "TIPO_ENDERECO", "Status_Pacote_Na_Rota", "Substatus_Pacote_Na_Rota", "Data_Promessa", "STATUS_PROMESSA"];
let records = [];
let filtered = [];
let deliveredPackageIds = new Set();
let page = 1;
const selectedSubstatuses = new Set();
const selectedCarriers = new Set();
const selectedRouteStatuses = new Set();

const $ = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat("pt-BR");

function label(value) {
  return LABELS[value] || value || "Não informado";
}

function formatCarrierName(value) {
  const name = String(value ?? "").trim().toLocaleLowerCase("pt-BR");
  return name.replace(/(^|[\s-])(\p{L})/gu, (_, separator, firstLetter) => `${separator}${firstLetter.toLocaleUpperCase("pt-BR")}`);
}

function safeText(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
}

function setNotice(message) {
  $("notice").textContent = message;
  $("notice").hidden = !message;
}

function excelDate(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") return new Date(Date.UTC(1899, 11, 30) + value * 86400000).toISOString().slice(0, 10);
  const parsed = new Date(value);
  return Number.isNaN(parsed.valueOf()) ? String(value) : parsed.toISOString().slice(0, 10);
}

function parseRows(rows) {
  const missing = REQUIRED_COLUMNS.filter((column) => !rows[0] || !(column in rows[0]));
  if (missing.length) throw new Error(`Colunas não encontradas: ${missing.join(", ")}`);
  return rows.filter((row) => row.ID_Pacote || row.ID_Rota).map((row) => ({
    base: row.Base_Origem || "",
    destination: row.Destino_XPT_Agencia || "",
    route: String(row.ID_Rota ?? ""),
    cycle: row.Ciclo_Rota || "",
    routeStatus: row.Status_Rota || "",
    carrier: row.Transportadora || "",
    vehicle: row.Tipo_Veiculo || "",
    package: String(row.ID_Pacote ?? ""),
    city: row.CIDADE_DESTINO || "",
    addressType: row.TIPO_ENDERECO || "",
    packageStatus: row.Status_Pacote_Na_Rota || "",
    substatus: row.Substatus_Pacote_Na_Rota || "",
    promiseDate: excelDate(row.Data_Promessa),
    promiseStatus: row.STATUS_PROMESSA || "",
  }));
}

function countBy(data, key) {
  return data.reduce((counts, record) => {
    const value = record[key] || "Não informado";
    counts.set(value, (counts.get(value) || 0) + 1);
    return counts;
  }, new Map());
}

function isDelivered(record) {
  return ["delivered", "delivered_place"].includes(record.substatus) || ["delivered", "delivered_place"].includes(record.packageStatus);
}

function renderMetrics(data) {
  const routes = new Map();
  data.forEach((record) => {
    if (!routes.has(record.route)) routes.set(record.route, record.routeStatus);
  });
  const active = [...routes.values()].filter((status) => status === "active").length;
  const delivered = data.filter((record) => record.substatus === "delivered" || record.packageStatus === "delivered").length;
  const delayed = data.filter((record) => record.promiseStatus === "Delay").length;
  $("metric-packages").textContent = fmt.format(data.length);
  $("metric-active").textContent = fmt.format(active);
  $("metric-delivered").textContent = fmt.format(delivered);
  $("metric-delayed").textContent = fmt.format(delayed);
  $("delivered-rate").textContent = `${data.length ? (delivered * 100 / data.length).toFixed(1).replace(".", ",") : "0"}%`;
  $("nav-route-count").textContent = fmt.format(routes.size);
  renderRouteStatus(routes);
  renderPromiseChart(data);
  renderAddressChart(data);
  renderCarriers(data);
}

function renderRouteStatus(routes) {
  const counts = new Map();
  routes.forEach((status) => counts.set(status || "Não informado", (counts.get(status || "Não informado") || 0) + 1));
  const sorted = [...counts].sort((a, b) => b[1] - a[1]);
  const total = routes.size;
  $("route-total").textContent = `${fmt.format(total)} rotas`;
  $("route-bar").innerHTML = sorted.map(([status, count]) => `<span title="${safeText(label(status))}: ${fmt.format(count)}" style="width:${count * 100 / (total || 1)}%;background:${STATUS_COLORS[status] || "#aab7b1"}"></span>`).join("");
  $("route-legend").innerHTML = sorted.map(([status, count]) => `<div class="legend-item"><span class="legend-dot" style="background:${STATUS_COLORS[status] || "#aab7b1"}"></span><span>${safeText(label(status))}</span><strong>${fmt.format(count)}</strong></div>`).join("");
}

function renderCarriers(data) {
  const carriers = [...countBy(data, "carrier")].filter(([name]) => name !== "Não informado").sort((a, b) => b[1] - a[1]);
  const max = carriers[0]?.[1] || 1;
  $("carrier-total").textContent = `${fmt.format(carriers.length)} transportadoras`;
  $("carrier-list").innerHTML = carriers.map(([name, count]) => `<div class="carrier-row"><span class="carrier-name" title="${safeText(formatCarrierName(name))}">${safeText(formatCarrierName(name))}</span><div class="carrier-track"><div class="carrier-fill" style="width:${count * 100 / max}%"></div></div><span class="carrier-value">${fmt.format(count)}</span></div>`).join("");
}

function renderPromiseChart(data) {
  renderBreakdownChart(data, "promiseStatus", "promise-chart", "promise-total", PROMISE_COLORS, PROMISE_ICONS, false);
}

function renderAddressChart(data) {
  const types = [...countBy(data, "addressType")].sort((a, b) => b[1] - a[1]);
  const total = data.length;
  $("address-total").textContent = `${fmt.format(types.length)} ${types.length === 1 ? "tipo" : "tipos"} · ${fmt.format(total)} pacotes`;
  $("address-chart").innerHTML = types.map(([type, count]) => {
    const typeRecords = data.filter((record) => (record.addressType || "Não informado") === type);
    const color = ADDRESS_COLORS[type] || "#718096";
    const share = total ? count * 100 / total : 0;
    const promises = new Map();
    typeRecords.forEach((record) => {
      const status = record.promiseStatus || "Sem Promessa";
      promises.set(status, (promises.get(status) || 0) + 1);
    });
    const promiseOrder = ["Early", "On Time", "Delay", "Sem Promessa"];
    const sortedPromises = [...promises].sort((a, b) => {
      const indexA = promiseOrder.indexOf(a[0]);
      const indexB = promiseOrder.indexOf(b[0]);
      return (indexA < 0 ? promiseOrder.length : indexA) - (indexB < 0 ? promiseOrder.length : indexB) || b[1] - a[1];
    });
    const promiseChips = sortedPromises.map(([status, amount]) => `<span class="address-promise-chip" style="--status-color:${PROMISE_COLORS[status] || "#8a9691"}"><span class="address-status-dot"></span><span>${safeText(status)}</span><strong>${fmt.format(amount)}</strong><small>${Math.round(amount * 100 / count)}%</small></span>`).join("");
    const onTime = (promises.get("Early") || 0) + (promises.get("On Time") || 0);
    const delayed = promises.get("Delay") || 0;
    const routes = new Map();
    typeRecords.forEach((record) => {
      const route = record.route || "Sem ID";
      if (!routes.has(route)) routes.set(route, { total: 0, promises: new Map() });
      const routeStats = routes.get(route);
      const status = record.promiseStatus || "Sem Promessa";
      routeStats.total += 1;
      routeStats.promises.set(status, (routeStats.promises.get(status) || 0) + 1);
    });
    const routeRows = [...routes].sort((a, b) => a[0].localeCompare(b[0], "pt-BR", { numeric: true })).map(([route, stats]) => `<tr><td>${safeText(route)}</td><td>${fmt.format(stats.total)}</td><td>${fmt.format(stats.promises.get("Early") || 0)}</td><td>${fmt.format(stats.promises.get("On Time") || 0)}</td><td>${fmt.format(stats.promises.get("Delay") || 0)}</td><td>${fmt.format(stats.promises.get("Sem Promessa") || 0)}</td></tr>`).join("");
    const name = type === "residential" ? "Residential" : type === "business" ? "Business" : formatCarrierName(type);
    const icon = ADDRESS_ICONS[type] || ADDRESS_ICONS.default;
    return `<article class="address-type-card" style="--type-accent:${color}"><div class="address-card-header"><div class="address-card-title"><span class="address-type-icon">${icon}</span><div><h3>${safeText(name)}</h3><p>${share.toFixed(1).replace(".", ",")}% do total · ${fmt.format(routes.size)} rotas</p></div></div><strong class="address-card-volume">${fmt.format(count)}</strong></div><div class="address-volume-track"><span style="width:${share}%"></span></div><div class="address-promise-chips">${promiseChips}</div><div class="address-timeliness"><span><span class="timeliness-check">✓</span>No prazo: <strong>${(onTime * 100 / count).toFixed(1).replace(".", ",")}%</strong></span><span><span class="timeliness-delay">▲</span>Atraso: <strong>${(delayed * 100 / count).toFixed(1).replace(".", ",")}%</strong></span></div><details class="address-route-details"><summary>Ver detalhes por rota (${fmt.format(routes.size)})</summary><div class="address-route-table-wrap"><table><thead><tr><th>ID_ROTA</th><th>PACOTES</th><th>EARLY</th><th>ON TIME</th><th>DELAY</th><th>SEM PROMESSA</th></tr></thead><tbody>${routeRows}</tbody></table></div></details></article>`;
  }).join("");
}

function renderCarrierPendingRoutes(data) {
  const targetCarriers = ["Envios Extra", "Agencias Kangu"];
  const routeStatuses = ["active", "close"];
  const routes = new Map();
  data.forEach((record) => {
    if (!targetCarriers.includes(record.carrier) || !routeStatuses.includes(record.routeStatus)) return;
    if (!record.package || deliveredPackageIds.has(record.package)) return;
    const key = `${record.carrier}\u0000${record.routeStatus}\u0000${record.route || "Sem ID"}`;
    if (!routes.has(key)) routes.set(key, { carrier: record.carrier, status: record.routeStatus, id: record.route || "Sem ID", packages: new Map() });
    const route = routes.get(key);
    if (!route.packages.has(record.package)) route.packages.set(record.package, record.substatus || record.packageStatus || "Não informado");
  });
  const carrierOrder = new Map(targetCarriers.map((carrier, index) => [carrier, index]));
  const statusOrder = new Map(routeStatuses.map((status, index) => [status, index]));
  const sortedRoutes = [...routes.values()].sort((a, b) => carrierOrder.get(a.carrier) - carrierOrder.get(b.carrier) || statusOrder.get(a.status) - statusOrder.get(b.status) || a.id.localeCompare(b.id, "pt-BR", { numeric: true }));
  const uniquePackages = new Set(sortedRoutes.flatMap((route) => [...route.packages.keys()]));
    $("pending-route-total").textContent = `${fmt.format(uniquePackages.size)} pacotes · ${fmt.format(sortedRoutes.length)} rotas`; // Updated to show package/route count
  $("pending-route-rows").innerHTML = sortedRoutes.length ? sortedRoutes.map((route) => {
    const statuses = new Map();
    route.packages.forEach((status) => statuses.set(status, (statuses.get(status) || 0) + 1));
    const statusList = [...statuses].sort((a, b) => b[1] - a[1]).map(([status, count]) => `<span class="pending-status">${safeText(label(status))}<strong>${fmt.format(count)}</strong></span>`).join("");
    const packageList = [...route.packages].map(([packageId, status]) => `<div class="pending-package"><span class="package-id">${safeText(packageId)}</span><span>${safeText(label(status))}</span></div>`).join("");
    return `<tr><td>${safeText(route.carrier)}</td><td><span class="status-pill ${statusClass(route.status)}">${route.status === "active" ? "Em aberto" : "Encerrada"}</span></td><td><span class="route-id">${safeText(route.id)}</span></td><td><strong class="pending-count">${fmt.format(route.packages.size)}</strong></td><td><div class="pending-status-list">${statusList}</div></td><td><details class="pending-package-details"><summary>Ver IDs (${fmt.format(route.packages.size)})</summary><div class="pending-package-list">${packageList}</div></details></td></tr>`;
  }).join("") : '<tr><td colspan="6" class="empty-state">Nenhuma pendência para essas transportadoras nos filtros atuais.</td></tr>';
}

function renderBreakdownChart(data, key, chartId, totalId, colors, icons, translateLabels) {
  const categories = [...countBy(data, key)].sort((a, b) => b[1] - a[1]);
  const total = data.length;
  const max = categories[0]?.[1] || 1;
  $(totalId).textContent = `${fmt.format(total)} ${total === 1 ? "pacote" : "pacotes"}`;
  $(chartId).innerHTML = categories.map(([value, count]) => {
    const color = colors[value] || "#8a9691";
    const percent = total ? (count * 100 / total).toFixed(1).replace(".", ",") : "0,0";
    const icon = icons[value] || icons.default;
    const text = translateLabels ? label(value) : value;
    return `<div class="promise-row"><span class="promise-icon" style="--promise-color:${color}">${icon}</span><div class="promise-detail"><div class="promise-meta"><span class="promise-label">${safeText(text)}</span><strong>${fmt.format(count)} <small>${percent}%</small></strong></div><div class="promise-track"><span style="width:${count * 100 / max}%;background:${color}"></span></div></div></div>`;
  }).join("");
}

function fillSelect(select, values, firstLabel, formatValue = label) {
  const previous = select.value;
  select.innerHTML = `<option value="">${firstLabel}</option>` + values.map((value) => `<option value="${safeText(value)}">${safeText(formatValue(value))}</option>`).join("");
  if (values.includes(previous)) select.value = previous;
}

function populateFilters() {
  populateRouteFilter();
  populateSubstatusFilter();
  fillSelect($("promise-filter"), [...new Set(records.map((record) => record.promiseStatus).filter(Boolean))].sort(), "Todas as previsões");
  fillSelect($("address-filter"), [...new Set(records.map((record) => record.addressType).filter(Boolean))].sort(), "Todos os tipos de endereço");
  fillSelect($("city-filter"), [...new Set(records.map((record) => record.city).filter(Boolean))].sort((a, b) => a.localeCompare(b, "pt-BR")), "Todas as cidades");
  populateCarrierFilter();
}

function routeStatusLabel(value) {
  if (value === "active") return "Em aberto";
  if (value === "close") return "Encerrada";
  return label(value);
}

function populateRouteFilter() {
  const values = [...new Set(records.map((record) => record.routeStatus).filter(Boolean))].sort((a, b) => routeStatusLabel(a).localeCompare(routeStatusLabel(b), "pt-BR"));
  selectedRouteStatuses.forEach((status) => {
    if (!values.includes(status)) selectedRouteStatuses.delete(status);
  });
  $("route-options").innerHTML = values.map((status) => `<label class="multi-filter-option"><input type="checkbox" value="${safeText(status)}"${selectedRouteStatuses.has(status) ? " checked" : ""}><span>${safeText(routeStatusLabel(status))}</span></label>`).join("");
  updateRouteSummary(values);
}

function updateRouteSummary(values = [...new Set(records.map((record) => record.routeStatus).filter(Boolean))]) {
  const count = selectedRouteStatuses.size;
  const selectAll = $("route-select-all");
  selectAll.checked = values.length > 0 && count === values.length;
  selectAll.indeterminate = count > 0 && count < values.length;
  selectAll.disabled = values.length === 0;
  $("route-summary").textContent = count === 0 ? "Todos os status de rota" : count === 1 ? routeStatusLabel([...selectedRouteStatuses][0]) : `${count} status selecionados`;
}

function populateCarrierFilter() {
  const values = [...new Set(records.map((record) => record.carrier).filter(Boolean))].sort((a, b) => a.localeCompare(b, "pt-BR"));
  selectedCarriers.forEach((carrier) => {
    if (!values.includes(carrier)) selectedCarriers.delete(carrier);
  });
  $("carrier-options").innerHTML = values.map((carrier) => `<label class="multi-filter-option"><input type="checkbox" value="${safeText(carrier)}"${selectedCarriers.has(carrier) ? " checked" : ""}><span>${safeText(formatCarrierName(carrier))}</span></label>`).join("");
  updateCarrierSummary(values);
}

function updateCarrierSummary(values = [...new Set(records.map((record) => record.carrier).filter(Boolean))]) {
  const count = selectedCarriers.size;
  const selectAll = $("carrier-select-all");
  selectAll.checked = values.length > 0 && count === values.length;
  selectAll.indeterminate = count > 0 && count < values.length;
  selectAll.disabled = values.length === 0;
  $("carrier-summary").textContent = count === 0 ? "Todas as transportadoras" : count === 1 ? formatCarrierName([...selectedCarriers][0]) : `${count} transportadoras selecionadas`;
}

function populateSubstatusFilter() {
  const values = [...new Set(records.map((record) => record.substatus))].sort((a, b) => label(a).localeCompare(label(b), "pt-BR"));
  selectedSubstatuses.forEach((value) => {
    if (!values.includes(value)) selectedSubstatuses.delete(value);
  });
  $("substatus-options").innerHTML = values.map((value) => `<label class="multi-filter-option"><input type="checkbox" value="${safeText(value)}"${selectedSubstatuses.has(value) ? " checked" : ""}><span>${safeText(label(value))}</span></label>`).join("");
  updateSubstatusSummary();
}

function updateSubstatusSummary() {
  const count = selectedSubstatuses.size;
  const total = new Set(records.map((record) => record.substatus)).size;
  const selectAll = $("substatus-select-all");
  selectAll.checked = total > 0 && count === total;
  selectAll.indeterminate = count > 0 && count < total;
  selectAll.disabled = total === 0;
  $("substatus-summary").textContent = count === 0 ? "Status do Pacote" : count === 1 ? label([...selectedSubstatuses][0]) : `${count} status selecionados`;
}

function statusClass(value) {
  if (["delivered", "delivered_place"].includes(value)) return "delivered";
  if (value === "Delay") return "delay";
  if (value === "On Time") return "on-time";
  if (value === "Early") return "early";
  if (["active", "on_route"].includes(value)) return value === "active" ? "active" : "on-route";
  if (["problem_solving", "return_to_station", "for_return"].includes(value)) return "problem";
  return value;
}

function renderTable() {
  const terms = $("search-input").value.toLocaleLowerCase("pt-BR").split(/[\s,;|]+/).map((term) => term.trim()).filter(Boolean);
  const pendingOnly = $("pending-filter").checked;
  const promiseStatus = $("promise-filter").value;
  const addressType = $("address-filter").value;
  const city = $("city-filter").value;
  const matchingRecords = records.filter((record) => {
    if (pendingOnly && (!record.package || deliveredPackageIds.has(record.package))) return false;
    if (selectedRouteStatuses.size && !selectedRouteStatuses.has(record.routeStatus)) return false;
    if (selectedSubstatuses.size && !selectedSubstatuses.has(record.substatus)) return false;
    if (promiseStatus && record.promiseStatus !== promiseStatus) return false;
    if (addressType && record.addressType !== addressType) return false;
    if (city && record.city !== city) return false;
    if (selectedCarriers.size && !selectedCarriers.has(record.carrier)) return false;
    const routeId = String(record.route).toLocaleLowerCase("pt-BR");
    const packageId = String(record.package).toLocaleLowerCase("pt-BR");
    const textFields = [record.city, record.carrier].map((value) => String(value).toLocaleLowerCase("pt-BR"));
    if (terms.length && !terms.some((term) => term === routeId || term === packageId || textFields.some((value) => value.includes(term)))) return false;
    return true;
  });
  if (terms.length) {
    const uniquePackages = new Map();
    matchingRecords.forEach((record, index) => {
      const packageId = record.package || `sem-id-${index}`;
      if (!uniquePackages.has(packageId)) uniquePackages.set(packageId, record);
    });
    filtered = [...uniquePackages.values()];
  } else {
    filtered = matchingRecords;
  }
  const packagesByRoute = new Map();
  filtered.forEach((record) => {
    const routeId = record.route || "Sem ID";
    if (!packagesByRoute.has(routeId)) packagesByRoute.set(routeId, new Set());
    if (record.package) packagesByRoute.get(routeId).add(record.package);
  });
  renderMetrics(filtered);
  const pages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  page = Math.min(page, pages);
  const start = (page - 1) * PAGE_SIZE;
  const pageRows = filtered.slice(start, start + PAGE_SIZE);
  const uniquePackageCount = new Set(filtered.map((record) => record.package).filter(Boolean)).size;
  const visibleRouteCount = new Set(filtered.map((record) => record.route).filter(Boolean)).size;
  $("result-count").textContent = pendingOnly ? `${fmt.format(uniquePackageCount)} pacotes · ${fmt.format(visibleRouteCount)} rotas` : terms.length ? `${fmt.format(filtered.length)} pacotes únicos` : `${fmt.format(filtered.length)} registros`;
  $("page-label").textContent = filtered.length ? `${fmt.format(start + 1)}–${fmt.format(Math.min(start + PAGE_SIZE, filtered.length))} de ${fmt.format(filtered.length)}` : "0 registros";
  $("prev-page").disabled = page <= 1;
  $("next-page").disabled = page >= pages;
  $("package-rows").innerHTML = pageRows.length ? pageRows.map((record) => {
    const status = record.substatus || record.packageStatus;
    const pillStatus = record.promiseStatus || status;
    const routePackageCount = packagesByRoute.get(record.route || "Sem ID")?.size || 0;
    return `<tr><td><span class="route-id">${safeText(record.route || "—")}</span><small class="route-package-count">${fmt.format(routePackageCount)} ${pendingOnly ? "pendentes" : "pacotes"}</small></td><td><span class="package-id">${safeText(record.package || "—")}</span></td><td class="city-cell" title="${safeText(record.city)}">${safeText(record.city || "—")}</td><td>${safeText(label(record.addressType) || "—")}</td><td class="carrier-cell" title="${safeText(formatCarrierName(record.carrier))}">${safeText(formatCarrierName(record.carrier) || "—")}</td><td><span class="status-pill ${statusClass(status)}">${safeText(label(status))}</span></td><td><span class="status-pill ${statusClass(pillStatus)}">${safeText(label(pillStatus))}</span></td></tr>`;
  }).join("") : '<tr><td colspan="7" class="empty-state">Nenhum pacote corresponde aos filtros.</td></tr>';
}

function displayData(nextRecords) {
  records = nextRecords;
  deliveredPackageIds = new Set(records.filter(isDelivered).map((record) => record.package).filter(Boolean));
  page = 1;
  populateFilters();
  renderTable();
  $("updated-at").textContent = `Planilha carregada às ${new Intl.DateTimeFormat("pt-BR", { hour: "2-digit", minute: "2-digit" }).format(new Date())}`;
  setNotice("");
}

function loadWorkbook(buffer) {
  if (!window.XLSX) throw new Error("Não foi possível carregar o leitor de planilhas. Verifique sua conexão e recarregue a página.");
  const workbook = window.XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  displayData(parseRows(rows));
}

async function loadDefaultWorkbook() {
  try {
    const response = await fetch(AUTOLOAD_PATH);
    if (!response.ok) throw new Error("A planilha padrão não está disponível nesta pasta.");
    loadWorkbook(await response.arrayBuffer());
  } catch (error) {
    $("package-rows").innerHTML = '<tr><td colspan="7" class="empty-state">Importe a planilha operacional para iniciar o monitor.</td></tr>';
    $("updated-at").textContent = "Aguardando planilha";
    setNotice(`${error.message} Use “Importar planilha” para selecionar um arquivo .xlsx ou .csv.`);
  }
}

$("file-input").addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    loadWorkbook(await file.arrayBuffer());
  } catch (error) {
    setNotice(`Não foi possível ler a planilha: ${error.message}`);
  }
  event.target.value = "";
});
$("refresh-button").addEventListener("click", loadDefaultWorkbook);
$("search-input").addEventListener("input", () => { page = 1; renderTable(); });
$("pending-filter").addEventListener("change", () => { page = 1; renderTable(); });
$("route-options").addEventListener("change", (event) => {
  const checkbox = event.target;
  if (checkbox.type !== "checkbox") return;
  if (checkbox.checked) selectedRouteStatuses.add(checkbox.value);
  else selectedRouteStatuses.delete(checkbox.value);
  updateRouteSummary();
  page = 1;
  renderTable();
});
$("route-clear").addEventListener("click", () => {
  selectedRouteStatuses.clear();
  populateRouteFilter();
  page = 1;
  renderTable();
});
$("route-select-all").addEventListener("change", (event) => {
  selectedRouteStatuses.clear();
  if (event.target.checked) records.forEach((record) => {
    if (record.routeStatus) selectedRouteStatuses.add(record.routeStatus);
  });
  populateRouteFilter();
  page = 1;
  renderTable();
});
$("substatus-options").addEventListener("change", (event) => {
  const checkbox = event.target;
  if (checkbox.type !== "checkbox") return;
  if (checkbox.checked) selectedSubstatuses.add(checkbox.value);
  else selectedSubstatuses.delete(checkbox.value);
  updateSubstatusSummary();
  page = 1;
  renderTable();
});
$("substatus-clear").addEventListener("click", () => {
  selectedSubstatuses.clear();
  populateSubstatusFilter();
  page = 1;
  renderTable();
});
$("substatus-select-all").addEventListener("change", (event) => {
  selectedSubstatuses.clear();
  if (event.target.checked) records.forEach((record) => selectedSubstatuses.add(record.substatus));
  populateSubstatusFilter();
  page = 1;
  renderTable();
});
$("carrier-options").addEventListener("change", (event) => {
  const checkbox = event.target;
  if (checkbox.type !== "checkbox") return;
  if (checkbox.checked) selectedCarriers.add(checkbox.value);
  else selectedCarriers.delete(checkbox.value);
  updateCarrierSummary();
  page = 1;
  renderTable();
});
$("carrier-clear").addEventListener("click", () => {
  selectedCarriers.clear();
  populateCarrierFilter();
  page = 1;
  renderTable();
});
$("carrier-select-all").addEventListener("change", (event) => {
  selectedCarriers.clear();
  if (event.target.checked) records.forEach((record) => {
    if (record.carrier) selectedCarriers.add(record.carrier);
  });
  populateCarrierFilter();
  page = 1;
  renderTable();
});
["promise-filter", "address-filter", "city-filter"].forEach((id) => $(id).addEventListener("change", () => { page = 1; renderTable(); }));
$("prev-page").addEventListener("click", () => { page -= 1; renderTable(); });
$("next-page").addEventListener("click", () => { page += 1; renderTable(); });
loadDefaultWorkbook();