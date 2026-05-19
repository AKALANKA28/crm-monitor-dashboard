const state = {
  page: 1,
  pageSize: 25,
  total: 0,
  columns: [],
  isLoading: false,
  lastColumnQuery: "",
};

const AUTO_REFRESH_MS = 10000;

const statusConfig = {
  Success: { countId: "successCount", className: "success" },
  Failed: { countId: "failedCount", className: "failed" },
  Pending: { countId: "pendingCount", className: "pending" },
  "In Progress": { countId: "inProgressCount", className: "progress" },
};

function normalizeColumnKey(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function findRequestIdColumn(columns) {
  const normalizedColumns = columns.map((column) => ({
    name: column,
    key: normalizeColumnKey(column),
  }));
  const exact = normalizedColumns.find((column) => column.key === "requestid");
  if (exact) return exact.name;

  const requestId = normalizedColumns.find(
    (column) => column.key.includes("request") && column.key.endsWith("id"),
  );
  if (requestId) return requestId.name;

  const id = normalizedColumns.find((column) => column.key === "id");
  return id ? id.name : null;
}

function findStatusColumn(columns, rows = []) {
  const normalizedColumns = columns.map((column) => ({
    name: column,
    key: normalizeColumnKey(column),
  }));
  const exactTargets = ["crmstatus", "status"];
  const exact = normalizedColumns.find((column) =>
    exactTargets.includes(column.key),
  );
  if (exact) return exact.name;

  const crmStatus = normalizedColumns.find(
    (column) => column.key.includes("crm") && column.key.includes("status"),
  );
  if (crmStatus) return crmStatus.name;

  const retryableStatusValues = new Set(["Failed", "In Progress"]);
  const statusLikeColumns = normalizedColumns.filter((column) =>
    column.key.includes("status"),
  );
  const byStatusValues = statusLikeColumns.find(({ name }) =>
    rows.some((row) => retryableStatusValues.has(canonicalStatus(row[name]))),
  );
  if (byStatusValues) return byStatusValues.name;

  const statusSuffix = statusLikeColumns.find(
    (column) => column.key.endsWith("status") && column.key !== "kycstatus",
  );
  if (statusSuffix) return statusSuffix.name;

  const byRetryableValues = normalizedColumns.find(({ name }) =>
    rows.some((row) => retryableStatusValues.has(canonicalStatus(row[name]))),
  );
  return byRetryableValues ? byRetryableValues.name : null;
}

function getRowValue(row, column, fallbackKeys = []) {
  if (column && hasOwn(row, column)) return row[column];

  for (const key of fallbackKeys) {
    if (hasOwn(row, key)) return row[key];
  }

  return "";
}

function hasOwn(object, key) {
  return Object.prototype.hasOwnProperty.call(object, key);
}

function findColumnMatch(query, columns) {
  const normalizedQuery = normalizeColumnKey(query);
  if (!normalizedQuery) return null;

  const normalizedColumns = columns.map((column) => ({
    name: column,
    key: normalizeColumnKey(column),
  }));

  const exact = normalizedColumns.find((col) => col.key === normalizedQuery);
  if (exact) return exact.name;

  const startsWith = normalizedColumns.find((col) =>
    col.key.startsWith(normalizedQuery),
  );
  if (startsWith) return startsWith.name;

  const includes = normalizedColumns.find((col) =>
    col.key.includes(normalizedQuery),
  );
  return includes ? includes.name : null;
}

function scrollToColumnMatch(query, columns) {
  const normalizedQuery = normalizeColumnKey(query);
  if (!normalizedQuery) {
    state.lastColumnQuery = "";
    return;
  }
  if (normalizedQuery === state.lastColumnQuery) return;

  const match = findColumnMatch(query, columns);
  state.lastColumnQuery = normalizedQuery;
  if (!match) return;

  const tableScroll = document.querySelector(".table-scroll");
  if (!tableScroll) return;

  const headerCell = tableScroll.querySelector(
    `thead th[data-col-key="${normalizeColumnKey(match)}"]`,
  );
  if (!headerCell) return;

  const stickyWidth = Array.from(
    tableScroll.querySelectorAll("thead th.sticky-col"),
  ).reduce((sum, cell) => sum + cell.offsetWidth, 0);

  const targetLeft = headerCell.offsetLeft - stickyWidth;
  tableScroll.scrollTo({ left: Math.max(0, targetLeft), behavior: "smooth" });
}

function formatNumber(value) {
  return new Intl.NumberFormat().format(Number(value || 0));
}

function todayIso() {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  return new Date(now.getTime() - offset * 60000).toISOString().slice(0, 10);
}

function setDefaultDates() {
  const today = todayIso();
  document.getElementById("dateFrom").value = today;
  document.getElementById("dateTo").value = today;
}

function getFilters(includePage = true) {
  const params = new URLSearchParams();
  const dateFrom = document.getElementById("dateFrom").value;
  const dateTo = document.getElementById("dateTo").value;
  const status = document.getElementById("status").value;
  const q = document.getElementById("searchText").value.trim();
  const pageSize = document.getElementById("pageSize").value;
  const columnMatch = findColumnMatch(q, state.columns);

  if (dateFrom) params.set("date_from", dateFrom);
  if (dateTo) params.set("date_to", dateTo);
  if (status) params.set("status", status);
  if (q && !columnMatch) params.set("q", q);
  params.set("page_size", pageSize);

  if (includePage) params.set("page", state.page);
  return params;
}

function setMessage(text, type = "success") {
  const message = document.getElementById("message");
  message.textContent = text;
  message.className = `message ${type}`;

  clearTimeout(window.messageTimer);
  window.messageTimer = setTimeout(() => {
    message.className = "message hidden";
    message.textContent = "";
  }, 3500);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function badgeForStatus(status) {
  const normalized = canonicalStatus(status);
  const css = statusConfig[normalized]?.className || "default";
  return `<span class="badge badge-${css}">${escapeHtml(normalized)}</span>`;
}

function updateSummary(data) {
  const counts = data.status_counts || {};
  const total = Number(data.total || 0);
  state.total = total;

  document.getElementById("totalCount").textContent = formatNumber(total);

  Object.entries(statusConfig).forEach(([status, config]) => {
    document.getElementById(config.countId).textContent = formatNumber(
      counts[status] || 0,
    );
  });
}

function renderTable(data) {
  const table = document.getElementById("requestsTable");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  const rows = data.rows || [];
  const columns = data.columns || [];
  state.columns = columns;
  const requestIdColumn = findRequestIdColumn(columns);
  const statusColumn = findStatusColumn(columns, rows);

  const total = Number(data.total || 0);
  const page = Number(data.page || 1);
  const pageSize = Number(data.page_size || 25);

  document.getElementById("tableMeta").textContent =
    `${formatNumber(total)} records`;
  document.getElementById("pageIndicatorNum").textContent = page;

  document.getElementById("prevBtn").disabled = page <= 1;
  document.getElementById("nextBtn").disabled = page * pageSize >= total;

  const visibleColumns = ["Actions", ...columns];

  thead.innerHTML = `
        <tr>
            ${visibleColumns
              .map((column) => {
                const className =
                  column === "Actions"
                    ? "sticky-col sticky-actions"
                    : column === requestIdColumn
                      ? "sticky-col sticky-request"
                      : "";
                const classAttr = className ? ` class="${className}"` : "";
                const keyAttr = ` data-col-key="${normalizeColumnKey(column)}"`;
                return `<th${classAttr}${keyAttr}>${escapeHtml(column)}</th>`;
              })
              .join("")}
        </tr>
    `;

  if (!rows.length) {
    tbody.innerHTML = `
            <tr>
                <td class="empty-state" colspan="${visibleColumns.length}" style="text-align: center; padding: 40px; color: #8ca3ba;">
                    No requests matched the selected filters.
                </td>
            </tr>
        `;
    return;
  }

  tbody.innerHTML = rows
    .map((row) => {
      const requestId = getRowValue(row, requestIdColumn, [
        "RequestId",
        "requestId",
        "request_id",
        "Id",
        "id",
      ]);
      const status = getRowValue(row, statusColumn, [
        "CRMStatus",
        "crmStatus",
        "crm_status",
        "Status",
        "status",
      ]);

      const retryButton = shouldAllowRetry(status)
        ? `<button class="btn btn-danger retry-btn" data-request-id="${escapeHtml(requestId)}">Retry</button>`
        : `<span class="text-disabled">No action</span>`;

      const actionCell = `<td class="sticky-col sticky-actions">${retryButton}</td>`;
      const cells = columns
        .map((column) => {
          const value = row[column];
          const className =
            column === requestIdColumn ? "sticky-col sticky-request" : "";
          const classAttr = className ? ` class="${className}"` : "";
          if (column === statusColumn) {
            return `<td${classAttr}>${badgeForStatus(value)}</td>`;
          }
          return `<td${classAttr} title="${escapeHtml(value)}">${escapeHtml(value)}</td>`;
        })
        .join("");

      return `<tr>${actionCell}${cells}</tr>`;
    })
    .join("");

  bindRetryButtons();
}

function shouldAllowRetry(status) {
  return ["Failed", "In Progress"].includes(canonicalStatus(status));
}

function canonicalStatus(status) {
  const text = String(status || "").trim();
  if (!text) return "Unknown";

  const normalized = text.toLowerCase().replace(/[\s_-]+/g, "");
  const statusMap = {
    success: "Success",
    failed: "Failed",
    fail: "Failed",
    failure: "Failed",
    error: "Failed",
    pending: "Pending",
    inprogress: "In Progress",
    processing: "In Progress",
  };

  return statusMap[normalized] || text;
}

function bindRetryButtons() {
  document.querySelectorAll(".retry-btn").forEach((button) => {
    button.addEventListener("click", async () => {
      const requestId = button.dataset.requestId;
      if (!requestId) return;

      const confirmRetry = confirm(
        `Move request ${requestId} back to Pending?`,
      );
      if (!confirmRetry) return;

      button.disabled = true;
      button.textContent = "Retrying...";

      try {
        const response = await fetch(
          `/api/requests/${encodeURIComponent(requestId)}/retry`,
          { method: "POST" },
        );
        if (!response.ok) {
          const error = await response.json().catch(() => ({}));
          throw new Error(error.detail || "Retry request failed");
        }
        setMessage(`Request ${requestId} moved back to Pending.`, "success");
        await loadDashboard(state.page);
      } catch (error) {
        setMessage(error.message, "error");
        button.disabled = false;
        button.textContent = "Retry";
      }
    });
  });
}

async function loadDashboard(page = 1) {
  if (state.isLoading) return;
  state.isLoading = true;
  state.page = page;
  state.pageSize = Number(document.getElementById("pageSize").value || 25);

  const params = getFilters(true);
  document.getElementById("tableMeta").textContent = "Loading...";

  try {
    const response = await fetch(`/api/requests?${params.toString()}`);
    if (!response.ok) throw new Error("Unable to load request data");
    const data = await response.json();
    updateSummary(data);
    renderTable(data);
    scrollToColumnMatch(
      document.getElementById("searchText").value.trim(),
      data.columns || [],
    );
  } catch (error) {
    setMessage(error.message, "error");
    document.getElementById("tableMeta").textContent = "Error loading";
  } finally {
    state.isLoading = false;
  }
}

function downloadExcel() {
  const params = getFilters(false);
  window.location.href = `/api/export?${params.toString()}`;
}

function resetFilters() {
  document.getElementById("status").value = "";
  document.getElementById("searchText").value = "";
  document.getElementById("pageSize").value = "25";
  setDefaultDates();
  loadDashboard(1);
}

function setupEvents() {
  // Mobile Drawer Toggle Logic
  const mobileFilterBtn = document.getElementById("mobileFilterBtn");
  const advancedFilters = document.getElementById("advancedFilters");
  const filterOverlay = document.getElementById("filterOverlay");
  const closeFilterBtn = document.getElementById("closeFilterBtn");

  function openFilters() {
    if (!advancedFilters) return;
    advancedFilters.classList.add("show-mobile");
    if (filterOverlay) filterOverlay.classList.add("active");
  }

  function closeFilters() {
    if (!advancedFilters) return;
    advancedFilters.classList.remove("show-mobile");
    if (filterOverlay) filterOverlay.classList.remove("active");
  }

  if (mobileFilterBtn) mobileFilterBtn.addEventListener("click", openFilters);
  if (closeFilterBtn) closeFilterBtn.addEventListener("click", closeFilters);
  if (filterOverlay) filterOverlay.addEventListener("click", closeFilters);

  // Submit Logic
  document.getElementById("filterForm").addEventListener("submit", (event) => {
    event.preventDefault();
    loadDashboard(1);

    // Automatically close drawer on mobile after hitting Apply
    if (window.innerWidth <= 768) {
      closeFilters();
    }
  });

  document.getElementById("resetBtn").addEventListener("click", resetFilters);
  document.getElementById("exportBtn").addEventListener("click", downloadExcel);
  document
    .getElementById("pageSize")
    .addEventListener("change", () => loadDashboard(1));
  document
    .getElementById("prevBtn")
    .addEventListener("click", () =>
      loadDashboard(Math.max(1, state.page - 1)),
    );
  document
    .getElementById("nextBtn")
    .addEventListener("click", () => loadDashboard(state.page + 1));
}

function startAutoRefresh() {
  if (!AUTO_REFRESH_MS) return;
  window.setInterval(() => loadDashboard(state.page), AUTO_REFRESH_MS);
}

setDefaultDates();
setupEvents();
loadDashboard(1);
startAutoRefresh();
