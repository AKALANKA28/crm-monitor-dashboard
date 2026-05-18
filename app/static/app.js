const state = {
  page: 1,
  pageSize: 25,
  total: 0,
  columns: [],
};

const statusConfig = {
  Success: { countId: "successCount", className: "success" },
  Failed: { countId: "failedCount", className: "failed" },
  Pending: { countId: "pendingCount", className: "pending" },
  "In Progress": { countId: "inProgressCount", className: "progress" },
};

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

  if (dateFrom) params.set("date_from", dateFrom);
  if (dateTo) params.set("date_to", dateTo);
  if (status) params.set("status", status);
  if (q) params.set("q", q);
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
  const normalized = String(status || "Unknown");
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

  const total = Number(data.total || 0);
  const page = Number(data.page || 1);
  const pageSize = Number(data.page_size || 25);

  document.getElementById("tableMeta").textContent =
    `${formatNumber(total)} records`;
  document.getElementById("pageIndicatorNum").textContent = page;

  document.getElementById("prevBtn").disabled = page <= 1;
  document.getElementById("nextBtn").disabled = page * pageSize >= total;

  if (!rows.length) {
    thead.innerHTML = "";
    tbody.innerHTML = `
            <tr>
                <td class="empty-state" style="text-align: center; padding: 40px; color: #8ca3ba;">
                    No requests matched the selected filters.
                </td>
            </tr>
        `;
    return;
  }

  const visibleColumns = ["Actions", ...columns];

  thead.innerHTML = `
        <tr>
            ${visibleColumns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}
        </tr>
    `;

  tbody.innerHTML = rows
    .map((row) => {
      const requestId = row.RequestId || row.requestId || row.request_id || "";
      const status = row.CRMStatus || row.crmStatus || row.crm_status || "";

      const retryButton = shouldAllowRetry(status)
        ? `<button class="btn btn-danger retry-btn" data-request-id="${escapeHtml(requestId)}">Retry</button>`
        : `<span class="text-disabled">No action</span>`;

      const cells = columns
        .map((column) => {
          const value = row[column];
          if (column === "CRMStatus") {
            return `<td>${badgeForStatus(value)}</td>`;
          }
          return `<td title="${escapeHtml(value)}">${escapeHtml(value)}</td>`;
        })
        .join("");

      return `<tr><td>${retryButton}</td>${cells}</tr>`;
    })
    .join("");

  bindRetryButtons();
}

function shouldAllowRetry(status) {
  return ["Failed", "In Progress"].includes(String(status || ""));
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
  } catch (error) {
    setMessage(error.message, "error");
    document.getElementById("tableMeta").textContent = "Error loading";
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

setDefaultDates();
setupEvents();
loadDashboard(1);
