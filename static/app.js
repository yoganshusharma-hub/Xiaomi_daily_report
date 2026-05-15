const authShell = document.querySelector("#authShell");
const appShell = document.querySelector("#appShell");
const loginForm = document.querySelector("#loginForm");
const loginButton = document.querySelector("#loginButton");
const loginEmail = document.querySelector("#loginEmail");
const loginPassword = document.querySelector("#loginPassword");
const authMessage = document.querySelector("#authMessage");
const logoutButton = document.querySelector("#logoutButton");
const form = document.querySelector("#reportForm");
const generateButton = document.querySelector("#generateButton");
const runState = document.querySelector("#runState");
const defaultList = document.querySelector("#defaultList");
const reportType = document.querySelector("#reportType");
const serviceFile = document.querySelector("#serviceFile");
const axioFile = document.querySelector("#axioFile");
const retailFile = document.querySelector("#retailFile");
const serviceFileName = document.querySelector("#serviceFileName");
const axioFileName = document.querySelector("#axioFileName");
const retailFileName = document.querySelector("#retailFileName");
const summaryGrid = document.querySelector("#summaryGrid");
const notice = document.querySelector("#notice");
const tableWrap = document.querySelector("#tableWrap");
const previewHead = document.querySelector("#previewHead");
const previewBody = document.querySelector("#previewBody");
const downloadReport = document.querySelector("#downloadReport");
const reportTabs = document.querySelectorAll(".report-tab");

const numberFormatter = new Intl.NumberFormat("en-IN");
let currentReports = {};
let activeReport = document.querySelector(".report-tab.active")?.dataset.report || "service";
let currentStatus = null;
let currentUser = null;

function setRunState(label) {
  runState.textContent = label;
}

function setAuthMessage(message = "") {
  authMessage.textContent = message;
  authMessage.hidden = message === "";
}

function showLogin(message = "") {
  currentUser = null;
  authShell.hidden = false;
  appShell.hidden = true;
  setAuthMessage(message);
  loginPassword.value = "";
}

function showApp(user) {
  currentUser = user;
  authShell.hidden = true;
  appShell.hidden = false;
  setAuthMessage("");
}

function formatBytes(bytes) {
  if (!bytes) return "0 KB";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** index).toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function renderFileStatus(label, status) {
  const dotClass = status.exists ? "ready" : "missing";

  return `
    <div class="file-status">
      <span class="status-dot ${dotClass}"></span>
      <div>
        <strong>${escapeHtml(label)}</strong>
      </div>
    </div>
  `;
}

function renderSummary(summary = {}, reportKey = activeReport) {
  const metrics = reportKey === "channel"
    ? [
        ["Total Units", summary.total_units ?? 0, "accent"],
        ["Total GWP", summary.total_gwp ?? 0, "accent"],
        ["AXIO Units", summary.axio_units ?? 0, ""],
        ["Retail Units", summary.retail_units ?? 0, ""],
        ["States", summary.states ?? 0, ""],
        ["Stores", summary.stores ?? 0, ""],
      ]
    : [
        ["Paid Units", summary.paid_rows ?? 0, "accent"],
        ["Total GWP", summary.total_gwp ?? 0, "accent"],
        ["Regions", summary.regions ?? summary.zones ?? 0, ""],
        ["Unmatched ASC", summary.unmatched_rows ?? 0, ""],
        ["Input Rows", summary.input_rows ?? 0, ""],
        ["Service Centers", summary.service_centers ?? 0, ""],
      ];

  summaryGrid.innerHTML = metrics
    .map(([label, value, className]) => `
      <div class="metric-tile ${className}">
        <small>${escapeHtml(label)}</small>
        <strong>${numberFormatter.format(value)}</strong>
      </div>
    `)
    .join("");
}

function renderPreview(columns, rows) {
  previewHead.innerHTML = `
    <tr>
      ${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}
    </tr>
  `;

  previewBody.innerHTML = rows
    .map((row) => {
      const isGrand = row.Region === "Grand Total" || row.Zone === "Grand Total" || row.State === "Grand Total";
      const isTotal = isGrand
        || String(row.Region || "").endsWith("Total")
        || String(row.Zone || "").endsWith("Total")
        || String(row.State || "").endsWith("Total")
        || String(row.DistributorName || "").endsWith("Total");
      const rowClass = isGrand ? "grand-row" : isTotal ? "total-row" : "";
      return `
        <tr class="${rowClass}">
          ${columns
            .map((column) => {
              const value = row[column] ?? "";
              const isNumeric = column.includes("Unit") || column.includes("GWP");
              const text = isNumeric && value !== "" ? numberFormatter.format(value) : value;
              return `<td class="${isNumeric ? "numeric" : ""}">${escapeHtml(text)}</td>`;
            })
            .join("")}
        </tr>
      `;
    })
    .join("");

  tableWrap.hidden = false;
}

function setDownloadEnabled(link, downloadValue, enabled) {
  link.href = "#";
  link.dataset.downloadValue = enabled ? String(downloadValue || "") : "";
  link.classList.toggle("disabled", !enabled);
  link.setAttribute("aria-disabled", enabled ? "false" : "true");
}

function setDownloads(downloads = {}, reportKey = activeReport) {
  const downloadValue = reportKey === "channel"
    ? downloads.channel_report
    : downloads.final_report;
  setDownloadEnabled(downloadReport, downloadValue, Boolean(downloadValue));
}

function renderDefaultStatus() {
  if (!currentStatus) return;

  const defaults = currentStatus.defaults;
  const rows = activeReport === "channel"
    ? [
        renderFileStatus("Upload Axio Report", defaults.axio),
        renderFileStatus("Upload Retail Report", defaults.retail),
      ]
    : [
        renderFileStatus("Upload Service Report", defaults.service),
      ];

  defaultList.innerHTML = rows.join("");
}

function setVisibleInputs(reportKey) {
  document.querySelectorAll("[data-report-input]").forEach((element) => {
    const isVisible = element.dataset.reportInput === reportKey;
    element.hidden = !isVisible;
    element.setAttribute("aria-hidden", isVisible ? "false" : "true");
  });

  reportType.value = reportKey;
}

function switchReport(reportKey, options = {}) {
  activeReport = reportKey;
  setVisibleInputs(reportKey);
  renderDefaultStatus();

  reportTabs.forEach((tab) => {
    const isActive = tab.dataset.report === reportKey;
    tab.classList.toggle("active", isActive);
    tab.setAttribute("aria-selected", isActive ? "true" : "false");
  });

  const report = currentReports[reportKey];
  if (!report) {
    renderSummary({}, reportKey);
    if (!options.keepPreview) {
      tableWrap.hidden = true;
    }
    setDownloads({}, reportKey);
    return;
  }

  renderSummary(report.summary, reportKey);
  renderPreview(report.columns, report.preview);
  setDownloads(report.downloads, reportKey);
}

function showNotice(message, isError = false) {
  notice.textContent = message;
  notice.classList.toggle("error", isError);
}

function getDownloadFilename(response, fallbackName) {
  const contentDisposition = response.headers.get("content-disposition") || "";
  const match = /filename="?([^"]+)"?/i.exec(contentDisposition);
  return match?.[1] || fallbackName;
}

async function triggerGeneratedDownload(downloadKey) {
  const fallbackNames = {
    final_report: "final_report.xlsx",
    zonal_report: "zonal_report.xlsx",
    channel_report: "final_channel_report.xlsx",
  };

  setRunState("Preparing");
  showNotice("Preparing download...");

  const formData = new FormData(form);
  formData.set("report_type", activeReport);
  formData.set("download_key", downloadKey);

  const response = await fetch("/api/download", {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    if (response.status === 401) {
      showLogin("Session expired. Please sign in again.");
      throw new Error("Please sign in again.");
    }
    const contentType = response.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
      const data = await response.json();
      throw new Error(data.error || data.detail || "Download failed.");
    }
    throw new Error(await response.text() || "Download failed.");
  }

  const blob = await response.blob();
  const filename = getDownloadFilename(response, fallbackNames[downloadKey] || "report.xlsx");
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(objectUrl);
}

async function getSession() {
  const response = await fetch("/api/auth/session", { cache: "no-store" });
  if (response.status === 401) {
    return null;
  }
  if (!response.ok) {
    throw new Error("Could not verify your session.");
  }
  return response.json();
}

async function loadStatus() {
  const response = await fetch("/api/status", { cache: "no-store" });
  if (response.status === 401) {
    showLogin("Session expired. Please sign in again.");
    throw new Error("Please sign in again.");
  }
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || data.detail || "Could not read local file status.");
  }
  currentStatus = await response.json();
  renderDefaultStatus();

  const downloads = {};
  if (currentStatus.outputs.final_report.exists) {
    downloads.final_report = "final_report";
  }
  if (currentStatus.outputs.channel_report.exists) {
    downloads.channel_report = "channel_report";
  }
  setDownloads(downloads, activeReport);
}

serviceFile.addEventListener("change", () => {
  serviceFileName.textContent = serviceFile.files[0]?.name || "Use local default or choose a CSV";
});

axioFile.addEventListener("change", () => {
  axioFileName.textContent = axioFile.files[0]?.name || "Use local default or choose a CSV";
});

retailFile.addEventListener("change", () => {
  retailFileName.textContent = retailFile.files[0]?.name || "Use local default or choose a CSV";
});

reportTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    switchReport(tab.dataset.report);
  });
});

downloadReport.addEventListener("click", async (event) => {
  const downloadValue = downloadReport.dataset.downloadValue || "";
  const isDisabled = downloadReport.getAttribute("aria-disabled") === "true";
  if (isDisabled || downloadValue === "") {
    event.preventDefault();
    return;
  }

  event.preventDefault();
  try {
    await triggerGeneratedDownload(downloadValue);
    showNotice("Downloaded report.");
    setRunState("Complete");
  } catch (error) {
    showNotice(error.message, true);
    setRunState("Error");
  }
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const email = loginEmail.value.trim().toLowerCase();
  const password = loginPassword.value;
  if (!email.endsWith("@zopper.com")) {
    setAuthMessage("Use your @zopper.com email address.");
    return;
  }

  loginButton.disabled = true;
  setAuthMessage("");

  try {
    const response = await fetch("/api/auth/login", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ email, password }),
    });

    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.detail || data.error || "Sign in failed.");
    }

    showApp(data.user);
    setRunState("Ready");
    showNotice("Signed in.");
    switchReport(activeReport, { keepPreview: true });
    await loadStatus();
  } catch (error) {
    setAuthMessage(error.message || "Sign in failed.");
  } finally {
    loginButton.disabled = false;
  }
});

logoutButton.addEventListener("click", async () => {
  await fetch("/api/auth/logout", { method: "POST" });
  currentReports = {};
  currentStatus = null;
  activeReport = "service";
  switchReport("service");
  renderSummary({}, "service");
  showLogin("Signed out.");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  // Vercel 4.5MB limit check
  const files = [serviceFile, axioFile, retailFile];
  let totalSize = 0;
  for (const f of files) {
    if (f.files[0]) totalSize += f.files[0].size;
  }

  if (totalSize > 4.4 * 1024 * 1024) {
    showNotice("Total upload size exceeds 4.5MB (Vercel limit). Please use smaller files.", true);
    return;
  }

  generateButton.disabled = true;
  setRunState("Running");
  showNotice("Generating report...");

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      body: new FormData(form),
    });

    if (response.status === 401) {
      showLogin("Session expired. Please sign in again.");
      throw new Error("Please sign in again.");
    }

    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || data.detail || "Report generation failed.");
    }

    currentReports = { ...currentReports, ...(data.reports || {}) };
    switchReport(data.active_report || "service");
    showNotice(`Generated ${data.summary.total_units} units at ${data.summary.generated_at}.`);
    setRunState("Complete");
  } catch (error) {
    showNotice(error.message, true);
    setRunState("Error");
  } finally {
    generateButton.disabled = false;
  }
});

async function initialiseApp() {
  switchReport(activeReport, { keepPreview: true });
  renderSummary({}, activeReport);

  try {
    const session = await getSession();
    if (!session?.user) {
      showLogin();
      return;
    }

    showApp(session.user);
    await loadStatus();
  } catch (error) {
    showLogin(error.message || "Please sign in.");
  }
}

initialiseApp();
