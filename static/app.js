const form = document.querySelector("#reportForm");
const generateButton = document.querySelector("#generateButton");
const runState = document.querySelector("#runState");
const defaultList = document.querySelector("#defaultList");
const reportType = document.querySelector("#reportType");
const serviceFile = document.querySelector("#serviceFile");
const axioFile = document.querySelector("#axioFile");
const retailFile = document.querySelector("#retailFile");
const masterFile = document.querySelector("#masterFile");
const serviceFileName = document.querySelector("#serviceFileName");
const axioFileName = document.querySelector("#axioFileName");
const retailFileName = document.querySelector("#retailFileName");
const masterFileLabel = document.querySelector("#masterFileLabel");
const masterFileName = document.querySelector("#masterFileName");
const summaryGrid = document.querySelector("#summaryGrid");
const notice = document.querySelector("#notice");
const tableWrap = document.querySelector("#tableWrap");
const previewHead = document.querySelector("#previewHead");
const previewBody = document.querySelector("#previewBody");
const downloadFinal = document.querySelector("#downloadFinal");
const downloadZonal = document.querySelector("#downloadZonal");
const downloadChannel = document.querySelector("#downloadChannel");
const reportTabs = document.querySelectorAll(".report-tab");

const numberFormatter = new Intl.NumberFormat("en-IN");
let currentReports = {};
let activeReport = "service";
let currentStatus = null;

function setRunState(label) {
  runState.textContent = label;
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
  const meta = status.exists
    ? `${status.name} · ${formatBytes(status.size)} · ${status.modified}`
    : `${status.name} not found`;

  return `
    <div class="file-status">
      <span class="status-dot ${dotClass}"></span>
      <div>
        <strong>${escapeHtml(label)}</strong>
        <small>${escapeHtml(meta)}</small>
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
  const directHref = enabled && String(downloadValue || "").startsWith("/") ? downloadValue : "#";
  link.href = directHref;
  link.dataset.downloadValue = enabled ? String(downloadValue || "") : "";
  link.classList.toggle("disabled", !enabled);
  link.setAttribute("aria-disabled", enabled ? "false" : "true");
}

function setDownloads(downloads = {}, reportKey = activeReport) {
  const isChannel = reportKey === "channel";
  downloadFinal.hidden = isChannel;
  downloadZonal.hidden = isChannel;
  downloadChannel.hidden = !isChannel;

  setDownloadEnabled(downloadFinal, downloads.final_report, Boolean(downloads.final_report));
  setDownloadEnabled(downloadZonal, downloads.zonal_report, Boolean(downloads.zonal_report));
  setDownloadEnabled(downloadChannel, downloads.channel_report, Boolean(downloads.channel_report));
}

function renderDefaultStatus() {
  if (!currentStatus) return;

  const defaults = currentStatus.defaults;
  const rows = activeReport === "channel"
    ? [
        renderFileStatus("AXIO report", defaults.axio),
        renderFileStatus("Retail report", defaults.retail),
        renderFileStatus("Channel master workbook", defaults.channel_master),
      ]
    : [
        renderFileStatus("Daily service report", defaults.service),
        renderFileStatus("Service master workbook", defaults.service_master),
      ];

  defaultList.innerHTML = rows.join("");
}

function setVisibleInputs(reportKey) {
  document.querySelectorAll("[data-report-input]").forEach((element) => {
    element.hidden = element.dataset.reportInput !== reportKey;
  });

  reportType.value = reportKey;
  masterFile.value = "";
  masterFileLabel.textContent = reportKey === "channel"
    ? "Channel master workbook"
    : "Service master workbook";
  masterFileName.textContent = reportKey === "channel"
    ? "Use local channel master or choose XLSB/XLSX"
    : "Use local service master or choose XLSX/XLSB";
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
    const contentType = response.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
      const data = await response.json();
      throw new Error(data.error || "Download failed.");
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

async function loadStatus() {
  const response = await fetch("/api/status", { cache: "no-store" });
  currentStatus = await response.json();
  renderDefaultStatus();

  const downloads = {};
  if (currentStatus.outputs.final_report.exists) {
    downloads.final_report = "/download/final_report.xlsx";
  }
  if (currentStatus.outputs.zonal_report.exists) {
    downloads.zonal_report = "/download/zonal_report.xlsx";
  }
  if (currentStatus.outputs.channel_report.exists) {
    downloads.channel_report = "/download/final_channel_report.xlsx";
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

masterFile.addEventListener("change", () => {
  if (masterFile.files[0]) {
    masterFileName.textContent = masterFile.files[0].name;
    return;
  }
  masterFileName.textContent = activeReport === "channel"
    ? "Use local channel master or choose XLSB/XLSX"
    : "Use local service master or choose XLSX/XLSB";
});

reportTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    switchReport(tab.dataset.report);
  });
});

[downloadFinal, downloadZonal, downloadChannel].forEach((link) => {
  link.addEventListener("click", async (event) => {
    const downloadValue = link.dataset.downloadValue || "";
    const isDisabled = link.getAttribute("aria-disabled") === "true";
    if (isDisabled || downloadValue === "") {
      event.preventDefault();
      return;
    }

    if (downloadValue.startsWith("/")) {
      return;
    }

    event.preventDefault();
    try {
      await triggerGeneratedDownload(downloadValue);
      showNotice(`Downloaded ${downloadValue.replaceAll("_", " ")}.`);
      setRunState("Complete");
    } catch (error) {
      showNotice(error.message, true);
      setRunState("Error");
    }
  });
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  generateButton.disabled = true;
  setRunState("Running");
  showNotice("Generating report...");

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      body: new FormData(form),
    });

    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Report generation failed.");
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

renderSummary({}, activeReport);
loadStatus().catch(() => {
  showNotice("Could not read local file status.", true);
});
