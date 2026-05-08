const readyText = document.getElementById("readyText");
const message = document.getElementById("message");
const refreshButton = document.getElementById("refreshButton");
const exportButton = document.getElementById("exportButton");
const captureButton = document.getElementById("captureButton");
const cancelButton = document.getElementById("cancelButton");

const rows = {
  native: document.getElementById("nativeStatus"),
  mo2: document.getElementById("mo2Status"),
  downloads: document.getElementById("downloadsStatus"),
  cookies: document.getElementById("cookiesStatus"),
};

function setRow(row, state, text) {
  row.classList.remove("ok", "warn", "bad");
  row.classList.add(state);
  row.querySelector("small").textContent = text;
}

function setMessage(text) {
  message.textContent = text || "";
}

function pendingSummary(response) {
  if (response.pendingDownload && response.pendingDownload.download) {
    const name = response.pendingDownload.download.name || "LoversLab file";
    return `Waiting for browser download: ${name}`;
  }

  if (response.externalCapture) {
    let label = "this page";
    try {
      label = new URL(response.externalCapture.pageUrl).hostname;
    } catch (_error) {
      label = response.externalCapture.pageUrl || label;
    }
    const count = response.externalCapture.capturedCount || 0;
    const notice = response.externalCapture.captureNotice ? ` ${response.externalCapture.captureNotice}` : "";
    return `Capture armed for archives from ${label} for 5 minutes. Captured: ${count}.${notice}`;
  }

  return "";
}

function summarizeReady(status) {
  return Boolean(
    status &&
    status.ok &&
    status.mo2 &&
    status.mo2.exists &&
    status.mo2.llPluginInstalled &&
    status.downloads &&
    status.downloads.exists &&
    status.cookies &&
    status.cookies.exists
  );
}

async function checkStatus() {
  setMessage("");
  readyText.textContent = "Checking status...";

  const response = await browser.runtime.sendMessage({ action: "popup_get_status" });
  if (!response || !response.ok) {
    setRow(rows.native, "bad", "Native app is not installed or not reachable.");
    setRow(rows.mo2, "bad", "Unavailable until native bridge works.");
    setRow(rows.downloads, "bad", "Unavailable until native bridge works.");
    setRow(rows.cookies, "bad", "Unavailable until native bridge works.");
    readyText.textContent = "Not ready";
    setMessage(response && response.error ? response.error : "Native messaging failed.");
    return;
  }

  setRow(rows.native, "ok", response.nativeApp.baseDir);

  if (response.mo2.exists && response.mo2.llPluginInstalled) {
    setRow(rows.mo2, "ok", response.mo2.llPluginPath);
  } else if (response.mo2.exists) {
    setRow(rows.mo2, "warn", "MO2 found, but LL plugin is not installed.");
  } else {
    setRow(rows.mo2, "bad", "ModOrganizer.exe was not found.");
  }

  setRow(
    rows.downloads,
    response.downloads.exists ? "ok" : "bad",
    response.downloads.path
  );

  setRow(
    rows.cookies,
    response.cookies.exists ? "ok" : "warn",
    response.cookies.exists ? response.cookies.path : "Click Export Cookies after logging into LoversLab."
  );

  readyText.textContent = summarizeReady(response) ? "Ready" : "Setup incomplete";
  const pending = pendingSummary(response);
  cancelButton.disabled = !pending;
  setMessage(pending);
}

async function exportCookies() {
  setMessage("Exporting LoversLab cookies...");
  const response = await browser.runtime.sendMessage({ action: "popup_export_cookies" });
  if (!response || !response.ok) {
    setMessage(response && response.error ? response.error : "Cookie export failed.");
    await checkStatus();
    return;
  }

  setMessage(`Cookies exported to ${response.savedTo}`);
  await checkStatus();
}

async function captureCurrentPage() {
  setMessage("Arming capture for current page...");

  let tab;
  try {
    const tabs = await browser.tabs.query({ active: true, currentWindow: true });
    tab = tabs && tabs[0];
  } catch (error) {
    setMessage(`Could not read current tab: ${error}`);
    return;
  }

  if (!tab || !tab.url || !/^https?:\/\//i.test(tab.url)) {
    setMessage("Open a web page before arming capture.");
    return;
  }

  const response = await browser.runtime.sendMessage({
    action: "popup_capture_current_page",
    pageUrl: tab.url,
    pageTitle: tab.title || "",
    tabId: tab.id
  });

  if (!response || !response.ok) {
    setMessage(response && response.error ? response.error : "Could not arm capture.");
    return;
  }

  await checkStatus();
}

async function cancelPending() {
  const response = await browser.runtime.sendMessage({ action: "popup_cancel_pending" });
  if (!response || !response.ok) {
    setMessage(response && response.error ? response.error : "Could not cancel pending download.");
    return;
  }

  setMessage(response.cancelled ? "Pending download cancelled." : "No pending download to cancel.");
  await checkStatus();
}

refreshButton.addEventListener("click", checkStatus);
exportButton.addEventListener("click", exportCookies);
captureButton.addEventListener("click", captureCurrentPage);
cancelButton.addEventListener("click", cancelPending);

checkStatus();
