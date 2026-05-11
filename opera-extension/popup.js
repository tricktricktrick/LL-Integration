const ext = globalThis.browser || globalThis.chrome;

const readyText = document.getElementById("readyText");
const message = document.getElementById("message");
const refreshButton = document.getElementById("refreshButton");
const exportButton = document.getElementById("exportButton");
const captureButton = document.getElementById("captureButton");
const cancelButton = document.getElementById("cancelButton");
const monitorButton = document.getElementById("monitorButton");

const rows = {
  native: document.getElementById("nativeStatus"),
  mo2: document.getElementById("mo2Status"),
  active: document.getElementById("activeStatus"),
  downloads: document.getElementById("downloadsStatus"),
  cookies: document.getElementById("cookiesStatus"),
};

function callbackPromise(call) {
  return new Promise((resolve, reject) => {
    call((result) => {
      const error = ext.runtime.lastError;
      if (error) {
        reject(new Error(error.message || String(error)));
        return;
      }
      resolve(result);
    });
  });
}

function sendMessage(payload) {
  return callbackPromise(done => ext.runtime.sendMessage(payload, done));
}

function queryTabs(query) {
  return callbackPromise(done => ext.tabs.query(query, done));
}

function setRow(row, state, text) {
  row.classList.remove("ok", "warn", "bad");
  row.classList.add(state);
  row.querySelector("small").textContent = text;
}

function setMessage(text) {
  message.textContent = text || "";
}

function activeTargetSummary(status) {
  const mo2 = status.mo2 || {};
  const instancePath = mo2.activeInstancePath || "";
  const game = mo2.activeGame ? `${mo2.activeGame} - ` : "";
  return instancePath ? `${game}${instancePath}` : "";
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
    status.downloads &&
    status.downloads.exists &&
    status.cookies &&
    status.cookies.exists
  );
}

async function checkStatus() {
  setMessage("");
  readyText.textContent = "Checking status...";

  let response;
  try {
    response = await sendMessage({ action: "popup_get_status" });
  } catch (error) {
    response = { ok: false, error: String(error) };
  }

  if (!response || !response.ok) {
    setRow(rows.native, "bad", "Native app is not installed or not reachable.");
    setRow(rows.mo2, "bad", "Unavailable until native bridge works.");
    setRow(rows.active, "bad", "Unavailable until native bridge works.");
    setRow(rows.downloads, "bad", "Unavailable until native bridge works.");
    setRow(rows.cookies, "bad", "Unavailable until native bridge works.");
    readyText.textContent = "Not ready";
    setMessage(response && response.error ? response.error : "Native messaging failed.");
    return;
  }

  setRow(rows.native, "ok", response.nativeApp.baseDir);

  const activePluginPath = response.mo2.activePluginPath || "";
  const hasNativeActiveFields = Boolean(response.mo2.activeInstancePath || activePluginPath);
  const pluginPath = activePluginPath || response.mo2.llPluginPath || "";
  if (activePluginPath || (response.mo2.exists && response.mo2.llPluginInstalled)) {
    setRow(
      rows.mo2,
      "ok",
      activePluginPath || (
        hasNativeActiveFields
          ? pluginPath
          : "Plugin installed; active target is shown below."
      )
    );
  } else if (response.mo2.exists) {
    setRow(rows.mo2, "warn", "MO2 found, but LL plugin is not installed.");
  } else if (response.downloads.exists) {
    setRow(rows.mo2, "ok", "Plugin status unavailable; target is synced below.");
  } else {
    setRow(rows.mo2, "bad", "ModOrganizer.exe was not found.");
  }

  const activeTarget = activeTargetSummary(response);
  const hasTargetDownloads = Boolean(response.downloads && response.downloads.exists);
  setRow(
    rows.active,
    activeTarget || hasTargetDownloads ? "ok" : "warn",
    activeTarget || (
      hasTargetDownloads
        ? "Synced through target downloads below."
        : "Open MO2 with LL Integration installed to sync the active target."
    )
  );

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
  const response = await sendMessage({ action: "popup_export_cookies" });
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
    const tabs = await queryTabs({ active: true, currentWindow: true });
    tab = tabs && tabs[0];
  } catch (error) {
    setMessage(`Could not read current tab: ${error}`);
    return;
  }

  if (!tab || !tab.url || !/^https?:\/\//i.test(tab.url)) {
    setMessage("Open a web page before arming capture.");
    return;
  }

  const response = await sendMessage({
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
  const response = await sendMessage({ action: "popup_cancel_pending" });
  if (!response || !response.ok) {
    setMessage(response && response.error ? response.error : "Could not cancel pending download.");
    return;
  }

  setMessage(response.cancelled ? "Pending download cancelled." : "No pending download to cancel.");
  await checkStatus();
}

async function openMonitor() {
  const originalText = monitorButton.textContent;
  monitorButton.disabled = true;
  monitorButton.textContent = "Opening...";
  setMessage("Opening floating controls. First launch can take a few seconds.");

  try {
    let tab = null;
    try {
      const tabs = await queryTabs({ active: true, currentWindow: true });
      tab = tabs && tabs[0] ? tabs[0] : null;
    } catch (_error) {
      tab = null;
    }

    const response = await sendMessage({
      action: "popup_open_monitor",
      targetTab: tab
        ? {
            id: tab.id,
            url: tab.url || "",
            title: tab.title || ""
          }
        : null
    });

    if (!response || !response.ok) {
      setMessage(response && response.error ? response.error : "Could not open floating controls.");
      return;
    }

    setMessage(response.closed ? "Floating controls closed." : "Floating controls opened. Arm follows the active tab.");
  } catch (error) {
    setMessage(`Could not open floating controls: ${error}`);
  } finally {
    monitorButton.disabled = false;
    monitorButton.textContent = originalText;
  }
}

refreshButton.addEventListener("click", checkStatus);
exportButton.addEventListener("click", exportCookies);
captureButton.addEventListener("click", captureCurrentPage);
cancelButton.addEventListener("click", cancelPending);
monitorButton.addEventListener("click", openMonitor);

checkStatus();
