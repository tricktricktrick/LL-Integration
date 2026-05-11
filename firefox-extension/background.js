const ARCHIVE_EXTENSIONS = [".7z", ".zip", ".rar"];
const PENDING_LL_DOWNLOAD_TIMEOUT_MS = 30 * 60 * 1000;
const EXTERNAL_CAPTURE_TIMEOUT_MS = 5 * 60 * 1000;
const EXTERNAL_CAPTURE_GRACE_MS = 60 * 1000;
const NATIVE_MESSAGE_TIMEOUT_MS = 8000;
let lastLLDownloadEvent = null;
let externalArchiveCapture = null;
const externalCaptureDownloads = new Map();
let monitorWindowId = null;
let monitorTabId = null;
let floatingControlsTimer = null;
let floatingControlsLastSeq = 0;
let floatingControlsFollow = false;
let floatingControlsPollInFlight = false;
let floatingControlsSyncInFlight = false;
let floatingControlsPort = null;
let floatingControlsNextRequestId = 1;
let floatingControlsTargetTab = null;
const floatingControlsRequests = new Map();

function isArchivePath(path) {
  const lower = (path || "").toLowerCase();
  return ARCHIVE_EXTENSIONS.some(ext => lower.endsWith(ext));
}

function sendNative(payload, okLog, errorLog) {
  const timeout = new Promise(resolve => {
    setTimeout(() => resolve({ ok: false, error: "Native messaging timed out." }), NATIVE_MESSAGE_TIMEOUT_MS);
  });
  return Promise.race([
    browser.runtime.sendNativeMessage("ll_integration_native", payload),
    timeout
  ])
    .then(response => {
      console.log(okLog, response);
      return response;
    })
    .catch(error => {
      console.error(errorLog, error);
      return { ok: false, error: String(error) };
    });
}

function rejectFloatingControlsRequests(error) {
  for (const pending of floatingControlsRequests.values()) {
    clearTimeout(pending.timeoutId);
    pending.reject(error);
  }
  floatingControlsRequests.clear();
}

function floatingControlsNativePort() {
  if (floatingControlsPort) {
    return floatingControlsPort;
  }

  floatingControlsPort = browser.runtime.connectNative("ll_integration_native");
  floatingControlsPort.onMessage.addListener((response) => {
    const requestId = response && response.requestId;
    const pending = floatingControlsRequests.get(requestId);
    if (!pending) {
      return;
    }

    clearTimeout(pending.timeoutId);
    floatingControlsRequests.delete(requestId);
    pending.resolve(response);
  });
  floatingControlsPort.onDisconnect.addListener(() => {
    const error = browser.runtime.lastError
      ? new Error(browser.runtime.lastError.message)
      : new Error("Floating controls native port disconnected.");
    floatingControlsPort = null;
    rejectFloatingControlsRequests(error);
    stopFloatingControlsPolling();
  });
  return floatingControlsPort;
}

function floatingControlsNativeRequest(payload) {
  return new Promise((resolve, reject) => {
    const requestId = floatingControlsNextRequestId++;
    const timeoutId = setTimeout(() => {
      floatingControlsRequests.delete(requestId);
      if (floatingControlsPort) {
        try {
          floatingControlsPort.disconnect();
        } catch (_error) {
          // The port may already be gone.
        }
        floatingControlsPort = null;
      }
      reject(new Error("Floating controls native request timed out."));
    }, NATIVE_MESSAGE_TIMEOUT_MS);

    floatingControlsRequests.set(requestId, { resolve, reject, timeoutId });

    try {
      floatingControlsNativePort().postMessage({ ...payload, requestId });
    } catch (error) {
      clearTimeout(timeoutId);
      floatingControlsRequests.delete(requestId);
      reject(error);
    }
  });
}

function isWebTab(tab) {
  return Boolean(tab && tab.id != null && tab.url && /^https?:\/\//i.test(tab.url));
}

function rememberFloatingControlsTarget(tab) {
  if (!isWebTab(tab)) {
    return null;
  }

  floatingControlsTargetTab = {
    id: tab.id,
    url: tab.url,
    title: tab.title || "",
    lastAccessed: tab.lastAccessed || Date.now()
  };
  return floatingControlsTargetTab;
}

function webTabFromList(tabs) {
  const candidates = (tabs || [])
    .filter(isWebTab)
    .sort((left, right) => (right.lastAccessed || 0) - (left.lastAccessed || 0));
  return candidates[0] || null;
}

async function currentActiveWebTab() {
  const focused = webTabFromList(await browser.tabs.query({ active: true, lastFocusedWindow: true }));
  if (focused) {
    return rememberFloatingControlsTarget(focused);
  }

  const active = webTabFromList(await browser.tabs.query({ active: true }));
  return rememberFloatingControlsTarget(active);
}

async function tabById(tabId) {
  try {
    const tab = await browser.tabs.get(tabId);
    return isWebTab(tab) ? tab : null;
  } catch (_error) {
    return null;
  }
}

async function armCaptureForTab(tab) {
  if (!tab || tab.id == null || !tab.url || !/^https?:\/\//i.test(tab.url)) {
    return null;
  }
  return setExternalArchiveCapture(tab.url, tab.title || "", tab.id);
}

async function retargetFloatingCapture(tab) {
  const target = rememberFloatingControlsTarget(tab)
    || await currentActiveWebTab()
    || floatingControlsTargetTab;
  if (!target) {
    return null;
  }
  return armCaptureForTab(target);
}

function floatingStatusLabel() {
  const capture = getExternalArchiveCapture();
  if (lastLLDownloadEvent && lastLLDownloadEvent.download) {
    return `Waiting: ${lastLLDownloadEvent.download.name || "archive"}`;
  }
  if (capture) {
    let label = capture.pageTitle || "";
    if (!label) {
      try {
        label = new URL(capture.pageUrl).hostname;
      } catch (_error) {
        label = capture.pageUrl || "page";
      }
    }
    return `${floatingControlsFollow ? "Follow armed" : "Armed"}: ${label}`;
  }
  return floatingControlsFollow ? "Armed: waiting for page" : "Idle";
}

function syncFloatingControlsStatus(visible = true) {
  if (floatingControlsSyncInFlight) {
    return Promise.resolve({ ok: true, skipped: true });
  }
  floatingControlsSyncInFlight = true;
  return floatingControlsNativeRequest(
    {
      action: "floating_controls_status",
      source: "firefox",
      armed: Boolean(getPendingLLDownload() || getExternalArchiveCapture() || floatingControlsFollow),
      follow: floatingControlsFollow,
      label: floatingStatusLabel(),
      visible
    }
  ).finally(() => {
    floatingControlsSyncInFlight = false;
  });
}

async function handleFloatingControlsCommand(state) {
  const command = state.command || "";
  if (!command) {
    return;
  }

  if (command === "arm") {
    floatingControlsFollow = true;
    await retargetFloatingCapture();
  } else if (command === "disarm") {
    floatingControlsFollow = false;
    cancelPendingCapture();
  } else if (command === "follow_on") {
    floatingControlsFollow = true;
    if (getExternalArchiveCapture()) {
      await retargetFloatingCapture();
    }
  } else if (command === "follow_off") {
    floatingControlsFollow = false;
  } else if (command === "close") {
    floatingControlsFollow = false;
    cancelPendingCapture();
    stopFloatingControlsPolling();
    return false;
  }
  return true;
}

async function pollFloatingControls() {
  if (floatingControlsPollInFlight) {
    return;
  }
  floatingControlsPollInFlight = true;
  try {
    const response = await floatingControlsNativeRequest({ action: "floating_controls_state", source: "firefox" });
    if (!response || !response.ok || !response.state) {
      return;
    }

    const state = response.state;
    const seq = Number(state.seq || 0);
    if (seq > floatingControlsLastSeq) {
      floatingControlsLastSeq = seq;
      const shouldSync = await handleFloatingControlsCommand(state);
      if (shouldSync === false) {
        return;
      }
    }
    if (state.visible === false) {
      stopFloatingControlsPolling();
      return;
    }
    await syncFloatingControlsStatus(true);
  } finally {
    floatingControlsPollInFlight = false;
  }
}

function startFloatingControlsPolling() {
  if (floatingControlsTimer) {
    return;
  }
  floatingControlsTimer = setInterval(() => {
    pollFloatingControls().catch(error => console.error("Floating controls poll failed:", error));
  }, 3000);
  pollFloatingControls().catch(error => console.error("Floating controls poll failed:", error));
}

function stopFloatingControlsPolling() {
  if (!floatingControlsTimer) {
    return;
  }
  clearInterval(floatingControlsTimer);
  floatingControlsTimer = null;
}

async function openFloatingControls(targetTab) {
  rememberFloatingControlsTarget(targetTab);
  const response = await floatingControlsNativeRequest({ action: "open_floating_controls", source: "firefox" });
  if (response && response.ok) {
    if (response.closed) {
      floatingControlsFollow = false;
      cancelPendingCapture();
      await syncFloatingControlsStatus(false);
      setTimeout(() => stopFloatingControlsPolling(), 1500);
      return response;
    }
    startFloatingControlsPolling();
    return response;
  }

  return response || { ok: false, error: "Could not open floating controls." };
}

function openMonitorWindow(targetTab) {
  const params = new URLSearchParams();
  if (targetTab && targetTab.id != null) {
    params.set("tabId", String(targetTab.id));
  }
  if (targetTab && targetTab.url) {
    params.set("pageUrl", targetTab.url);
  }
  if (targetTab && targetTab.title) {
    params.set("pageTitle", targetTab.title);
  }

  const url = browser.runtime.getURL(`monitor.html?${params.toString()}`);
  return browser.windows.create({
    url,
    type: "popup",
    width: 380,
    height: 220
  }).catch(error => {
    console.warn("Hook monitor popup window failed, trying normal window.", error);
    return browser.windows.create({
      url,
      type: "normal",
      width: 380,
      height: 220
    });
  }).catch(error => {
    console.warn("Hook monitor window failed, opening tab fallback.", error);
    return browser.tabs.create({ url });
  });
}

async function toggleMonitorWindow(targetTab) {
  if (monitorWindowId != null || monitorTabId != null) {
    try {
      if (monitorWindowId != null) {
        await browser.windows.remove(monitorWindowId);
      } else {
        await browser.tabs.remove(monitorTabId);
      }
    } catch (_error) {
      // The monitor may already be gone; clearing the ids is enough.
    }
    monitorWindowId = null;
    monitorTabId = null;
    cancelPendingCapture();
    return { ok: true, closed: true, disarmed: true };
  }

  const opened = await openMonitorWindow(targetTab);
  if (opened && opened.url) {
    monitorTabId = opened.id;
    monitorWindowId = null;
  } else {
    monitorWindowId = opened && opened.id;
    monitorTabId = null;
  }
  return { ok: true, windowId: monitorWindowId, tabId: monitorTabId };
}

function setPendingLLDownload(eventPayload) {
  externalArchiveCapture = null;
  lastLLDownloadEvent = {
    ...eventPayload,
    pendingStartedAt: new Date().toISOString(),
    pendingExpiresAt: new Date(Date.now() + PENDING_LL_DOWNLOAD_TIMEOUT_MS).toISOString()
  };
}

function urlOrigin(url) {
  try {
    return new URL(url).origin;
  } catch (_error) {
    return "";
  }
}

function isLoversLabUrl(url) {
  return isSiteUrl(url, "loverslab.com");
}

function isDwemerModsUrl(url) {
  return isSiteUrl(url, "dwemermods.com");
}

function isSiteUrl(url, domain) {
  const host = (() => {
    try {
      return new URL(url).hostname.toLowerCase();
    } catch (_error) {
      return "";
    }
  })();
  return host === domain || host.endsWith(`.${domain}`);
}

function isSupportedSourceUrl(url) {
  return isLoversLabUrl(url) || isDwemerModsUrl(url);
}

function getPendingLLDownload() {
  if (!lastLLDownloadEvent) {
    return null;
  }

  const expiresAt = Date.parse(lastLLDownloadEvent.pendingExpiresAt || "");
  if (Number.isFinite(expiresAt) && Date.now() > expiresAt) {
    console.log("LL pending download expired:", lastLLDownloadEvent.download && lastLLDownloadEvent.download.name);
    lastLLDownloadEvent = null;
    return null;
  }

  return lastLLDownloadEvent;
}

function setExternalArchiveCapture(pageUrl, pageTitle, tabId) {
  lastLLDownloadEvent = null;
  externalArchiveCapture = {
    action: "capture_external_archive",
    source: "firefox",
    sourceType: "external",
    capturedAt: new Date().toISOString(),
    pageUrl,
    pageTitle: pageTitle || "",
    pageOrigin: urlOrigin(pageUrl),
    tabId,
    capturedCount: 0,
    pendingStartedAt: new Date().toISOString(),
    pendingExpiresAt: new Date(Date.now() + EXTERNAL_CAPTURE_TIMEOUT_MS).toISOString(),
    download: {
      name: "",
      version: "",
      url: "",
      size: "",
      date_iso: ""
    }
  };
  return externalArchiveCapture;
}

function shortenExternalCapture(reason) {
  if (!externalArchiveCapture) {
    return;
  }

  const graceExpiresAt = Date.now() + EXTERNAL_CAPTURE_GRACE_MS;
  const currentExpiresAt = Date.parse(externalArchiveCapture.pendingExpiresAt || "");
  if (!Number.isFinite(currentExpiresAt) || currentExpiresAt > graceExpiresAt) {
    externalArchiveCapture.pendingExpiresAt = new Date(graceExpiresAt).toISOString();
    externalArchiveCapture.captureNotice = reason;
    console.log("External archive capture entering grace period:", reason);
  }
}

function getExternalArchiveCapture() {
  if (!externalArchiveCapture) {
    return null;
  }

  const expiresAt = Date.parse(externalArchiveCapture.pendingExpiresAt || "");
  if (Number.isFinite(expiresAt) && Date.now() > expiresAt) {
    console.log("External archive capture expired:", externalArchiveCapture.pageUrl);
    externalArchiveCapture = null;
    return null;
  }

  return externalArchiveCapture;
}

function cancelPendingCapture() {
  const hadPending = Boolean(lastLLDownloadEvent || externalArchiveCapture);
  lastLLDownloadEvent = null;
  externalArchiveCapture = null;
  externalCaptureDownloads.clear();
  return { ok: true, cancelled: hadPending };
}

function archiveNameFromPath(path) {
  return String(path || "").split(/[\\/]/).pop();
}

function externalEventForCompletedArchive(item, captureSnapshot) {
  const capture = captureSnapshot || getExternalArchiveCapture();
  if (!capture) {
    return null;
  }

  return {
    kind: "external",
    event: {
      ...capture,
      capturedCount: (capture.capturedCount || 0) + 1,
      download: {
        name: archiveNameFromPath(item.filename),
        version: "",
        url: item.finalUrl || item.url || "",
        size: "",
        date_iso: ""
      }
    }
  };
}

function eventForCompletedArchive(item) {
  const externalSnapshot = externalCaptureDownloads.get(item.id);
  if (externalSnapshot) {
    return externalEventForCompletedArchive(item, externalSnapshot);
  }

  const pendingLLDownload = getPendingLLDownload();
  if (pendingLLDownload) {
    return { kind: "ll", event: pendingLLDownload };
  }

  return null;
}

async function shouldMarkExternalDownload(item) {
  const capture = getExternalArchiveCapture();
  if (!capture || getPendingLLDownload()) {
    return null;
  }

  let capturedTab = null;
  try {
    capturedTab = await browser.tabs.get(capture.tabId);
  } catch (_error) {
    return null;
  }

  const tabUrl = capturedTab && capturedTab.url ? capturedTab.url : "";
  const referrer = item.referrer || "";

  const tabOrigin = urlOrigin(tabUrl);
  if (tabOrigin === capture.pageOrigin) {
    return { ...capture };
  }

  if (referrer && urlOrigin(referrer) === capture.pageOrigin) {
    return { ...capture };
  }

  return null;
}

async function exportLLCookies() {
  const allCookies = await browser.cookies.getAll({});
  const llCookies = allCookies.filter(c =>
    c.domain.includes("loverslab.com")
  );

  console.log("LL cookies found:", llCookies.length);

  const payload = {
    action: "save_ll_cookies",
    source: "firefox",
    exportedAt: new Date().toISOString(),
    cookies: llCookies.map(c => ({
      name: c.name,
      value: c.value,
      domain: c.domain,
      path: c.path,
      secure: c.secure,
      httpOnly: c.httpOnly,
      expirationDate: c.expirationDate || null
    }))
  };

  return sendNative(payload, "Native response:", "Native error:");
}

browser.runtime.onMessage.addListener((message) => {
  if (!message) {
    return;
  }

  if (message.action === "popup_get_status") {
    return sendNative({ action: "status", source: "firefox" }, "Native status:", "Native status error:")
      .then(response => ({
        ...response,
        pendingDownload: getPendingLLDownload(),
        externalCapture: getExternalArchiveCapture()
      }));
  }

  if (message.action === "popup_export_cookies") {
    return exportLLCookies();
  }

  if (message.action === "popup_capture_current_page") {
    const capture = setExternalArchiveCapture(message.pageUrl, message.pageTitle, message.tabId);
    return Promise.resolve({ ok: true, capture });
  }

  if (message.action === "popup_cancel_pending") {
    return Promise.resolve(cancelPendingCapture());
  }

  if (message.action === "popup_open_monitor") {
    try {
      return openFloatingControls(message.targetTab)
        .catch(error => ({ ok: false, error: String(error) }));
    } catch (error) {
      return Promise.resolve({ ok: false, error: String(error) });
    }
  }

  if (message.action !== "ll_download_clicked") {
    return;
  }

  const payload = {
    action: "save_ll_download_event",
    source: "firefox",
    capturedAt: new Date().toISOString(),
    sourceType: message.sourceType || "loverslab",
    pageUrl: message.pageUrl,
    pageTitle: message.pageTitle || "",
    download: message.download
  };
  setPendingLLDownload(payload);

  return sendNative(payload, "LL download event saved:", "LL download event error:");
});

browser.downloads.onChanged.addListener(async (delta) => {
  if (!delta.state || delta.state.current !== "complete") {
    return;
  }

  const items = await browser.downloads.search({ id: delta.id });
  const item = items && items[0];
  if (!item || !isArchivePath(item.filename)) {
    externalCaptureDownloads.delete(delta.id);
    return;
  }

  const pending = eventForCompletedArchive(item);
  if (!pending) {
    return;
  }

  const payload = {
    action: "save_ll_download_completed",
    source: "firefox",
    completedAt: new Date().toISOString(),
    archivePath: item.filename,
    browserDownloadUrl: item.finalUrl || item.url || "",
    event: pending.event
  };

  const response = await sendNative(
    payload,
    "LL download sidecar saved:",
    "LL download sidecar error:"
  );

  if (pending.kind === "ll") {
    if (response && response.ok) {
      lastLLDownloadEvent = null;
    }
  } else {
    externalCaptureDownloads.delete(item.id);
    if (response && response.ok) {
      if (externalArchiveCapture) {
        externalArchiveCapture.capturedCount = (externalArchiveCapture.capturedCount || 0) + 1;
      }
    }
  }
});

browser.downloads.onCreated.addListener(async (item) => {
  const captureSnapshot = await shouldMarkExternalDownload(item);
  if (!captureSnapshot) {
    return;
  }

  externalCaptureDownloads.set(item.id, captureSnapshot);
});

browser.tabs.onUpdated.addListener((tabId, changeInfo, tabInfo) => {
  if (tabInfo && tabInfo.active && isWebTab(tabInfo)) {
    rememberFloatingControlsTarget(tabInfo);
    if (floatingControlsFollow) {
      tabById(tabId).then(tab => {
        if (tab) {
          armCaptureForTab(tab);
          syncFloatingControlsStatus(true);
        }
      });
    }
  }

  const capture = getExternalArchiveCapture();
  if (!capture || capture.tabId !== tabId || (!changeInfo.url && !changeInfo.title)) {
    return;
  }

  if (floatingControlsFollow) {
    tabById(tabId).then(tab => {
      if (tab) {
        armCaptureForTab(tab);
        syncFloatingControlsStatus();
      }
    });
    return;
  }

  if (changeInfo.url !== capture.pageUrl) {
    shortenExternalCapture("Captured page changed; capture will stop soon.");
  }
});

browser.tabs.onActivated.addListener((activeInfo) => {
  tabById(activeInfo.tabId).then(tab => {
    if (tab) {
      rememberFloatingControlsTarget(tab);
    } else {
      floatingControlsTargetTab = null;
    }

    const capture = getExternalArchiveCapture();
    if (floatingControlsFollow) {
      if (tab) {
        armCaptureForTab(tab);
        syncFloatingControlsStatus(true);
      }
      return;
    }

    if (!capture || capture.tabId === activeInfo.tabId) {
      return;
    }

    shortenExternalCapture("Captured tab is no longer active; capture will stop soon.");
  });
});

browser.windows.onRemoved.addListener((windowId) => {
  if (windowId === monitorWindowId) {
    monitorWindowId = null;
    cancelPendingCapture();
  }
});

browser.tabs.onRemoved.addListener((tabId) => {
  if (tabId === monitorTabId) {
    monitorTabId = null;
    cancelPendingCapture();
  }
});
