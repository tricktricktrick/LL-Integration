const ARCHIVE_EXTENSIONS = [".7z", ".zip", ".rar"];
const PENDING_LL_DOWNLOAD_TIMEOUT_MS = 30 * 60 * 1000;
const EXTERNAL_CAPTURE_TIMEOUT_MS = 5 * 60 * 1000;
const EXTERNAL_CAPTURE_GRACE_MS = 60 * 1000;
let lastLLDownloadEvent = null;
let externalArchiveCapture = null;
const externalCaptureDownloads = new Map();

function isArchivePath(path) {
  const lower = (path || "").toLowerCase();
  return ARCHIVE_EXTENSIONS.some(ext => lower.endsWith(ext));
}

function sendNative(payload, okLog, errorLog) {
  return browser.runtime.sendNativeMessage("ll_integration_native", payload)
    .then(response => {
      console.log(okLog, response);
      return response;
    })
    .catch(error => {
      console.error(errorLog, error);
      return { ok: false, error: String(error) };
    });
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
  const host = (() => {
    try {
      return new URL(url).hostname.toLowerCase();
    } catch (_error) {
      return "";
    }
  })();
  return host === "loverslab.com" || host.endsWith(".loverslab.com");
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

  const downloadUrl = item.finalUrl || item.url || "";
  const referrer = item.referrer || "";
  if (isLoversLabUrl(downloadUrl) || isLoversLabUrl(referrer)) {
    return null;
  }

  let capturedTab = null;
  try {
    capturedTab = await browser.tabs.get(capture.tabId);
  } catch (_error) {
    return null;
  }

  const tabUrl = capturedTab && capturedTab.url ? capturedTab.url : "";
  if (isLoversLabUrl(tabUrl)) {
    externalArchiveCapture = null;
    return null;
  }

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

  if (message.action !== "ll_download_clicked") {
    return;
  }

  const payload = {
    action: "save_ll_download_event",
    source: "firefox",
    capturedAt: new Date().toISOString(),
    pageUrl: message.pageUrl,
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

browser.tabs.onUpdated.addListener((tabId, changeInfo) => {
  const capture = getExternalArchiveCapture();
  if (!capture || capture.tabId !== tabId || !changeInfo.url) {
    return;
  }

  if (changeInfo.url !== capture.pageUrl) {
    shortenExternalCapture("Captured page changed; capture will stop soon.");
  }
});

browser.tabs.onActivated.addListener((activeInfo) => {
  const capture = getExternalArchiveCapture();
  if (!capture || capture.tabId === activeInfo.tabId) {
    return;
  }

  shortenExternalCapture("Captured tab is no longer active; capture will stop soon.");
});
