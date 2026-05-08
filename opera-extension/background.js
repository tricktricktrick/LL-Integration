const ext = globalThis.browser || globalThis.chrome;

const ARCHIVE_EXTENSIONS = [".7z", ".zip", ".rar"];
const PENDING_LL_DOWNLOAD_TIMEOUT_MS = 30 * 60 * 1000;
const EXTERNAL_CAPTURE_TIMEOUT_MS = 5 * 60 * 1000;
const EXTERNAL_CAPTURE_GRACE_MS = 60 * 1000;

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

function storageGet(keys) {
  return callbackPromise(done => ext.storage.local.get(keys, done));
}

function storageSet(values) {
  return callbackPromise(done => ext.storage.local.set(values, done));
}

function storageRemove(keys) {
  return callbackPromise(done => ext.storage.local.remove(keys, done));
}

function getDownloads(query) {
  return callbackPromise(done => ext.downloads.search(query, done));
}

function getTab(tabId) {
  return callbackPromise(done => ext.tabs.get(tabId, done));
}

function getCookies(details) {
  return callbackPromise(done => ext.cookies.getAll(details, done));
}

function isArchivePath(path) {
  const lower = (path || "").toLowerCase();
  return ARCHIVE_EXTENSIONS.some(extName => lower.endsWith(extName));
}

function sendNative(payload, okLog, errorLog) {
  return new Promise((resolve) => {
    ext.runtime.sendNativeMessage("ll_integration_native", payload, (response) => {
      const error = ext.runtime.lastError;
      if (error) {
        console.error(errorLog, error.message || String(error));
        resolve({ ok: false, error: error.message || String(error) });
        return;
      }

      console.log(okLog, response);
      resolve(response || { ok: false, error: "Native bridge returned no response." });
    });
  });
}

async function getExternalCaptureDownloads() {
  const state = await storageGet(["externalCaptureDownloads"]);
  return state.externalCaptureDownloads || {};
}

async function setExternalCaptureDownload(downloadId, captureSnapshot) {
  const downloads = await getExternalCaptureDownloads();
  downloads[String(downloadId)] = captureSnapshot;
  await storageSet({ externalCaptureDownloads: downloads });
}

async function getExternalCaptureDownload(downloadId) {
  const downloads = await getExternalCaptureDownloads();
  return downloads[String(downloadId)] || null;
}

async function deleteExternalCaptureDownload(downloadId) {
  const downloads = await getExternalCaptureDownloads();
  delete downloads[String(downloadId)];
  await storageSet({ externalCaptureDownloads: downloads });
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

async function setPendingLLDownload(eventPayload) {
  const lastLLDownloadEvent = {
    ...eventPayload,
    pendingStartedAt: new Date().toISOString(),
    pendingExpiresAt: new Date(Date.now() + PENDING_LL_DOWNLOAD_TIMEOUT_MS).toISOString()
  };
  await storageSet({
    lastLLDownloadEvent,
    externalArchiveCapture: null,
    externalCaptureDownloads: {}
  });
}

async function getPendingLLDownload() {
  const state = await storageGet(["lastLLDownloadEvent"]);
  const lastLLDownloadEvent = state.lastLLDownloadEvent || null;
  if (!lastLLDownloadEvent) {
    return null;
  }

  const expiresAt = Date.parse(lastLLDownloadEvent.pendingExpiresAt || "");
  if (Number.isFinite(expiresAt) && Date.now() > expiresAt) {
    console.log("LL pending download expired:", lastLLDownloadEvent.download && lastLLDownloadEvent.download.name);
    await storageRemove(["lastLLDownloadEvent"]);
    return null;
  }

  return lastLLDownloadEvent;
}

async function setExternalArchiveCapture(pageUrl, pageTitle, tabId) {
  const externalArchiveCapture = {
    action: "capture_external_archive",
    source: "opera",
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

  await storageSet({
    lastLLDownloadEvent: null,
    externalArchiveCapture,
    externalCaptureDownloads: {}
  });
  return externalArchiveCapture;
}

async function shortenExternalCapture(reason) {
  const capture = await getExternalArchiveCapture();
  if (!capture) {
    return;
  }

  const graceExpiresAt = Date.now() + EXTERNAL_CAPTURE_GRACE_MS;
  const currentExpiresAt = Date.parse(capture.pendingExpiresAt || "");
  if (!Number.isFinite(currentExpiresAt) || currentExpiresAt > graceExpiresAt) {
    capture.pendingExpiresAt = new Date(graceExpiresAt).toISOString();
    capture.captureNotice = reason;
    await storageSet({ externalArchiveCapture: capture });
    console.log("External archive capture entering grace period:", reason);
  }
}

async function getExternalArchiveCapture() {
  const state = await storageGet(["externalArchiveCapture"]);
  const externalArchiveCapture = state.externalArchiveCapture || null;
  if (!externalArchiveCapture) {
    return null;
  }

  const expiresAt = Date.parse(externalArchiveCapture.pendingExpiresAt || "");
  if (Number.isFinite(expiresAt) && Date.now() > expiresAt) {
    console.log("External archive capture expired:", externalArchiveCapture.pageUrl);
    await storageRemove(["externalArchiveCapture"]);
    return null;
  }

  return externalArchiveCapture;
}

async function cancelPendingCapture() {
  const state = await storageGet(["lastLLDownloadEvent", "externalArchiveCapture"]);
  const hadPending = Boolean(state.lastLLDownloadEvent || state.externalArchiveCapture);
  await storageSet({
    lastLLDownloadEvent: null,
    externalArchiveCapture: null,
    externalCaptureDownloads: {}
  });
  return { ok: true, cancelled: hadPending };
}

function archiveNameFromPath(path) {
  return String(path || "").split(/[\\/]/).pop();
}

async function externalEventForCompletedArchive(item, captureSnapshot) {
  const capture = captureSnapshot || await getExternalArchiveCapture();
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

async function eventForCompletedArchive(item) {
  const externalSnapshot = await getExternalCaptureDownload(item.id);
  if (externalSnapshot) {
    return externalEventForCompletedArchive(item, externalSnapshot);
  }

  const pendingLLDownload = await getPendingLLDownload();
  if (pendingLLDownload) {
    return { kind: "ll", event: pendingLLDownload };
  }

  return null;
}

async function shouldMarkExternalDownload(item) {
  const capture = await getExternalArchiveCapture();
  const pendingLLDownload = await getPendingLLDownload();
  if (!capture || pendingLLDownload) {
    return null;
  }

  const downloadUrl = item.finalUrl || item.url || "";
  const referrer = item.referrer || "";
  if (isLoversLabUrl(downloadUrl) || isLoversLabUrl(referrer)) {
    return null;
  }

  let capturedTab = null;
  try {
    capturedTab = await getTab(capture.tabId);
  } catch (_error) {
    return null;
  }

  const tabUrl = capturedTab && capturedTab.url ? capturedTab.url : "";
  if (isLoversLabUrl(tabUrl)) {
    await storageRemove(["externalArchiveCapture"]);
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
  const allCookies = await getCookies({});
  const llCookies = allCookies.filter(c =>
    c.domain.includes("loverslab.com")
  );

  console.log("LL cookies found:", llCookies.length);

  const payload = {
    action: "save_ll_cookies",
    source: "opera",
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

async function handleRuntimeMessage(message) {
  if (!message) {
    return { ok: false, error: "Empty message." };
  }

  if (message.action === "popup_get_status") {
    const response = await sendNative({ action: "status", source: "opera" }, "Native status:", "Native status error:");
    return {
      ...response,
      pendingDownload: await getPendingLLDownload(),
      externalCapture: await getExternalArchiveCapture()
    };
  }

  if (message.action === "popup_export_cookies") {
    return exportLLCookies();
  }

  if (message.action === "popup_capture_current_page") {
    const capture = await setExternalArchiveCapture(message.pageUrl, message.pageTitle, message.tabId);
    return { ok: true, capture };
  }

  if (message.action === "popup_cancel_pending") {
    return cancelPendingCapture();
  }

  if (message.action !== "ll_download_clicked") {
    return { ok: false, error: `Unknown action: ${message.action}` };
  }

  const payload = {
    action: "save_ll_download_event",
    source: "opera",
    capturedAt: new Date().toISOString(),
    pageUrl: message.pageUrl,
    download: message.download
  };
  await setPendingLLDownload(payload);

  return sendNative(payload, "LL download event saved:", "LL download event error:");
}

ext.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  handleRuntimeMessage(message)
    .then(sendResponse)
    .catch(error => sendResponse({ ok: false, error: String(error) }));
  return true;
});

ext.downloads.onChanged.addListener(async (delta) => {
  if (!delta.state || delta.state.current !== "complete") {
    return;
  }

  const items = await getDownloads({ id: delta.id });
  const item = items && items[0];
  if (!item || !isArchivePath(item.filename)) {
    await deleteExternalCaptureDownload(delta.id);
    return;
  }

  const pending = await eventForCompletedArchive(item);
  if (!pending) {
    return;
  }

  const payload = {
    action: "save_ll_download_completed",
    source: "opera",
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
      await storageRemove(["lastLLDownloadEvent"]);
    }
  } else {
    await deleteExternalCaptureDownload(item.id);
    if (response && response.ok) {
      const capture = await getExternalArchiveCapture();
      if (capture) {
        capture.capturedCount = (capture.capturedCount || 0) + 1;
        await storageSet({ externalArchiveCapture: capture });
      }
    }
  }
});

ext.downloads.onCreated.addListener(async (item) => {
  const captureSnapshot = await shouldMarkExternalDownload(item);
  if (!captureSnapshot) {
    return;
  }

  await setExternalCaptureDownload(item.id, captureSnapshot);
});

ext.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (!changeInfo.url) {
    return;
  }

  getExternalArchiveCapture().then((capture) => {
    if (!capture || capture.tabId !== tabId) {
      return;
    }

    if (changeInfo.url !== capture.pageUrl) {
      shortenExternalCapture("Captured page changed; capture will stop soon.");
    }
  });
});

ext.tabs.onActivated.addListener((activeInfo) => {
  getExternalArchiveCapture().then((capture) => {
    if (!capture || capture.tabId === activeInfo.tabId) {
      return;
    }

    shortenExternalCapture("Captured tab is no longer active; capture will stop soon.");
  });
});
