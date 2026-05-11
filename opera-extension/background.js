const ext = globalThis.browser || globalThis.chrome;

const ARCHIVE_EXTENSIONS = [".7z", ".zip", ".rar"];
const PENDING_LL_DOWNLOAD_TIMEOUT_MS = 30 * 60 * 1000;
const EXTERNAL_CAPTURE_TIMEOUT_MS = 5 * 60 * 1000;
const EXTERNAL_CAPTURE_GRACE_MS = 60 * 1000;
const DOWNLOAD_LOOKUP_RETRIES = 8;
const DOWNLOAD_LOOKUP_DELAY_MS = 250;
const NATIVE_MESSAGE_TIMEOUT_MS = 8000;
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

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
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

function extractVersion(name) {
  if (!name) {
    return "";
  }

  const stem = String(name).replace(/\.(7z|zip|rar|tar|gz|bz2|xz)$/i, "");
  const match = stem.match(/(?<!\d)v?(\d+(?:[.-]\d+){1,3})(?:\b|(?=\D))/i);
  return match ? match[1].replace(/-/g, ".") : "";
}

async function findCompletedDownload(id) {
  for (let attempt = 0; attempt < DOWNLOAD_LOOKUP_RETRIES; attempt += 1) {
    const items = await getDownloads({ id });
    const item = items && items[0];
    if (item && item.filename) {
      console.log("Opera download lookup:", { id, attempt, filename: item.filename, state: item.state });
      return item;
    }
    await sleep(DOWNLOAD_LOOKUP_DELAY_MS);
  }

  const items = await getDownloads({ id });
  const item = items && items[0] ? items[0] : null;
  console.log("Opera download lookup final:", { id, found: Boolean(item), filename: item && item.filename, state: item && item.state });
  return item;
}

function sendNative(payload, okLog, errorLog) {
  return new Promise((resolve) => {
    let settled = false;
    const timeoutId = setTimeout(() => {
      if (settled) {
        return;
      }
      settled = true;
      resolve({ ok: false, error: "Native messaging timed out." });
    }, NATIVE_MESSAGE_TIMEOUT_MS);

    ext.runtime.sendNativeMessage("ll_integration_native", payload, (response) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeoutId);
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

  floatingControlsPort = ext.runtime.connectNative("ll_integration_native");
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
    const error = ext.runtime.lastError
      ? new Error(ext.runtime.lastError.message)
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
  const focused = webTabFromList(await callbackPromise(done => ext.tabs.query({ active: true, lastFocusedWindow: true }, done)));
  if (focused) {
    return rememberFloatingControlsTarget(focused);
  }

  const active = webTabFromList(await callbackPromise(done => ext.tabs.query({ active: true }, done)));
  return rememberFloatingControlsTarget(active);
}

async function tabById(tabId) {
  try {
    const tab = await getTab(tabId);
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

async function floatingStatusLabel() {
  const pending = await getPendingLLDownload();
  if (pending && pending.download) {
    return `Waiting: ${pending.download.name || "archive"}`;
  }

  const capture = await getExternalArchiveCapture();
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
  return floatingControlsFollow ? "Follow idle" : "Idle";
}

async function syncFloatingControlsStatus(visible = true) {
  if (floatingControlsSyncInFlight) {
    return { ok: true, skipped: true };
  }
  floatingControlsSyncInFlight = true;
  try {
    const pending = await getPendingLLDownload();
    const capture = await getExternalArchiveCapture();
    return floatingControlsNativeRequest(
      {
        action: "floating_controls_status",
        source: "opera",
        armed: Boolean(pending || capture),
        follow: floatingControlsFollow,
        label: await floatingStatusLabel(),
        visible
      }
    );
  } finally {
    floatingControlsSyncInFlight = false;
  }
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
    await cancelPendingCapture();
  } else if (command === "follow_on") {
    floatingControlsFollow = true;
    if (await getExternalArchiveCapture()) {
      await retargetFloatingCapture();
    }
  } else if (command === "follow_off") {
    floatingControlsFollow = false;
  } else if (command === "close") {
    floatingControlsFollow = false;
    await cancelPendingCapture();
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
    const response = await floatingControlsNativeRequest({ action: "floating_controls_state", source: "opera" });
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
  const response = await floatingControlsNativeRequest({ action: "open_floating_controls", source: "opera" });
  if (response && response.ok) {
    if (response.closed) {
      floatingControlsFollow = false;
      await cancelPendingCapture();
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

  const url = ext.runtime.getURL(`monitor.html?${params.toString()}`);
  return callbackPromise(done => ext.windows.create({
    url,
    type: "popup",
    width: 380,
    height: 220
  }, done)).catch(error => {
    console.warn("Hook monitor popup window failed, trying normal window.", error);
    return callbackPromise(done => ext.windows.create({
      url,
      type: "normal",
      width: 380,
      height: 220
    }, done));
  }).catch(error => {
    console.warn("Hook monitor window failed, opening tab fallback.", error);
    return callbackPromise(done => ext.tabs.create({ url }, done));
  });
}

async function toggleMonitorWindow(targetTab) {
  if (monitorWindowId != null || monitorTabId != null) {
    try {
      if (monitorWindowId != null) {
        await callbackPromise(done => ext.windows.remove(monitorWindowId, done));
      } else {
        await callbackPromise(done => ext.tabs.remove(monitorTabId, done));
      }
    } catch (_error) {
      // The monitor may already be gone; clearing the ids is enough.
    }
    monitorWindowId = null;
    monitorTabId = null;
    await cancelPendingCapture();
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

async function getExternalCaptureDownloads() {
  const state = await storageGet(["externalCaptureDownloads"]);
  return state.externalCaptureDownloads || {};
}

async function setExternalCaptureDownload(downloadId, captureSnapshot) {
  const downloads = await getExternalCaptureDownloads();
  downloads[String(downloadId)] = captureSnapshot;
  await storageSet({ externalCaptureDownloads: downloads });
}

async function setPendingDownloadPath(downloadId, archivePath) {
  const state = await storageGet(["pendingDownloadPaths"]);
  const paths = state.pendingDownloadPaths || {};
  paths[String(downloadId)] = archivePath;
  await storageSet({ pendingDownloadPaths: paths });
}

async function getPendingDownloadPath(downloadId) {
  const state = await storageGet(["pendingDownloadPaths"]);
  const paths = state.pendingDownloadPaths || {};
  return paths[String(downloadId)] || "";
}

async function deletePendingDownloadPath(downloadId) {
  const state = await storageGet(["pendingDownloadPaths"]);
  const paths = state.pendingDownloadPaths || {};
  delete paths[String(downloadId)];
  await storageSet({ pendingDownloadPaths: paths });
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
    externalCaptureDownloads: {},
    pendingDownloadPaths: {}
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

  const fallback = llEventFromCompletedDownload(item);
  if (fallback) {
    console.warn("Opera reconstructed LL event from completed download because no pending click event was available.", {
      id: item.id,
      url: item.finalUrl || item.url || "",
      referrer: item.referrer || ""
    });
    return { kind: "ll", event: fallback };
  }

  return null;
}

function llEventFromCompletedDownload(item) {
  const downloadUrl = item.finalUrl || item.url || "";
  const referrer = item.referrer || "";
  const sourceType = isDwemerModsUrl(downloadUrl) || isDwemerModsUrl(referrer) ? "dwemermods" : "loverslab";
  if (!isSupportedSourceUrl(downloadUrl) && !isSupportedSourceUrl(referrer)) {
    return null;
  }

  const archiveName = archiveNameFromPath(item.filename);
  if (!archiveName || !isArchivePath(archiveName)) {
    return null;
  }

  return {
    action: "save_ll_download_event",
    source: "opera",
    sourceType,
    capturedAt: new Date().toISOString(),
    pageUrl: referrer || downloadUrl,
    download: {
      name: archiveName,
      version: extractVersion(archiveName),
      url: downloadUrl,
      size: "",
      date_iso: ""
    }
  };
}

async function shouldMarkExternalDownload(item) {
  const capture = await getExternalArchiveCapture();
  const pendingLLDownload = await getPendingLLDownload();
  if (!capture || pendingLLDownload) {
    return null;
  }

  let capturedTab = null;
  try {
    capturedTab = await getTab(capture.tabId);
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

  if (message.action === "popup_open_monitor") {
    try {
      return await openFloatingControls(message.targetTab);
    } catch (error) {
      return { ok: false, error: String(error) };
    }
  }

  if (message.action !== "ll_download_clicked") {
    return { ok: false, error: `Unknown action: ${message.action}` };
  }

  const payload = {
    action: "save_ll_download_event",
    source: "opera",
    capturedAt: new Date().toISOString(),
    sourceType: message.sourceType || "loverslab",
    pageUrl: message.pageUrl,
    pageTitle: message.pageTitle || "",
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

  const item = await findCompletedDownload(delta.id);
  const archivePath = (item && item.filename) || await getPendingDownloadPath(delta.id);
  if (!item || !isArchivePath(archivePath)) {
    console.log("Opera completed download ignored:", { id: delta.id, found: Boolean(item), archivePath });
    await deleteExternalCaptureDownload(delta.id);
    await deletePendingDownloadPath(delta.id);
    return;
  }

  const pending = await eventForCompletedArchive(item);
  if (!pending) {
    console.log("Opera completed archive has no pending LL/capture event:", { id: delta.id, archivePath });
    return;
  }

  const payload = {
    action: "save_ll_download_completed",
    source: "opera",
    completedAt: new Date().toISOString(),
    archivePath,
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
  await deletePendingDownloadPath(item.id);
});

ext.downloads.onCreated.addListener(async (item) => {
  if (item && item.id != null && isArchivePath(item.filename)) {
    await setPendingDownloadPath(item.id, item.filename);
  }

  const captureSnapshot = await shouldMarkExternalDownload(item);
  if (!captureSnapshot) {
    return;
  }

  await setExternalCaptureDownload(item.id, captureSnapshot);
});

ext.tabs.onUpdated.addListener((tabId, changeInfo, tabInfo) => {
  if (tabInfo && tabInfo.active && isWebTab(tabInfo)) {
    rememberFloatingControlsTarget(tabInfo);
  }

  if (!changeInfo.url && !changeInfo.title) {
    return;
  }

  getExternalArchiveCapture().then((capture) => {
    if (!capture || capture.tabId !== tabId) {
      return;
    }

    if (floatingControlsFollow) {
      tabById(tabId).then((tab) => {
        if (tab) {
          armCaptureForTab(tab).then(() => syncFloatingControlsStatus(true));
        }
      });
      return;
    }

    if (changeInfo.url !== capture.pageUrl) {
      shortenExternalCapture("Captured page changed; capture will stop soon.");
    }
  });
});

ext.tabs.onActivated.addListener((activeInfo) => {
  tabById(activeInfo.tabId).then((tab) => {
    if (tab) {
      rememberFloatingControlsTarget(tab);
    } else {
      floatingControlsTargetTab = null;
    }

    return getExternalArchiveCapture().then((capture) => ({ capture, tab }));
  }).then(({ capture, tab }) => {
    if (floatingControlsFollow && capture) {
      if (tab) {
        armCaptureForTab(tab).then(() => syncFloatingControlsStatus(true));
      }
      return;
    }

    if (!capture || capture.tabId === activeInfo.tabId) {
      return;
    }

    shortenExternalCapture("Captured tab is no longer active; capture will stop soon.");
  });
});

ext.windows.onRemoved.addListener((windowId) => {
  if (windowId === monitorWindowId) {
    monitorWindowId = null;
    cancelPendingCapture();
  }
});

ext.tabs.onRemoved.addListener((tabId) => {
  if (tabId === monitorTabId) {
    monitorTabId = null;
    cancelPendingCapture();
  }
});
