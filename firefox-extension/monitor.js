const readyText = document.getElementById("readyText");
const message = document.getElementById("message");
const armButton = document.getElementById("armButton");
const disarmButton = document.getElementById("disarmButton");
const followButton = document.getElementById("followButton");

const params = new URLSearchParams(window.location.search);
let targetTab = {
  id: Number(params.get("tabId")),
  url: params.get("pageUrl") || "",
  title: params.get("pageTitle") || ""
};
let followActiveTab = false;
let retargeting = false;

function hasTarget() {
  return Number.isFinite(targetTab.id) && targetTab.url && /^https?:\/\//i.test(targetTab.url);
}

function targetLabel() {
  try {
    return new URL(targetTab.url).hostname;
  } catch (_error) {
    return targetTab.title || "target page";
  }
}

function setTargetFromTab(tab) {
  if (!tab || tab.id == null || !tab.url || !/^https?:\/\//i.test(tab.url)) {
    return false;
  }

  targetTab = {
    id: tab.id,
    url: tab.url,
    title: tab.title || ""
  };
  return true;
}

async function activeWebTab() {
  const tabs = await browser.tabs.query({ active: true });
  const candidates = (tabs || [])
    .filter(tab => tab && tab.id != null && tab.url && /^https?:\/\//i.test(tab.url))
    .sort((left, right) => (right.lastAccessed || 0) - (left.lastAccessed || 0));
  return candidates[0] || null;
}

async function followTarget() {
  if (!followActiveTab) {
    return false;
  }

  try {
    const tab = await activeWebTab();
    return setTargetFromTab(tab);
  } catch (_error) {
    return false;
  }
}

function captureLabel(response) {
  if (response.pendingDownload && response.pendingDownload.download) {
    const name = response.pendingDownload.download.name || "archive";
    return `Waiting: ${name}`;
  }

  if (response.externalCapture) {
    const count = response.externalCapture.capturedCount || 0;
    return `${followActiveTab ? "Follow " : ""}Armed on ${targetLabel()}. Captured: ${count}`;
  }

  return hasTarget()
    ? `${followActiveTab ? "Follow " : ""}Idle on ${targetLabel()}`
    : "No target tab";
}

async function refreshStatus() {
  await followTarget();
  let response;
  try {
    response = await browser.runtime.sendMessage({ action: "popup_get_status" });
  } catch (error) {
    response = { ok: false, error: String(error) };
  }

  if (!response || !response.ok) {
    readyText.textContent = "Not ready";
    message.textContent = response && response.error ? response.error : "Native messaging failed.";
    armButton.disabled = true;
    disarmButton.disabled = true;
    return;
  }

  const armed = Boolean(response.externalCapture || response.pendingDownload);
  if (
    followActiveTab &&
    response.externalCapture &&
    !response.pendingDownload &&
    response.externalCapture.tabId !== targetTab.id &&
    hasTarget() &&
    !retargeting
  ) {
    retargeting = true;
    try {
      await browser.runtime.sendMessage({
        action: "popup_capture_current_page",
        pageUrl: targetTab.url,
        pageTitle: targetTab.title,
        tabId: targetTab.id
      });
      response = await browser.runtime.sendMessage({ action: "popup_get_status" });
    } finally {
      retargeting = false;
    }
  }

  readyText.textContent = armed ? "Hook armed" : "Hook idle";
  message.textContent = captureLabel(response);
  armButton.disabled = !hasTarget() || armed;
  disarmButton.disabled = !armed;
  followButton.textContent = followActiveTab ? "Follow On" : "Follow Off";
}

async function armCapture() {
  if (!hasTarget()) {
    message.textContent = "Open this monitor from a web page tab.";
    return;
  }

  message.textContent = "Arming...";
  const response = await browser.runtime.sendMessage({
    action: "popup_capture_current_page",
    pageUrl: targetTab.url,
    pageTitle: targetTab.title,
    tabId: targetTab.id
  });

  if (!response || !response.ok) {
    message.textContent = response && response.error ? response.error : "Could not arm capture.";
  }
  await refreshStatus();
}

async function disarmCapture() {
  message.textContent = "Disarming...";
  const response = await browser.runtime.sendMessage({ action: "popup_cancel_pending" });
  if (!response || !response.ok) {
    message.textContent = response && response.error ? response.error : "Could not disarm capture.";
  }
  await refreshStatus();
}

async function toggleFollow() {
  followActiveTab = !followActiveTab;
  followButton.textContent = followActiveTab ? "Follow On" : "Follow Off";
  await refreshStatus();
}

armButton.addEventListener("click", armCapture);
disarmButton.addEventListener("click", disarmCapture);
followButton.addEventListener("click", toggleFollow);

refreshStatus();
setInterval(refreshStatus, 2000);
