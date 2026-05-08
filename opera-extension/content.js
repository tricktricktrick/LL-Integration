const ext = globalThis.browser || globalThis.chrome;

function textOrNull(root, selector) {
  const element = root.querySelector(selector);
  const text = element ? element.textContent.trim().replace(/\s+/g, " ") : "";
  return text || null;
}

function attrOrNull(root, selector, attribute) {
  const element = root.querySelector(selector);
  return element ? element.getAttribute(attribute) : null;
}

function extractVersion(name) {
  if (!name) {
    return null;
  }

  const stem = name.replace(/\.(7z|zip|rar|tar|gz|bz2|xz)$/i, "");
  const match = stem.match(/\bv?(\d+(?:\.\d+){1,3})(?:\b|(?=\D))/i);
  return match ? match[1] : null;
}

function isLLDownloadLink(link) {
  if (!link || !link.href) {
    return false;
  }

  try {
    const url = new URL(link.href);
    return url.hostname.includes("loverslab.com") && url.searchParams.get("do") === "download";
  } catch (_error) {
    return false;
  }
}

function pageTitleFallback() {
  const title =
    textOrNull(document, "h1.ipsType_pageTitle") ||
    textOrNull(document, "[data-pageTitle]") ||
    document.title.replace(/- LoversLab$/i, "").trim();
  return title || null;
}

function fileInfoDateFallback() {
  return attrOrNull(document, "time[datetime]", "datetime");
}

document.addEventListener("click", (event) => {
  const link = event.target.closest('a[data-action="download"], a[href*="do=download"]');
  if (!isLLDownloadLink(link)) {
    return;
  }

  const item = link.closest(".ipsDataItem") || document;
  const name =
    textOrNull(item, ".ipsDataItem_title .ipsType_break, .ipsType_break") ||
    pageTitleFallback();
  const meta = textOrNull(item, ".ipsDataItem_meta");

  ext.runtime.sendMessage({
    action: "ll_download_clicked",
    pageUrl: window.location.href,
    download: {
      name,
      version: extractVersion(name),
      url: link.href,
      size: meta ? meta.split("/")[0].trim() : null,
      date_iso: attrOrNull(item, "time[datetime]", "datetime") || fileInfoDateFallback()
    }
  });
}, true);
