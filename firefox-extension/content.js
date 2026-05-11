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

function hostMatches(url, domain) {
  try {
    const host = new URL(url).hostname.toLowerCase();
    return host === domain || host.endsWith(`.${domain}`);
  } catch (_error) {
    return false;
  }
}

function isLoversLabDownloadLink(link) {
  if (!link || !link.href) {
    return false;
  }

  try {
    const url = new URL(link.href);
    return hostMatches(url.href, "loverslab.com") && url.searchParams.get("do") === "download";
  } catch (_error) {
    return false;
  }
}

function isDwemerModsPage() {
  return hostMatches(window.location.href, "dwemermods.com");
}

function isDwemerDownloadControl(control) {
  if (!control || !isDwemerModsPage()) {
    return false;
  }

  if (control.matches(".download-btn, [class*='download-btn']")) {
    return true;
  }

  if (control.tagName && control.tagName.toLowerCase() === "a") {
    try {
      const url = new URL(control.href, window.location.href);
      if (!hostMatches(url.href, "dwemermods.com")) {
        return false;
      }
      const path = url.pathname.toLowerCase();
      if (path.includes("/download") || path.includes("/files") || path.includes("/mods/")) {
        return true;
      }
    } catch (_error) {
      return false;
    }
  }

  const text = (control.textContent || control.getAttribute("aria-label") || control.value || "").trim().toLowerCase();
  return /\bdownload\b/.test(text);
}

function pageTitleFallback() {
  const title =
    textOrNull(document, "h1.ipsType_pageTitle") ||
    textOrNull(document, "[data-pageTitle]") ||
    textOrNull(document, "h1") ||
    document.title.replace(/- LoversLab$/i, "").trim();
  return title || null;
}

function fileInfoDateFallback() {
  return attrOrNull(document, "time[datetime]", "datetime");
}

function nearestDwemerFileItem(control) {
  let node = control;
  while (node && node !== document.body) {
    const text = (node.textContent || "").replace(/\s+/g, " ");
    if (/\.(?:7z|zip|rar)\b/i.test(text) || /\b\d+(?:[.,]\d+)?\s*(?:KiB|MiB|GiB|KB|MB|GB)\b/i.test(text)) {
      return node;
    }
    node = node.parentElement;
  }

  return control.closest("[class*='file'], li, tr, article, section") || document;
}

function buildLoversLabDownload(link) {
  const item = link.closest(".ipsDataItem") || document;
  const name =
    textOrNull(item, ".ipsDataItem_title .ipsType_break, .ipsType_break") ||
    pageTitleFallback();
  const meta = textOrNull(item, ".ipsDataItem_meta");

  return {
    sourceType: "loverslab",
    download: {
      name,
      version: extractVersion(name),
      url: link.href,
      size: meta ? meta.split("/")[0].trim() : null,
      date_iso: attrOrNull(item, "time[datetime]", "datetime") || fileInfoDateFallback()
    }
  };
}

function buildDwemerModsDownload(control) {
  const item = nearestDwemerFileItem(control);
  const text = (item.textContent || "").replace(/\s+/g, " ").trim();
  const archiveName = text.match(/[^\\/:*?"<>|\s][^\\/:*?"<>|]*\.(?:7z|zip|rar)\b/i);
  const explicitSize = textOrNull(item, ".file-size");
  const size = explicitSize || (text.match(/\b\d+(?:[.,]\d+)?\s*(?:KiB|MiB|GiB|KB|MB|GB)\b/i) || [null])[0];
  const explicitName =
    textOrNull(item, ".file-name") ||
    textOrNull(item, "h1, h2, h3, h4, [class*='title'], [class*='name']") ||
    (archiveName ? archiveName[0].trim() : null);
  const href = control.href || control.getAttribute("formaction") || (control.form && control.form.action) || window.location.href;
  const name = archiveName ? archiveName[0].trim() : (explicitName || pageTitleFallback());

  return {
    sourceType: "dwemermods",
    pageTitle: pageTitleFallback(),
    download: {
      name,
      version: extractVersion(name),
      url: new URL(href, window.location.href).href,
      size: size ? String(size).replace(",", ".") : null,
      date_iso: attrOrNull(item, "time[datetime]", "datetime") || fileInfoDateFallback()
    }
  };
}

document.addEventListener("click", (event) => {
  const control = event.target.closest('a[data-action="download"], a[href*="do=download"], a[href], button, input[type="submit"], input[type="button"]');
  const eventData = isLoversLabDownloadLink(control)
    ? buildLoversLabDownload(control)
    : isDwemerDownloadControl(control)
      ? buildDwemerModsDownload(control)
      : null;

  if (!eventData) {
    return;
  }

  browser.runtime.sendMessage({
    action: "ll_download_clicked",
    sourceType: eventData.sourceType,
    pageUrl: window.location.href,
    pageTitle: eventData.pageTitle || pageTitleFallback(),
    download: eventData.download
  });
}, true);
