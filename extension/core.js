export const PASSTHROUGH_HOST = "mediap.svc.ms";
export const PASSTHROUGH_PATH = "/transform/passthrough";

export function isPassthroughRequest(url) {
  try {
    const parsed = new URL(url);
    return parsed.protocol === "https:" &&
      parsed.hostname === PASSTHROUGH_HOST &&
      parsed.pathname.startsWith(PASSTHROUGH_PATH);
  } catch {
    return false;
  }
}

export function cleanFilename(name) {
  const fallback = "sharepoint-document.pdf";
  const value = String(name || "")
    .replace(/[\r\n]/g, "")
    .trim();

  if (!value) {
    return fallback;
  }

  const cleaned = value
    .replace(/[\\/*?:"<>|]/g, "")
    .replace(/\s+/g, "_");

  if (!cleaned) {
    return fallback;
  }

  return cleaned.toLowerCase().endsWith(".pdf") ? cleaned : `${cleaned}.pdf`;
}

export function getHeaderValue(headers, wantedName) {
  const wanted = wantedName.toLowerCase();
  const match = (headers || []).find((header) => header.name.toLowerCase() === wanted);
  return match?.value || null;
}

function getRawQueryParam(url, names) {
  const parsed = new URL(url);
  const wanted = new Set(names.map((name) => name.toLowerCase()));
  const query = parsed.search.startsWith("?") ? parsed.search.slice(1) : parsed.search;

  for (const part of query.split("&")) {
    const separator = part.indexOf("=");
    const rawName = separator >= 0 ? part.slice(0, separator) : part;
    const rawValue = separator >= 0 ? part.slice(separator + 1) : "";
    if (wanted.has(decodeURIComponent(rawName).toLowerCase()) && rawValue) {
      return rawValue;
    }
  }

  return null;
}

function decodeQueryValue(value) {
  return decodeURIComponent(String(value).replace(/\+/g, " "));
}

function deriveSharePointWebPath(decodedSourcePath) {
  const parts = decodedSourcePath.split("/").filter(Boolean);
  if (parts.length < 2) {
    return null;
  }

  if (parts[0] === "personal") {
    return `/personal/${parts[1]}`;
  }

  if (parts[0] === "sites" || parts[0] === "teams") {
    return `/${parts[0]}/${parts[1]}`;
  }

  return null;
}

export function buildSharePointDownload(pageUrl) {
  try {
    const parsed = new URL(pageUrl);
    const rawSourcePath = getRawQueryParam(pageUrl, ["id", "SourceUrl"]);
    if (!rawSourcePath) {
      return null;
    }

    const decodedSourcePath = decodeQueryValue(rawSourcePath);
    if (!decodedSourcePath.startsWith("/") || !/\.pdf$/i.test(decodedSourcePath)) {
      return null;
    }

    const webPath = deriveSharePointWebPath(decodedSourcePath);
    if (!webPath) {
      return null;
    }

    const filename = cleanFilename(decodedSourcePath.split("/").pop());
    return {
      url: `${parsed.origin}${webPath}/_layouts/15/download.aspx?SourceUrl=${rawSourcePath}`,
      filename
    };
  } catch {
    return null;
  }
}

export function buildSharePointCookieHeader(cookies) {
  const byName = new Map((cookies || []).map((cookie) => [cookie.name, cookie.value]));
  const fedAuth = byName.get("FedAuth");
  const rtFa = byName.get("rtFa");

  if (!fedAuth || !rtFa) {
    return null;
  }

  return `FedAuth=${fedAuth}; rtFa=${rtFa}`;
}

export class DownloadArmory {
  constructor({ ttlMs = 90_000 } = {}) {
    this.ttlMs = ttlMs;
    this.tabs = new Map();
  }

  arm(tabId, filename, now = Date.now()) {
    this.tabs.set(Number(tabId), {
      filename: cleanFilename(filename),
      armedAt: now,
      claimed: false
    });
  }

  has(tabId, now = Date.now()) {
    const item = this.tabs.get(Number(tabId));
    if (!item) {
      return false;
    }
    if (now - item.armedAt > this.ttlMs) {
      this.tabs.delete(Number(tabId));
      return false;
    }
    return !item.claimed;
  }

  claim(tabId, now = Date.now()) {
    const numericTabId = Number(tabId);
    const item = this.tabs.get(numericTabId);
    if (!item) {
      return null;
    }
    if (now - item.armedAt > this.ttlMs || item.claimed) {
      this.tabs.delete(numericTabId);
      return null;
    }

    item.claimed = true;
    return { filename: item.filename };
  }

  clear(tabId) {
    this.tabs.delete(Number(tabId));
  }
}
