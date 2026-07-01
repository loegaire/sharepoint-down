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

export function buildBrowserDownloadOptions({ url, filename, headers = [] }) {
  const options = {
    url,
    filename: cleanFilename(filename),
    conflictAction: "uniquify"
  };

  if (headers.length > 0) {
    options.headers = headers;
  }

  return options;
}

const REPLAYABLE_DOWNLOAD_HEADERS = new Map([
  ["authorization", "Authorization"],
  ["x-spopactoken", "X-SPOPacToken"]
]);

export function pickBrowserDownloadHeaders(headers) {
  const entries = Array.isArray(headers)
    ? headers.map((header) => [header?.name, header?.value])
    : Object.entries(headers || {});
  const result = [];
  const seen = new Set();

  for (const [rawName, rawValue] of entries) {
    const lowerName = String(rawName || "").toLowerCase();
    const canonicalName = REPLAYABLE_DOWNLOAD_HEADERS.get(lowerName);
    if (!canonicalName || seen.has(lowerName) || rawValue == null || rawValue === "") {
      continue;
    }

    seen.add(lowerName);
    result.push({ name: canonicalName, value: String(rawValue) });
  }

  return result;
}

function base64FromText(value) {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(String(value), "utf8").toString("base64");
  }

  const bytes = new TextEncoder().encode(String(value));
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary);
}

function textPrefixFromBase64(value, maxBytes = 512) {
  const compact = String(value).replace(/\s+/g, "");
  const prefix = compact.slice(0, Math.ceil(maxBytes / 3) * 4);

  if (typeof Buffer !== "undefined") {
    return Buffer.from(prefix, "base64").toString("utf8");
  }

  const binary = atob(prefix);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
}

function looksLikeHtmlPayload(payload) {
  const prefix = String(payload || "").slice(0, 256).trimStart().toLowerCase();
  return prefix.startsWith("<html") ||
    prefix.startsWith("<!doctype") ||
    prefix.startsWith("<script");
}

export function buildCapturedDownloadUrl({ body, base64Encoded, mimeType = "application/pdf" }) {
  if (!body) {
    throw new Error("Captured body is empty");
  }

  if (base64Encoded) {
    if (looksLikeHtmlPayload(textPrefixFromBase64(body))) {
      throw new Error("Captured response is HTML, not the raw file");
    }
    return `data:${mimeType};base64,${String(body).replace(/\s+/g, "")}`;
  }

  if (looksLikeHtmlPayload(body)) {
    throw new Error("Captured response is HTML, not the raw file");
  }
  return `data:${mimeType};base64,${base64FromText(body)}`;
}

function lowerHeaderMap(headers) {
  const result = {};
  for (const [name, value] of Object.entries(headers || {})) {
    result[name.toLowerCase()] = String(value || "").toLowerCase();
  }
  return result;
}

export function isCaptureCandidate(response) {
  const url = String(response?.url || "").toLowerCase();
  const mimeType = String(response?.mimeType || "").toLowerCase();
  const headers = lowerHeaderMap(response?.headers);
  const contentType = headers["content-type"] || mimeType;

  if (contentType.includes("text/html") || contentType.includes("application/json")) {
    return false;
  }

  return (
    mimeType.includes("application/pdf") ||
    contentType.includes("application/pdf") ||
    url.includes("mediap.svc.ms/transform/passthrough") ||
    url.includes("/_layouts/15/download.aspx") ||
    url.includes("/_api/") && url.includes("$value")
  );
}

function filenameFromUrlParam(url) {
  try {
    const rawSourcePath = getRawQueryParam(url, ["id", "SourceUrl"]);
    if (!rawSourcePath) {
      return null;
    }
    const decodedSourcePath = decodeQueryValue(rawSourcePath);
    if (!/\.pdf$/i.test(decodedSourcePath)) {
      return null;
    }
    return cleanFilename(decodedSourcePath.split("/").pop());
  } catch {
    return null;
  }
}

function filenameFromPath(url) {
  try {
    const parsed = new URL(url);
    const decoded = decodeURIComponent(parsed.pathname.split("/").pop() || "");
    return /\.pdf$/i.test(decoded) ? cleanFilename(decoded) : null;
  } catch {
    return null;
  }
}

export function chooseCaptureFilename({ pageUrl, responseUrl, fallbackTitle }) {
  return (
    filenameFromUrlParam(pageUrl) ||
    filenameFromUrlParam(responseUrl) ||
    filenameFromPath(responseUrl) ||
    cleanFilename(fallbackTitle)
  );
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
