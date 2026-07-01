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
