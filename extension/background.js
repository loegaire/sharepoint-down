import {
  DownloadArmory,
  buildSharePointCookieHeader,
  buildSharePointDownload,
  cleanFilename,
  getHeaderValue,
  isPassthroughRequest
} from "./core.js";

const armory = new DownloadArmory({ ttlMs: 90_000 });
const MEDIA_FILTER = { urls: ["https://mediap.svc.ms/transform/passthrough*"] };
const WEB_REQUEST_OPTIONS = ["requestHeaders", "extraHeaders"];

function callbackToPromise(fn) {
  return new Promise((resolve, reject) => {
    fn((result) => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message));
      } else {
        resolve(result);
      }
    });
  });
}

function setBadge(tabId, text, color) {
  chrome.action.setBadgeText({ tabId, text });
  chrome.action.setBadgeBackgroundColor({ tabId, color });
}

function clearBadgeSoon(tabId) {
  setTimeout(() => chrome.action.setBadgeText({ tabId, text: "" }), 3000);
}

async function ensureContentScript(tabId) {
  try {
    await chrome.tabs.sendMessage(tabId, { type: "collect-document-info" });
  } catch {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"]
    });
  }
}

async function collectDocumentInfo(tab) {
  await ensureContentScript(tab.id);
  try {
    return await chrome.tabs.sendMessage(tab.id, { type: "collect-document-info" });
  } catch {
    return {
      filename: tab.title || "sharepoint-document.pdf",
      url: tab.url || "",
      title: tab.title || ""
    };
  }
}

async function reloadTab(tabId) {
  await callbackToPromise((done) => chrome.tabs.reload(tabId, { bypassCache: true }, done));
}

async function getSharePointCookieHeader(url) {
  try {
    const cookies = await chrome.cookies.getAll({ url });
    return buildSharePointCookieHeader(cookies);
  } catch (error) {
    console.warn("SharePoint downloader could not read auth cookies:", error.message);
    return null;
  }
}

async function downloadWithNativeHost({ url, filename, cookieHeader }) {
  if (!cookieHeader) {
    return { ok: false, error: "SharePoint auth cookies are not available" };
  }

  return new Promise((resolve) => {
    chrome.runtime.sendNativeMessage(
      "com.sp_automation.downloader",
      { url, filename, cookieHeader },
      (response) => {
        const error = chrome.runtime.lastError;
        if (error) {
          resolve({ ok: false, error: error.message });
          return;
        }
        resolve(response || { ok: false, error: "Native helper returned no response" });
      }
    );
  });
}

function downloadWithToken({ tabId, url, token, filename }) {
  chrome.downloads.download(
    {
      url,
      filename,
      conflictAction: "uniquify",
      saveAs: false,
      headers: [{ name: "X-SPOPacToken", value: token }]
    },
    (downloadId) => {
      const error = chrome.runtime.lastError;
      if (error) {
        console.error("SharePoint downloader failed:", error.message);
        setBadge(tabId, "ERR", "#dc2626");
        clearBadgeSoon(tabId);
        return;
      }

      console.log(`SharePoint downloader saved ${filename} as download ${downloadId}`);
      setBadge(tabId, "OK", "#16a34a");
      clearBadgeSoon(tabId);
    }
  );
}

function downloadDirect({ tabId, url, filename }) {
  chrome.downloads.download(
    {
      url,
      filename,
      conflictAction: "uniquify",
      saveAs: false
    },
    (downloadId) => {
      const error = chrome.runtime.lastError;
      if (error) {
        console.error("SharePoint direct download failed:", error.message);
        setBadge(tabId, "ERR", "#dc2626");
        clearBadgeSoon(tabId);
        return;
      }

      console.log(`SharePoint direct download saved ${filename} as download ${downloadId}`);
      setBadge(tabId, "OK", "#16a34a");
      clearBadgeSoon(tabId);
    }
  );
}

chrome.action.onClicked.addListener(async (tab) => {
  if (!tab?.id) {
    return;
  }

  try {
    const info = await collectDocumentInfo(tab);
    const filename = cleanFilename(info?.filename || tab.title);
    const directDownload = buildSharePointDownload(info?.url || tab.url || "");

    if (directDownload) {
      const cookieHeader = await getSharePointCookieHeader(directDownload.url);
      const nativeResult = await downloadWithNativeHost({
        url: directDownload.url,
        filename: directDownload.filename,
        cookieHeader
      });

      if (nativeResult.ok) {
        setBadge(tab.id, "OK", "#16a34a");
        clearBadgeSoon(tab.id);
        return;
      }

      console.warn("SharePoint native download failed, falling back to browser download:", nativeResult.error);
      setBadge(tab.id, "DL", "#0f766e");
      downloadDirect({
        tabId: tab.id,
        url: directDownload.url,
        filename: directDownload.filename
      });
      return;
    }

    armory.arm(tab.id, filename);
    setBadge(tab.id, "ON", "#2563eb");
    await reloadTab(tab.id);
  } catch (error) {
    console.error("SharePoint downloader could not arm this tab:", error);
    setBadge(tab.id, "ERR", "#dc2626");
    clearBadgeSoon(tab.id);
  }
});

chrome.webRequest.onBeforeSendHeaders.addListener(
  (details) => {
    if (details.tabId < 0 || !armory.has(details.tabId) || !isPassthroughRequest(details.url)) {
      return;
    }

    const token = getHeaderValue(details.requestHeaders, "X-SPOPacToken");
    if (!token) {
      return;
    }

    const claim = armory.claim(details.tabId);
    if (!claim) {
      return;
    }

    armory.clear(details.tabId);
    setBadge(details.tabId, "DL", "#0f766e");
    downloadWithToken({
      tabId: details.tabId,
      url: details.url,
      token,
      filename: claim.filename
    });
  },
  MEDIA_FILTER,
  WEB_REQUEST_OPTIONS
);
