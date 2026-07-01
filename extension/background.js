import {
  DownloadArmory,
  buildSharePointCookieHeader,
  buildSharePointDownload,
  chooseCaptureFilename,
  cleanFilename,
  getHeaderValue,
  isCaptureCandidate,
  isPassthroughRequest
} from "./core.js";

const armory = new DownloadArmory({ ttlMs: 90_000 });
const MEDIA_FILTER = { urls: ["https://mediap.svc.ms/transform/passthrough*"] };
const WEB_REQUEST_OPTIONS = ["requestHeaders", "extraHeaders"];
const DEBUGGER_PROTOCOL_VERSION = "1.3";

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

async function writeCapturedWithNativeHost({ filename, body, base64Encoded }) {
  return new Promise((resolve) => {
    chrome.runtime.sendNativeMessage(
      "com.sp_automation.downloader",
      { mode: "writeCaptured", filename, body, base64Encoded },
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

function debuggerCommand(tabId, method, params = {}) {
  return new Promise((resolve, reject) => {
    chrome.debugger.sendCommand({ tabId }, method, params, (result) => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message));
      } else {
        resolve(result);
      }
    });
  });
}

function attachDebugger(tabId) {
  return new Promise((resolve, reject) => {
    chrome.debugger.attach({ tabId }, DEBUGGER_PROTOCOL_VERSION, () => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message));
      } else {
        resolve();
      }
    });
  });
}

function detachDebugger(tabId) {
  return new Promise((resolve) => {
    chrome.debugger.detach({ tabId }, () => resolve());
  });
}

async function captureViewerResponse(tab, info) {
  const tabId = tab.id;
  const pendingResponses = new Map();
  let settled = false;

  await attachDebugger(tabId);

  return new Promise(async (resolve) => {
    const cleanup = async (result) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeoutId);
      chrome.debugger.onEvent.removeListener(onEvent);
      await detachDebugger(tabId);
      resolve(result);
    };

    const timeoutId = setTimeout(() => {
      cleanup({ ok: false, error: "Timed out waiting for a raw viewer response" });
    }, 90_000);

    const onEvent = async (source, method, params) => {
      if (settled || source.tabId !== tabId) {
        return;
      }

      if (method === "Network.responseReceived" && isCaptureCandidate(params.response)) {
        pendingResponses.set(params.requestId, params.response);
        return;
      }

      if (method !== "Network.loadingFinished" || !pendingResponses.has(params.requestId)) {
        return;
      }

      const response = pendingResponses.get(params.requestId);
      pendingResponses.delete(params.requestId);

      try {
        const body = await debuggerCommand(tabId, "Network.getResponseBody", {
          requestId: params.requestId
        });
        if (!body?.body) {
          return;
        }

        const filename = chooseCaptureFilename({
          pageUrl: info?.url || tab.url || "",
          responseUrl: response.url,
          fallbackTitle: info?.filename || tab.title || "sharepoint-document.pdf"
        });
        const nativeResult = await writeCapturedWithNativeHost({
          filename,
          body: body.body,
          base64Encoded: Boolean(body.base64Encoded)
        });

        if (nativeResult.ok) {
          await cleanup(nativeResult);
        } else {
          console.warn("Captured response was rejected:", nativeResult.error);
        }
      } catch (error) {
        console.warn("Could not read captured response body:", error.message);
      }
    };

    chrome.debugger.onEvent.addListener(onEvent);

    try {
      await debuggerCommand(tabId, "Network.enable", {
        maxTotalBufferSize: 100_000_000,
        maxResourceBufferSize: 100_000_000
      });
      await debuggerCommand(tabId, "Page.enable");
      await debuggerCommand(tabId, "Page.reload", { ignoreCache: false });
    } catch (error) {
      await cleanup({ ok: false, error: error.message });
    }
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

chrome.action.onClicked.addListener(async (tab) => {
  if (!tab?.id) {
    return;
  }

  try {
    const info = await collectDocumentInfo(tab);
    const filename = cleanFilename(info?.filename || tab.title);
    const directDownload = buildSharePointDownload(info?.url || tab.url || "");

    setBadge(tab.id, "CAP", "#7c3aed");
    const captureResult = await captureViewerResponse(tab, info);
    if (captureResult.ok) {
      setBadge(tab.id, "OK", "#16a34a");
      clearBadgeSoon(tab.id);
      return;
    }

    console.warn("SharePoint debugger capture failed:", captureResult.error);

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

      console.warn("SharePoint native direct download failed:", nativeResult.error);
      setBadge(tab.id, "ERR", "#dc2626");
      clearBadgeSoon(tab.id);
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
