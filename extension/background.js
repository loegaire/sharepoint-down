import {
  DownloadArmory,
  buildBrowserDownloadOptions,
  buildCapturedDownloadUrl,
  buildSharePointDownload,
  chooseCaptureFilename,
  cleanFilename,
  getHeaderValue,
  isCaptureCandidate,
  isPassthroughRequest,
  pickBrowserDownloadHeaders
} from "./core.js";

const armory = new DownloadArmory({ ttlMs: 90_000 });
const MEDIA_FILTER = { urls: ["https://mediap.svc.ms/transform/passthrough*"] };
const WEB_REQUEST_OPTIONS = ["requestHeaders", "extraHeaders"];
const DEBUGGER_PROTOCOL_VERSION = "1.3";
const activeDownloads = new Map();

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

function rememberDownload(tabId, downloadId, filename) {
  activeDownloads.set(downloadId, { tabId, filename });
  setBadge(tabId, "DL", "#0f766e");
}

async function startBrowserDownload({ tabId, url, filename, headers = [] }) {
  const options = buildBrowserDownloadOptions({ url, filename, headers });

  return new Promise((resolve, reject) => {
    chrome.downloads.download(options, (downloadId) => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message));
        return;
      }

      rememberDownload(tabId, downloadId, options.filename);
      console.log(`SharePoint downloader started browser download ${downloadId} for ${options.filename}`);
      resolve({ ok: true, downloadId, filename: options.filename });
    });
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

function mergeRequestHeaders(store, requestId, headers) {
  const current = store.get(requestId) || {};
  for (const [name, value] of Object.entries(headers || {})) {
    current[name] = value;
  }
  store.set(requestId, current);
}

function captureFilename(tab, info, response) {
  return chooseCaptureFilename({
    pageUrl: info?.url || tab.url || "",
    responseUrl: response.url,
    fallbackTitle: info?.filename || tab.title || "sharepoint-document.pdf"
  });
}

function needsReplayHeader(response) {
  return isPassthroughRequest(response?.url || "");
}

async function captureViewerResponse(tab, info) {
  const tabId = tab.id;
  const pendingResponses = new Map();
  const requestHeaders = new Map();
  const streamAttempts = new Set();
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

    const attemptBrowserStream = async (requestId) => {
      if (settled || streamAttempts.has(requestId) || !pendingResponses.has(requestId)) {
        return false;
      }

      const response = pendingResponses.get(requestId);
      const headers = pickBrowserDownloadHeaders(requestHeaders.get(requestId));
      if (needsReplayHeader(response) && headers.length === 0) {
        return false;
      }

      streamAttempts.add(requestId);
      try {
        const result = await startBrowserDownload({
          tabId,
          url: response.url,
          filename: captureFilename(tab, info, response),
          headers
        });
        await cleanup(result);
        return true;
      } catch (error) {
        console.warn("Could not start browser stream for captured response:", error.message);
        return false;
      }
    };

    const onEvent = async (source, method, params) => {
      if (settled || source.tabId !== tabId) {
        return;
      }

      if (method === "Network.requestWillBeSent") {
        mergeRequestHeaders(requestHeaders, params.requestId, params.request?.headers);
        return;
      }

      if (method === "Network.requestWillBeSentExtraInfo") {
        mergeRequestHeaders(requestHeaders, params.requestId, params.headers);
        await attemptBrowserStream(params.requestId);
        return;
      }

      if (method === "Network.responseReceived" && isCaptureCandidate(params.response)) {
        pendingResponses.set(params.requestId, params.response);
        await attemptBrowserStream(params.requestId);
        return;
      }

      if (method !== "Network.loadingFinished" || !pendingResponses.has(params.requestId)) {
        return;
      }

      if (await attemptBrowserStream(params.requestId)) {
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

        const downloadUrl = buildCapturedDownloadUrl({
          body: body.body,
          base64Encoded: Boolean(body.base64Encoded),
          mimeType: response.mimeType || "application/pdf"
        });
        const result = await startBrowserDownload({
          tabId,
          url: downloadUrl,
          filename: captureFilename(tab, info, response)
        });
        await cleanup(result);
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

async function downloadWithToken({ tabId, url, token, filename }) {
  try {
    await startBrowserDownload({
      tabId,
      url,
      filename,
      headers: [{ name: "X-SPOPacToken", value: token }]
    });
  } catch (error) {
    console.error("SharePoint downloader failed:", error.message);
    setBadge(tabId, "ERR", "#dc2626");
    clearBadgeSoon(tabId);
  }
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
      return;
    }

    console.warn("SharePoint debugger capture failed:", captureResult.error);

    if (directDownload) {
      try {
        await startBrowserDownload({
          tabId: tab.id,
          url: directDownload.url,
          filename: directDownload.filename
        });
        return;
      } catch (error) {
        console.warn("SharePoint browser direct download failed:", error.message);
        setBadge(tab.id, "ERR", "#dc2626");
        clearBadgeSoon(tab.id);
        return;
      }
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

chrome.downloads.onChanged.addListener((delta) => {
  const item = activeDownloads.get(delta.id);
  if (!item) {
    return;
  }

  if (delta.error?.current || delta.state?.current === "interrupted") {
    console.warn(
      `SharePoint downloader interrupted ${item.filename}: ${delta.error?.current || "interrupted"}`
    );
    activeDownloads.delete(delta.id);
    setBadge(item.tabId, "ERR", "#dc2626");
    clearBadgeSoon(item.tabId);
    return;
  }

  if (delta.state?.current === "complete") {
    activeDownloads.delete(delta.id);
    setBadge(item.tabId, "OK", "#16a34a");
    clearBadgeSoon(item.tabId);
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
