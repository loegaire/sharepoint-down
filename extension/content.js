const FILE_LINK_SELECTORS = [
  'a[data-automationid="DetailsRowLink"]',
  'button[data-automationid="FieldRenderer-name"]',
  'span[data-automationid="FieldRenderer-name"]',
  'a[data-automationid="Tile-link"]',
  '[role="row"] a',
  '.ms-List-cell a'
];

const GENERIC_CLICK_SELECTORS = ["a", "button", '[role="row"]'];

const LAST_FILENAME_KEY = "spDownloader.lastFilename";

function bestText(element) {
  return element?.innerText?.trim() || element?.textContent?.trim() || "";
}

function titleFromDocument() {
  const title = document.title
    .replace(/\s*[-|]\s*(OneDrive|SharePoint|Microsoft 365|Office).*$/i, "")
    .replace(/\.(pdf)\s*[-|].*$/i, ".pdf")
    .trim();

  return title || "";
}

function titleFromUrl() {
  const url = new URL(location.href);
  const candidates = [
    url.searchParams.get("file"),
    url.searchParams.get("FileName"),
    url.searchParams.get("filename"),
    url.searchParams.get("SourceDoc")
  ];

  for (const candidate of candidates) {
    if (candidate && /\.pdf($|[?#])/i.test(candidate)) {
      return decodeURIComponent(candidate);
    }
  }

  const pathName = decodeURIComponent(url.pathname.split("/").pop() || "");
  return /\.pdf$/i.test(pathName) ? pathName : "";
}

function collectDocumentInfo() {
  const visibleTitleSelectors = [
    '[data-automationid="TitleText"]',
    '[data-automation-id="TitleText"]',
    '[aria-label$=".pdf"]',
    'button[title$=".pdf"]',
    'span[title$=".pdf"]'
  ];

  const visibleTitle = visibleTitleSelectors
    .map((selector) => document.querySelector(selector))
    .map(bestText)
    .find(Boolean);

  return {
    filename:
      visibleTitle ||
      sessionStorage.getItem(LAST_FILENAME_KEY) ||
      titleFromDocument() ||
      titleFromUrl() ||
      "sharepoint-document.pdf",
    url: location.href,
    title: document.title
  };
}

document.addEventListener(
  "click",
  (event) => {
    const clicked = event.target instanceof Element ? event.target : event.target?.parentElement;
    if (!clicked) {
      return;
    }

    for (const selector of FILE_LINK_SELECTORS) {
      const element = clicked.closest(selector);
      const text = bestText(element);
      if (text) {
        sessionStorage.setItem(LAST_FILENAME_KEY, text);
        return;
      }
    }

    for (const selector of GENERIC_CLICK_SELECTORS) {
      const element = clicked.closest(selector);
      const text = bestText(element);
      if (/\.pdf\b/i.test(text)) {
        sessionStorage.setItem(LAST_FILENAME_KEY, text);
        return;
      }
    }
  },
  true
);

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type !== "collect-document-info") {
    return false;
  }

  sendResponse(collectDocumentInfo());
  return true;
});
