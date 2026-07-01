import assert from "node:assert/strict";
import test from "node:test";

import {
  DownloadArmory,
  buildSharePointCookieHeader,
  buildSharePointDownload,
  chooseCaptureFilename,
  cleanFilename,
  getHeaderValue,
  isCaptureCandidate,
  isPassthroughRequest
} from "../extension/core.js";

test("recognizes SharePoint media passthrough requests", () => {
  assert.equal(
    isPassthroughRequest("https://mediap.svc.ms/transform/passthrough?docid=abc"),
    true
  );
  assert.equal(isPassthroughRequest("https://example.com/transform/passthrough"), false);
  assert.equal(isPassthroughRequest("not a url"), false);
});

test("cleans filenames for Chromium downloads", () => {
  assert.equal(cleanFilename("  Lecture 01 / Intro.pdf\n"), "Lecture_01_Intro.pdf");
  assert.equal(cleanFilename("No extension"), "No_extension.pdf");
  assert.equal(cleanFilename(""), "sharepoint-document.pdf");
  assert.equal(cleanFilename(':"<>|'), "sharepoint-document.pdf");
});

test("finds request headers case-insensitively", () => {
  const headers = [
    { name: "Accept", value: "*/*" },
    { name: "x-spopactoken", value: "token-123" }
  ];

  assert.equal(getHeaderValue(headers, "X-SPOPacToken"), "token-123");
  assert.equal(getHeaderValue(headers, "Authorization"), null);
});

test("builds SharePoint direct download URL from opened PDF id parameter", () => {
  const openedUrl = "https://husteduvn-my.sharepoint.com/shared?listurl=https%3A%2F%2Fhusteduvn-my.sharepoint.com%2Fpersonal%2Fdat_tt194246_sis_hust_edu_vn%2FDocuments&id=%2Fpersonal%2Fdat_tt194246_sis_hust_edu_vn%2FDocuments%2F%C3%94n%20thi%20cu%E1%BB%91i%20k%C3%AC%2FDatabase%20-%20cu%E1%BB%91i%20k%C3%AC%2020191.pdf&parent=%2Fpersonal%2Fdat_tt194246_sis_hust_edu_vn%2FDocuments";

  assert.deepEqual(buildSharePointDownload(openedUrl), {
    url: "https://husteduvn-my.sharepoint.com/personal/dat_tt194246_sis_hust_edu_vn/_layouts/15/download.aspx?SourceUrl=%2Fpersonal%2Fdat_tt194246_sis_hust_edu_vn%2FDocuments%2F%C3%94n%20thi%20cu%E1%BB%91i%20k%C3%AC%2FDatabase%20-%20cu%E1%BB%91i%20k%C3%AC%2020191.pdf",
    filename: "Database_-_cuối_kì_20191.pdf"
  });
});

test("does not build direct download URL for folder ids", () => {
  const folderUrl = "https://husteduvn-my.sharepoint.com/personal/dat_tt194246_sis_hust_edu_vn/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fdat_tt194246_sis_hust_edu_vn%2FDocuments%2F%C3%94n%20thi%20cu%E1%BB%91i%20k%C3%AC&FolderCTID=0x012000";

  assert.equal(buildSharePointDownload(folderUrl), null);
});

test("builds cookie header from SharePoint auth cookies only", () => {
  const cookies = [
    { name: "SIMI", value: "ignored" },
    { name: "FedAuth", value: "fed" },
    { name: "rtFa", value: "rt" }
  ];

  assert.equal(buildSharePointCookieHeader(cookies), "FedAuth=fed; rtFa=rt");
  assert.equal(buildSharePointCookieHeader([{ name: "FedAuth", value: "fed" }]), null);
});

test("identifies raw viewer response candidates", () => {
  assert.equal(isCaptureCandidate({
    url: "https://mediap.svc.ms/transform/passthrough?docid=x",
    mimeType: "application/octet-stream",
    headers: {}
  }), true);
  assert.equal(isCaptureCandidate({
    url: "https://husteduvn-my.sharepoint.com/anything",
    mimeType: "application/pdf",
    headers: {}
  }), true);
  assert.equal(isCaptureCandidate({
    url: "https://husteduvn-my.sharepoint.com/error",
    mimeType: "text/html",
    headers: { "content-type": "text/html" }
  }), false);
});

test("chooses capture filename from current URL before generic response URL", () => {
  const pageUrl = "https://husteduvn-my.sharepoint.com/shared?id=%2Fpersonal%2Fdat%2FDocuments%2FDatabase%20-%20cu%E1%BB%91i%20k%C3%AC%2020191.pdf";
  const responseUrl = "https://mediap.svc.ms/transform/passthrough?docid=x";

  assert.equal(
    chooseCaptureFilename({ pageUrl, responseUrl, fallbackTitle: "Viewer" }),
    "Database_-_cuối_kì_20191.pdf"
  );
});

test("allows one download claim per extension click", () => {
  const armory = new DownloadArmory({ ttlMs: 60_000 });

  armory.arm(7, "First File.pdf", 1_000);

  assert.deepEqual(armory.claim(7, 1_100), { filename: "First_File.pdf" });
  assert.equal(armory.claim(7, 1_200), null);

  armory.arm(7, "Second File.pdf", 2_000);
  assert.deepEqual(armory.claim(7, 2_100), { filename: "Second_File.pdf" });
});

test("expires stale armed tabs", () => {
  const armory = new DownloadArmory({ ttlMs: 500 });

  armory.arm(3, "Old File.pdf", 1_000);

  assert.equal(armory.claim(3, 1_600), null);
  assert.equal(armory.has(3, 1_600), false);
});
