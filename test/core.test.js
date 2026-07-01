import assert from "node:assert/strict";
import test from "node:test";

import {
  DownloadArmory,
  cleanFilename,
  getHeaderValue,
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
