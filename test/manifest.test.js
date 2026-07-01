import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const ROOT = dirname(dirname(fileURLToPath(import.meta.url)));
const manifest = JSON.parse(readFileSync(join(ROOT, "extension", "manifest.json"), "utf8"));

test("manifest wires generated icon assets", () => {
  const expected = {
    16: "icons/icon16.png",
    32: "icons/icon32.png",
    48: "icons/icon48.png",
    128: "icons/icon128.png"
  };

  assert.deepEqual(manifest.icons, expected);
  assert.deepEqual(manifest.action.default_icon, expected);

  for (const iconPath of Object.values(expected)) {
    assert.equal(existsSync(join(ROOT, "extension", iconPath)), true, iconPath);
  }
});
