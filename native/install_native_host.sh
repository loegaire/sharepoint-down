#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
HOST_PATH="$ROOT_DIR/native/sp_downloader_host.py"
EXTENSION_PATH="$ROOT_DIR/extension"
HOST_NAME="com.sp_automation.downloader"

chmod +x "$HOST_PATH"

mapfile -t EXTENSION_IDS < <(python3 - "$EXTENSION_PATH" "$@" <<'PY'
import json
import sys
from pathlib import Path

extension_path = str(Path(sys.argv[1]).resolve())
explicit_ids = sys.argv[2:]
ids = set(explicit_ids)

preference_paths = [
    Path.home() / ".config/google-chrome/Default/Preferences",
    Path.home() / ".config/chromium/Default/Preferences",
]

for preferences_path in preference_paths:
    if not preferences_path.exists():
        continue
    try:
        data = json.loads(preferences_path.read_text())
    except Exception:
        continue
    settings = data.get("extensions", {}).get("settings", {})
    for extension_id, setting in settings.items():
        path = setting.get("path")
        if path and str(Path(path).resolve()) == extension_path:
            ids.add(extension_id)

for extension_id in sorted(ids):
    print(extension_id)
PY
)

if [ "${#EXTENSION_IDS[@]}" -eq 0 ]; then
  echo "No unpacked extension ID found."
  echo "Load the extension from $EXTENSION_PATH first, then rerun this script."
  echo "Or pass the extension ID explicitly: $0 <extension-id>"
  exit 1
fi

ALLOWED_ORIGINS=$(printf '%s\n' "${EXTENSION_IDS[@]}" | python3 -c 'import json,sys; print(json.dumps([f"chrome-extension://{line.strip()}/" for line in sys.stdin if line.strip()]))')

for BASE in "$HOME/.config/google-chrome" "$HOME/.config/chromium"; do
  MANIFEST_DIR="$BASE/NativeMessagingHosts"
  mkdir -p "$MANIFEST_DIR"
  python3 - "$HOST_NAME" "$HOST_PATH" "$ALLOWED_ORIGINS" > "$MANIFEST_DIR/$HOST_NAME.json" <<'PY'
import json
import sys

name, path, origins_json = sys.argv[1], sys.argv[2], sys.argv[3]
print(json.dumps({
    "name": name,
    "description": "SharePoint downloader native helper",
    "path": path,
    "type": "stdio",
    "allowed_origins": json.loads(origins_json),
}, indent=2))
PY
  echo "Installed $MANIFEST_DIR/$HOST_NAME.json"
done

echo "Allowed extension IDs: ${EXTENSION_IDS[*]}"
