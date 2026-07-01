#!/usr/bin/env python3
import json
import os
import re
import struct
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path
from urllib.parse import urlparse


HOST_SUFFIX = ".sharepoint.com"
HOST_NAME = "com.sp_automation.downloader"
FALLBACK_NAME = "sharepoint-document.pdf"


def sanitize_filename(filename):
    value = str(filename or "").strip().replace("\r", "").replace("\n", "")
    value = value.replace("/", "").replace("\\", "")
    value = re.sub(r'[\\/*?:"<>|]', "", value)
    value = re.sub(r"\s+", "_", value)
    value = value.lstrip(".")
    return value or FALLBACK_NAME


def validate_message(message):
    url = str(message.get("url") or "")
    cookie_header = str(message.get("cookieHeader") or "")
    parsed = urlparse(url)

    if parsed.scheme != "https" or not parsed.hostname or not parsed.hostname.endswith(HOST_SUFFIX):
        raise ValueError("URL must be an HTTPS SharePoint URL")
    if "/_layouts/15/download.aspx" not in parsed.path:
        raise ValueError("URL must target SharePoint download.aspx")
    if "SourceUrl=" not in parsed.query:
        raise ValueError("URL must include SourceUrl")
    if "FedAuth=" not in cookie_header or "rtFa=" not in cookie_header:
        raise ValueError("Missing SharePoint auth cookies")


def read_message():
    raw_length = sys.stdin.buffer.read(4)
    if not raw_length:
        return None
    message_length = struct.unpack("<I", raw_length)[0]
    return json.loads(sys.stdin.buffer.read(message_length).decode("utf-8"))


def write_message(payload):
    encoded = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    sys.stdout.buffer.write(struct.pack("<I", len(encoded)))
    sys.stdout.buffer.write(encoded)
    sys.stdout.buffer.flush()


def reject_html(payload):
    prefix = payload[:256].lstrip().lower()
    return prefix.startswith(b"<html") or prefix.startswith(b"<!doctype") or prefix.startswith(b"<script")


def download_file(message):
    validate_message(message)

    filename = sanitize_filename(message.get("filename"))
    downloads_dir = Path.home() / "Downloads"
    downloads_dir.mkdir(parents=True, exist_ok=True)
    destination = downloads_dir / filename

    headers = {
        "Cookie": message["cookieHeader"],
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/126 Safari/537.36",
    }
    request = urllib.request.Request(message["url"], headers=headers, method="GET")

    with urllib.request.urlopen(request, timeout=120) as response:
        content_type = response.headers.get("content-type", "")
        payload = response.read()

    if not payload:
        raise ValueError("SharePoint returned an empty file")
    if content_type.startswith("text/html") or reject_html(payload):
        raise ValueError("SharePoint returned HTML instead of a document")

    with tempfile.NamedTemporaryFile("wb", delete=False, dir=str(downloads_dir), prefix=".sp-download-") as handle:
        handle.write(payload)
        temp_path = Path(handle.name)

    os.replace(temp_path, destination)
    return {"ok": True, "path": str(destination), "bytes": len(payload)}


def main():
    try:
        message = read_message()
        if message is None:
            return
        write_message(download_file(message))
    except urllib.error.HTTPError as error:
        write_message({"ok": False, "error": f"HTTP {error.code}"})
    except Exception as error:
        write_message({"ok": False, "error": str(error)})


if __name__ == "__main__":
    main()
