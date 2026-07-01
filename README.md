# SharePoint Opened PDF Downloader

Chromium extension for downloading the PDF currently opened in a SharePoint or
OneDrive viewer tab.

## Install

1. Open `chrome://extensions` or `chromium://extensions`.
2. Enable **Developer mode**.
3. Click **Load unpacked**.
4. Select this repo's `extension/` directory.
5. Install the local native helper:

```bash
cd /home/thinh/sp-automation
./native/install_native_host.sh
```

If Chrome asks for new permissions after updates, reload the extension from
`chrome://extensions` and run the installer again.

## Use

1. Log in to SharePoint in your normal Chromium profile.
2. Open the PDF you want in the SharePoint viewer.
3. Click the extension toolbar button once.
4. The extension reads the current SharePoint file URL, sends the current
   `FedAuth`/`rtFa` cookies to the local helper, and downloads one PDF to
   `~/Downloads`. The badge should end at `OK`.

Click the extension again for the next opened PDF. The extension is idle until
you click it, so repeated viewer network requests do not trigger repeated
downloads.

If the native helper is not installed, the extension falls back to the older
browser download/media-token path, which is less reliable for view-only shared
files.

## Legacy Python Reference

The previous Playwright/cURL implementation is preserved on the
`legacy-python-reference` branch.

## Development

Run the unit tests:

```bash
npm test
```
