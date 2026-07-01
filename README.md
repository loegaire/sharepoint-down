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

The current version must request `cookies` and `nativeMessaging`. If those do
not appear in Chrome's permission prompt, the old extension is still loaded.

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

## Troubleshooting

If clicking the extension appears to do nothing:

1. Open `chrome://extensions`.
2. Find **SharePoint Opened PDF Downloader**.
3. Click the circular reload button on that extension card.
4. Accept the new permissions.
5. Run `./native/install_native_host.sh` again.

If the extension was not reloaded after version `0.2.0`, Chrome will still be
running the old code that cannot call the native helper.

## Legacy Python Reference

The previous Playwright/cURL implementation is preserved on the
`legacy-python-reference` branch.

## Development

Run the unit tests:

```bash
npm test
```
