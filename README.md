# SharePoint Opened PDF Downloader

Chromium extension for downloading the PDF currently opened in a SharePoint or
OneDrive viewer tab.

## Demo

[Watch the video demo](./2026-07-01%2004-17-27.mp4)

## Install

1. Open `chrome://extensions` or `chromium://extensions`.
2. Enable **Developer mode**.
3. Click **Load unpacked**.
4. Select this repo's `extension/` directory.
5. If Chrome asks for new permissions after updates, reload the extension from
   `chrome://extensions` and accept the prompt.

The current version must request `debugger`, `downloads`, and `webRequest`.
It does not need the native helper for normal downloads.

## Use

1. Log in to SharePoint in your normal Chromium profile.
2. Open the PDF you want in the SharePoint viewer.
3. Click the extension toolbar button once.
4. The extension attaches to the current tab, reloads the viewer, captures the
   PDF response or its replayable viewer URL, and starts one Chrome browser
   download. The badge should move through `CAP`, then `DL`, and end at `OK`.

Chrome owns the final save location. If Chrome is configured to ask where to
save each file, the normal file chooser opens. Otherwise Chrome saves to the
profile's configured download directory and shows progress in the browser
download UI.

Click the extension again for the next opened PDF. The extension is idle until
you click it, so repeated viewer network requests do not trigger repeated
downloads.

If raw capture fails, the extension tries the older SharePoint `download.aspx`
URL through Chrome's download manager.

## Troubleshooting

If clicking the extension appears to do nothing:

1. Open `chrome://extensions`.
2. Find **SharePoint Opened PDF Downloader**.
3. Click the circular reload button on that extension card.
4. Accept the new permissions.
5. Make sure the extension version shown by Chrome is `0.4.0` or newer.

If the extension was not reloaded after version `0.4.0`, Chrome may still be
running older code that writes directly to `~/Downloads`.

## Legacy Python Reference

The previous Playwright/cURL implementation is preserved on the
`legacy-python-reference` branch.

## Development

Run the unit tests:

```bash
npm test
```
