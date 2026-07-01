# SharePoint Opened PDF Downloader

Chromium extension for downloading the PDF currently opened in a SharePoint or
OneDrive viewer tab.

## Install

1. Open `chrome://extensions` or `chromium://extensions`.
2. Enable **Developer mode**.
3. Click **Load unpacked**.
4. Select this repo's `extension/` directory.

## Use

1. Log in to SharePoint in your normal Chromium profile.
2. Open the PDF you want in the SharePoint viewer.
3. Click the extension toolbar button once.
4. The extension reloads the current tab, captures the viewer's media token, and
   downloads one PDF. The tab badge moves through `ON`, `DL`, then `OK`.

Click the extension again for the next opened PDF. The extension is idle until
you click it, so repeated viewer network requests do not trigger repeated
downloads.

## Legacy Python Reference

The previous Playwright/cURL implementation is preserved on the
`legacy-python-reference` branch.

## Development

Run the unit tests:

```bash
npm test
```
