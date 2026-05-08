# LL Integration Opera/Chromium Extension

Experimental Manifest V3 build for Opera and other Chromium-based browsers.

Stable Chromium extension ID for unpacked testing:

```text
ndnmgkboipaepgndebnikcnicechokln
```

The Windows installer registers the native host under the Chromium native messaging registry path with this origin:

```text
chrome-extension://ndnmgkboipaepgndebnikcnicechokln/
```

## Test Locally

1. Run `LLIntegrationInstaller.exe` or `LLIntegrationInstaller-WithToolbar.exe`.
2. Open Opera extension management.
3. Enable developer mode.
4. Load `opera-extension/` as an unpacked extension.
5. Open the popup and click `Check Status`.
6. Log into LoversLab in Opera and click `Export Cookies`.

If the Opera store assigns a different extension ID, update the installer constant before release.
