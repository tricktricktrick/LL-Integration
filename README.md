# LL Integration

Alpha tooling for connecting Firefox or Opera/Chromium downloads from LoversLab and selected external mod pages to Mod Organizer 2.

## What It Does

- Exports LoversLab cookies from Firefox or Opera/Chromium through Native Messaging.
- Detects LoversLab and Dwemer Mods download clicks and writes source metadata.
- Captures manually armed external archive downloads from pages such as Patreon.
- Optionally shows floating desktop capture controls for Arm, Disarm, and Follow mode.
- Copies supported archives into the configured MO2 downloads folder.
- Adds MO2 tools for managing source links, opening source pages, editing manual links, and checking updates.
- Stores installed mod source links in MO2 `meta.ini` under `[LoversLab]`, so the metadata does not appear as a mod file conflict.
- When installed into multiple MO2 instances, the active instance keeps the browser bridge downloads path in sync.

Supported archive extensions are `.7z`, `.zip`, and `.rar`.

## Components

- `firefox-extension/`: Firefox extension popup, LoversLab click detection, external capture mode.
- `opera-extension/`: experimental Opera/Chromium Manifest V3 extension.
- `native-app/`: Python Native Messaging bridge used by the browser extensions.
- `native-app/overlay.py`: optional floating capture controls launched by the native bridge.
- `mo2-plugin/`: Python MO2 plugin.
- `installer.py`: Windows GUI installer for the native bridge and MO2 plugin.

## Alpha Install

1. Run `installer.py`.
2. Select `ModOrganizer.exe`.
3. Select the MO2 downloads folder.
4. Enable optional floating capture controls if you want a topmost Arm / Disarm / Follow window.
5. Install the Firefox extension from `firefox-extension/` or load the experimental Opera extension from `opera-extension/`.
6. Restart the browser and MO2.
7. In the browser, log into LoversLab and click `Export Cookies`.
8. In MO2, open `Tools > LL Integration`.

The installer writes the native bridge to `%LOCALAPPDATA%\LLIntegration` and registers Firefox plus Chromium Native Messaging manifests for the current Windows user.

For portable or multi-game MO2 setups, run the installer once for each MO2 instance where you want the plugin. After that, opening or switching to an instance makes it the active target for new browser downloads.

## Firefox Workflow

### LoversLab And Dwemer Mods

1. Open a LoversLab file page or a Dwemer Mods mod page.
2. Click a supported download.
3. Let Firefox complete the archive download.
4. The native bridge writes `.ll.ini` and `.ll.json` sidecars, copies the archive to MO2 downloads, and the MO2 plugin stores the installed mod link in `meta.ini`.

Dwemer Mods downloads are captured through the browser so Cloudflare can stay in the normal page flow. They are stored as fixed/manual links and skipped by automatic LoversLab update checks.

### External Pages

For Patreon, SubscribeStar, Mega, Google Drive, or similar pages:

1. Open the source page.
2. Click `Capture Archives From This Page`.
3. Download one or more archives within the capture window.
4. Click `Cancel Capture` when finished.

External captures are marked as manual/fixed links, so MO2 will not try to fetch update data automatically.

### Floating Capture Controls

If enabled in the installer, the browser popup can open a small topmost desktop window with Arm, Disarm, and Follow controls. Arm follows the active browser tab by default; Follow Off freezes capture on the current target. Closing the floating window disarms capture and keeps the overlay process cached briefly for a faster reopen. The floating window does not read pages, capture input, or monitor the desktop. It only sends explicit capture commands through the installed native bridge; the browser extension still handles tabs and download events.

### Voice Finder

`Find Voice Packs` scans installed LoversLab-linked mods as base mods, compares them with all installed MO2 mods that look like DBVO/DVO/voice packs, and marks matches as installed, possible, missing, or ignored. Rows and candidate lists are color-coded by confidence. Use `Classify` when a mod is detected in the wrong role, for example a `VoiceFiles` mod shown as a base mod; forced base/voice choices are persisted. Double-click any cell in a row to inspect every local voice-like mod in full, sorted by score, filter the candidate list, open the selected mod folder, or fix the selected pack as a manual match for that base mod. You can add LoversLab voice source URLs; normal file-page URLs are converted to `?do=download` automatically. Fetching those pages reads every download entry, scores likely archives, keeps a searchable `All downloads` list for manual picks, and opens a download window that stays visible while the selected archive downloads into the MO2 downloads folder. `False local match` only hides installed-mod comparisons; online false matches are hidden from the online candidate dialog. `Voice mods` shows a filterable inventory of local voice-like mods sorted by name or install date. Manual matches, classification overrides, false matches, and ignored base mods are saved in the MO2 plugin data folder; use `False matches` to remove accidental blacklists.

## MO2 Workflow

Open `Tools > LL Integration`.

- `Manage LoversLab Links`: list linked mods, open source pages, edit metadata, purge bad links, and optionally fetch update information.
- `Create Manual Link`: create source metadata manually for multipart archives or manually installed mods.
- `Purge Suspicious Links`: clean accidental LoversLab metadata from Nexus-identified mods.
- `Integration Paths`: show install paths and the update timing log.

`Fetch Updates` is manual on purpose. It uses pacing to avoid hammering LoversLab.

## Safety Notes

- Do not commit real cookies, generated `.ll.ini` sidecars, download events, metadata, or timing logs.
- The repo `.gitignore` excludes runtime storage and generated files.
- `native-app/config.json` and `native-app/manifest.json` are generated by the installer and intentionally ignored.

## Privacy And License

- Privacy policy: see `PRIVACY.md`.
- License: source-available, all rights reserved. See `LICENSE.md`.

## Current Limitations

- Update checks depend on a valid LoversLab session cookie export.
- Some LoversLab pages use non-standard version names or separate full/patch files. Use `Edit` and `Download file pattern` for those cases.
- External sites are tracked as manual source links; update checking is not attempted automatically.
- The MO2 toolbar icon is not pinned directly; the plugin is exposed through `Tools > LL Integration`.
