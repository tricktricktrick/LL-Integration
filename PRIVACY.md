# Privacy Policy

Last updated: May 8, 2026

LL Integration is a local companion tool for connecting supported browser downloads with Mod Organizer 2.

This project does not operate a server, does not use analytics, does not sell data, and does not send collected data to the developer.

## Data Collected Locally

LL Integration may store the following data on your own computer:

- LoversLab session cookies exported through the browser extension.
- LoversLab file page URLs and download URLs.
- Downloaded archive names, sizes, timestamps, and quick file hashes.
- Source metadata used to link downloaded archives to installed MO2 mods.
- Local paths for the native bridge, MO2 plugin, MO2 downloads folder, and installed mod folders.
- Update check cache and timing information used by the MO2 plugin.

This data is used only to:

- authenticate update checks against LoversLab from your local machine;
- copy supported archives into your configured MO2 downloads folder;
- associate downloaded archives with installed MO2 mods;
- open source pages and check for available updates.

## Browser Extension Permissions

The browser extension requests permissions for cookies, downloads, tabs, active tab access, native messaging, and LoversLab URLs.

These permissions are used to:

- export LoversLab cookies to the local native bridge;
- detect LoversLab download actions;
- detect completed archive downloads;
- optionally capture archive downloads from a page you manually arm.

External page capture is user-triggered and temporary. It is intended for pages such as Patreon, SubscribeStar, Mega, Google Drive, or similar pages where mod archives may be hosted outside LoversLab.

## Storage Location

Runtime data is stored locally, commonly under:

```text
%LOCALAPPDATA%\LLIntegration
```

The MO2 plugin also stores source metadata inside MO2 mod `meta.ini` files under a `[LoversLab]` section.

## Data Sharing

LL Integration does not transmit your cookies, download history, or local paths to the developer.

Network requests are made only as part of the tool's functionality, such as:

- browser requests made by the websites you visit;
- local update checks to LoversLab when you manually run them from MO2;
- archive downloads you start yourself in the browser.

## Deleting Data

You can delete LL Integration data by:

- uninstalling or clearing data through the LL Integration installer when available;
- deleting `%LOCALAPPDATA%\LLIntegration`;
- removing the browser extension and its local extension data;
- using the MO2 plugin tools to purge LL metadata from selected mods.

## Third Parties

LL Integration is not affiliated with LoversLab, Mozilla, Google, Opera, Microsoft, Nexus Mods, or Mod Organizer 2.

Websites you visit and download from have their own privacy policies and data practices.

## Contact

For privacy questions or bug reports, use the issue tracker on the LL Integration GitHub repository.
