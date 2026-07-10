# officesuite-frontend

Shared static frontend (HTML/CSS/JS, no build step) originally developed for
[OfficeSuite](https://github.com/Decumano/OfficeSuite), a Tauri desktop app.
It's split out here so it can be reused as-is by other projects.

## Structure

- `index.html`, `folio-office-suite.html` — app shell / markup
- `main.js` — application logic
- `styles.css` — styling
- `platform.js` — the native-backend adapter (see below)
- `vendor/` — third-party libraries (see [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md))
- `assets/` — static assets

## Platform adapter

`platform.js` isolates every call to a native file-system backend behind a
single `Platform` object (`pickWorkFolder`, `readWorkFile`, `writeWorkFile`,
`listWorkFolder`, `createWorkFolder`, `deleteWorkEntry`, `moveWorkEntry`,
`defaultFile`, `saveFile`). `main.js` never talks to a native backend
directly.

- **On Tauri**, `platform.js` detects `window.__TAURI__` and delegates to
  `invoke(...)` calls against the host app's Rust commands (see
  `src-tauri/src/lib.rs` in the OfficeSuite repo for the reference
  implementation of those commands).
- **Without a native backend**, work-folder operations are unavailable and
  the app runs off in-memory/`localStorage` state instead (already the
  existing behavior before a work folder is picked). `saveFile` falls back
  to a standard browser download instead of a native save dialog.

To embed this frontend in a new host, either provide the same native
backend (matching the command names/params in `platform.js`) or implement
an equivalent by editing `platform.js`'s `tauri` branch.

## Usage in a consuming repo

Add as a submodule:

```
git submodule add https://github.com/Decumano/officesuite-frontend.git src
git submodule update --init --recursive
```
